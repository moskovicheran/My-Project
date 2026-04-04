from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required, current_user
from functools import wraps
from app.models import (db, AdminNote, MoneyTransfer, SAHierarchy, SARakeConfig,
                        RakeConfig, SharedExpense, ExpenseCharge, User, LoginLog)

admin_bp = Blueprint('admin', __name__, url_prefix='/admin')


def admin_required(f):
    @wraps(f)
    @login_required
    def decorated(*args, **kwargs):
        if not hasattr(current_user, 'role') or current_user.role != 'admin':
            flash('אין לך הרשאה לדף זה.', 'danger')
            return redirect(url_for('main.dashboard'))
        return f(*args, **kwargs)
    return decorated


@admin_bp.route('/')
@admin_required
def overview():
    from app.union_data import get_union_overview, get_cumulative_totals
    from app.models import User, DailyPlayerStats
    from sqlalchemy import func as sqlfunc
    meta, _, _ = get_union_overview()
    ct = get_cumulative_totals()
    meta['period'] = ct['period']

    # Agent stats - use shared function (same logic as agent dashboard)
    from app.union_data import get_agent_totals
    agent_users = User.query.filter_by(role='agent').filter(User.player_id.isnot(None)).all()
    agents_data = []
    for u in agent_users:
        totals = get_agent_totals(u.player_id)
        agents_data.append({
            'username': u.username, 'player_id': u.player_id,
            'players': totals['player_count'], 'rake': totals['total_rake'],
            'pnl': totals['total_pnl'], 'hands': totals['total_hands'],
        })
    agents_data.sort(key=lambda a: a['rake'], reverse=True)

    return render_template('admin/overview.html',
                           meta=meta, clubs=ct['clubs'],
                           total={'active_players': ct['total_players'],
                                  'total_hands': ct['total_hands'],
                                  'total_fee': ct['total_rake'], 'pnl': ct['total_pnl']},
                           tables_count=ct['uploads_count'],
                           total_rake=ct['total_rake'], total_pnl=ct['total_pnl'],
                           total_hands=ct['total_hands'],
                           agents=agents_data)


@admin_bp.route('/agent-view/<sa_id>')
@admin_required
def agent_view(sa_id):
    """Admin view of a specific agent's dashboard."""
    from app.union_data import get_super_agent_tables, get_members_hierarchy, get_cumulative_stats
    from app.models import DailyPlayerStats
    from sqlalchemy import func as sqlfunc

    sa_tables = get_super_agent_tables()
    my_sas = [sa for sa in sa_tables if sa['sa_id'] == sa_id]
    child_sa_ids = [h.child_sa_id for h in SAHierarchy.query.filter_by(parent_sa_id=sa_id).all()]
    child_sas = [sa for sa in sa_tables if sa['sa_id'] in child_sa_ids]
    all_sa_ids = [sa_id] + child_sa_ids

    # Get players from cumulative DB
    my_players_db = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname),
        sqlfunc.max(DailyPlayerStats.club), sqlfunc.max(DailyPlayerStats.agent_id),
        sqlfunc.max(DailyPlayerStats.role),
        sqlfunc.sum(DailyPlayerStats.pnl), sqlfunc.sum(DailyPlayerStats.rake),
        sqlfunc.sum(DailyPlayerStats.hands),
    ).filter(DailyPlayerStats.sa_id.in_(all_sa_ids)
    ).group_by(DailyPlayerStats.player_id).all()

    # Transfer adjustments
    from app.union_data import get_transfer_adjustments
    xfer_adj = get_transfer_adjustments([p[0] for p in my_players_db])

    agents_map = {}
    direct_players = []
    for pid, nick, club, ag_id, role, pnl, rake, hands in my_players_db:
        pnl = round(float(pnl or 0) + xfer_adj.get(pid, 0), 2)
        rake = round(float(rake or 0), 2)
        hands = int(hands or 0)
        member = {'player_id': pid, 'nickname': nick, 'pnl': pnl, 'rake': rake, 'hands': hands}
        if ag_id and ag_id != '-' and ag_id != sa_id:
            if ag_id not in agents_map:
                agents_map[ag_id] = {'id': ag_id, 'nick': ag_id, 'members': [],
                                     'total_pnl': 0, 'total_rake': 0, 'total_hands': 0}
            agents_map[ag_id]['members'].append(member)
            agents_map[ag_id]['total_pnl'] += pnl
            agents_map[ag_id]['total_rake'] += rake
            agents_map[ag_id]['total_hands'] += hands
        else:
            direct_players.append(member)

    # Fetch missing players for agents found in the initial query
    if agents_map:
        found_agent_ids = list(agents_map.keys())
        all_found_pids = set(p[0] for p in my_players_db)
        missing_players = DailyPlayerStats.query.with_entities(
            DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname),
            sqlfunc.max(DailyPlayerStats.club), sqlfunc.max(DailyPlayerStats.agent_id),
            sqlfunc.max(DailyPlayerStats.role),
            sqlfunc.sum(DailyPlayerStats.pnl), sqlfunc.sum(DailyPlayerStats.rake),
            sqlfunc.sum(DailyPlayerStats.hands),
        ).filter(
            DailyPlayerStats.agent_id.in_(found_agent_ids),
            DailyPlayerStats.player_id.notin_(list(all_found_pids))
        ).group_by(DailyPlayerStats.player_id).all()
        for pid, nick, club, ag_id, role, pnl, rake, hands in missing_players:
            if (role or '').lower() in ('name entry',):
                continue
            pnl = round(float(pnl or 0) + xfer_adj.get(pid, 0), 2)
            rake = round(float(rake or 0), 2)
            hands = int(hands or 0)
            member = {'player_id': pid, 'nickname': nick, 'pnl': pnl, 'rake': rake, 'hands': hands}
            if ag_id in agents_map:
                agents_map[ag_id]['members'].append(member)
                agents_map[ag_id]['total_pnl'] += pnl
                agents_map[ag_id]['total_rake'] += rake
                agents_map[ag_id]['total_hands'] += hands

    # Agent nicknames from Excel
    for sa in my_sas + child_sas:
        for ag_id, ag in sa.get('agents', {}).items():
            if ag_id in agents_map:
                agents_map[ag_id]['nick'] = ag['nick']

    # Override child_sas with cumulative DB data
    all_child_pids = set()
    for cs in child_sas:
        for m in cs.get('direct', []):
            all_child_pids.add(m['player_id'])
        for ag in cs.get('agents', {}).values():
            for m in ag.get('members', []):
                all_child_pids.add(m['player_id'])
    if all_child_pids:
        cumul = get_cumulative_stats(list(all_child_pids))
        for cs in child_sas:
            cs_rake = cs_pnl = cs_hands = 0
            for m in cs.get('direct', []):
                c = cumul.get(m['player_id'])
                if c:
                    m['pnl'] = c['pnl']; m['rake'] = c['rake']; m['hands'] = c.get('hands', 0)
                cs_rake += m.get('rake', 0); cs_pnl += m.get('pnl', 0); cs_hands += m.get('hands', 0)
            for ag in cs.get('agents', {}).values():
                ag_r = ag_p = ag_h = 0
                for m in ag.get('members', []):
                    c = cumul.get(m['player_id'])
                    if c:
                        m['pnl'] = c['pnl']; m['rake'] = c['rake']; m['hands'] = c.get('hands', 0)
                    ag_r += m.get('rake', 0); ag_p += m.get('pnl', 0); ag_h += m.get('hands', 0)
                ag['total_rake'] = round(ag_r, 2); ag['total_pnl'] = round(ag_p, 2); ag['total_hands'] = ag_h
                cs_rake += ag_r; cs_pnl += ag_p; cs_hands += ag_h
            cs['total_rake'] = round(cs_rake, 2); cs['total_pnl'] = round(cs_pnl, 2); cs['total_hands'] = cs_hands

    total_rake = sum(m['rake'] for m in direct_players) + sum(a['total_rake'] for a in agents_map.values())
    total_pnl = sum(m['pnl'] for m in direct_players) + sum(a['total_pnl'] for a in agents_map.values())
    total_hands = sum(m['hands'] for m in direct_players) + sum(a['total_hands'] for a in agents_map.values())
    player_count = len(my_players_db)

    agents_sorted = dict(sorted(agents_map.items(), key=lambda x: x[1].get('total_rake', 0), reverse=True))

    sa_nick = my_sas[0]['sa_nick'] if my_sas else sa_id
    my_sa = {
        'sa_id': sa_id, 'sa_nick': sa_nick,
        'club': my_sas[0]['club'] if my_sas else '',
        'agents': agents_sorted, 'direct': direct_players,
        'total_pnl': total_pnl, 'total_rake': total_rake, 'total_hands': total_hands,
    }

    # Managed clubs
    rake_cfgs = SARakeConfig.query.filter_by(sa_id=sa_id).filter(SARakeConfig.managed_club_id.isnot(None)).all()
    managed_clubs = []
    if rake_cfgs:
        clubs_data, _ = get_members_hierarchy()
        club_id_to_name = {c['club_id']: c['name'] for c in clubs_data}
        all_nicks = dict(DailyPlayerStats.query.with_entities(
            DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname)
        ).group_by(DailyPlayerStats.player_id).all())
        for cfg in rake_cfgs:
            cname = club_id_to_name.get(cfg.managed_club_id, '')
            if not cname:
                continue
            cp = DailyPlayerStats.query.with_entities(
                DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname),
                sqlfunc.max(DailyPlayerStats.sa_id), sqlfunc.max(DailyPlayerStats.agent_id),
                sqlfunc.sum(DailyPlayerStats.pnl), sqlfunc.sum(DailyPlayerStats.rake),
            ).filter(DailyPlayerStats.club == cname
            ).group_by(DailyPlayerStats.player_id).all()
            club_sas = {}
            no_sa = []
            cr = cp_pnl = 0
            for pid, nick, sid, aid, pv, rv in cp:
                p = round(float(pv or 0), 2); r = round(float(rv or 0), 2)
                cr += r; cp_pnl += p
                mem = {'player_id': pid, 'nickname': nick, 'pnl_total': p, 'rake_total': r}
                if sid and sid != '-':
                    if sid not in club_sas:
                        club_sas[sid] = {'nick': all_nicks.get(sid, sid), 'id': sid, 'agents': {}, 'direct_members': []}
                    sa = club_sas[sid]
                    if aid and aid != '-' and aid != sid:
                        if aid not in sa['agents']:
                            sa['agents'][aid] = {'nick': all_nicks.get(aid, aid), 'members': []}
                        sa['agents'][aid]['members'].append(mem)
                    else:
                        sa['direct_members'].append(mem)
                else:
                    no_sa.append(mem)
            managed_clubs.append({'name': cname, 'club_id': cfg.managed_club_id,
                                  'total_rake': round(cr, 2), 'total_pnl': round(cp_pnl, 2),
                                  'super_agents': club_sas, 'no_sa_members': no_sa})

    return render_template('admin/agent_view.html',
                           sa_nick=sa_nick, sa_id=sa_id,
                           my_sa=my_sa, child_sas=child_sas,
                           managed_clubs=managed_clubs,
                           total_rake=round(total_rake, 2),
                           total_pnl=round(total_pnl, 2),
                           total_hands=int(total_hands),
                           player_count=player_count)


@admin_bp.route('/transfers', methods=['GET', 'POST'])
@admin_required
def transfers():
    from app.union_data import get_all_members, get_player_balance, get_all_balances

    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add':
            from_key = request.form.get('from_key', '').strip()
            to_key = request.form.get('to_key', '').strip()
            description = request.form.get('description', '').strip()
            try:
                amount = float(request.form.get('amount', 0))
            except ValueError:
                flash('סכום לא תקין.', 'danger')
                return redirect(url_for('admin.transfers'))

            if not from_key or not to_key or '|' not in from_key or '|' not in to_key:
                flash('יש לבחור שולח ומקבל.', 'danger')
            elif from_key == to_key:
                flash('לא ניתן להעביר לאותו שחקן.', 'warning')
            elif amount <= 0:
                flash('הסכום חייב להיות חיובי.', 'danger')
            else:
                from_pid, from_name = from_key.split('|', 1)
                to_pid, to_name = to_key.split('|', 1)
                from_balance = get_player_balance(from_pid)
                to_balance = get_player_balance(to_pid)
                max_transfer = min(abs(from_balance), to_balance)
                if from_balance >= 0:
                    flash(f'{from_name} לא במינוס - אין חוב להעביר.', 'danger')
                elif to_balance <= 0:
                    flash(f'{to_name} לא בפלוס - אין זכות לקבל.', 'danger')
                elif amount > max_transfer:
                    flash(f'חריגה! מקסימום להעברה: {max_transfer:.2f} (חוב: {abs(from_balance):.2f}, זכות: {to_balance:.2f}).', 'danger')
                else:
                    t = MoneyTransfer(user_id=current_user.id,
                                      from_player_id=from_pid, from_name=from_name,
                                      to_player_id=to_pid, to_name=to_name,
                                      amount=amount, description=description)
                    db.session.add(t)
                    db.session.commit()
                    flash(f'העברה של {amount} מ-{from_name} ל-{to_name} בוצעה.', 'success')
        elif action == 'delete':
            tid = request.form.get('transfer_id')
            t = MoneyTransfer.query.get(tid)
            if t:
                db.session.delete(t)
                db.session.commit()
                flash('העברה נמחקה.', 'success')
        return redirect(url_for('admin.transfers'))

    members = get_all_members()
    balances = get_all_balances()
    all_transfers = MoneyTransfer.query.order_by(MoneyTransfer.created_at.desc()).all()
    return render_template('admin/transfers.html',
                           transfers=all_transfers, members=members, balances=balances)


@admin_bp.route('/notes', methods=['GET', 'POST'])
@admin_required
def notes():
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add':
            content = request.form.get('content', '').strip()
            if content:
                note = AdminNote(user_id=current_user.id, content=content)
                db.session.add(note)
                db.session.commit()
                flash('הערה נוספה.', 'success')
        elif action == 'delete':
            nid = request.form.get('note_id')
            note = AdminNote.query.get(nid)
            if note:
                db.session.delete(note)
                db.session.commit()
                flash('הערה נמחקה.', 'success')
        return redirect(url_for('admin.notes'))

    all_notes = AdminNote.query.order_by(AdminNote.created_at.desc()).all()
    return render_template('admin/notes.html', notes=all_notes)


@admin_bp.route('/upload')
@admin_required
def upload():
    return redirect(url_for('upload.index'))


@admin_bp.route('/users')
@admin_required
def users():
    return redirect(url_for('auth.users'))


@admin_bp.route('/rake')
@admin_required
def rake():
    from app.union_data import get_cumulative_totals
    ct = get_cumulative_totals()
    club_rake = {}
    for c in ct['clubs']:
        club_rake[c['club_name']] = {
            'rake': c['total_fee'], 'pnl': c['pnl'],
            'hands': c['total_hands'], 'sessions': c['active_players'],
        }
    return render_template('admin/rake.html', club_rake=club_rake,
                           total_rake=ct['total_rake'], total_pnl=ct['total_pnl'])


@admin_bp.route('/clubs')
@admin_required
def clubs():
    from app.union_data import get_members_hierarchy, get_cumulative_totals
    clubs, grand = get_members_hierarchy()
    ct = get_cumulative_totals()
    grand = {'rake': ct['total_rake'], 'pnl': ct['total_pnl']}
    # Update club totals from cumulative
    club_totals = {c['club_name']: c for c in ct['clubs']}
    for club in clubs:
        ct_club = club_totals.get(club['name'])
        if ct_club:
            club['total_rake'] = ct_club['total_fee']
            club['total_pnl'] = ct_club['pnl']
            club['total_hands'] = ct_club['total_hands']
            club['active_players'] = ct_club['active_players']
    return render_template('admin/clubs.html', clubs=clubs, grand=grand,
                           total_hands=ct.get('total_hands', 0),
                           total_players=ct.get('total_players', 0),
                           uploads_count=ct.get('uploads_count', 0))


@admin_bp.route('/lost-players')
@admin_required
def lost_players():
    return render_template('admin/lost_players.html')


@admin_bp.route('/agents', methods=['GET', 'POST'])
@admin_required
def agents():
    from app.union_data import get_all_super_agents, get_all_clubs, get_super_agent_tables

    if request.method == 'POST':
        action = request.form.get('action')

        # SA → SA hierarchy
        if action == 'add_hierarchy':
            parent_id = request.form.get('parent_sa_id')
            child_id = request.form.get('child_sa_id')
            force = request.form.get('force') == '1'
            if parent_id and child_id and parent_id != child_id:
                existing = SAHierarchy.query.filter_by(child_sa_id=child_id).first()
                if existing and existing.parent_sa_id != parent_id:
                    if force:
                        existing.parent_sa_id = parent_id
                        db.session.commit()
                        flash(f'SA הועבר בהצלחה.', 'success')
                    else:
                        # Return info about current assignment for JS confirm
                        sa_map = {sa['id']: sa['nick'] for sa in get_all_super_agents()}
                        old_parent = sa_map.get(existing.parent_sa_id, existing.parent_sa_id)
                        child_name = sa_map.get(child_id, child_id)
                        flash(f'TRANSFER_CONFIRM:SA:{child_id}:{parent_id}:{child_name} כבר משויך ל-{old_parent}. להעביר?', 'warning')
                elif existing and existing.parent_sa_id == parent_id:
                    flash('שיוך זה כבר קיים.', 'warning')
                else:
                    db.session.add(SAHierarchy(parent_sa_id=parent_id, child_sa_id=child_id))
                    db.session.commit()
                    flash('שיוך SA → SA נוסף.', 'success')
            else:
                flash('יש לבחור שני Super Agents שונים.', 'warning')

        elif action == 'delete_hierarchy':
            link = SAHierarchy.query.get(request.form.get('link_id'))
            if link:
                db.session.delete(link)
                db.session.commit()
                flash('שיוך SA → SA נמחק.', 'success')

        # SA → Club mapping
        elif action == 'set_club':
            sa_id = request.form.get('sa_id')
            club_id = request.form.get('club_id', '').strip()
            force = request.form.get('force') == '1'
            if sa_id and club_id:
                # Check if this club is already assigned to another SA
                existing_other = SARakeConfig.query.filter_by(managed_club_id=club_id).filter(SARakeConfig.sa_id != sa_id).first()
                existing_same = SARakeConfig.query.filter_by(sa_id=sa_id, managed_club_id=club_id).first()
                if existing_same:
                    flash('שיוך זה כבר קיים.', 'warning')
                elif existing_other and not force:
                    sa_map = {sa['id']: sa['nick'] for sa in get_all_super_agents()}
                    old_sa = sa_map.get(existing_other.sa_id, existing_other.sa_id)
                    flash(f'TRANSFER_CONFIRM:CLUB:{club_id}:{sa_id}:מועדון זה כבר משויך ל-{old_sa}. להעביר?', 'warning')
                elif existing_other and force:
                    existing_other.sa_id = sa_id
                    db.session.commit()
                    flash('מועדון הועבר בהצלחה.', 'success')
                else:
                    db.session.add(SARakeConfig(sa_id=sa_id, rake_percent=0, managed_club_id=club_id))
                    db.session.commit()
                    flash('שיוך SA → מועדון נוסף.', 'success')
            else:
                flash('יש לבחור SA ומועדון.', 'warning')

        # Rake %
        elif action == 'set_rake':
            sa_id = request.form.get('sa_id')
            try:
                pct = float(request.form.get('rake_percent', 0))
            except ValueError:
                pct = 0
            if sa_id:
                config = SARakeConfig.query.filter_by(sa_id=sa_id).first()
                if not config:
                    config = SARakeConfig(sa_id=sa_id)
                    db.session.add(config)
                config.rake_percent = pct
                db.session.commit()
                flash(f'אחוז רייק עודכן ל-{pct}%.', 'success')

        elif action == 'delete_config':
            config = SARakeConfig.query.get(request.form.get('config_id'))
            if config:
                db.session.delete(config)
                db.session.commit()
                flash('הגדרת SA נמחקה.', 'success')

        # Rake config for clubs/agents/players
        elif action == 'add_rake_config':
            entity_type = request.form.get('entity_type', '')
            entity_key = request.form.get('entity_key', '').strip()
            try:
                pct = float(request.form.get('rake_percent', 0))
            except ValueError:
                pct = 0
            if entity_type and entity_key and '|' in entity_key:
                eid, ename = entity_key.split('|', 1)
                existing = RakeConfig.query.filter_by(entity_type=entity_type, entity_id=eid).first()
                if existing:
                    existing.rake_percent = pct
                    flash(f'אחוז רייק ל-{ename} עודכן ל-{pct}%.', 'success')
                else:
                    db.session.add(RakeConfig(entity_type=entity_type, entity_id=eid,
                                             entity_name=ename, rake_percent=pct))
                    flash(f'רייק {pct}% הוגדר ל-{ename}.', 'success')
                db.session.commit()
            else:
                flash('יש לבחור ישות ולהזין אחוז.', 'warning')

        elif action == 'delete_rake_config':
            rc = RakeConfig.query.get(request.form.get('rc_id'))
            if rc:
                db.session.delete(rc)
                db.session.commit()
                flash('הגדרת רייק נמחקה.', 'success')

        return redirect(url_for('admin.agents'))

    # GET - build all data
    from app.models import DailyPlayerStats
    from sqlalchemy import func as sqlfunc
    all_sa = get_all_super_agents()
    all_clubs = get_all_clubs()
    # Also add SAs from DB that aren't in Excel
    sa_ids_excel = {sa['id'] for sa in all_sa}
    db_sas = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.sa_id, sqlfunc.max(DailyPlayerStats.nickname), sqlfunc.max(DailyPlayerStats.club)
    ).filter(DailyPlayerStats.sa_id != '-', DailyPlayerStats.sa_id != '', DailyPlayerStats.sa_id.isnot(None)
    ).group_by(DailyPlayerStats.sa_id).all()
    for sid, nick, club in db_sas:
        if sid and sid not in sa_ids_excel:
            all_sa.append({'id': sid, 'nick': nick or sid, 'club': club or ''})
    all_sa.sort(key=lambda x: x['nick'].lower())
    sa_name_map = {sa['id']: sa for sa in all_sa}

    # SA hierarchy links
    hierarchy_links = []
    for link in SAHierarchy.query.all():
        p = sa_name_map.get(link.parent_sa_id, {})
        c = sa_name_map.get(link.child_sa_id, {})
        hierarchy_links.append({
            'id': link.id,
            'parent_id': link.parent_sa_id, 'parent_nick': p.get('nick', link.parent_sa_id),
            'child_id': link.child_sa_id, 'child_nick': c.get('nick', link.child_sa_id),
        })

    # Rake configs
    club_map = {c['club_id']: c['name'] for c in all_clubs}
    configs = []
    for cfg in SARakeConfig.query.all():
        sa = sa_name_map.get(cfg.sa_id, {})
        configs.append({
            'id': cfg.id, 'sa_id': cfg.sa_id,
            'sa_nick': sa.get('nick', cfg.sa_id),
            'sa_club': sa.get('club', ''),
            'rake_percent': cfg.rake_percent,
            'managed_club_id': cfg.managed_club_id or '',
            'managed_club_name': club_map.get(cfg.managed_club_id, ''),
        })

    # SA stats from cumulative DB
    from app.models import DailyPlayerStats
    from sqlalchemy import func as sqlfunc
    sa_stats_db = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.sa_id,
        sqlfunc.sum(DailyPlayerStats.pnl),
        sqlfunc.sum(DailyPlayerStats.rake),
        sqlfunc.sum(DailyPlayerStats.hands),
    ).filter(DailyPlayerStats.sa_id != '', DailyPlayerStats.sa_id != '-', DailyPlayerStats.role != 'Name Entry'
    ).group_by(DailyPlayerStats.sa_id).all()
    sa_stats = {}
    for sid, pnl, rake, hands in sa_stats_db:
        sa_stats[sid] = {
            'total_rake': round(float(rake or 0), 2),
            'total_pnl': round(float(pnl or 0), 2),
            'total_hands': int(hands or 0),
        }

    # Rake configs for clubs/agents/players
    from app.union_data import get_all_members
    all_members = get_all_members()
    rake_configs = RakeConfig.query.order_by(RakeConfig.entity_type).all()

    # All agents (non-SA) from DB — get real agent name from their player entry
    agent_ids_db = [r[0] for r in DailyPlayerStats.query.with_entities(
        DailyPlayerStats.agent_id
    ).filter(DailyPlayerStats.agent_id != '-', DailyPlayerStats.agent_id != '', DailyPlayerStats.agent_id.isnot(None)
    ).group_by(DailyPlayerStats.agent_id).all() if r[0]]
    # Get agent nicknames from their own player records
    agent_nicks = dict(DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname)
    ).filter(DailyPlayerStats.player_id.in_(agent_ids_db)
    ).group_by(DailyPlayerStats.player_id).all()) if agent_ids_db else {}
    # Get club info
    agent_clubs = dict(DailyPlayerStats.query.with_entities(
        DailyPlayerStats.agent_id, sqlfunc.max(DailyPlayerStats.club)
    ).filter(DailyPlayerStats.agent_id.in_(agent_ids_db)
    ).group_by(DailyPlayerStats.agent_id).all()) if agent_ids_db else {}
    all_sub_agents = [{'id': aid, 'nick': agent_nicks.get(aid, aid), 'club': agent_clubs.get(aid, '')} for aid in agent_ids_db]

    return render_template('admin/agents.html',
                           all_sa=all_sa, all_clubs=all_clubs,
                           all_members=all_members,
                           all_sub_agents=all_sub_agents,
                           hierarchy_links=hierarchy_links,
                           configs=configs, sa_stats=sa_stats,
                           rake_configs=rake_configs)


@admin_bp.route('/expenses', methods=['GET', 'POST'])
@admin_required
def expenses():
    if request.method == 'POST':
        action = request.form.get('action')

        try:
            if action == 'add':
                description = request.form.get('description', '').strip()
                try:
                    amount = float(request.form.get('amount', 0))
                except ValueError:
                    amount = 0
                if description and amount > 0:
                    exp = SharedExpense(user_id=current_user.id, description=description, amount=amount)
                    db.session.add(exp)
                    db.session.commit()
                    flash(f'הוצאה ({amount}) נוספה.', 'success')
                else:
                    flash('יש למלא תיאור וסכום.', 'danger')

            elif action == 'charge':
                exp_id = request.form.get('expense_id')
                exp = SharedExpense.query.get(exp_id)
                if exp and not exp.charged:
                    agents = User.query.filter_by(role='agent').filter(User.player_id.isnot(None)).all()
                    if not agents:
                        flash('אין סוכנים במערכת לחייב.', 'warning')
                    else:
                        share = round(exp.amount / len(agents), 2)
                        for agent in agents:
                            charge = ExpenseCharge(expense_id=exp.id,
                                                  agent_player_id=agent.player_id,
                                                  agent_name=agent.username,
                                                  charge_amount=share)
                            db.session.add(charge)
                        exp.charged = True
                        db.session.commit()
                        flash(f'חויבו {len(agents)} סוכנים, {share} לכל אחד.', 'success')

            elif action == 'delete':
                exp_id = request.form.get('expense_id')
                exp = SharedExpense.query.get(exp_id)
                if exp and not exp.charged:
                    db.session.delete(exp)
                    db.session.commit()
                    flash('הוצאה נמחקה.', 'success')
                elif exp and exp.charged:
                    flash('לא ניתן למחוק הוצאה שכבר חויבה.', 'warning')
        except Exception as e:
            db.session.rollback()
            flash(f'שגיאה: {str(e)[:100]}', 'danger')

        return redirect(url_for('admin.expenses'))

    all_expenses = SharedExpense.query.order_by(SharedExpense.created_at.desc()).all()
    recent_charges = ExpenseCharge.query.order_by(ExpenseCharge.created_at.desc()).limit(50).all()
    agents_count = User.query.filter_by(role='agent').filter(User.player_id.isnot(None)).count()
    return render_template('admin/expenses.html',
                           expenses=all_expenses, charges=recent_charges,
                           agents_count=agents_count)


@admin_bp.route('/top-players')
@admin_required
def top_players():
    from app.union_data import get_cumulative_stats

    # Get all cumulative stats from DB
    all_cumulative = get_cumulative_stats()
    all_players = []
    for pid, cs in all_cumulative.items():
        if cs.get('hands', 0) == 0 and cs.get('pnl', 0) == 0:
            continue  # Skip name-only entries
        all_players.append({
            'player_id': pid, 'member_id': pid,
            'nickname': cs['nickname'],
            'club': cs['club'],
            'pnl': cs['pnl'], 'pnl_total': cs['pnl'],
            'rake': cs['rake'], 'rake_total': cs['rake'],
            'hands': cs['hands'], 'hands_total': cs['hands'],
        })

    top_winners = sorted(all_players, key=lambda x: x['pnl'], reverse=True)[:10]
    top_losers = sorted(all_players, key=lambda x: x['pnl'])[:10]
    top_rake = sorted(all_players, key=lambda x: x['rake'], reverse=True)[:10]
    top_active = sorted(all_players, key=lambda x: x['hands'], reverse=True)[:10]

    biggest_winner = top_winners[0]['pnl'] if top_winners else 0
    biggest_loser = top_losers[0]['pnl'] if top_losers else 0

    return render_template('admin/top_players.html',
                           top_winners=top_winners, top_losers=top_losers,
                           top_rake=top_rake, top_active=top_active,
                           total_players=len(all_players),
                           biggest_winner=biggest_winner,
                           biggest_loser=biggest_loser)


@admin_bp.route('/reports')
@admin_required
def reports():
    from app.union_data import get_all_members
    members = get_all_members()
    return render_template('admin/reports.html', members=members)


@admin_bp.route('/logins')
@admin_required
def logins():
    import re
    from datetime import timedelta

    def parse_ua(ua):
        if not ua:
            return '-', '-'
        # Device
        if 'iPhone' in ua:
            device = 'iPhone'
        elif 'iPad' in ua:
            device = 'iPad'
        elif 'Android' in ua:
            m = re.search(r'Android[^;]*;\s*([^)]+)', ua)
            raw = m.group(1).strip().split(' Build')[0].strip() if m else ''
            device = raw if len(raw) > 2 else 'Android'
        elif 'Macintosh' in ua:
            device = 'Mac'
        elif 'Windows' in ua:
            device = 'Windows PC'
        elif 'Linux' in ua:
            device = 'Linux PC'
        else:
            device = '-'
        # Browser
        if 'Edg/' in ua:
            browser = 'Edge'
        elif 'OPR/' in ua or 'Opera' in ua:
            browser = 'Opera'
        elif 'Chrome/' in ua and 'Safari/' in ua:
            browser = 'Chrome'
        elif 'Safari/' in ua and 'Chrome' not in ua:
            browser = 'Safari'
        elif 'Firefox/' in ua:
            browser = 'Firefox'
        else:
            browser = '-'
        return device, browser

    import urllib.request, json

    def get_ip_city(ip):
        if not ip or ip.startswith('127.') or ip.startswith('192.168.'):
            return '-'
        try:
            resp = urllib.request.urlopen(f'https://ipinfo.io/{ip}/json', timeout=3)
            data = json.loads(resp.read())
            return data.get('city', '-')
        except Exception:
            return '-'

    logs = LoginLog.query.order_by(LoginLog.created_at.desc()).limit(100).all()
    # Cache IP→city to avoid duplicate lookups
    ip_city_cache = {}
    for log in logs:
        log.created_at = log.created_at + timedelta(hours=3)
        log.device, log.browser = parse_ua(log.user_agent)
        ip = log.ip_address
        if ip not in ip_city_cache:
            ip_city_cache[ip] = get_ip_city(ip)
        log.city = ip_city_cache[ip]
    return render_template('admin/logins.html', logs=logs)
