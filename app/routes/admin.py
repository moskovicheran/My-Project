from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required, current_user
from functools import wraps
from app.models import (db, AdminNote, MoneyTransfer, SAHierarchy, SARakeConfig,
                        RakeConfig, SharedExpense, ExpenseCharge, User)

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
    from app.union_data import get_union_overview, get_ring_game_detail
    meta, clubs, total = get_union_overview()
    tables = get_ring_game_detail()
    total_rake = round(sum(p['rake'] for t in tables for p in t['players']), 2)
    total_pnl = round(sum(p['pnl'] for t in tables for p in t['players']), 2)
    total_hands = sum(p['hands'] for t in tables for p in t['players'])
    return render_template('admin/overview.html',
                           meta=meta, clubs=clubs, total=total,
                           tables_count=len(tables),
                           total_rake=total_rake, total_pnl=total_pnl,
                           total_hands=int(total_hands))


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
    from app.union_data import get_ring_game_detail, get_union_overview
    _, clubs, _ = get_union_overview()
    tables = get_ring_game_detail()

    # Rake by club
    club_rake = {}
    for t in tables:
        for p in t['players']:
            club = p['club_name']
            if club not in club_rake:
                club_rake[club] = {'rake': 0, 'pnl': 0, 'hands': 0, 'sessions': 0}
            club_rake[club]['rake'] = round(club_rake[club]['rake'] + p['rake'], 2)
            club_rake[club]['pnl'] = round(club_rake[club]['pnl'] + p['pnl'], 2)
            club_rake[club]['hands'] += int(p['hands'])
        for club in set(p['club_name'] for p in t['players']):
            club_rake[club]['sessions'] += 1

    total_rake = round(sum(c['rake'] for c in club_rake.values()), 2)
    total_pnl = round(sum(c['pnl'] for c in club_rake.values()), 2)
    return render_template('admin/rake.html', club_rake=club_rake,
                           total_rake=total_rake, total_pnl=total_pnl)


@admin_bp.route('/clubs')
@admin_required
def clubs():
    from app.union_data import get_members_hierarchy
    clubs, grand = get_members_hierarchy()
    return render_template('admin/clubs.html', clubs=clubs, grand=grand)


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
            if parent_id and child_id and parent_id != child_id:
                if SAHierarchy.query.filter_by(child_sa_id=child_id).first():
                    flash('Super Agent זה כבר משויך למישהו אחר.', 'warning')
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
            if sa_id:
                config = SARakeConfig.query.filter_by(sa_id=sa_id).first()
                if not config:
                    config = SARakeConfig(sa_id=sa_id, rake_percent=0)
                    db.session.add(config)
                config.managed_club_id = club_id if club_id else None
                db.session.commit()
                flash('שיוך SA → מועדון עודכן.', 'success')

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
    all_sa = get_all_super_agents()
    all_clubs = get_all_clubs()
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

    # SA stats from Excel
    sa_tables = get_super_agent_tables()
    sa_stats = {}
    for sa in sa_tables:
        sa_stats[sa['sa_id']] = {
            'total_rake': sa['total_rake'],
            'total_pnl': sa['total_pnl'],
            'total_hands': sa['total_hands'],
        }

    # Rake configs for clubs/agents/players
    from app.union_data import get_all_members
    all_members = get_all_members()
    rake_configs = RakeConfig.query.order_by(RakeConfig.entity_type).all()

    return render_template('admin/agents.html',
                           all_sa=all_sa, all_clubs=all_clubs,
                           all_members=all_members,
                           hierarchy_links=hierarchy_links,
                           configs=configs, sa_stats=sa_stats,
                           rake_configs=rake_configs)


@admin_bp.route('/expenses', methods=['GET', 'POST'])
@admin_required
def expenses():
    if request.method == 'POST':
        action = request.form.get('action')

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
                flash(f'הוצאה "{description}" ({amount}) נוספה.', 'success')
            else:
                flash('יש למלא תיאור וסכום.', 'danger')

        elif action == 'charge':
            exp_id = request.form.get('expense_id')
            exp = SharedExpense.query.get(exp_id)
            if exp and not exp.charged:
                # Get all agent users
                agents = User.query.filter_by(role='agent').filter(User.player_id.isnot(None)).all()
                if not agents:
                    flash('אין סוכנים (מנהלים) במערכת לחייב.', 'warning')
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
    from app.union_data import get_top_members
    top_winners, top_losers = get_top_members(10)

    # Top by rake
    from app.union_data import _read_sheets, _num
    sheets = _read_sheets()
    all_players = []
    if 'Union Member Statistics' in sheets:
        df = sheets['Union Member Statistics']
        current_club = ''
        for i in range(6, len(df)):
            row = df.iloc[i]
            if '(ID:' in str(row.iloc[0]):
                current_club = str(row.iloc[0]).split(' (ID:')[0]
            nickname = str(row.iloc[9])
            if nickname in ('nan', '-'):
                continue
            all_players.append({
                'player_id': str(row.iloc[8]),
                'nickname': nickname,
                'club': current_club,
                'pnl': _num(row.iloc[37]),
                'rake': _num(row.iloc[64]),
                'hands': _num(row.iloc[151]),
            })

    top_rake = sorted(all_players, key=lambda x: x['rake'], reverse=True)[:10]
    top_active = sorted(all_players, key=lambda x: x['hands'], reverse=True)[:10]

    biggest_winner = top_winners[0]['pnl_total'] if top_winners else 0
    biggest_loser = top_losers[0]['pnl_total'] if top_losers else 0

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
