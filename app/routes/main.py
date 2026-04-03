import io
from datetime import date
from flask import Blueprint, render_template, redirect, url_for, flash, request, jsonify, send_file
from flask_login import login_required, current_user
from sqlalchemy import func
from app.models import db, Transaction

main_bp = Blueprint('main', __name__)

INCOME_CATEGORIES = ['משכורת', 'פרילנס', 'השקעות', 'מתנה', 'אחר']
EXPENSE_CATEGORIES = ['מזון', 'דיור', 'תחבורה', 'בריאות', 'בידור', 'קניות', 'חינוך', 'חשבונות', 'אחר']


@main_bp.route('/')
def index():
    if current_user.is_authenticated:
        return redirect(url_for('main.dashboard'))
    return redirect(url_for('auth.login'))


@main_bp.route('/dashboard')
@login_required
def dashboard():
    if hasattr(current_user, 'role') and current_user.role == 'admin':
        from app.union_data import get_union_overview, get_cumulative_totals
        meta, _, _ = get_union_overview()
        ct = get_cumulative_totals()
        meta['period'] = ct['period']
        return render_template('main/admin_dashboard.html',
                               meta=meta, clubs=ct['clubs'],
                               total={'active_players': ct['total_players'],
                                      'total_hands': ct['total_hands'],
                                      'total_fee': ct['total_rake'], 'pnl': ct['total_pnl']},
                               tables_count=ct['uploads_count'],
                               total_rake=ct['total_rake'], total_pnl=ct['total_pnl'],
                               total_hands=ct['total_hands'],
                               ring_rake=ct.get('ring_rake', 0),
                               mtt_rake=ct.get('mtt_rake', 0))

    if hasattr(current_user, 'role') and current_user.role == 'club' and current_user.player_id:
        from app.models import DailyPlayerStats, DailyUpload
        from app.union_data import get_members_hierarchy
        from sqlalchemy import func as sqlfunc
        from datetime import datetime as dt

        club_id = current_user.player_id
        # Find club name
        clubs_data, _ = get_members_hierarchy()
        club_name = None
        club_obj = None
        for c in clubs_data:
            if c['club_id'] == club_id:
                club_name = c['name']
                club_obj = c
                break

        # Available upload dates
        available_dates = [u[0].strftime('%Y-%m-%d') for u in
                           DailyUpload.query.with_entities(DailyUpload.upload_date)
                           .distinct().order_by(DailyUpload.upload_date.desc()).all()]

        # Date filter — supports multiple dates: ?dates=2026-03-30,2026-03-31
        selected_dates = [d.strip() for d in request.args.get('dates', '').split(',') if d.strip()]
        upload_ids_filter = []
        if selected_dates:
            valid_dates = []
            for ds in selected_dates:
                try:
                    sel = dt.strptime(ds, '%Y-%m-%d').date()
                    upload = DailyUpload.query.filter_by(upload_date=sel).first()
                    if upload:
                        upload_ids_filter.append(upload.id)
                        valid_dates.append(ds)
                except ValueError:
                    pass
            selected_dates = valid_dates

        if club_name:
            # Base query
            base_filters = [DailyPlayerStats.club == club_name,
                            DailyPlayerStats.role != 'Name Entry']
            if upload_ids_filter:
                base_filters.append(DailyPlayerStats.upload_id.in_(upload_ids_filter))

            club_players_db = DailyPlayerStats.query.with_entities(
                DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname),
                sqlfunc.max(DailyPlayerStats.sa_id), sqlfunc.max(DailyPlayerStats.agent_id),
                sqlfunc.sum(DailyPlayerStats.pnl), sqlfunc.sum(DailyPlayerStats.rake),
                sqlfunc.sum(DailyPlayerStats.hands),
            ).filter(*base_filters).group_by(DailyPlayerStats.player_id).all()

            # Nickname map
            all_nicks = dict(DailyPlayerStats.query.with_entities(
                DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname)
            ).group_by(DailyPlayerStats.player_id).all())

            # Build SA structure
            club_sas = {}
            no_sa = []
            total_rake = 0
            total_pnl = 0
            total_hands = 0
            for pid, nick, sa_id_val, ag_id_val, pnl_val, rake_val, hands_val in club_players_db:
                p = round(float(pnl_val or 0), 2)
                r = round(float(rake_val or 0), 2)
                h = int(hands_val or 0)
                total_rake += r
                total_pnl += p
                total_hands += h
                member = {'player_id': pid, 'nickname': nick, 'pnl_total': p, 'rake_total': r, 'hands': h}

                if sa_id_val and sa_id_val != '-':
                    if sa_id_val not in club_sas:
                        sa_nick = all_nicks.get(sa_id_val, sa_id_val)
                        club_sas[sa_id_val] = {'nick': sa_nick, 'id': sa_id_val,
                                                'agents': {}, 'direct_members': []}
                    sa = club_sas[sa_id_val]
                    if ag_id_val and ag_id_val != '-' and ag_id_val != sa_id_val:
                        if ag_id_val not in sa['agents']:
                            ag_nick = all_nicks.get(ag_id_val, ag_id_val)
                            sa['agents'][ag_id_val] = {'nick': ag_nick, 'members': []}
                        sa['agents'][ag_id_val]['members'].append(member)
                    else:
                        sa['direct_members'].append(member)
                else:
                    no_sa.append(member)

            managed_club = {
                'name': club_name, 'club_id': club_id,
                'total_rake': round(total_rake, 2), 'total_pnl': round(total_pnl, 2),
                'super_agents': club_sas, 'no_sa_members': no_sa,
            }
            player_count = len(club_players_db)

            # Net rake calculation (club's percentage)
            from app.models import RakeConfig
            club_rc = RakeConfig.query.filter_by(entity_type='club', entity_id=club_id).first()
            rake_pct = club_rc.rake_percent if club_rc else 100
            net_rake = round(total_rake * rake_pct / 100, 2)

            return render_template('main/club_dashboard.html',
                                   managed_club=managed_club,
                                   total_rake=round(total_rake, 2),
                                   net_rake=net_rake,
                                   rake_pct=rake_pct,
                                   total_pnl=round(total_pnl, 2),
                                   total_hands=total_hands,
                                   player_count=player_count,
                                   available_dates=available_dates,
                                   selected_dates=selected_dates)

        # Club not found in data
        return render_template('main/club_dashboard.html',
                               managed_club=None, total_rake=0, net_rake=0,
                               rake_pct=100, total_pnl=0,
                               total_hands=0, player_count=0,
                               available_dates=available_dates,
                               selected_dates=[])

    if hasattr(current_user, 'role') and current_user.role == 'agent' and current_user.player_id:
        from app.union_data import get_super_agent_tables, get_members_hierarchy
        from app.models import SAHierarchy, SARakeConfig, RakeConfig, ExpenseCharge, DailyPlayerStats, DailyUpload
        from sqlalchemy import func as sqlfunc
        from datetime import datetime as dt
        sa_id = current_user.player_id

        # Available upload dates
        available_dates = [u[0].strftime('%Y-%m-%d') for u in
                           DailyUpload.query.with_entities(DailyUpload.upload_date)
                           .distinct().order_by(DailyUpload.upload_date.desc()).all()]

        # Date filter
        selected_dates = [d.strip() for d in request.args.get('dates', '').split(',') if d.strip()]
        upload_ids_filter = []
        if selected_dates:
            valid_dates = []
            for ds in selected_dates:
                try:
                    sel = dt.strptime(ds, '%Y-%m-%d').date()
                    upload = DailyUpload.query.filter_by(upload_date=sel).first()
                    if upload:
                        upload_ids_filter.append(upload.id)
                        valid_dates.append(ds)
                except ValueError:
                    pass
            selected_dates = valid_dates

        # Resolve the actual SA/Agent ID for this user
        from sqlalchemy import or_
        known_ids = {sa_id}

        # 1) Check if player_id is directly used as sa_id or agent_id
        is_sa = DailyPlayerStats.query.filter(DailyPlayerStats.sa_id == sa_id).first() is not None
        is_agent = DailyPlayerStats.query.filter(DailyPlayerStats.agent_id == sa_id).first() is not None

        if not is_sa and not is_agent:
            # 2) Player ID doesn't match directly - look up their role to find real ID
            own_row = DailyPlayerStats.query.filter(DailyPlayerStats.player_id == sa_id).first()
            if own_row:
                role_lower = (own_row.role or '').lower()
                if 'super' in role_lower or role_lower in ('sa',):
                    # They're an SA - use sa_id from their row
                    if own_row.sa_id and own_row.sa_id != '-':
                        known_ids.add(own_row.sa_id)
                elif 'agent' in role_lower:
                    # They're a sub-agent - use agent_id from their row
                    if own_row.agent_id and own_row.agent_id != '-':
                        known_ids.add(own_row.agent_id)

        known_ids.discard('')
        known_ids.discard('-')

        # Get SA structure from Excel (for hierarchy display)
        sa_tables = get_super_agent_tables()
        my_sas = []
        for kid in known_ids:
            my_sas.extend([sa for sa in sa_tables if sa['sa_id'] == kid])
        child_sa_ids = []
        for kid in known_ids:
            child_sa_ids.extend([h.child_sa_id for h in SAHierarchy.query.filter_by(parent_sa_id=kid).all()])
        child_sas = [sa for sa in sa_tables if sa['sa_id'] in child_sa_ids]
        all_sa_ids = list(known_ids) + child_sa_ids

        # Get ALL players that ever belonged to this SA/Agent from CUMULATIVE DB
        base_agent_filters = [or_(
            DailyPlayerStats.sa_id.in_(all_sa_ids),
            DailyPlayerStats.agent_id.in_(all_sa_ids)
        )]
        if upload_ids_filter:
            base_agent_filters.append(DailyPlayerStats.upload_id.in_(upload_ids_filter))

        my_players_db = DailyPlayerStats.query.with_entities(
            DailyPlayerStats.player_id,
            sqlfunc.max(DailyPlayerStats.nickname),
            sqlfunc.max(DailyPlayerStats.club),
            sqlfunc.max(DailyPlayerStats.agent_id),
            sqlfunc.max(DailyPlayerStats.role),
            sqlfunc.sum(DailyPlayerStats.pnl),
            sqlfunc.sum(DailyPlayerStats.rake),
            sqlfunc.sum(DailyPlayerStats.hands),
        ).filter(*base_agent_filters
        ).group_by(DailyPlayerStats.player_id).all()

        # Build agent structure from DB data
        all_my_player_ids = set()
        agents_map = {}  # agent_id -> {nick, members, totals}
        direct_players = []
        for pid, nick, club, ag_id, role, pnl, rake, hands in my_players_db:
            pnl = round(float(pnl or 0), 2)
            rake = round(float(rake or 0), 2)
            hands = int(hands or 0)
            all_my_player_ids.add(pid)
            member = {'player_id': pid, 'nickname': nick, 'role': role or 'Player',
                      'pnl': pnl, 'rake': rake, 'hands': hands}
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

        # Override child_sas Excel data with cumulative DB data
        from app.union_data import get_cumulative_stats
        all_child_player_ids = set()
        for cs in child_sas:
            for m in cs.get('direct', []):
                all_child_player_ids.add(m['player_id'])
            for ag in cs.get('agents', {}).values():
                for m in ag.get('members', []):
                    all_child_player_ids.add(m['player_id'])
        if all_child_player_ids:
            cumul = get_cumulative_stats(list(all_child_player_ids))
            for cs in child_sas:
                cs_rake = 0
                cs_pnl = 0
                cs_hands = 0
                for m in cs.get('direct', []):
                    c = cumul.get(m['player_id'])
                    if c:
                        m['pnl'] = c['pnl']
                        m['rake'] = c['rake']
                        m['hands'] = c.get('hands', 0)
                    cs_rake += m.get('rake', 0)
                    cs_pnl += m.get('pnl', 0)
                    cs_hands += m.get('hands', 0)
                for ag in cs.get('agents', {}).values():
                    ag_rake = 0
                    ag_pnl = 0
                    ag_hands = 0
                    for m in ag.get('members', []):
                        c = cumul.get(m['player_id'])
                        if c:
                            m['pnl'] = c['pnl']
                            m['rake'] = c['rake']
                            m['hands'] = c.get('hands', 0)
                        ag_rake += m.get('rake', 0)
                        ag_pnl += m.get('pnl', 0)
                        ag_hands += m.get('hands', 0)
                    ag['total_rake'] = round(ag_rake, 2)
                    ag['total_pnl'] = round(ag_pnl, 2)
                    ag['total_hands'] = ag_hands
                    cs_rake += ag_rake
                    cs_pnl += ag_pnl
                    cs_hands += ag_hands
                cs['total_rake'] = round(cs_rake, 2)
                cs['total_pnl'] = round(cs_pnl, 2)
                cs['total_hands'] = cs_hands

        # Find agent nicknames from Excel
        for sa in my_sas + child_sas:
            for ag_id, ag in sa.get('agents', {}).items():
                if ag_id in agents_map:
                    agents_map[ag_id]['nick'] = ag['nick']

        # Query rake configs for sub-agents
        agent_ids = list(agents_map.keys())
        agent_rake_configs = {rc.entity_id: rc.rake_percent
                              for rc in RakeConfig.query.filter(
                                  RakeConfig.entity_type.in_(['sub_agent', 'agent']),
                                  RakeConfig.entity_id.in_(agent_ids)).all()} if agent_ids else {}
        for ag_id, ag in agents_map.items():
            pct = agent_rake_configs.get(ag_id, 0)
            ag['rake_pct'] = pct
            ag['agent_net_rake'] = round(ag['total_rake'] * pct / 100, 2)
            ag['sa_keeps'] = round(ag['total_rake'] - ag['agent_net_rake'], 2)

        # Query rake configs for players
        all_player_ids_list = list(all_my_player_ids)
        player_rake_configs = {rc.entity_id: rc.rake_percent
                               for rc in RakeConfig.query.filter(
                                   RakeConfig.entity_type == 'player',
                                   RakeConfig.entity_id.in_(all_player_ids_list)).all()} if all_player_ids_list else {}
        players_with_rake = []
        for m in direct_players:
            pct = player_rake_configs.get(m['player_id'], 0)
            if pct:
                refund = round(m['rake'] * pct / 100, 2)
                players_with_rake.append({'nick': m['nickname'], 'rake_pct': pct,
                                          'total_rake': m['rake'], 'refund': refund})
        for ag in agents_map.values():
            for m in ag['members']:
                pct = player_rake_configs.get(m['player_id'], 0)
                if pct:
                    refund = round(m['rake'] * pct / 100, 2)
                    players_with_rake.append({'nick': m['nickname'], 'rake_pct': pct,
                                              'total_rake': m['rake'], 'refund': refund})

        # Combined rake refund list (agents + players)
        rake_refund_list = []
        for ag in agents_map.values():
            if ag.get('rake_pct'):
                rake_refund_list.append({'nick': ag['nick'], 'rake_pct': ag['rake_pct'],
                                         'total_rake': ag['total_rake'], 'refund': ag['agent_net_rake'],
                                         'type': 'agent'})
        for p in players_with_rake:
            rake_refund_list.append({'nick': p['nick'], 'rake_pct': p['rake_pct'],
                                     'total_rake': p['total_rake'], 'refund': p['refund'],
                                     'type': 'player'})
        total_rake_refund = round(sum(r['refund'] for r in rake_refund_list), 2)

        # Build a single SA structure with cumulative data
        total_rake = sum(m['rake'] for m in direct_players) + sum(a['total_rake'] for a in agents_map.values())
        total_pnl = sum(m['pnl'] for m in direct_players) + sum(a['total_pnl'] for a in agents_map.values())
        total_hands = sum(m['hands'] for m in direct_players) + sum(a['total_hands'] for a in agents_map.values())

        # Create a single SA object for template
        my_sa_combined = {
            'sa_id': sa_id, 'sa_nick': current_user.username,
            'club': my_sas[0]['club'] if my_sas else '',
            'agents': agents_map, 'direct': direct_players,
            'total_pnl': total_pnl, 'total_rake': total_rake, 'total_hands': total_hands,
        }
        my_sas = [my_sa_combined]

        # Managed clubs (multiple) - built from cumulative DB data
        rake_cfgs = SARakeConfig.query.filter_by(sa_id=sa_id).filter(SARakeConfig.managed_club_id.isnot(None)).all()
        rake_pct = rake_cfgs[0].rake_percent if rake_cfgs else 0
        managed_clubs = []
        club_net_rake = 0
        club_keeps_pct = 0
        if rake_cfgs:
            # Get club names from hierarchy (for club_id -> name mapping)
            clubs_data, _ = get_members_hierarchy()
            club_id_to_name = {c['club_id']: c['name'] for c in clubs_data}

            for cfg in rake_cfgs:
                club_name = club_id_to_name.get(cfg.managed_club_id, '')
                if not club_name:
                    continue

                # Build ID → nickname map from cumulative DB (includes SA/Agent name entries)
                all_nicknames = dict(DailyPlayerStats.query.with_entities(
                    DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname)
                ).group_by(DailyPlayerStats.player_id).all())

                # Get ALL players in this club from cumulative DB
                club_filters = [DailyPlayerStats.club == club_name]
                if upload_ids_filter:
                    club_filters.append(DailyPlayerStats.upload_id.in_(upload_ids_filter))
                club_players_db = DailyPlayerStats.query.with_entities(
                    DailyPlayerStats.player_id,
                    sqlfunc.max(DailyPlayerStats.nickname),
                    sqlfunc.max(DailyPlayerStats.sa_id),
                    sqlfunc.max(DailyPlayerStats.agent_id),
                    sqlfunc.sum(DailyPlayerStats.pnl),
                    sqlfunc.sum(DailyPlayerStats.rake),
                ).filter(*club_filters
                ).group_by(DailyPlayerStats.player_id).all()

                # Build SA structure from DB data
                club_sas = {}
                no_sa = []
                club_rake = 0
                club_pnl = 0
                for pid, nick, sa_id_val, ag_id_val, pnl_val, rake_val in club_players_db:
                    p = round(float(pnl_val or 0), 2)
                    r = round(float(rake_val or 0), 2)
                    club_rake += r
                    club_pnl += p
                    member = {'player_id': pid, 'nickname': nick, 'pnl_total': p, 'rake_total': r}

                    if sa_id_val and sa_id_val != '-':
                        if sa_id_val not in club_sas:
                            sa_nick = all_nicknames.get(sa_id_val, sa_id_val)
                            club_sas[sa_id_val] = {'nick': sa_nick, 'id': sa_id_val,
                                                    'agents': {}, 'direct_members': []}
                        sa = club_sas[sa_id_val]
                        if ag_id_val and ag_id_val != '-' and ag_id_val != sa_id_val:
                            if ag_id_val not in sa['agents']:
                                ag_nick = all_nicknames.get(ag_id_val, ag_id_val)
                                sa['agents'][ag_id_val] = {'nick': ag_nick, 'members': []}
                            sa['agents'][ag_id_val]['members'].append(member)
                        else:
                            sa['direct_members'].append(member)
                    else:
                        no_sa.append(member)

                club_rake = round(club_rake, 2)
                club_pnl = round(club_pnl, 2)

                club_obj = {
                    'name': club_name, 'club_id': cfg.managed_club_id,
                    'total_rake': club_rake, 'total_pnl': club_pnl,
                    'super_agents': club_sas, 'no_sa_members': no_sa,
                }
                total_rake += club_rake
                total_pnl += club_pnl
                club_rc = RakeConfig.query.filter_by(entity_type='club', entity_id=cfg.managed_club_id).first()
                keeps_pct = club_rc.rake_percent if club_rc else 0
                net = round(club_rake * (100 - keeps_pct) / 100, 2)
                club_net_rake += net
                club_keeps_pct = keeps_pct
                managed_clubs.append(club_obj)

        # Sort managed clubs by rake (high to low)
        managed_clubs.sort(key=lambda c: c.get('total_rake', 0), reverse=True)

        # Sort agents by rake (high to low)
        agents_sorted = dict(sorted(my_sa_combined['agents'].items(),
                                     key=lambda x: x[1].get('total_rake', 0), reverse=True))
        my_sa_combined['agents'] = agents_sorted

        personal_rake = round(my_sa_combined['total_rake'], 2)
        clubs_total_rake = round(sum(c.get('total_rake', 0) for c in managed_clubs), 2)
        sa_net_rake = round(personal_rake * rake_pct / 100, 2) if rake_pct else 0
        net_rake = round(sa_net_rake + club_net_rake, 2)
        player_count = len(all_my_player_ids)

        # My own rake percentage (if configured as sub_agent or agent)
        my_rake_rc = RakeConfig.query.filter(
            RakeConfig.entity_type.in_(['sub_agent', 'agent']),
            RakeConfig.entity_id == sa_id).first()
        my_rake_pct = my_rake_rc.rake_percent if my_rake_rc else 0
        my_rake_earning = round(personal_rake * my_rake_pct / 100, 2) if my_rake_pct else 0

        # Expense charges for this agent
        expense_charges = ExpenseCharge.query.filter_by(agent_player_id=sa_id).all()
        total_expenses = round(sum(c.charge_amount for c in expense_charges), 2)
        net_rake_after_expenses = round(net_rake - total_expenses, 2)

        return render_template('main/agent_dashboard.html',
                               my_sas=my_sas, child_sas=child_sas,
                               managed_clubs=managed_clubs,
                               total_rake=total_rake, total_pnl=total_pnl,
                               total_hands=int(total_hands), net_rake=net_rake,
                               personal_rake=personal_rake,
                               clubs_total_rake=clubs_total_rake,
                               net_rake_after_expenses=net_rake_after_expenses,
                               total_expenses=total_expenses,
                               expense_charges=expense_charges,
                               rake_refund_list=rake_refund_list,
                               total_rake_refund=total_rake_refund,
                               my_rake_pct=my_rake_pct,
                               my_rake_earning=my_rake_earning,
                               rake_pct=rake_pct, player_count=player_count,
                               club_net_rake=club_net_rake,
                               club_keeps_pct=club_keeps_pct,
                               available_dates=available_dates,
                               selected_dates=selected_dates)

    if hasattr(current_user, 'role') and current_user.role == 'player' and current_user.player_id:
        from app.union_data import get_cumulative_stats
        from app.models import PlayerSession

        player_id = current_user.player_id
        cs = get_cumulative_stats([player_id]).get(player_id)
        sessions = PlayerSession.query.filter_by(player_id=player_id).all()
        session_list = [{'table_name': s.table_name, 'game_type': s.game_type,
                         'blinds': s.blinds or '', 'pnl': round(s.pnl, 2)} for s in sessions]

        return render_template('main/player_dashboard.html',
                               player=cs or {'nickname': current_user.username, 'club': '-', 'pnl': 0, 'rake': 0, 'hands': 0},
                               sessions=session_list)

    transactions = (Transaction.query
                    .filter_by(user_id=current_user.id)
                    .order_by(Transaction.date.desc())
                    .limit(5)
                    .all())

    total_income = db.session.query(func.sum(Transaction.amount)).filter_by(
        user_id=current_user.id, type='income').scalar() or 0

    total_expense = db.session.query(func.sum(Transaction.amount)).filter_by(
        user_id=current_user.id, type='expense').scalar() or 0

    balance = total_income - total_expense

    return render_template('main/dashboard.html',
                           transactions=transactions,
                           total_income=total_income,
                           total_expense=total_expense,
                           balance=balance)


@main_bp.route('/agent/reports')
@login_required
def agent_reports():
    if not hasattr(current_user, 'role') or current_user.role != 'agent' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.models import SAHierarchy, SARakeConfig, DailyPlayerStats
    from sqlalchemy import func as sqlfunc

    sa_id = current_user.player_id

    # Collect ALL player IDs in the box: direct SA players + child SAs + managed clubs
    all_sa_ids = [sa_id]
    child_sa_ids = [h.child_sa_id for h in SAHierarchy.query.filter_by(parent_sa_id=sa_id).all()]
    all_sa_ids.extend(child_sa_ids)

    # Players under my SAs (from cumulative DB)
    sa_players = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname)
    ).filter(
        DailyPlayerStats.sa_id.in_(all_sa_ids),
        DailyPlayerStats.role != 'Name Entry'
    ).group_by(DailyPlayerStats.player_id).all()

    my_players = []
    my_player_ids = set()
    for pid, nick in sa_players:
        my_players.append({'player_id': pid, 'nickname': nick})
        my_player_ids.add(pid)

    # Players from managed clubs
    from app.union_data import get_members_hierarchy
    rake_cfgs = SARakeConfig.query.filter_by(sa_id=sa_id).filter(SARakeConfig.managed_club_id.isnot(None)).all()
    if rake_cfgs:
        clubs_data, _ = get_members_hierarchy()
        club_id_to_name = {c['club_id']: c['name'] for c in clubs_data}
        for cfg in rake_cfgs:
            club_name = club_id_to_name.get(cfg.managed_club_id)
            if club_name:
                club_players = DailyPlayerStats.query.with_entities(
                    DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname)
                ).filter(
                    DailyPlayerStats.club == club_name,
                    DailyPlayerStats.role != 'Name Entry'
                ).group_by(DailyPlayerStats.player_id).all()
                for pid, nick in club_players:
                    if pid not in my_player_ids:
                        my_players.append({'player_id': pid, 'nickname': nick})
                        my_player_ids.add(pid)

    my_players.sort(key=lambda x: x['nickname'].lower())

    return render_template('main/agent_reports.html',
                           players=my_players, player_ids=list(my_player_ids))


# ═══════════════════════ EXCEL EXPORTS ═══════════════════════

def _make_excel(sheets_data, filename):
    """Create Excel file from dict of {sheet_name: [{col: val, ...}]}."""
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sheet_name, rows in sheets_data.items():
        import re
        safe_name = re.sub(r'[\[\]\*\?:/\\]', '', sheet_name)[:31] or 'Sheet'
        ws = wb.create_sheet(title=safe_name)
        if not rows:
            continue
        # Headers
        headers = list(rows[0].keys())
        header_font = Font(bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color='4361EE', end_color='4361EE', fill_type='solid')
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')
        # Data with color formatting
        green_font = Font(color='2EC4B6', bold=True)
        red_font = Font(color='EF233C', bold=True)
        bold_font = Font(bold=True)
        for row_idx, row_data in enumerate(rows, 2):
            for col_idx, key in enumerate(headers, 1):
                val = row_data.get(key, '')
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                # Color P&L values
                if key in ('P&L', 'נטו שלי') and isinstance(val, (int, float)):
                    if val > 0:
                        cell.font = green_font
                    elif val < 0:
                        cell.font = red_font
                # Bold totals row
                if row_data.get(headers[0]) == 'סה"כ':
                    cell.font = Font(bold=True, color=cell.font.color if cell.font.color else '000000')
        # Auto-width
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 40)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@main_bp.route('/export/player/<player_id>')
@login_required
def export_player(player_id):
    """Export player personal report - all games, P&L, record."""
    from app.union_data import get_cumulative_stats
    from app.models import PlayerSession

    cs = get_cumulative_stats([player_id]).get(player_id)
    if not cs:
        flash('שחקן לא נמצא.', 'danger')
        return redirect(url_for('main.dashboard'))

    sessions = PlayerSession.query.filter_by(player_id=player_id).all()
    session_rows = [{'משחק': s.table_name, 'סוג': s.game_type,
                     'בליינדס': s.blinds or '', 'P&L': round(s.pnl, 2)} for s in sessions]

    summary = [{'שחקן': cs['nickname'], 'קלאב': cs['club'],
                'P&L': cs['pnl'], 'Hands': cs['hands']}]

    return _make_excel({
        'סיכום': summary,
        'רקורד משחקים': session_rows,
    }, f'{cs["nickname"]}_report.xlsx')


@main_bp.route('/export/agent/account')
@login_required
def export_agent_account():
    """Export agent account summary - personal rake, club rake, expenses, net."""
    if current_user.role != 'agent' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.models import SAHierarchy, SARakeConfig, RakeConfig, ExpenseCharge, DailyPlayerStats
    from app.union_data import get_members_hierarchy
    from sqlalchemy import func as sqlfunc

    sa_id = current_user.player_id

    # Personal rake
    personal = DailyPlayerStats.query.with_entities(
        sqlfunc.sum(DailyPlayerStats.rake), sqlfunc.sum(DailyPlayerStats.pnl)
    ).filter(DailyPlayerStats.sa_id == sa_id, DailyPlayerStats.role != 'Name Entry').first()
    personal_rake = round(float(personal[0] or 0), 2)
    personal_pnl = round(float(personal[1] or 0), 2)

    # Club rakes
    rake_cfgs = SARakeConfig.query.filter_by(sa_id=sa_id).filter(SARakeConfig.managed_club_id.isnot(None)).all()
    clubs_data, _ = get_members_hierarchy()
    club_id_to_name = {c['club_id']: c['name'] for c in clubs_data}
    club_rows = []
    total_club_rake = 0
    for cfg in rake_cfgs:
        name = club_id_to_name.get(cfg.managed_club_id, '')
        if name:
            cr = DailyPlayerStats.query.with_entities(
                sqlfunc.sum(DailyPlayerStats.rake), sqlfunc.sum(DailyPlayerStats.pnl)
            ).filter(DailyPlayerStats.club == name).first()
            rake = round(float(cr[0] or 0), 2)
            pnl = round(float(cr[1] or 0), 2)
            club_rc = RakeConfig.query.filter_by(entity_type='club', entity_id=cfg.managed_club_id).first()
            keeps = club_rc.rake_percent if club_rc else 0
            net = round(rake * (100 - keeps) / 100, 2)
            club_rows.append({'מועדון': name, 'Rake': rake, 'P&L': pnl,
                              'מועדון מקבל %': keeps, 'נטו שלי': net})
            total_club_rake += net

    # Expenses
    charges = ExpenseCharge.query.filter_by(agent_player_id=sa_id).all()
    expense_rows = [{'הוצאה': c.expense.description if c.expense else '', 'סכום': c.charge_amount,
                     'תאריך': c.created_at.strftime('%d/%m/%Y')} for c in charges]
    total_expenses = round(sum(c.charge_amount for c in charges), 2)

    summary = [{'סוכן': current_user.username, 'רייק אישי': personal_rake,
                'רייק מועדונים (נטו)': total_club_rake, 'הוצאות משותפות': total_expenses,
                'P&L': personal_pnl}]

    sheets = {'סיכום חשבון': summary}
    if club_rows:
        sheets['מועדונים'] = club_rows
    if expense_rows:
        sheets['הוצאות'] = expense_rows
    return _make_excel(sheets, f'{current_user.username}_account.xlsx')


@main_bp.route('/export/agent/players')
@login_required
def export_agent_players():
    """Export all agent's players, agents, SAs, clubs with rake % and totals."""
    if current_user.role != 'agent' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.models import SAHierarchy, SARakeConfig, DailyPlayerStats, RakeConfig
    from app.union_data import get_members_hierarchy
    from sqlalchemy import func as sqlfunc

    sa_id = current_user.player_id
    all_sa_ids = [sa_id]
    child_sa_ids = [h.child_sa_id for h in SAHierarchy.query.filter_by(parent_sa_id=sa_id).all()]
    all_sa_ids.extend(child_sa_ids)

    # Nickname map
    all_nicks = dict(DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname)
    ).group_by(DailyPlayerStats.player_id).all())

    sheets = {}

    # ── Sheet 1: My Players (direct) ──
    players = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname),
        sqlfunc.max(DailyPlayerStats.club), sqlfunc.max(DailyPlayerStats.sa_id),
        sqlfunc.max(DailyPlayerStats.agent_id),
        sqlfunc.sum(DailyPlayerStats.pnl), sqlfunc.sum(DailyPlayerStats.rake),
        sqlfunc.sum(DailyPlayerStats.hands),
    ).filter(DailyPlayerStats.sa_id.in_(all_sa_ids), DailyPlayerStats.role != 'Name Entry'
    ).group_by(DailyPlayerStats.player_id).all()

    # Group players by agent - each agent gets its own sheet
    agent_groups = {}  # agent_name -> [players]
    direct_players = []
    for p in players:
        ag_id = p[4] if p[4] and p[4] != '-' else None
        ag_name = all_nicks.get(ag_id, ag_id) if ag_id else None
        row = {
            'שחקן': p[1], 'ID': p[0], 'קלאב': p[2],
            'P&L': round(float(p[5] or 0), 2),
            'Rake': round(float(p[6] or 0), 2),
            'Hands': int(p[7] or 0),
        }
        if ag_name and ag_name != all_nicks.get(sa_id, sa_id):
            if ag_name not in agent_groups:
                agent_groups[ag_name] = []
            agent_groups[ag_name].append(row)
        else:
            direct_players.append(row)

    # Create sheet per agent
    for ag_name, ag_players in sorted(agent_groups.items(), key=lambda x: sum(r['Rake'] for r in x[1]), reverse=True):
        ag_players.sort(key=lambda x: x['Rake'], reverse=True)
        ag_players.append({
            'שחקן': 'סה"כ', 'ID': '', 'קלאב': '',
            'P&L': round(sum(r['P&L'] for r in ag_players), 2),
            'Rake': round(sum(r['Rake'] for r in ag_players), 2),
            'Hands': sum(r['Hands'] for r in ag_players),
        })
        sheets[ag_name[:31]] = ag_players

    # Direct players sheet
    if direct_players:
        direct_players.sort(key=lambda x: x['Rake'], reverse=True)
        direct_players.append({
            'שחקן': 'סה"כ', 'ID': '', 'קלאב': '',
            'P&L': round(sum(r['P&L'] for r in direct_players), 2),
            'Rake': round(sum(r['Rake'] for r in direct_players), 2),
            'Hands': sum(r['Hands'] for r in direct_players),
        })
        sheets['שחקנים ישירים'] = direct_players

    # ── Sheet 2: My Agents ──
    agent_stats = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.agent_id,
        sqlfunc.sum(DailyPlayerStats.pnl), sqlfunc.sum(DailyPlayerStats.rake),
        sqlfunc.sum(DailyPlayerStats.hands), sqlfunc.count(sqlfunc.distinct(DailyPlayerStats.player_id)),
    ).filter(
        DailyPlayerStats.sa_id.in_(all_sa_ids), DailyPlayerStats.role != 'Name Entry',
        DailyPlayerStats.agent_id != '', DailyPlayerStats.agent_id != '-'
    ).group_by(DailyPlayerStats.agent_id).all()

    agent_rows = []
    for ag in agent_stats:
        ag_name = all_nicks.get(ag[0], ag[0])
        rc = RakeConfig.query.filter_by(entity_type='agent', entity_id=ag[0]).first()
        rake_pct = rc.rake_percent if rc else 0
        rake = round(float(ag[2] or 0), 2)
        agent_rows.append({
            'סוכן': ag_name, 'ID': ag[0], 'שחקנים': int(ag[4] or 0),
            'P&L': round(float(ag[1] or 0), 2), 'Rake': rake,
            'אחוז רייק %': rake_pct,
            'Hands': int(ag[3] or 0),
        })
    agent_rows.sort(key=lambda x: x['Rake'], reverse=True)
    if agent_rows:
        agent_rows.append({
            'סוכן': 'סה"כ', 'ID': '', 'שחקנים': sum(r['שחקנים'] for r in agent_rows),
            'P&L': round(sum(r['P&L'] for r in agent_rows), 2),
            'Rake': round(sum(r['Rake'] for r in agent_rows), 2),
            'אחוז רייק %': '', 'Hands': sum(r['Hands'] for r in agent_rows),
        })
    sheets['סוכנים'] = agent_rows

    # ── Sheet 3: My Super Agents ──
    sa_rows = []
    for csa_id in child_sa_ids:
        sa_data = DailyPlayerStats.query.with_entities(
            sqlfunc.sum(DailyPlayerStats.pnl), sqlfunc.sum(DailyPlayerStats.rake),
            sqlfunc.sum(DailyPlayerStats.hands), sqlfunc.count(sqlfunc.distinct(DailyPlayerStats.player_id)),
        ).filter(DailyPlayerStats.sa_id == csa_id, DailyPlayerStats.role != 'Name Entry').first()
        sa_name = all_nicks.get(csa_id, csa_id)
        rc = RakeConfig.query.filter_by(entity_type='agent', entity_id=csa_id).first()
        rake_pct = rc.rake_percent if rc else 0
        rake = round(float(sa_data[1] or 0), 2)
        sa_rows.append({
            'Super Agent': sa_name, 'ID': csa_id, 'שחקנים': int(sa_data[3] or 0),
            'P&L': round(float(sa_data[0] or 0), 2), 'Rake': rake,
            'אחוז רייק %': rake_pct,
            'Hands': int(sa_data[2] or 0),
        })
    sa_rows.sort(key=lambda x: x['Rake'], reverse=True)
    if sa_rows:
        sa_rows.append({
            'Super Agent': 'סה"כ', 'ID': '', 'שחקנים': sum(r['שחקנים'] for r in sa_rows),
            'P&L': round(sum(r['P&L'] for r in sa_rows), 2),
            'Rake': round(sum(r['Rake'] for r in sa_rows), 2),
            'אחוז רייק %': '', 'Hands': sum(r['Hands'] for r in sa_rows),
        })
    if sa_rows:
        sheets['Super Agents'] = sa_rows

    # ── Sheet 4: My Clubs ──
    rake_cfgs = SARakeConfig.query.filter_by(sa_id=sa_id).filter(SARakeConfig.managed_club_id.isnot(None)).all()
    if rake_cfgs:
        clubs_data, _ = get_members_hierarchy()
        club_id_to_name = {c['club_id']: c['name'] for c in clubs_data}
        club_rows = []
        for cfg in rake_cfgs:
            name = club_id_to_name.get(cfg.managed_club_id)
            if not name:
                continue
            cr = DailyPlayerStats.query.with_entities(
                sqlfunc.sum(DailyPlayerStats.pnl), sqlfunc.sum(DailyPlayerStats.rake),
                sqlfunc.sum(DailyPlayerStats.hands), sqlfunc.count(sqlfunc.distinct(DailyPlayerStats.player_id)),
            ).filter(DailyPlayerStats.club == name, DailyPlayerStats.role != 'Name Entry').first()
            club_rc = RakeConfig.query.filter_by(entity_type='club', entity_id=cfg.managed_club_id).first()
            keeps = club_rc.rake_percent if club_rc else 0
            rake = round(float(cr[1] or 0), 2)
            net = round(rake * (100 - keeps) / 100, 2)
            club_rows.append({
                'מועדון': name, 'שחקנים': int(cr[3] or 0),
                'P&L': round(float(cr[0] or 0), 2), 'Rake': rake,
                'מועדון מקבל %': keeps, 'נטו שלי': net,
                'Hands': int(cr[2] or 0),
            })
        club_rows.sort(key=lambda x: x['Rake'], reverse=True)
        if club_rows:
            club_rows.append({
                'מועדון': 'סה"כ', 'שחקנים': sum(r['שחקנים'] for r in club_rows),
                'P&L': round(sum(r['P&L'] for r in club_rows), 2),
                'Rake': round(sum(r['Rake'] for r in club_rows), 2),
                'מועדון מקבל %': '', 'נטו שלי': round(sum(r['נטו שלי'] for r in club_rows), 2),
                'Hands': sum(r['Hands'] for r in club_rows),
            })
        sheets['מועדונים'] = club_rows

    return _make_excel(sheets, f'{current_user.username}_players.xlsx')


@main_bp.route('/export/agent/club/<club_id>')
@login_required
def export_agent_club(club_id):
    """Export specific club details - SAs, Agents, Players."""
    if current_user.role != 'agent' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.models import DailyPlayerStats
    from app.union_data import get_members_hierarchy
    from sqlalchemy import func as sqlfunc

    clubs_data, _ = get_members_hierarchy()
    club_name = None
    for c in clubs_data:
        if c['club_id'] == club_id:
            club_name = c['name']
            break
    if not club_name:
        flash('מועדון לא נמצא.', 'danger')
        return redirect(url_for('main.dashboard'))

    players = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname),
        sqlfunc.max(DailyPlayerStats.sa_id), sqlfunc.max(DailyPlayerStats.agent_id),
        sqlfunc.max(DailyPlayerStats.role), sqlfunc.sum(DailyPlayerStats.pnl),
        sqlfunc.sum(DailyPlayerStats.rake), sqlfunc.sum(DailyPlayerStats.hands),
    ).filter(DailyPlayerStats.club == club_name, DailyPlayerStats.role != 'Name Entry'
    ).group_by(DailyPlayerStats.player_id).all()

    # Build nickname map
    all_nicks = dict(DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname)
    ).group_by(DailyPlayerStats.player_id).all())

    rows = [{'שחקן': p[1], 'ID': p[0],
             'Super Agent': all_nicks.get(p[2], p[2]) if p[2] and p[2] != '-' else '',
             'Agent': all_nicks.get(p[3], p[3]) if p[3] and p[3] != '-' else '',
             'תפקיד': p[4], 'P&L': round(float(p[5] or 0), 2),
             'Rake': round(float(p[6] or 0), 2), 'Hands': int(p[7] or 0)} for p in players]
    rows.sort(key=lambda x: x['Rake'], reverse=True)

    return _make_excel({club_name: rows}, f'{club_name}_report.xlsx')


@main_bp.route('/export/agent/period')
@login_required
def export_agent_period():
    """Export agent data for specific date range."""
    if current_user.role != 'agent' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.models import SAHierarchy, SARakeConfig, DailyPlayerStats, DailyUpload
    from app.union_data import get_members_hierarchy
    from sqlalchemy import func as sqlfunc
    from datetime import datetime

    from_date = request.args.get('from', '')
    to_date = request.args.get('to', '')
    if not from_date or not to_date:
        flash('יש לבחור תאריכים.', 'danger')
        return redirect(url_for('main.agent_reports'))

    fd = datetime.strptime(from_date, '%Y-%m-%d').date()
    td = datetime.strptime(to_date, '%Y-%m-%d').date()

    sa_id = current_user.player_id
    all_sa_ids = [sa_id]
    child_sa_ids = [h.child_sa_id for h in SAHierarchy.query.filter_by(parent_sa_id=sa_id).all()]
    all_sa_ids.extend(child_sa_ids)

    # Get uploads in range
    uploads = DailyUpload.query.filter(DailyUpload.upload_date >= fd, DailyUpload.upload_date <= td).all()
    upload_ids = [u.id for u in uploads]
    if not upload_ids:
        flash('אין נתונים בטווח התאריכים.', 'warning')
        return redirect(url_for('main.agent_reports'))

    players = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname),
        sqlfunc.max(DailyPlayerStats.club), sqlfunc.sum(DailyPlayerStats.pnl),
        sqlfunc.sum(DailyPlayerStats.rake), sqlfunc.sum(DailyPlayerStats.hands),
    ).filter(
        DailyPlayerStats.upload_id.in_(upload_ids),
        DailyPlayerStats.sa_id.in_(all_sa_ids),
        DailyPlayerStats.role != 'Name Entry'
    ).group_by(DailyPlayerStats.player_id).all()

    rows = [{'שחקן': p[1], 'ID': p[0], 'קלאב': p[2],
             'P&L': round(float(p[3] or 0), 2), 'Rake': round(float(p[4] or 0), 2),
             'Hands': int(p[5] or 0)} for p in players]
    rows.sort(key=lambda x: x['Rake'], reverse=True)

    return _make_excel({f'{from_date} - {to_date}': rows},
                       f'{current_user.username}_{from_date}_{to_date}.xlsx')


@main_bp.route('/export/club/report')
@login_required
def export_club_report():
    """Export club report - all SAs, Agents, Players with balances."""
    if current_user.role != 'club' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.models import DailyPlayerStats, DailyUpload
    from app.union_data import get_members_hierarchy
    from sqlalchemy import func as sqlfunc

    club_id = current_user.player_id
    clubs_data, _ = get_members_hierarchy()
    club_name = None
    for c in clubs_data:
        if c['club_id'] == club_id:
            club_name = c['name']
            break
    if not club_name:
        flash('מועדון לא נמצא.', 'danger')
        return redirect(url_for('main.dashboard'))

    # Date filter
    dates_str = request.args.get('dates', '')
    base_filters = [DailyPlayerStats.club == club_name, DailyPlayerStats.role != 'Name Entry']
    filename_suffix = ''
    if dates_str:
        from datetime import datetime as dt
        upload_ids = []
        for ds in dates_str.split(','):
            ds = ds.strip()
            try:
                sel = dt.strptime(ds, '%Y-%m-%d').date()
                upload = DailyUpload.query.filter_by(upload_date=sel).first()
                if upload:
                    upload_ids.append(upload.id)
            except ValueError:
                pass
        if upload_ids:
            base_filters.append(DailyPlayerStats.upload_id.in_(upload_ids))
            filename_suffix = f'_{dates_str.replace(",", "_")}'

    # All players in this club
    players = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname),
        sqlfunc.max(DailyPlayerStats.sa_id), sqlfunc.max(DailyPlayerStats.agent_id),
        sqlfunc.sum(DailyPlayerStats.pnl), sqlfunc.sum(DailyPlayerStats.rake),
        sqlfunc.sum(DailyPlayerStats.hands),
    ).filter(*base_filters).group_by(DailyPlayerStats.player_id).all()

    # Nickname map
    all_nicks = dict(DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname)
    ).group_by(DailyPlayerStats.player_id).all())

    import re
    sheets = {}

    # Group by SA - each SA gets its own sheet with all their players
    sa_groups = {}  # sa_id -> [players]
    no_sa_players = []
    for p in players:
        sa_id = p[2] if p[2] and p[2] != '-' else None
        ag_id = p[3] if p[3] and p[3] != '-' else None
        ag_name = all_nicks.get(ag_id, ag_id) if ag_id else ''
        row = {
            'שחקן': p[1], 'ID': p[0],
            'Agent': ag_name,
            'P&L': round(float(p[4] or 0), 2),
            'Rake': round(float(p[5] or 0), 2),
            'Hands': int(p[6] or 0),
        }
        if sa_id:
            if sa_id not in sa_groups:
                sa_groups[sa_id] = []
            sa_groups[sa_id].append(row)
        else:
            no_sa_players.append(row)

    # Sheet per SA
    for sa_id, sa_players in sorted(sa_groups.items(), key=lambda x: sum(r['Rake'] for r in x[1]), reverse=True):
        sa_name = all_nicks.get(sa_id, sa_id)
        sa_players.sort(key=lambda x: x['Rake'], reverse=True)
        sa_players.append({
            'שחקן': 'סה"כ', 'ID': '', 'Agent': '',
            'P&L': round(sum(r['P&L'] for r in sa_players), 2),
            'Rake': round(sum(r['Rake'] for r in sa_players), 2),
            'Hands': sum(r['Hands'] for r in sa_players),
        })
        safe_name = re.sub(r'[\[\]\*\?:/\\]', '', sa_name)[:31] or 'SA'
        sheets[safe_name] = sa_players

    # Players without SA
    if no_sa_players:
        no_sa_players.sort(key=lambda x: x['Rake'], reverse=True)
        no_sa_players.append({
            'שחקן': 'סה"כ', 'ID': '', 'Agent': '',
            'P&L': round(sum(r['P&L'] for r in no_sa_players), 2),
            'Rake': round(sum(r['Rake'] for r in no_sa_players), 2),
            'Hands': sum(r['Hands'] for r in no_sa_players),
        })
        sheets['ללא SA'] = no_sa_players

    # Summary sheet - all SAs
    sa_rows = []
    for sa_id, sa_players_list in sa_groups.items():
        real_players = [p for p in sa_players_list if p['שחקן'] != 'סה"כ']
        sa_rows.append({
            'Super Agent': all_nicks.get(sa_id, sa_id), 'ID': sa_id,
            'שחקנים': len(real_players),
            'P&L': round(sum(r['P&L'] for r in real_players), 2),
            'Rake': round(sum(r['Rake'] for r in real_players), 2),
            'Hands': sum(r['Hands'] for r in real_players),
        })
    sa_rows.sort(key=lambda x: x['Rake'], reverse=True)
    if sa_rows:
        sa_rows.append({
            'Super Agent': 'סה"כ', 'ID': '', 'שחקנים': sum(r['שחקנים'] for r in sa_rows),
            'P&L': round(sum(r['P&L'] for r in sa_rows), 2),
            'Rake': round(sum(r['Rake'] for r in sa_rows), 2),
            'Hands': sum(r['Hands'] for r in sa_rows),
        })
        sheets['Super Agents'] = sa_rows

    return _make_excel(sheets, f'{club_name}_report{filename_suffix}.xlsx')


@main_bp.route('/club/reports')
@login_required
def club_reports():
    if not hasattr(current_user, 'role') or current_user.role != 'club' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.models import DailyPlayerStats
    from app.union_data import get_members_hierarchy
    from sqlalchemy import func as sqlfunc

    club_id = current_user.player_id
    clubs_data, _ = get_members_hierarchy()
    club_name = None
    for c in clubs_data:
        if c['club_id'] == club_id:
            club_name = c['name']
            break
    if not club_name:
        flash('מועדון לא נמצא.', 'danger')
        return redirect(url_for('main.dashboard'))

    # All players in this club
    club_players = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname)
    ).filter(
        DailyPlayerStats.club == club_name,
        DailyPlayerStats.role != 'Name Entry'
    ).group_by(DailyPlayerStats.player_id).all()

    players = [{'player_id': pid, 'nickname': nick} for pid, nick in club_players]
    player_ids = [pid for pid, _ in club_players]

    return render_template('main/club_reports.html', players=players, player_ids=player_ids)


@main_bp.route('/export/club/period')
@login_required
def export_club_period():
    """Export club data for specific date range."""
    if current_user.role != 'club' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.models import DailyPlayerStats, DailyUpload
    from app.union_data import get_members_hierarchy
    from sqlalchemy import func as sqlfunc
    from datetime import datetime

    from_date = request.args.get('from', '')
    to_date = request.args.get('to', '')
    if not from_date or not to_date:
        flash('יש לבחור תאריכים.', 'danger')
        return redirect(url_for('main.club_reports'))

    fd = datetime.strptime(from_date, '%Y-%m-%d').date()
    td = datetime.strptime(to_date, '%Y-%m-%d').date()

    club_id = current_user.player_id
    clubs_data, _ = get_members_hierarchy()
    club_name = None
    for c in clubs_data:
        if c['club_id'] == club_id:
            club_name = c['name']
            break
    if not club_name:
        flash('מועדון לא נמצא.', 'danger')
        return redirect(url_for('main.club_reports'))

    uploads = DailyUpload.query.filter(DailyUpload.upload_date >= fd, DailyUpload.upload_date <= td).all()
    upload_ids = [u.id for u in uploads]
    if not upload_ids:
        flash('אין נתונים בטווח התאריכים.', 'warning')
        return redirect(url_for('main.club_reports'))

    players = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname),
        sqlfunc.max(DailyPlayerStats.club), sqlfunc.sum(DailyPlayerStats.pnl),
        sqlfunc.sum(DailyPlayerStats.rake), sqlfunc.sum(DailyPlayerStats.hands),
    ).filter(
        DailyPlayerStats.upload_id.in_(upload_ids),
        DailyPlayerStats.club == club_name,
        DailyPlayerStats.role != 'Name Entry'
    ).group_by(DailyPlayerStats.player_id).all()

    rows = [{'שחקן': p[1], 'ID': p[0], 'קלאב': p[2],
             'P&L': round(float(p[3] or 0), 2), 'Rake': round(float(p[4] or 0), 2),
             'Hands': int(p[5] or 0)} for p in players]
    rows.sort(key=lambda x: x['Rake'], reverse=True)

    return _make_excel({f'{from_date} - {to_date}': rows},
                       f'{club_name}_{from_date}_{to_date}.xlsx')


@main_bp.route('/club/transfers', methods=['GET', 'POST'])
@login_required
def club_transfers():
    if not hasattr(current_user, 'role') or current_user.role != 'club' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.union_data import get_player_balance, get_all_balances, get_members_hierarchy
    from app.models import MoneyTransfer, DailyPlayerStats
    from sqlalchemy import func as sqlfunc

    club_id = current_user.player_id
    clubs_data, _ = get_members_hierarchy()
    club_name = None
    for c in clubs_data:
        if c['club_id'] == club_id:
            club_name = c['name']
            break
    if not club_name:
        flash('מועדון לא נמצא.', 'danger')
        return redirect(url_for('main.dashboard'))

    # All players in this club
    club_players_db = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname)
    ).filter(
        DailyPlayerStats.club == club_name,
        DailyPlayerStats.role != 'Name Entry'
    ).group_by(DailyPlayerStats.player_id).all()

    my_player_ids = set()
    my_players = []
    for pid, nick in club_players_db:
        my_player_ids.add(pid)
        my_players.append({'player_id': pid, 'nickname': nick})

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
                return redirect(url_for('main.club_transfers'))

            if not from_key or not to_key or '|' not in from_key or '|' not in to_key:
                flash('יש לבחור שולח ומקבל.', 'danger')
            elif from_key == to_key:
                flash('לא ניתן להעביר לאותו שחקן.', 'warning')
            elif amount <= 0:
                flash('הסכום חייב להיות חיובי.', 'danger')
            else:
                from_pid = from_key.split('|', 1)[0]
                to_pid = to_key.split('|', 1)[0]
                from_name = from_key.split('|', 1)[1]
                to_name = to_key.split('|', 1)[1]
                if from_pid not in my_player_ids or to_pid not in my_player_ids:
                    flash('אין הרשאה להעביר לשחקן שלא שייך למועדון.', 'danger')
                else:
                    from_balance = get_player_balance(from_pid)
                    to_balance = get_player_balance(to_pid)
                    max_transfer = min(abs(from_balance), to_balance)
                    if from_balance >= 0:
                        flash(f'{from_name} לא במינוס.', 'danger')
                    elif to_balance <= 0:
                        flash(f'{to_name} לא בפלוס.', 'danger')
                    elif amount > max_transfer:
                        flash(f'חריגה! מקסימום: {max_transfer:.2f}', 'danger')
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
            if t and (t.from_player_id in my_player_ids or t.to_player_id in my_player_ids):
                db.session.delete(t)
                db.session.commit()
                flash('העברה נמחקה.', 'success')
        return redirect(url_for('main.club_transfers'))

    balances = get_all_balances(my_player_ids)
    my_transfers = MoneyTransfer.query.filter(
        db.or_(
            MoneyTransfer.from_player_id.in_(my_player_ids),
            MoneyTransfer.to_player_id.in_(my_player_ids)
        )
    ).order_by(MoneyTransfer.created_at.desc()).all()

    return render_template('main/club_transfers.html',
                           players=my_players, balances=balances,
                           transfers=my_transfers)


@main_bp.route('/agent/transfers', methods=['GET', 'POST'])
@login_required
def agent_transfers():
    if not hasattr(current_user, 'role') or current_user.role != 'agent' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.union_data import get_super_agent_tables, get_player_balance, get_all_balances
    from app.models import MoneyTransfer, SAHierarchy

    sa_id = current_user.player_id
    sa_tables = get_super_agent_tables()

    # Collect all player IDs under this agent
    my_player_ids = set()
    my_players = []  # [{player_id, nickname, club}]
    my_sas = [sa for sa in sa_tables if sa['sa_id'] == sa_id]
    child_sa_ids = [h.child_sa_id for h in SAHierarchy.query.filter_by(parent_sa_id=sa_id).all()]
    child_sas = [sa for sa in sa_tables if sa['sa_id'] in child_sa_ids]

    for sa in my_sas + child_sas:
        for m in sa['direct']:
            my_player_ids.add(m['player_id'])
            my_players.append({'player_id': m['player_id'], 'nickname': m['nickname'], 'club': sa['club']})
        for ag in sa['agents'].values():
            for m in ag['members']:
                my_player_ids.add(m['player_id'])
                my_players.append({'player_id': m['player_id'], 'nickname': m['nickname'], 'club': sa['club']})

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
                return redirect(url_for('main.agent_transfers'))

            if not from_key or not to_key or '|' not in from_key or '|' not in to_key:
                flash('יש לבחור שולח ומקבל.', 'danger')
            elif from_key == to_key:
                flash('לא ניתן להעביר לאותו שחקן.', 'warning')
            elif amount <= 0:
                flash('הסכום חייב להיות חיובי.', 'danger')
            else:
                from_pid = from_key.split('|', 1)[0]
                to_pid = to_key.split('|', 1)[0]
                from_name = from_key.split('|', 1)[1]
                to_name = to_key.split('|', 1)[1]
                # Verify both players belong to this agent
                if from_pid not in my_player_ids or to_pid not in my_player_ids:
                    flash('אין הרשאה להעביר לשחקן שלא שייך אליך.', 'danger')
                else:
                    from_balance = get_player_balance(from_pid)
                    to_balance = get_player_balance(to_pid)
                    max_transfer = min(abs(from_balance), to_balance)
                    if from_balance >= 0:
                        flash(f'{from_name} לא במינוס.', 'danger')
                    elif to_balance <= 0:
                        flash(f'{to_name} לא בפלוס.', 'danger')
                    elif amount > max_transfer:
                        flash(f'חריגה! מקסימום: {max_transfer:.2f}', 'danger')
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
            if t and (t.from_player_id in my_player_ids or t.to_player_id in my_player_ids):
                db.session.delete(t)
                db.session.commit()
                flash('העברה נמחקה.', 'success')
        return redirect(url_for('main.agent_transfers'))

    balances = get_all_balances(my_player_ids)
    # Get transfers for my players only
    my_transfers = MoneyTransfer.query.filter(
        db.or_(
            MoneyTransfer.from_player_id.in_(my_player_ids),
            MoneyTransfer.to_player_id.in_(my_player_ids)
        )
    ).order_by(MoneyTransfer.created_at.desc()).all()

    return render_template('main/agent_transfers.html',
                           players=my_players, balances=balances,
                           transfers=my_transfers)


@main_bp.route('/api/report')
@login_required
def report_api():
    from app.models import DailyPlayerStats, DailyUpload
    from datetime import datetime
    from sqlalchemy import func

    from_date = request.args.get('from')
    to_date = request.args.get('to')
    player_id = request.args.get('player_id', '')

    if not from_date or not to_date:
        return jsonify({'error': 'missing dates'}), 400

    try:
        fd = datetime.strptime(from_date, '%Y-%m-%d').date()
        td = datetime.strptime(to_date, '%Y-%m-%d').date()
    except ValueError:
        return jsonify({'error': 'invalid date format'}), 400

    # Get upload IDs in range
    uploads = DailyUpload.query.filter(
        DailyUpload.upload_date >= fd,
        DailyUpload.upload_date <= td
    ).all()
    upload_ids = [u.id for u in uploads]

    if not upload_ids:
        return jsonify({'players': [], 'totals': {'pnl': 0, 'rake': 0, 'hands': 0}, 'days': 0})

    query = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id,
        func.max(DailyPlayerStats.nickname),
        func.max(DailyPlayerStats.club),
        func.sum(DailyPlayerStats.pnl),
        func.sum(DailyPlayerStats.rake),
        func.sum(DailyPlayerStats.hands),
    ).filter(DailyPlayerStats.upload_id.in_(upload_ids))

    if player_id:
        query = query.filter(DailyPlayerStats.player_id == player_id)

    query = query.group_by(DailyPlayerStats.player_id)
    results = query.all()

    players = []
    total_pnl = 0
    total_rake = 0
    total_hands = 0
    for pid, nick, club, pnl, rake, hands in results:
        p = round(float(pnl or 0), 2)
        r = round(float(rake or 0), 2)
        h = int(hands or 0)
        players.append({'player_id': pid, 'nickname': nick, 'club': club,
                        'pnl': p, 'rake': r, 'hands': h})
        total_pnl += p
        total_rake += r
        total_hands += h

    players.sort(key=lambda x: x['pnl'], reverse=True)

    return jsonify({
        'players': players,
        'totals': {'pnl': round(total_pnl, 2), 'rake': round(total_rake, 2), 'hands': total_hands},
        'days': len(upload_ids)
    })


@main_bp.route('/api/report-dates')
@login_required
def report_dates_api():
    """Return list of dates that have upload data."""
    from app.models import DailyUpload
    uploads = DailyUpload.query.with_entities(DailyUpload.upload_date).distinct().all()
    dates = [u[0].strftime('%Y-%m-%d') for u in uploads]
    return jsonify({'dates': dates})


@main_bp.route('/api/tournament-players')
@login_required
def tournament_players_api():
    """Return players who played in a specific tournament."""
    from app.models import PlayerSession
    title = request.args.get('title', '')
    if not title:
        return jsonify({'players': []})
    sessions = PlayerSession.query.filter_by(table_name=title, game_type='MTT').all()
    players = []
    for s in sessions:
        players.append({
            'player_id': s.player_id,
            'pnl': round(s.pnl, 2),
        })
    # Get nicknames
    from app.models import DailyPlayerStats
    nicks = dict(DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, func.max(DailyPlayerStats.nickname)
    ).group_by(DailyPlayerStats.player_id).all())
    for p in players:
        p['nickname'] = nicks.get(p['player_id'], p['player_id'])
    players.sort(key=lambda x: x['pnl'], reverse=True)
    return jsonify({'players': players})


@main_bp.route('/api/player-record/<player_id>')
@login_required
def player_record_api(player_id):
    from app.models import PlayerSession
    # Read ALL sessions from cumulative DB
    db_sessions = PlayerSession.query.filter_by(player_id=player_id).all()
    sessions = []
    for s in db_sessions:
        sessions.append({
            'table': s.table_name,
            'game': s.game_type,
            'blinds': s.blinds or '',
            'pnl': round(s.pnl, 2),
        })
    total_pnl = round(sum(s['pnl'] for s in sessions), 2)
    return jsonify({'sessions': sessions, 'total_pnl': total_pnl})


@main_bp.route('/transactions')
@login_required
def transactions():
    tx_type = request.args.get('type', '')
    category = request.args.get('category', '')

    query = Transaction.query.filter_by(user_id=current_user.id)

    if tx_type in ('income', 'expense'):
        query = query.filter_by(type=tx_type)
    if category:
        query = query.filter_by(category=category)

    all_transactions = query.order_by(Transaction.date.desc()).all()

    return render_template('main/transactions.html',
                           transactions=all_transactions,
                           income_categories=INCOME_CATEGORIES,
                           expense_categories=EXPENSE_CATEGORIES,
                           selected_type=tx_type,
                           selected_category=category)


@main_bp.route('/transactions/add', methods=['GET', 'POST'])
@login_required
def add_transaction():
    if request.method == 'POST':
        tx_type = request.form.get('type')
        amount_str = request.form.get('amount', '')
        category = request.form.get('category', '').strip()
        description = request.form.get('description', '').strip()
        date_str = request.form.get('date', '')

        error = None
        try:
            amount = float(amount_str)
            if amount <= 0:
                error = 'הסכום חייב להיות חיובי.'
        except ValueError:
            error = 'סכום לא תקין.'

        if not error:
            if tx_type not in ('income', 'expense'):
                error = 'סוג עסקה לא תקין.'
            elif not category:
                error = 'יש לבחור קטגוריה.'

        if not error:
            try:
                tx_date = date.fromisoformat(date_str)
            except ValueError:
                tx_date = date.today()

            transaction = Transaction(
                user_id=current_user.id,
                type=tx_type,
                amount=amount,
                category=category,
                description=description,
                date=tx_date
            )
            db.session.add(transaction)
            db.session.commit()
            flash('העסקה נוספה בהצלחה.', 'success')
            return redirect(url_for('main.transactions'))

        flash(error, 'danger')

    return render_template('main/add_transaction.html',
                           income_categories=INCOME_CATEGORIES,
                           expense_categories=EXPENSE_CATEGORIES,
                           today=date.today().isoformat())


@main_bp.route('/transactions/delete/<int:tx_id>', methods=['POST'])
@login_required
def delete_transaction(tx_id):
    transaction = Transaction.query.filter_by(id=tx_id, user_id=current_user.id).first_or_404()
    db.session.delete(transaction)
    db.session.commit()
    flash('העסקה נמחקה.', 'info')
    return redirect(url_for('main.transactions'))
