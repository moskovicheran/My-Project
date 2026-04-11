from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required
from app.models import db, SAHierarchy
from app.union_data import (get_union_overview, get_ring_games, get_mtts,
                            get_top_members, get_members_hierarchy,
                            get_player_detail, get_super_agent_tables,
                            get_all_super_agents)

union_bp = Blueprint('union', __name__, url_prefix='/union')


@union_bp.route('/')
@login_required
def overview():
    meta, clubs, total = get_union_overview()
    return render_template('union/overview.html', meta=meta, clubs=clubs, total=total)


@union_bp.route('/ring-games')
@login_required
def ring_games():
    games, totals = get_ring_games()
    # Override with cumulative Ring Rake from Member Statistics (matches dashboard)
    from app.union_data import get_cumulative_totals
    ct = get_cumulative_totals()
    totals['rake'] = ct.get('ring_rake', totals['rake'])
    return render_template('union/ring_games.html', games=games, totals=totals)


@union_bp.route('/mtts')
@login_required
def mtts():
    from app.models import TournamentStats
    from sqlalchemy import func as sqlfunc

    # Try cumulative DB first
    db_mtts = TournamentStats.query.all()
    if db_mtts:
        # Each tournament as separate row
        mtt_list = [{'title': t.title, 'game_type': t.game_type,
                     'buyin': t.buyin, 'fee': t.fee, 'reentry': t.reentry,
                     'gtd': t.gtd, 'entries': t.entries, 'prize_pool': t.prize_pool,
                     'start': t.start, 'duration': t.duration} for t in db_mtts]
        total_entries = sum(m['entries'] for m in mtt_list)
        total_prize = round(sum(m['prize_pool'] for m in mtt_list), 2)
        total_rake = round(sum(m['fee'] * m['entries'] for m in mtt_list if m['fee'] > 0), 2)
        totals = {'entries': int(total_entries), 'prize_pool': total_prize, 'total_buyin': 0, 'total_rake': total_rake}
        return render_template('union/mtts.html', mtts=mtt_list, totals=totals)

    # Fallback to Excel
    mtt_list, totals = get_mtts()
    return render_template('union/mtts.html', mtts=mtt_list, totals=totals)


@union_bp.route('/members')
@login_required
def members():
    from app.union_data import get_cumulative_stats
    all_cumulative = get_cumulative_stats()
    all_players = []
    for pid, cs in all_cumulative.items():
        if cs.get('hands', 0) == 0 and cs.get('pnl', 0) == 0:
            continue
        all_players.append({
            'member_id': pid,
            'nickname': cs['nickname'],
            'club': cs['club'],
            'pnl_total': cs['pnl'],
            'rake_total': cs['rake'],
            'hands_total': cs['hands'],
        })
    top_winners = sorted(all_players, key=lambda x: x['pnl_total'], reverse=True)[:20]
    top_losers = sorted(all_players, key=lambda x: x['pnl_total'])[:20]
    return render_template('union/members.html',
                           top_winners=top_winners,
                           top_losers=top_losers)


@union_bp.route('/cash')
@login_required
def cash():
    from app.union_data import get_ring_game_detail
    tables = get_ring_game_detail()
    total_rake  = round(sum(p['rake'] for t in tables for p in t['players']), 2)
    total_pnl   = round(sum(p['pnl']  for t in tables for p in t['players']), 2)
    total_buyin = round(sum(p['buyin'] for t in tables for p in t['players']), 2)
    total_hands = sum(p['hands'] for t in tables for p in t['players'])
    return render_template('union/cash.html', tables=tables,
                           total_rake=total_rake, total_pnl=total_pnl,
                           total_buyin=total_buyin, total_hands=int(total_hands))


@union_bp.route('/agents')
@login_required
def agents():
    from app.union_data import get_cumulative_stats
    sa_tables = get_super_agent_tables()
    # Override Excel data with cumulative DB data
    all_pids = set()
    for sa in sa_tables:
        for m in sa.get('direct', []):
            all_pids.add(m['player_id'])
        for ag in sa.get('agents', {}).values():
            for m in ag.get('members', []):
                all_pids.add(m['player_id'])
    if all_pids:
        cumul = get_cumulative_stats(list(all_pids))
        from app.union_data import get_transfer_adjustments
        xfer_adj = get_transfer_adjustments(list(all_pids))
        for sa in sa_tables:
            sa_rake = 0
            sa_pnl = 0
            sa_hands = 0
            for m in sa.get('direct', []):
                c = cumul.get(m['player_id'])
                if c:
                    m['pnl'] = round(c['pnl'] + xfer_adj.get(m['player_id'], 0), 2)
                    m['rake'] = c['rake']
                    m['hands'] = c.get('hands', 0)
                sa_rake += m.get('rake', 0)
                sa_pnl += m.get('pnl', 0)
                sa_hands += m.get('hands', 0)
            for ag in sa.get('agents', {}).values():
                ag_rake = 0
                ag_pnl = 0
                ag_hands = 0
                for m in ag.get('members', []):
                    c = cumul.get(m['player_id'])
                    if c:
                        m['pnl'] = round(c['pnl'] + xfer_adj.get(m['player_id'], 0), 2)
                        m['rake'] = c['rake']
                        m['hands'] = c.get('hands', 0)
                    ag_rake += m.get('rake', 0)
                    ag_pnl += m.get('pnl', 0)
                    ag_hands += m.get('hands', 0)
                ag['total_rake'] = round(ag_rake, 2)
                ag['total_pnl'] = round(ag_pnl, 2)
                ag['total_hands'] = ag_hands
                sa_rake += ag_rake
                sa_pnl += ag_pnl
                sa_hands += ag_hands
            sa['total_rake'] = round(sa_rake, 2)
            sa['total_pnl'] = round(sa_pnl, 2)
            sa['total_hands'] = sa_hands
    return render_template('union/agents.html', sa_tables=sa_tables)


@union_bp.route('/players')
@login_required
def players():
    from app.union_data import get_cumulative_stats
    clubs, grand = get_members_hierarchy()

    # Override Excel pnl/rake with cumulative DB values
    from app.union_data import get_transfer_adjustments
    cumulative = get_cumulative_stats()
    # Collect all player IDs for transfer adjustments
    all_pids = list(cumulative.keys())
    xfer_adj = get_transfer_adjustments(all_pids)

    # Helper to update members with cumulative data
    def _update_members(members, excel_pids):
        total_pnl = 0
        total_rake = 0
        for m in members:
            excel_pids.add(m['player_id'])
            c = cumulative.get(m['player_id'])
            if c:
                m['pnl_total'] = round(c['pnl'] + xfer_adj.get(m['player_id'], 0), 2)
                m['rake_total'] = c['rake']
            total_pnl += m['pnl_total']
            total_rake += m['rake_total']
        return total_pnl, total_rake

    def _update_sa(sa, excel_pids):
        sa_pnl = 0
        sa_rake = 0
        for ag in sa.get('agents', {}).values():
            p, r = _update_members(ag.get('members', []), excel_pids)
            sa_pnl += p
            sa_rake += r
        p, r = _update_members(sa.get('direct_members', []), excel_pids)
        sa_pnl += p
        sa_rake += r
        # Child super agents (nested)
        for child_sa in sa.get('child_super_agents', {}).values():
            cp, cr = _update_sa(child_sa, excel_pids)
            sa_pnl += cp
            sa_rake += cr
        return sa_pnl, sa_rake

    # Track which player_ids are already in the hierarchy (from Excel)
    excel_pids = set()
    grand_pnl = 0
    grand_rake = 0
    for club in clubs:
        club_pnl = 0
        club_rake = 0
        for sa in club.get('super_agents', {}).values():
            sp, sr = _update_sa(sa, excel_pids)
            club_pnl += sp
            club_rake += sr
        p, r = _update_members(club.get('no_sa_members', []), excel_pids)
        club_pnl += p
        club_rake += r
        club['total_pnl'] = round(club_pnl, 2)
        club['total_rake'] = round(club_rake, 2)
        grand_pnl += club_pnl
        grand_rake += club_rake

    # Add players from cumulative DB that are NOT in the active Excel
    from app.models import DailyPlayerStats
    from sqlalchemy import func as sqlfunc
    db_only_players = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname),
        sqlfunc.max(DailyPlayerStats.club), sqlfunc.max(DailyPlayerStats.sa_id),
        sqlfunc.max(DailyPlayerStats.agent_id),
    ).filter(
        DailyPlayerStats.player_id.notin_(list(excel_pids)) if excel_pids else True,
        DailyPlayerStats.role != 'Name Entry'
    ).group_by(DailyPlayerStats.player_id).all()

    club_map = {c['name']: c for c in clubs}
    for pid, nick, club_name, sa_id, ag_id in db_only_players:
        c = cumulative.get(pid)
        if not c or (c['pnl'] == 0 and c['rake'] == 0 and c['hands'] == 0):
            continue
        pnl = round(c['pnl'] + xfer_adj.get(pid, 0), 2)
        member = {
            'player_id': pid, 'nickname': nick or pid,
            'role': 'Player', 'country': '-',
            'sa_id': sa_id or '-', 'sa_nick': '-',
            'agent_id': ag_id or '-', 'agent_nick': '-',
            'pnl_total': pnl, 'rake_total': c['rake'], 'hands_total': c['hands'],
        }
        # Place in correct club
        if club_name and club_name in club_map:
            club_map[club_name]['no_sa_members'].append(member)
            club_map[club_name]['total_pnl'] = round(club_map[club_name].get('total_pnl', 0) + pnl, 2)
            club_map[club_name]['total_rake'] = round(club_map[club_name].get('total_rake', 0) + c['rake'], 2)
        else:
            # Unknown club — create or use catch-all
            if club_name and club_name not in club_map:
                new_club = {'name': club_name, 'club_id': '', 'super_agents': {}, 'no_sa_members': [],
                            'total_pnl': 0, 'total_rake': 0}
                clubs.append(new_club)
                club_map[club_name] = new_club
            target = club_map.get(club_name, clubs[0] if clubs else None)
            if target:
                target['no_sa_members'].append(member)
                target['total_pnl'] = round(target.get('total_pnl', 0) + pnl, 2)
                target['total_rake'] = round(target.get('total_rake', 0) + c['rake'], 2)
        grand_pnl += pnl
        grand_rake += c['rake']

    grand = {'pnl': round(grand_pnl, 2), 'rake': round(grand_rake, 2)}

    return render_template('union/players.html', clubs=clubs, grand=grand)


@union_bp.route('/player/<player_id>')
@login_required
def player_detail(player_id):
    from app.union_data import get_cumulative_stats
    member_info, sessions, club_entries = get_player_detail(player_id)

    # Get cumulative data from DB
    cumulative = get_cumulative_stats([player_id])
    cs = cumulative.get(player_id)

    # If player not in Excel but exists in DB - build member_info from DB
    if not member_info and cs:
        member_info = {
            'player_id': player_id,
            'nickname': cs.get('nickname', player_id),
            'role': 'Player',
            'country': '-',
            'club': cs.get('club', '-'),
            'sa_nick': '-',
            'agent_nick': '-',
            'pnl_total': cs['pnl'],
            'rake_total': cs['rake'],
            'hands_total': cs['hands'],
        }

    if not member_info:
        return render_template('union/player_detail.html',
                               member=None, sessions=[], club_entries=[],
                               player_id=player_id)

    # Transfer adjustment for this player
    from app.union_data import get_transfer_adjustments
    xfer_adj = get_transfer_adjustments([player_id])

    # Use cumulative data
    if cs:
        total_rake = cs['rake']
        total_pnl = round(cs['pnl'] + xfer_adj.get(player_id, 0), 2)
        total_hands = cs['hands']
        club_entries = [{'club': cs.get('club', member_info.get('club', '-')),
                        'pnl_total': cs['pnl'], 'rake_total': cs['rake'],
                        'hands_total': cs['hands'],
                        'sa_nick': member_info.get('sa_nick', '-'),
                        'agent_nick': member_info.get('agent_nick', '-')}]
    else:
        total_rake = sum(s['rake'] for s in sessions)
        total_pnl = sum(s['pnl'] for s in sessions)
        total_hands = sum(s['hands'] for s in sessions)

    # Load sessions from DB (cumulative, all uploads)
    from app.models import PlayerSession, DailyUpload
    db_sessions = (PlayerSession.query
                   .join(DailyUpload, PlayerSession.upload_id == DailyUpload.id)
                   .add_columns(DailyUpload.upload_date)
                   .filter(PlayerSession.player_id == player_id)
                   .order_by(DailyUpload.upload_date.asc())
                   .all())
    if db_sessions:
        sessions = []
        for s, upload_date in db_sessions:
            date_fmt = upload_date.strftime('%d/%m/%Y') if upload_date else ''
            sessions.append({
                'table_name': s.table_name,
                'game_type': s.game_type,
                'blinds': s.blinds or '',
                'date': date_fmt,
                'club_name': '',
                'buyin': 0, 'cashout': 0, 'hands': 0, 'rake': 0,
                'pnl': round(s.pnl, 2),
            })

    # Build game type stats from sessions
    game_stats = {}
    for s in sessions:
        gt = s.get('game_type', 'Other') or 'Other'
        if gt not in game_stats:
            game_stats[gt] = {'count': 0, 'pnl': 0, 'wins': 0, 'losses': 0, 'blinds': {}}
        gs = game_stats[gt]
        gs['count'] += 1
        gs['pnl'] = round(gs['pnl'] + s['pnl'], 2)
        if s['pnl'] >= 0:
            gs['wins'] += 1
        else:
            gs['losses'] += 1
        b = s.get('blinds', '-') or '-'
        if b not in gs['blinds']:
            gs['blinds'][b] = {'count': 0, 'pnl': 0, 'wins': 0, 'losses': 0}
        gs['blinds'][b]['count'] += 1
        gs['blinds'][b]['pnl'] = round(gs['blinds'][b]['pnl'] + s['pnl'], 2)
        if s['pnl'] >= 0:
            gs['blinds'][b]['wins'] += 1
        else:
            gs['blinds'][b]['losses'] += 1

    # Calculate win rates
    total_wins = sum(g['wins'] for g in game_stats.values())
    total_losses = sum(g['losses'] for g in game_stats.values())
    total_sessions = sum(g['count'] for g in game_stats.values())

    return render_template('union/player_detail.html',
                           member=member_info,
                           sessions=sessions,
                           club_entries=club_entries,
                           total_rake=total_rake,
                           total_pnl=total_pnl,
                           total_hands=total_hands,
                           game_stats=game_stats,
                           total_sessions=total_sessions,
                           total_wins=total_wins,
                           total_losses=total_losses)


@union_bp.route('/sa-hierarchy', methods=['GET', 'POST'])
@login_required
def sa_hierarchy():
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add':
            parent_id = request.form.get('parent_sa_id')
            child_id = request.form.get('child_sa_id')
            if parent_id and child_id and parent_id != child_id:
                existing = SAHierarchy.query.filter_by(child_sa_id=child_id).first()
                if existing:
                    flash('Super Agent זה כבר משויך למישהו אחר', 'warning')
                else:
                    db.session.add(SAHierarchy(parent_sa_id=parent_id, child_sa_id=child_id))
                    db.session.commit()
                    flash('שיוך נוסף בהצלחה', 'success')
            else:
                flash('יש לבחור שני Super Agents שונים', 'warning')
        elif action == 'delete':
            link_id = request.form.get('link_id')
            link = SAHierarchy.query.get(link_id)
            if link:
                db.session.delete(link)
                db.session.commit()
                flash('שיוך נמחק', 'success')
        return redirect(url_for('union.sa_hierarchy'))

    all_sa = get_all_super_agents()
    sa_name_map = {sa['id']: sa for sa in all_sa}
    links = SAHierarchy.query.all()
    link_list = []
    for link in links:
        parent = sa_name_map.get(link.parent_sa_id, {})
        child = sa_name_map.get(link.child_sa_id, {})
        link_list.append({
            'id': link.id,
            'parent_id': link.parent_sa_id,
            'parent_nick': parent.get('nick', link.parent_sa_id),
            'parent_club': parent.get('club', ''),
            'child_id': link.child_sa_id,
            'child_nick': child.get('nick', link.child_sa_id),
            'child_club': child.get('club', ''),
        })
    return render_template('union/sa_hierarchy.html',
                           all_sa=all_sa, links=link_list)
