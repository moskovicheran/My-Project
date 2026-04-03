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
        for sa in sa_tables:
            sa_rake = 0
            sa_pnl = 0
            sa_hands = 0
            for m in sa.get('direct', []):
                c = cumul.get(m['player_id'])
                if c:
                    m['pnl'] = c['pnl']
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
                        m['pnl'] = c['pnl']
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
    cumulative = get_cumulative_stats()
    grand_pnl = 0
    grand_rake = 0
    for club in clubs:
        club_pnl = 0
        club_rake = 0
        for sa in club.get('super_agents', {}).values():
            for ag in sa.get('agents', {}).values():
                for m in ag.get('members', []):
                    c = cumulative.get(m['player_id'])
                    if c:
                        m['pnl_total'] = c['pnl']
                        m['rake_total'] = c['rake']
                    club_pnl += m['pnl_total']
                    club_rake += m['rake_total']
            for m in sa.get('direct_members', []):
                c = cumulative.get(m['player_id'])
                if c:
                    m['pnl_total'] = c['pnl']
                    m['rake_total'] = c['rake']
                club_pnl += m['pnl_total']
                club_rake += m['rake_total']
        for m in club.get('no_sa_members', []):
            c = cumulative.get(m['player_id'])
            if c:
                m['pnl_total'] = c['pnl']
                m['rake_total'] = c['rake']
            club_pnl += m['pnl_total']
            club_rake += m['rake_total']
        club['total_pnl'] = round(club_pnl, 2)
        club['total_rake'] = round(club_rake, 2)
        grand_pnl += club_pnl
        grand_rake += club_rake
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

    # Use cumulative data
    if cs:
        total_rake = cs['rake']
        total_pnl = cs['pnl']
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
    from app.models import PlayerSession
    db_sessions = PlayerSession.query.filter_by(player_id=player_id).all()
    if db_sessions:
        sessions = []
        for s in db_sessions:
            sessions.append({
                'table_name': s.table_name,
                'game_type': s.game_type,
                'blinds': s.blinds or '',
                'club_name': '',
                'buyin': 0, 'cashout': 0, 'hands': 0, 'rake': 0,
                'pnl': round(s.pnl, 2),
            })

    return render_template('union/player_detail.html',
                           member=member_info,
                           sessions=sessions,
                           club_entries=club_entries,
                           total_rake=total_rake,
                           total_pnl=total_pnl,
                           total_hands=total_hands)


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
