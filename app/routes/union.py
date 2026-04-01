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
    return render_template('union/ring_games.html', games=games, totals=totals)


@union_bp.route('/mtts')
@login_required
def mtts():
    mtt_list, totals = get_mtts()
    return render_template('union/mtts.html', mtts=mtt_list, totals=totals)


@union_bp.route('/members')
@login_required
def members():
    top_winners, top_losers = get_top_members(20)
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
    sa_tables = get_super_agent_tables()
    return render_template('union/agents.html', sa_tables=sa_tables)


@union_bp.route('/players')
@login_required
def players():
    clubs, grand = get_members_hierarchy()
    return render_template('union/players.html', clubs=clubs, grand=grand)


@union_bp.route('/player/<player_id>')
@login_required
def player_detail(player_id):
    from app.union_data import get_cumulative_stats
    member_info, sessions, club_entries = get_player_detail(player_id)
    if not member_info:
        return render_template('union/player_detail.html',
                               member=None, sessions=[], club_entries=[],
                               player_id=player_id)

    # Use cumulative data from DB instead of single file
    cumulative = get_cumulative_stats([player_id])
    if player_id in cumulative:
        cs = cumulative[player_id]
        total_rake = cs['rake']
        total_pnl = cs['pnl']
        total_hands = cs['hands']
        # Update club entries with cumulative data
        if club_entries:
            # Scale proportionally or just show cumulative total
            club_entries = [{'club': cs.get('club', club_entries[0].get('club', '')),
                            'pnl_total': cs['pnl'], 'rake_total': cs['rake'],
                            'hands_total': cs['hands'],
                            'sa_nick': club_entries[0].get('sa_nick', '-'),
                            'agent_nick': club_entries[0].get('agent_nick', '-')}]
    else:
        total_rake = sum(s['rake'] for s in sessions)
        total_pnl = sum(s['pnl'] for s in sessions)
        total_hands = sum(s['hands'] for s in sessions)

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
