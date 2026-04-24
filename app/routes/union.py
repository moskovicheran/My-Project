from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required, current_user
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

    # Add SA's own game stats as a direct player (if they also play)
    from app.models import DailyPlayerStats
    from sqlalchemy import func as sqlfunc
    for sa in sa_tables:
        sa_id_val = sa.get('sa_id')
        if sa_id_val:
            existing_pids = set(m['player_id'] for m in sa.get('direct', []))
            for ag in sa.get('agents', {}).values():
                for m in ag.get('members', []):
                    existing_pids.add(m['player_id'])
            if sa_id_val not in existing_pids:
                sa_own = DailyPlayerStats.query.with_entities(
                    sqlfunc.max(DailyPlayerStats.nickname),
                    sqlfunc.sum(DailyPlayerStats.pnl),
                    sqlfunc.sum(DailyPlayerStats.rake),
                    sqlfunc.sum(DailyPlayerStats.hands),
                ).filter(
                    DailyPlayerStats.player_id == sa_id_val,
                    DailyPlayerStats.role != 'Name Entry'
                ).first()
                if sa_own and (float(sa_own[1] or 0) != 0 or float(sa_own[2] or 0) != 0):
                    sa['direct'].insert(0, {
                        'player_id': sa_id_val,
                        'nickname': sa_own[0] or sa_id_val,
                        'role': 'Player',
                        'pnl': round(float(sa_own[1] or 0), 2),
                        'rake': round(float(sa_own[2] or 0), 2),
                        'hands': int(sa_own[3] or 0),
                    })

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
    # Only include rows whose role is actually 'Player'. Super Agent / Agent /
    # Manager / Master rows appear in the Excel as aggregate manager entries
    # (0 hands, non-zero rake/pnl representing manager take) and must not be
    # rendered as regular players — it produces entries like "0 hands, -223 PNL"
    # that look like data bugs.
    db_only_players = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname),
        sqlfunc.max(DailyPlayerStats.club), sqlfunc.max(DailyPlayerStats.sa_id),
        sqlfunc.max(DailyPlayerStats.agent_id),
    ).filter(
        DailyPlayerStats.player_id.notin_(list(excel_pids)) if excel_pids else True,
        DailyPlayerStats.role == 'Player'
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

    # Also include SAs/Agents referenced in DB rows but missing from the
    # Excel hierarchy (e.g. Redflag 3106-0170 — appears only as sa_id of
    # other players, no own row in Union Member Statistics). Aggregate their
    # activity via their sa_id/agent_id hierarchy so the search can find
    # them. Rendered with role="Super Agent"/"Agent" so they're clearly
    # distinguishable from regular players.
    from sqlalchemy import or_ as _or
    ref_sa_ids = {r[0] for r in DailyPlayerStats.query.with_entities(
        sqlfunc.distinct(DailyPlayerStats.sa_id)
    ).filter(DailyPlayerStats.sa_id.isnot(None),
             DailyPlayerStats.sa_id != '',
             DailyPlayerStats.sa_id != '-').all() if r[0]}
    ref_ag_ids = {r[0] for r in DailyPlayerStats.query.with_entities(
        sqlfunc.distinct(DailyPlayerStats.agent_id)
    ).filter(DailyPlayerStats.agent_id.isnot(None),
             DailyPlayerStats.agent_id != '',
             DailyPlayerStats.agent_id != '-').all() if r[0]}
    referenced_ids = ref_sa_ids | ref_ag_ids
    # Skip those already on the page (Excel or player-role path)
    seen = set(excel_pids)
    for club in clubs:
        for sa in club.get('super_agents', {}).values():
            seen.add(sa.get('id'))
            for ag in sa.get('agents', {}).values():
                seen.add(ag.get('id'))
        for m in club.get('no_sa_members', []):
            seen.add(m.get('player_id'))
    missing = [pid for pid in referenced_ids if pid not in seen]

    for pid in missing:
        # Aggregate activity where this pid acts as SA or Agent
        agg = DailyPlayerStats.query.with_entities(
            sqlfunc.max(DailyPlayerStats.club),
            sqlfunc.sum(DailyPlayerStats.rake),
            sqlfunc.sum(DailyPlayerStats.pnl),
            sqlfunc.sum(DailyPlayerStats.hands),
        ).filter(
            _or(DailyPlayerStats.sa_id == pid, DailyPlayerStats.agent_id == pid),
            DailyPlayerStats.role != 'Name Entry',
        ).first()
        if not agg:
            continue
        club_name, rake_sum, pnl_sum, hands_sum = agg
        rake_sum = round(float(rake_sum or 0), 2)
        pnl_sum = round(float(pnl_sum or 0), 2)
        hands_sum = int(hands_sum or 0)
        if rake_sum == 0 and pnl_sum == 0 and hands_sum == 0:
            continue
        # Nickname from own row if exists (e.g. Name Entry with nickname)
        own = DailyPlayerStats.query.with_entities(
            sqlfunc.max(DailyPlayerStats.nickname)
        ).filter(DailyPlayerStats.player_id == pid).scalar()
        role = 'Super Agent' if pid in ref_sa_ids else 'Agent'
        member = {
            'player_id': pid, 'nickname': own or pid,
            'role': role, 'country': '-',
            'sa_id': '-', 'sa_nick': '-',
            'agent_id': '-', 'agent_nick': '-',
            'pnl_total': pnl_sum, 'rake_total': rake_sum, 'hands_total': hands_sum,
        }
        if club_name and club_name in club_map:
            club_map[club_name]['no_sa_members'].append(member)
        elif club_name:
            new_club = {'name': club_name, 'club_id': '', 'super_agents': {},
                        'no_sa_members': [member], 'total_pnl': 0, 'total_rake': 0}
            clubs.append(new_club)
            club_map[club_name] = new_club
        # Note: NOT added to club/grand totals — the activity is already
        # reflected in the player rows aggregated above. This is purely a
        # search/visibility entry.

    grand = {'pnl': round(grand_pnl, 2), 'rake': round(grand_rake, 2)}

    # Get available game types and per-player game types for filters
    from app.models import PlayerSession
    from sqlalchemy import func as sqlfunc
    game_types_q = PlayerSession.query.with_entities(
        sqlfunc.distinct(PlayerSession.game_type)
    ).filter(PlayerSession.game_type.isnot(None)).all()
    game_types = sorted([r[0] for r in game_types_q if r[0]])

    # Build player->game_types mapping
    player_games = {}
    pg_rows = PlayerSession.query.with_entities(
        PlayerSession.player_id, PlayerSession.game_type
    ).filter(PlayerSession.game_type.isnot(None)).group_by(
        PlayerSession.player_id, PlayerSession.game_type
    ).all()
    for pid, gt in pg_rows:
        if pid not in player_games:
            player_games[pid] = []
        player_games[pid].append(gt)

    return render_template('union/players.html', clubs=clubs, grand=grand,
                           game_types=game_types, player_games=player_games)


@union_bp.route('/player/<player_id>')
@login_required
def player_detail(player_id):
    # Authorization: players see only themselves; admin sees all
    if current_user.role == 'player':
        if current_user.player_id != player_id:
            flash('אין לך הרשאה לצפות בשחקן זה.', 'danger')
            return redirect(url_for('main.dashboard'))

    from app.union_data import get_cumulative_stats, get_agent_scope_predicate
    member_info, sessions, club_entries = get_player_detail(player_id)

    # Determine viewer scope: when an agent (or admin acting as agent via
    # ?view_as=) looks at a player who plays in multiple places, show only
    # the rows attributed to THAT agent's card. `?full=1` opts out.
    # Admin without view_as → global view (backwards-compatible).
    scope_sa_id = None
    if request.args.get('full') != '1':
        if current_user.role == 'admin':
            scope_sa_id = request.args.get('view_as')
        elif current_user.role == 'agent':
            scope_sa_id = current_user.player_id
    # Self-view never scoped: a player is their own player_id — the scope
    # would filter them out of their own card.
    if scope_sa_id == player_id:
        scope_sa_id = None
    scope_pred = get_agent_scope_predicate(scope_sa_id) if scope_sa_id else None
    # Optional ?club=<name> — when a player is clicked from a managed-club
    # section, filter to that club only (a player may have rows in multiple
    # clubs per upload, so the scope predicate alone is not enough).
    club_filter = request.args.get('club', '').strip()
    # If a club filter is set we always run the "scoped" branch below so the
    # totals/sessions slice picks up the club constraint.
    scope_applied = (scope_pred is not None) or bool(club_filter)

    # Scoped-view helpers: rebuild totals/clubs/upload_filter from
    # DailyPlayerStats restricted to the viewer's scope predicate / club.
    scope_upload_ids = None  # None = no filter, set = filter sessions
    if scope_applied:
        from app.models import DailyPlayerStats as _DPS
        from sqlalchemy import func as _sf
        _scope_filters = [
            _DPS.player_id == player_id,
            _DPS.role != 'Name Entry',
        ]
        if scope_pred is not None:
            _scope_filters.append(scope_pred)
        if club_filter:
            _scope_filters.append(_DPS.club == club_filter)
        scoped_rows = _DPS.query.with_entities(
            _DPS.upload_id, _DPS.club,
            _sf.sum(_DPS.rake), _sf.sum(_DPS.pnl), _sf.sum(_DPS.hands),
        ).filter(*_scope_filters).group_by(_DPS.upload_id, _DPS.club).all()
        scope_upload_ids = {uid for uid, _, _, _, _ in scoped_rows}

    # Get cumulative data from DB
    cumulative = get_cumulative_stats([player_id])
    cs = cumulative.get(player_id)

    # If player not in Excel but exists in DB - build member_info from DB.
    # Use the actual role from DailyPlayerStats rather than hardcoding 'Player':
    # the hierarchy includes Super Agents / Agents / Managers whose stat rows
    # are administrative (0 hands + non-zero rake/pnl), not real play data.
    if not member_info and cs:
        from app.models import DailyPlayerStats as _DPS
        from sqlalchemy import func as _sqlfunc
        actual_role = _DPS.query.with_entities(_sqlfunc.max(_DPS.role)).filter(
            _DPS.player_id == player_id,
            _DPS.role != 'Name Entry'
        ).scalar() or 'Player'
        member_info = {
            'player_id': player_id,
            'nickname': cs.get('nickname', player_id),
            'role': actual_role,
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

    # Use cumulative data — if scope is applied, recompute from the scoped
    # DailyPlayerStats slice so totals/clubs reflect only this viewer's
    # card (e.g. areyoufold under niroha shows SPC T only, not POKER GARDEN).
    if scope_applied:
        total_rake = 0.0
        total_pnl = 0.0
        total_hands = 0
        _clubs_map = {}
        for uid, club, r, p, h in scoped_rows:
            total_rake += float(r or 0)
            total_pnl += float(p or 0)
            total_hands += int(h or 0)
            e = _clubs_map.setdefault(club, {
                'club': club, 'pnl_total': 0.0, 'rake_total': 0.0,
                'hands_total': 0,
                'sa_nick': member_info.get('sa_nick', '-'),
                'agent_nick': member_info.get('agent_nick', '-'),
            })
            e['pnl_total'] += float(p or 0)
            e['rake_total'] += float(r or 0)
            e['hands_total'] += int(h or 0)
        total_rake = round(total_rake, 2)
        total_pnl = round(total_pnl, 2)
        for e in _clubs_map.values():
            e['pnl_total'] = round(e['pnl_total'], 2)
            e['rake_total'] = round(e['rake_total'], 2)
        club_entries = list(_clubs_map.values())
        # Skip transfer adjustment in scoped view — transfers aren't tied
        # to a specific card, so including them would misattribute.
    elif cs:
        total_rake = cs['rake']
        total_pnl = round(cs['pnl'] + xfer_adj.get(player_id, 0), 2)
        total_hands = cs['hands']
        # Build per-club club_entries from DailyPlayerStats so the page shows
        # the "פירוט לפי קלאב" breakdown (e.g. SPC T + SPC Un) instead of a
        # single combined card. Totals stay from cs so they match the agent
        # dashboard card this page was reached from.
        from app.models import DailyPlayerStats as _DPS_ce
        from sqlalchemy import func as _sf_ce
        _club_rows = _DPS_ce.query.with_entities(
            _DPS_ce.club,
            _sf_ce.sum(_DPS_ce.rake),
            _sf_ce.sum(_DPS_ce.pnl),
            _sf_ce.sum(_DPS_ce.hands),
        ).filter(
            _DPS_ce.player_id == player_id,
            _DPS_ce.role != 'Name Entry',
            _DPS_ce.club != '',
        ).group_by(_DPS_ce.club).all()
        if _club_rows:
            club_entries = [{
                'club': c,
                'pnl_total': round(float(p or 0), 2),
                'rake_total': round(float(r or 0), 2),
                'hands_total': int(h or 0),
                'sa_nick': member_info.get('sa_nick', '-'),
                'agent_nick': member_info.get('agent_nick', '-'),
            } for c, r, p, h in _club_rows]
            club_entries.sort(key=lambda e: e['rake_total'], reverse=True)
        else:
            club_entries = [{'club': cs.get('club', member_info.get('club', '-')),
                            'pnl_total': cs['pnl'], 'rake_total': cs['rake'],
                            'hands_total': cs['hands'],
                            'sa_nick': member_info.get('sa_nick', '-'),
                            'agent_nick': member_info.get('agent_nick', '-')}]
    else:
        total_rake = sum(s['rake'] for s in sessions)
        total_pnl = sum(s['pnl'] for s in sessions)
        total_hands = sum(s['hands'] for s in sessions)

    # Load sessions from DB (cumulative, all uploads).
    # In scoped view, restrict to uploads where the player had in-scope rows
    # (a player has at most one club per upload, so this is a clean filter).
    # Master/SA exception: when Member Statistics shows pnl=0 across all
    # scoped rows (Member Stats doesn't aggregate their play), the upload-id
    # proxy breaks — their real play sessions may live in uploads where
    # they have no Member Stats row for this club. Fall back to showing all
    # of the player's sessions so the card reflects reality.
    from app.models import PlayerSession, DailyUpload
    # Master/SA fallback: in club-scoped view, trigger the fallback when
    # EITHER (a) scoped Member Stats rows exist but sum to pnl=0, OR
    # (b) no scoped Member Stats rows at all — both indicate the upload-id
    # proxy can't represent the player's real play. Show all their
    # sessions instead of forcing an empty list.
    _master_fallback = False
    if scope_applied:
        if scope_upload_ids:
            _master_fallback = all(float(p or 0) == 0 for _, _, _, p, _ in scoped_rows)
        else:
            _master_fallback = True
    _sess_q = (PlayerSession.query
               .join(DailyUpload, PlayerSession.upload_id == DailyUpload.id)
               .add_columns(DailyUpload.upload_date)
               .filter(PlayerSession.player_id == player_id))
    if scope_applied and not _master_fallback:
        _sess_q = _sess_q.filter(PlayerSession.upload_id.in_(list(scope_upload_ids)))
    db_sessions = _sess_q.order_by(DailyUpload.upload_date.asc()).all()
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

    # Masters / SAs whose Member Statistics row carries pnl=0 (their play
    # isn't aggregated there, only their management data) — derive total_pnl
    # from actual game sessions so top cards and the record-table total
    # reflect reality instead of 0. Applies in both full and scoped views;
    # transfers are already suppressed in scoped view so xfer_adj is skipped.
    if total_pnl == 0 and db_sessions:
        sess_pnl = sum(float(s.pnl or 0) for s, _ in db_sessions)
        if sess_pnl != 0:
            if scope_applied:
                total_pnl = round(sess_pnl, 2)
            else:
                total_pnl = round(sess_pnl + xfer_adj.get(player_id, 0), 2)

    # Add money transfers as special session entries — only in full view
    # (transfers aren't tied to a specific card, so in scoped view we
    # suppress them to avoid misattribution).
    from app.models import MoneyTransfer
    from datetime import timedelta
    transfers_out = [] if scope_applied else MoneyTransfer.query.filter_by(from_player_id=player_id).all()
    transfers_in = [] if scope_applied else MoneyTransfer.query.filter_by(to_player_id=player_id).all()
    for t in transfers_out:
        il_time = t.created_at + timedelta(hours=3) if t.created_at else None
        sessions.append({
            'table_name': f'העברה ל-{t.to_name}',
            'game_type': 'העברה',
            'blinds': t.description or '',
            'date': il_time.strftime('%d/%m/%Y') if il_time else '',
            'date_sort': il_time if il_time else t.created_at,
            'pnl': round(-t.amount, 2),
            'is_transfer': True,
        })
    for t in transfers_in:
        il_time = t.created_at + timedelta(hours=3) if t.created_at else None
        sessions.append({
            'table_name': f'קיבלת מ-{t.from_name}',
            'game_type': 'העברה',
            'blinds': t.description or '',
            'date': il_time.strftime('%d/%m/%Y') if il_time else '',
            'date_sort': il_time if il_time else t.created_at,
            'pnl': round(t.amount, 2),
            'is_transfer': True,
        })

    # Sort sessions by date (newest first)
    from datetime import datetime
    def _sort_key(s):
        if s.get('date_sort'):
            return s['date_sort']
        try:
            return datetime.strptime(s.get('date', ''), '%d/%m/%Y')
        except (ValueError, TypeError):
            return datetime.min
    sessions.sort(key=_sort_key, reverse=True)

    # Build game type stats from sessions (exclude transfers)
    game_stats = {}
    for s in sessions:
        if s.get('is_transfer'):
            continue
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

    # Hierarchy breadcrumb — chain of (Agent → SA → ... → Manager) above
    # this player. Renders as "מנוהל תחת: X → Y → Z" on the detail page.
    hierarchy_chain = _build_hierarchy_chain(player_id)

    # Resolve the scope SA's display name for the badge in the template
    scope_sa_nick = None
    if scope_applied:
        from app.models import DailyPlayerStats as _DPS
        from sqlalchemy import func as _sf
        scope_sa_nick = _DPS.query.with_entities(_sf.max(_DPS.nickname)).filter(
            _DPS.player_id == scope_sa_id,
            _DPS.nickname.isnot(None), _DPS.nickname != ''
        ).scalar() or scope_sa_id

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
                           total_losses=total_losses,
                           hierarchy_chain=hierarchy_chain,
                           scope_applied=scope_applied,
                           scope_sa_id=scope_sa_id,
                           scope_sa_nick=scope_sa_nick,
                           club_filter=club_filter)


def _build_hierarchy_chain(player_id):
    """Return a list of {'role','id','label','is_manager'} describing the
    management chain above the player — Agent → SA → (parent SAs) → Manager.

    - If the viewed entity is themselves an SA (appears as sa_id of others or
      as child in SAHierarchy), we walk upward from their own player_id.
    - If they're an Agent (appears as agent_id of others), we walk from the
      SA they typically operate under.
    - Otherwise (regular Player), we use their raw sa_id/agent_id (honoring
      any PlayerAssignment override) as the starting node.
    """
    from app.models import DailyPlayerStats, SAHierarchy, PlayerAssignment
    from app.routes.admin import OVERVIEW_MANAGERS
    from sqlalchemy import func as sqlfunc

    def _nick(pid):
        if not pid:
            return ''
        n = DailyPlayerStats.query.with_entities(
            sqlfunc.max(DailyPlayerStats.nickname)
        ).filter(DailyPlayerStats.player_id == pid,
                 DailyPlayerStats.nickname.isnot(None),
                 DailyPlayerStats.nickname != '').scalar()
        return n or pid

    managers_map = {p: u for u, p in OVERVIEW_MANAGERS}

    # Is the viewed entity themselves an SA / Agent?
    is_sa = (SAHierarchy.query.filter_by(child_sa_id=player_id).first() is not None
             or DailyPlayerStats.query.filter(
                 DailyPlayerStats.sa_id == player_id).first() is not None)
    is_agent_role = DailyPlayerStats.query.filter(
        DailyPlayerStats.agent_id == player_id,
        DailyPlayerStats.role != 'Name Entry').first() is not None

    chain = []
    seen = {player_id}

    # Determine starting node on the SA walk
    player_club = DailyPlayerStats.query.with_entities(
        sqlfunc.max(DailyPlayerStats.club)
    ).filter(DailyPlayerStats.player_id == player_id,
             DailyPlayerStats.role != 'Name Entry').scalar() or ''

    if is_sa:
        # Player themselves IS an SA — walk UPWARD from their parent in
        # SAHierarchy (self isn't shown; the chain displays who manages them).
        link = SAHierarchy.query.filter_by(child_sa_id=player_id).first()
        cur = link.parent_sa_id if link else ''
    elif is_agent_role:
        # Player is an Agent — walk from the SA they typically operate under.
        sa_of_agent = DailyPlayerStats.query.with_entities(
            sqlfunc.max(DailyPlayerStats.sa_id)
        ).filter(DailyPlayerStats.agent_id == player_id,
                 DailyPlayerStats.sa_id.isnot(None),
                 DailyPlayerStats.sa_id != '',
                 DailyPlayerStats.sa_id != '-').scalar()
        cur = sa_of_agent or ''
    else:
        # Regular player — start from raw sa_id/agent_id (honor override)
        pa = PlayerAssignment.query.filter_by(player_id=player_id).first()
        own_sa, own_ag = DailyPlayerStats.query.with_entities(
            sqlfunc.max(DailyPlayerStats.sa_id),
            sqlfunc.max(DailyPlayerStats.agent_id),
        ).filter(DailyPlayerStats.player_id == player_id,
                 DailyPlayerStats.role != 'Name Entry').first() or ('', '')
        sa_id = (pa.assigned_sa_id if pa and pa.assigned_sa_id else (own_sa or ''))
        ag_id = (pa.assigned_agent_id if pa and pa.assigned_agent_id else (own_ag or ''))
        # Agent node first (if distinct from SA)
        if ag_id and ag_id not in ('-', '', sa_id):
            chain.append({'role': 'Agent', 'id': ag_id,
                          'label': _nick(ag_id), 'is_manager': False})
            seen.add(ag_id)
        cur = sa_id if sa_id and sa_id not in ('-', '') else ''

    # Walk upward through SA hierarchy until hitting a Manager or dead end
    safety = 10  # cycle guard
    while cur and cur not in seen and safety > 0:
        safety -= 1
        seen.add(cur)
        if cur in managers_map:
            chain.append({'role': 'Manager', 'id': cur,
                          'label': managers_map[cur], 'is_manager': True})
            return chain
        chain.append({'role': 'Super Agent', 'id': cur,
                      'label': _nick(cur), 'is_manager': False})
        link = SAHierarchy.query.filter_by(child_sa_id=cur).first()
        cur = link.parent_sa_id if link else ''

    # Fallback — no Manager found via SA hierarchy. Try matching the player's
    # club against SARakeConfig.managed_club_id (e.g. Spc o → Riko via
    # literal club-name mapping). Covers players attributed via managed
    # clubs rather than SA/agent hierarchy.
    if not any(n['is_manager'] for n in chain) and player_club:
        from app.models import SARakeConfig
        from app.union_data import get_members_hierarchy as _gmh
        _cd, _ = _gmh()
        _cid_to_name = {c['club_id']: c['name'] for c in _cd}
        for cfg in SARakeConfig.query.filter(
                SARakeConfig.managed_club_id.isnot(None)).all():
            resolved = _cid_to_name.get(cfg.managed_club_id) or cfg.managed_club_id
            if resolved == player_club and cfg.sa_id in managers_map:
                chain.append({'role': 'Manager', 'id': cfg.sa_id,
                              'label': managers_map[cfg.sa_id],
                              'is_manager': True})
                break
    return chain


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
