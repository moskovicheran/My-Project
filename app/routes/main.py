import io
from datetime import date
from flask import Blueprint, render_template, redirect, url_for, flash, request, jsonify, send_file
from flask_login import login_required, current_user
from sqlalchemy import func
from app.models import db, Transaction

main_bp = Blueprint('main', __name__)

INCOME_CATEGORIES = ['משכורת', 'פרילנס', 'השקעות', 'מתנה', 'אחר']

# Agents whose dashboard should hide the "רייק אישי" total and the
# percentage badge next to "הרייק שלי" — they only see their own earning.
AGENTS_HIDE_PERSONAL_BREAKDOWN = {'9319-6677'}  # Shlomi (sarbuvx)


def _resolve_date_uploads(selected_dates):
    """Resolve selected date strings to upload IDs, checking both active and archived data.

    Returns (active_upload_ids, archive_period_id, archive_upload_ids, valid_dates).

    IMPORTANT: If the caller passed at least one date string but NONE of them
    resolved to an upload, active_upload_ids is returned as [-1] (a sentinel
    non-existent upload id). This way, callers that use
    `if upload_ids_filter: filter.append(upload_id.in_(ids))` will still
    apply the filter and get zero rows, instead of silently falling back to
    all-time data. Passing an empty input list still returns empty (no filter)."""
    from app.models import DailyUpload, ArchivedUpload
    from datetime import datetime as dt
    active_upload_ids = []
    archive_period_id = None
    archive_upload_ids = []
    valid_dates = []
    for ds in selected_dates:
        try:
            sel = dt.strptime(ds, '%Y-%m-%d').date()
            # Check active first
            upload = DailyUpload.query.filter_by(upload_date=sel).first()
            if upload:
                active_upload_ids.append(upload.id)
                valid_dates.append(ds)
            else:
                # Check archive
                archived = ArchivedUpload.query.filter(ArchivedUpload.upload_date == sel).first()
                if archived:
                    archive_period_id = archived.period_id
                    archive_upload_ids.append(archived.original_id)
                    valid_dates.append(ds)
        except ValueError:
            pass
    # Sentinel for "user asked for a filter that matched nothing"
    if selected_dates and not active_upload_ids and not archive_upload_ids:
        active_upload_ids = [-1]
    return active_upload_ids, archive_period_id, archive_upload_ids, valid_dates


def _format_period_label(selected_dates):
    """Human-readable date label for Excel banners.
    Single date → DD/MM/YYYY. Range → DD/MM/YYYY — DD/MM/YYYY. Multiple non-contiguous → list."""
    if not selected_dates:
        return None
    from datetime import datetime as dt
    try:
        parsed = sorted({dt.strptime(d, '%Y-%m-%d').date() for d in selected_dates})
    except ValueError:
        return ', '.join(selected_dates)
    if len(parsed) == 1:
        return parsed[0].strftime('%d/%m/%Y')
    # If the dates are a contiguous run, render as a range
    from datetime import timedelta
    is_contiguous = all((parsed[i] - parsed[i - 1]) == timedelta(days=1) for i in range(1, len(parsed)))
    if is_contiguous:
        return f"{parsed[0].strftime('%d/%m/%Y')} — {parsed[-1].strftime('%d/%m/%Y')}"
    return ', '.join(d.strftime('%d/%m/%Y') for d in parsed)


EXPENSE_CATEGORIES = ['מזון', 'דיור', 'תחבורה', 'בריאות', 'בידור', 'קניות', 'חינוך', 'חשבונות', 'אחר']


@main_bp.route('/')
def index():
    if current_user.is_authenticated:
        return redirect(url_for('main.dashboard'))
    return redirect(url_for('auth.login'))


@main_bp.route('/dashboard')
@login_required
def dashboard():
    if (hasattr(current_user, 'role') and current_user.role == 'admin'
            and not request.args.get('view_as')
            and not request.args.get('view_player')):
        # Admin home dashboard mirrors the /admin/ overview (managers,
        # tracked clubs, date picker, totals) via the shared context builder.
        from app.routes.admin import build_overview_context
        return render_template('main/admin_dashboard.html', **build_overview_context())

    # Admin may view any club's dashboard via ?view_as=<club_id>.
    # Ambiguity: some club IDs contain a '-' (e.g. 5481-5364), which used to
    # collide with the agent view_as format. Resolve by checking User table:
    # if view_as matches a known user's player_id → agent view, otherwise → club.
    _admin_view_as_club = None
    if (hasattr(current_user, 'role') and current_user.role == 'admin'
            and request.args.get('view_as')):
        _va = request.args.get('view_as')
        if _va:
            from app.models import User as _User
            _matches_user = _User.query.filter_by(player_id=_va).first() is not None
            if not _matches_user:
                _admin_view_as_club = _va

    if (hasattr(current_user, 'role') and current_user.role == 'club' and current_user.player_id) \
            or _admin_view_as_club:
        from app.models import DailyPlayerStats, DailyUpload
        from app.union_data import get_members_hierarchy
        from sqlalchemy import func as sqlfunc
        from datetime import datetime as dt

        club_id = _admin_view_as_club if _admin_view_as_club else current_user.player_id
        # Find club name
        clubs_data, _ = get_members_hierarchy()
        club_name = None
        club_obj = None
        for c in clubs_data:
            if c['club_id'] == club_id:
                club_name = c['name']
                club_obj = c
                break

        # Available upload dates (active + archived)
        from app.models import ArchivedUpload
        active_dates = {u[0].strftime('%Y-%m-%d') for u in
                        DailyUpload.query.with_entities(DailyUpload.upload_date).distinct().all()}
        archive_dates = {u[0].strftime('%Y-%m-%d') for u in
                         ArchivedUpload.query.with_entities(ArchivedUpload.upload_date).distinct().all()}
        available_dates = sorted(active_dates | archive_dates, reverse=True)

        # Date filter — supports multiple dates: ?dates=2026-03-30,2026-03-31
        requested_dates = [d.strip() for d in request.args.get('dates', '').split(',') if d.strip()]
        had_date_filter = bool(requested_dates)
        selected_dates = requested_dates
        upload_ids_filter = []
        use_archive = False
        archive_period_id = None
        archive_upload_ids = []
        if selected_dates:
            upload_ids_filter, archive_period_id, archive_upload_ids, selected_dates = _resolve_date_uploads(selected_dates)
            use_archive = bool(archive_upload_ids)

        if club_name:
            if use_archive and archive_period_id:
                # Query from archived data
                from app.models import ArchivedPlayerStats
                base_filters = [ArchivedPlayerStats.club == club_name,
                                ArchivedPlayerStats.role != 'Name Entry',
                                ArchivedPlayerStats.period_id == archive_period_id,
                                ArchivedPlayerStats.upload_id.in_(archive_upload_ids)]
                StatsModel = ArchivedPlayerStats
            else:
                # Base query (active data)
                base_filters = [DailyPlayerStats.club == club_name,
                                DailyPlayerStats.role != 'Name Entry']
                if upload_ids_filter:
                    base_filters.append(DailyPlayerStats.upload_id.in_(upload_ids_filter))
                elif had_date_filter:
                    # Dates requested but none resolved to uploads → return empty, don't silently show all-time
                    base_filters.append(DailyPlayerStats.upload_id == -1)
                StatsModel = DailyPlayerStats

            club_players_db = StatsModel.query.with_entities(
                StatsModel.player_id, sqlfunc.max(StatsModel.nickname),
                sqlfunc.max(StatsModel.sa_id), sqlfunc.max(StatsModel.agent_id),
                sqlfunc.sum(StatsModel.pnl), sqlfunc.sum(StatsModel.rake),
                sqlfunc.sum(StatsModel.hands),
            ).filter(*base_filters).group_by(StatsModel.player_id).all()

            # Nickname map
            all_nicks = dict(StatsModel.query.with_entities(
                StatsModel.player_id, sqlfunc.max(StatsModel.nickname)
            ).group_by(StatsModel.player_id).all())

            # Transfer adjustments
            from app.union_data import get_transfer_adjustments, get_player_overrides
            xfer_adj = get_transfer_adjustments([p[0] for p in club_players_db])

            # Manual overrides (/admin/lost-players) — replace natural sa_id/agent_id
            # so club view matches the agent's personal dashboard.
            overrides_map = get_player_overrides()

            # Build SA structure
            club_sas = {}
            agents_no_sa = {}
            no_sa = []
            total_rake = 0
            total_pnl = 0
            total_hands = 0
            for pid, nick, sa_id_val, ag_id_val, pnl_val, rake_val, hands_val in club_players_db:
                _ov = overrides_map.get(pid)
                if _ov:
                    if _ov.get('sa_id'):
                        sa_id_val = _ov['sa_id']
                    if _ov.get('agent_id'):
                        ag_id_val = _ov['agent_id']
                p = round(float(pnl_val or 0) + xfer_adj.get(pid, 0), 2)
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
                    # No SA — check if there's an agent
                    if ag_id_val and ag_id_val != '-':
                        if ag_id_val not in agents_no_sa:
                            ag_nick = all_nicks.get(ag_id_val, ag_id_val)
                            agents_no_sa[ag_id_val] = {'nick': ag_nick, 'members': []}
                        agents_no_sa[ag_id_val]['members'].append(member)
                    else:
                        no_sa.append(member)

            managed_club = {
                'name': club_name, 'club_id': club_id,
                'total_rake': round(total_rake, 2), 'total_pnl': round(total_pnl, 2),
                'super_agents': club_sas, 'agents_no_sa': agents_no_sa,
                'no_sa_members': no_sa,
            }
            player_count = len(club_players_db)

            # Net rake calculation (club's percentage)
            from app.models import RakeConfig
            club_rc = RakeConfig.query.filter_by(entity_type='club', entity_id=club_id).first()
            rake_pct = club_rc.rake_percent if club_rc else 100
            net_rake = round(total_rake * rake_pct / 100, 2)

            # Sort all player lists by PnL for the club dashboard:
            # positives first (largest win), negatives next (biggest loss first), zeros last.
            def _sort_by_pnl_club(lst, attr='pnl_total'):
                if not lst:
                    return
                def key(m):
                    v = m.get(attr, 0) or 0
                    if v > 0: return (0, -v)
                    if v < 0: return (1, v)
                    return (2, 0)
                lst.sort(key=key)
            for _sa in managed_club.get('super_agents', {}).values():
                _sort_by_pnl_club(_sa.get('direct_members'))
                for _ag in _sa.get('agents', {}).values():
                    _sort_by_pnl_club(_ag.get('members'))
            _sort_by_pnl_club(managed_club.get('no_sa_members'))

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

    # Admin viewing agent dashboard via ?view_as or agent's own dashboard
    view_as_id = request.args.get('view_as') if current_user.role == 'admin' else None
    if (current_user.role == 'agent' and current_user.player_id) or view_as_id:
        from app.union_data import get_super_agent_tables, get_members_hierarchy, get_child_sa_entries
        from app.models import SAHierarchy, SARakeConfig, RakeConfig, ExpenseCharge, DailyPlayerStats, DailyUpload, MoneyTransfer, User
        from sqlalchemy import func as sqlfunc
        from datetime import datetime as dt

        if view_as_id:
            agent_user = User.query.filter_by(player_id=view_as_id).first()
            sa_id = view_as_id
            view_as_username = agent_user.username if agent_user else view_as_id
        else:
            sa_id = current_user.player_id
            view_as_username = None

        # Available upload dates (active + archived)
        from app.models import ArchivedUpload
        active_dates = {u[0].strftime('%Y-%m-%d') for u in
                        DailyUpload.query.with_entities(DailyUpload.upload_date).distinct().all()}
        archive_dates = {u[0].strftime('%Y-%m-%d') for u in
                         ArchivedUpload.query.with_entities(ArchivedUpload.upload_date).distinct().all()}
        available_dates = sorted(active_dates | archive_dates, reverse=True)

        # Date filter
        requested_dates = [d.strip() for d in request.args.get('dates', '').split(',') if d.strip()]
        had_date_filter = bool(requested_dates)
        selected_dates = requested_dates
        upload_ids_filter = []
        use_archive = False
        archive_period_id = None
        archive_upload_ids = []
        if selected_dates:
            upload_ids_filter, archive_period_id, archive_upload_ids, selected_dates = _resolve_date_uploads(selected_dates)
            use_archive = bool(archive_upload_ids)

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

        # Get SA structure from Excel (for hierarchy display of THIS agent)
        sa_tables = get_super_agent_tables()
        my_sas = []
        for kid in known_ids:
            my_sas.extend([sa for sa in sa_tables if sa['sa_id'] == kid])

        # Managed club names — used both to skip overlapping child SAs and
        # to exclude managed-club rows from the hier-tree aggregations
        # below (avoids double-counting overlap players).
        managed_club_names = set()
        rake_cfgs_early = SARakeConfig.query.filter_by(sa_id=sa_id).filter(SARakeConfig.managed_club_id.isnot(None)).all()
        if rake_cfgs_early:
            clubs_data_early, _ = get_members_hierarchy()
            for cfg in rake_cfgs_early:
                for c in clubs_data_early:
                    if c['club_id'] == cfg.managed_club_id:
                        managed_club_names.add(c['name'])
        managed_club_names_list = list(managed_club_names)

        # Child SAs — DB-first (SAHierarchy), Excel-enriched. The helper
        # handles dedup + DB-only backfill so this dashboard can't silently
        # drop an SA assigned only via the admin control panel.
        child_sas = get_child_sa_entries(list(known_ids), managed_club_names)
        child_sa_ids = [cs['sa_id'] for cs in child_sas]

        all_sa_ids = list(known_ids) + child_sa_ids

        # Determine which stats model to use (active or archived)
        if use_archive and archive_period_id:
            from app.models import ArchivedPlayerStats
            SM = ArchivedPlayerStats
        else:
            SM = DailyPlayerStats

        # Manual overrides: players the admin has attached to one of our SAs/agents
        # via /admin/lost-players (or the overrides section of /admin/agents).
        # Also collect agent_ids that sit under our SAs — regular agents (not
        # child SAs) aren't in all_sa_ids, so assignments to them would be
        # missed without this extra set.
        from app.models import PlayerAssignment
        _my_agent_ids_rows = SM.query.with_entities(SM.agent_id).filter(
            SM.sa_id.in_(all_sa_ids),
            SM.agent_id.isnot(None),
            SM.agent_id != '',
            SM.agent_id != '-',
        ).distinct().all()
        my_known_agent_ids = {r[0] for r in _my_agent_ids_rows if r[0]}
        _assign_targets = list(set(all_sa_ids) | my_known_agent_ids)
        _override_rows = PlayerAssignment.query.filter(
            or_(
                PlayerAssignment.assigned_sa_id.in_(_assign_targets),
                PlayerAssignment.assigned_agent_id.in_(_assign_targets),
            )
        ).all()
        override_player_ids = {r.player_id for r in _override_rows}
        # Also build a global overrides map (pid → {sa_id, agent_id}) for all
        # players — used below to replace natural sa_id/agent_id on display.
        from app.union_data import get_player_overrides
        overrides_map = get_player_overrides()

        # Get ALL players that ever belonged to this SA/Agent
        # Step 1: Find all player_ids that belong to this SA
        _pid_filters = [or_(SM.sa_id.in_(all_sa_ids), SM.agent_id.in_(all_sa_ids)), SM.role != 'Name Entry']
        if use_archive and archive_period_id:
            _pid_filters += [SM.period_id == archive_period_id, SM.upload_id.in_(archive_upload_ids)]
        elif upload_ids_filter:
            _pid_filters.append(SM.upload_id.in_(upload_ids_filter))
        my_player_ids_query = SM.query.with_entities(SM.player_id).filter(*_pid_filters).distinct()
        my_player_id_list = [r[0] for r in my_player_ids_query.all()]
        # Union with override-assigned players (they may not belong naturally)
        if override_player_ids:
            my_player_id_list = list(set(my_player_id_list) | override_player_ids)

        # Step 2: Get cumulative stats — hier channel ONLY.
        # Row-level scope:
        #  - sa_id OR agent_id in hierarchy (keeps only this agent's channel)
        #  - club NOT IN managed_clubs (clubs bucket handles those)
        # Override players (manually attached via PlayerAssignment) are
        # exempt — include their rows even if they don't match the scope,
        # because that's the whole point of an override.
        _hier_row_pred = or_(SM.sa_id.in_(all_sa_ids), SM.agent_id.in_(all_sa_ids))
        if override_player_ids:
            _row_scope = or_(_hier_row_pred, SM.player_id.in_(list(override_player_ids)))
        else:
            _row_scope = _hier_row_pred
        base_agent_filters = [SM.player_id.in_(my_player_id_list),
                              SM.role != 'Name Entry',
                              _row_scope]
        if managed_club_names_list:
            base_agent_filters.append(SM.club.notin_(managed_club_names_list))
        if use_archive and archive_period_id:
            base_agent_filters += [SM.period_id == archive_period_id, SM.upload_id.in_(archive_upload_ids)]
        elif upload_ids_filter:
            # Active data with date filter — was missing, causing direct players
            # to aggregate across ALL active uploads instead of only the filtered ones
            base_agent_filters.append(SM.upload_id.in_(upload_ids_filter))

        my_players_db = SM.query.with_entities(
            SM.player_id,
            sqlfunc.max(SM.nickname),
            sqlfunc.max(SM.club),
            sqlfunc.max(SM.agent_id),
            sqlfunc.max(SM.role),
            sqlfunc.sum(SM.pnl),
            sqlfunc.sum(SM.rake),
            sqlfunc.sum(SM.hands),
        ).filter(*base_agent_filters
        ).group_by(SM.player_id).all()

        # Build agent structure from DB data
        # First, get actual sa_id per player (for correct direct player filtering)
        _sa_lookup_filters = [or_(SM.sa_id.in_(all_sa_ids), SM.agent_id.in_(all_sa_ids))]
        if use_archive and archive_period_id:
            _sa_lookup_filters.append(SM.period_id == archive_period_id)
            _sa_lookup_filters.append(SM.upload_id.in_(archive_upload_ids))
        player_sa_lookup = dict(SM.query.with_entities(
            SM.player_id, sqlfunc.max(SM.sa_id)
        ).filter(*_sa_lookup_filters).group_by(SM.player_id).all())
        # Apply sa overrides to the lookup
        for _pid, _ov in overrides_map.items():
            if _ov.get('sa_id'):
                player_sa_lookup[_pid] = _ov['sa_id']

        has_child_sas = len(child_sa_ids) > 0
        all_my_player_ids = set()
        agents_map = {}  # agent_id -> {nick, members, totals}
        direct_players = []
        for pid, nick, club, ag_id, role, pnl, rake, hands in my_players_db:
            pnl = round(float(pnl or 0), 2)
            rake = round(float(rake or 0), 2)
            hands = int(hands or 0)
            all_my_player_ids.add(pid)
            # Apply agent_id override — if admin attached this player to a specific agent
            _ov = overrides_map.get(pid)
            if _ov and _ov.get('agent_id'):
                ag_id = _ov['agent_id']
            _is_overridden = bool(_ov and (_ov.get('sa_id') or _ov.get('agent_id')))
            member = {'player_id': pid, 'nickname': nick, 'role': role or 'Player',
                      'pnl': pnl, 'rake': rake, 'hands': hands,
                      'overridden': _is_overridden}
            actual_sa = player_sa_lookup.get(pid, '')
            # SA filtering only needed when user has child SAs (to prevent duplicates).
            # Overridden players bypass this check — admin explicitly attached them
            # via /admin/lost-players, so their natural sa_id is irrelevant.
            sa_ok = True if (_is_overridden or not has_child_sas) else (actual_sa in known_ids)
            if ag_id and ag_id != '-' and ag_id != sa_id and ag_id not in child_sa_ids and sa_ok:
                if ag_id not in agents_map:
                    agents_map[ag_id] = {'id': ag_id, 'nick': ag_id, 'members': [],
                                         'total_pnl': 0, 'total_rake': 0, 'total_hands': 0}
                agents_map[ag_id]['members'].append(member)
                agents_map[ag_id]['total_pnl'] += pnl
                agents_map[ag_id]['total_rake'] += rake
                agents_map[ag_id]['total_hands'] += hands
            elif (not ag_id or ag_id == '-' or ag_id == sa_id) and sa_ok:
                direct_players.append(member)
            # else: belongs to child SA, handled by child_sas section

        # Fetch missing players for agents found in the initial query
        # Only for agents whose sa_id is directly ours (not child SAs - those are handled separately)
        if agents_map:
            # Filter: only agents that belong directly to our SA, not to child SAs
            direct_agent_ids = [ag_id for ag_id in agents_map.keys()
                                if player_sa_lookup.get(ag_id, '') in known_ids]
            if direct_agent_ids:
                _miss_filters = [
                    or_(SM.agent_id.in_(direct_agent_ids), SM.sa_id.in_(direct_agent_ids)),
                    SM.player_id.notin_(list(all_my_player_ids)),
                    SM.role != 'Name Entry'
                ]
                if managed_club_names_list:
                    _miss_filters.append(SM.club.notin_(managed_club_names_list))
                if use_archive and archive_period_id:
                    _miss_filters.append(SM.period_id == archive_period_id)
                    _miss_filters.append(SM.upload_id.in_(archive_upload_ids))
                elif upload_ids_filter:
                    _miss_filters.append(SM.upload_id.in_(upload_ids_filter))
                missing_players = SM.query.with_entities(
                    SM.player_id, sqlfunc.max(SM.nickname),
                    sqlfunc.max(SM.club), sqlfunc.max(SM.agent_id),
                    sqlfunc.max(SM.sa_id),
                    sqlfunc.max(SM.role),
                    sqlfunc.sum(SM.pnl), sqlfunc.sum(SM.rake),
                    sqlfunc.sum(SM.hands),
                ).filter(*_miss_filters).group_by(SM.player_id).all()
                for pid, nick, club, ag_id, sa_id_val, role, pnl, rake, hands in missing_players:
                    pnl = round(float(pnl or 0), 2)
                    rake = round(float(rake or 0), 2)
                    hands = int(hands or 0)
                    all_my_player_ids.add(pid)
                    member = {'player_id': pid, 'nickname': nick, 'role': role or 'Player',
                              'pnl': pnl, 'rake': rake, 'hands': hands}
                    target_ag = ag_id if ag_id in agents_map else (sa_id_val if sa_id_val in agents_map else None)
                    if target_ag:
                        agents_map[target_ag]['members'].append(member)
                        agents_map[target_ag]['total_pnl'] += pnl
                        agents_map[target_ag]['total_rake'] += rake
                        agents_map[target_ag]['total_hands'] += hands

        # Adjust PnL by transfers (settlements)
        from app.union_data import get_transfer_adjustments
        xfer_adj = get_transfer_adjustments(all_my_player_ids)
        for m in direct_players:
            m['pnl'] = round(m['pnl'] + xfer_adj.get(m['player_id'], 0), 2)
        for ag in agents_map.values():
            ag['total_pnl'] = 0
            for m in ag['members']:
                m['pnl'] = round(m['pnl'] + xfer_adj.get(m['player_id'], 0), 2)
                ag['total_pnl'] += m['pnl']
            ag['total_pnl'] = round(ag['total_pnl'], 2)

        # Add agent's own game stats if not already in members (for agents who also play)
        for ag_id, ag in agents_map.items():
            existing_pids = set(m['player_id'] for m in ag['members'])
            if ag_id not in existing_pids:
                _own_filters = [SM.player_id == ag_id, SM.role != 'Name Entry']
                if managed_club_names_list:
                    _own_filters.append(SM.club.notin_(managed_club_names_list))
                if use_archive and archive_period_id:
                    _own_filters += [SM.period_id == archive_period_id, SM.upload_id.in_(archive_upload_ids)]
                elif upload_ids_filter:
                    _own_filters.append(SM.upload_id.in_(upload_ids_filter))
                own_stats = SM.query.with_entities(
                    sqlfunc.sum(SM.pnl),
                    sqlfunc.sum(SM.rake),
                    sqlfunc.sum(SM.hands),
                ).filter(*_own_filters).first()
                if own_stats and (float(own_stats[0] or 0) != 0 or float(own_stats[1] or 0) != 0):
                    ag_nick = ag.get('nick', ag_id)
                    own_pnl = round(float(own_stats[0] or 0) + xfer_adj.get(ag_id, 0), 2)
                    own_rake = round(float(own_stats[1] or 0), 2)
                    own_hands = int(own_stats[2] or 0)
                    member = {'player_id': ag_id, 'nickname': ag_nick, 'role': 'Player',
                              'pnl': own_pnl, 'rake': own_rake, 'hands': own_hands}
                    ag['members'].insert(0, member)
                    ag['total_pnl'] = round(ag['total_pnl'] + own_pnl, 2)
                    ag['total_rake'] = round(ag['total_rake'] + own_rake, 2)
                    ag['total_hands'] += own_hands

        # Add SA's own game stats as a direct player (if not already there)
        for sid in known_ids:
            if sid not in all_my_player_ids:
                _sa_filters = [SM.player_id == sid, SM.role != 'Name Entry']
                if managed_club_names_list:
                    _sa_filters.append(SM.club.notin_(managed_club_names_list))
                if use_archive and archive_period_id:
                    _sa_filters += [SM.period_id == archive_period_id, SM.upload_id.in_(archive_upload_ids)]
                elif upload_ids_filter:
                    _sa_filters.append(SM.upload_id.in_(upload_ids_filter))
                sa_own = SM.query.with_entities(
                    sqlfunc.max(SM.nickname),
                    sqlfunc.sum(SM.pnl),
                    sqlfunc.sum(SM.rake),
                    sqlfunc.sum(SM.hands),
                ).filter(*_sa_filters).first()
                if sa_own and (float(sa_own[1] or 0) != 0 or float(sa_own[2] or 0) != 0):
                    sa_pnl = round(float(sa_own[1] or 0) + xfer_adj.get(sid, 0), 2)
                    sa_member = {'player_id': sid, 'nickname': sa_own[0] or sid, 'role': 'Player',
                                 'pnl': sa_pnl,
                                 'rake': round(float(sa_own[2] or 0), 2),
                                 'hands': int(sa_own[3] or 0)}
                    direct_players.insert(0, sa_member)
                    all_my_player_ids.add(sid)

        # Fetch missing agents and players for child_sas from DB
        for cs in child_sas:
            sa_id_val = cs.get('sa_id')
            if sa_id_val:
                existing_agent_ids = set(cs.get('agents', {}).keys())
                # Find agents in DB that are missing from Excel
                _cs_filters = [SM.sa_id == sa_id_val, SM.agent_id != '', SM.agent_id != '-', SM.agent_id != sa_id_val]
                if use_archive and archive_period_id:
                    _cs_filters += [SM.period_id == archive_period_id, SM.upload_id.in_(archive_upload_ids)]
                elif upload_ids_filter:
                    _cs_filters.append(SM.upload_id.in_(upload_ids_filter))
                db_agents = SM.query.with_entities(sqlfunc.distinct(SM.agent_id)).filter(*_cs_filters).all()
                all_nicks_map = dict(SM.query.with_entities(
                    SM.player_id, sqlfunc.max(SM.nickname)
                ).group_by(SM.player_id).all())
                for (ag_id_db,) in db_agents:
                    if ag_id_db not in existing_agent_ids:
                        ag_nick = all_nicks_map.get(ag_id_db, ag_id_db)
                        cs['agents'][ag_id_db] = {'id': ag_id_db, 'nick': ag_nick, 'members': [],
                                                   'total_pnl': 0, 'total_rake': 0, 'total_hands': 0}

        for cs in child_sas:
            for ag_id, ag in cs.get('agents', {}).items():
                existing_pids = set(m['player_id'] for m in ag.get('members', []))
                _mem_filters = [SM.agent_id == ag_id, SM.role != 'Name Entry']
                if existing_pids:
                    _mem_filters.append(SM.player_id.notin_(list(existing_pids)))
                if managed_club_names_list:
                    _mem_filters.append(SM.club.notin_(managed_club_names_list))
                if use_archive and archive_period_id:
                    _mem_filters += [SM.period_id == archive_period_id, SM.upload_id.in_(archive_upload_ids)]
                elif upload_ids_filter:
                    _mem_filters.append(SM.upload_id.in_(upload_ids_filter))
                db_members = SM.query.with_entities(
                    SM.player_id, sqlfunc.max(SM.nickname),
                    sqlfunc.max(SM.role),
                    sqlfunc.sum(SM.pnl), sqlfunc.sum(SM.rake),
                    sqlfunc.sum(SM.hands),
                ).filter(*_mem_filters).group_by(SM.player_id).all()
                for pid, nick, role, pnl, rake, hands in db_members:
                    ag['members'].append({
                        'player_id': pid, 'nickname': nick, 'role': role or 'Player',
                        'pnl': round(float(pnl or 0), 2),
                        'rake': round(float(rake or 0), 2),
                        'hands': int(hands or 0),
                    })
            # Also check direct players under SA
            sa_id_val = cs.get('sa_id')
            if sa_id_val:
                existing_direct_pids = set(m['player_id'] for m in cs.get('direct', []))
                existing_agent_pids = set()
                for ag in cs.get('agents', {}).values():
                    for m in ag.get('members', []):
                        existing_agent_pids.add(m['player_id'])
                all_existing = existing_direct_pids | existing_agent_pids | {sa_id_val}
                _dir_filters = [SM.sa_id == sa_id_val, SM.agent_id.in_(['', '-']),
                                SM.player_id.notin_(list(all_existing)), SM.role != 'Name Entry']
                if managed_club_names_list:
                    _dir_filters.append(SM.club.notin_(managed_club_names_list))
                if use_archive and archive_period_id:
                    _dir_filters += [SM.period_id == archive_period_id, SM.upload_id.in_(archive_upload_ids)]
                elif upload_ids_filter:
                    _dir_filters.append(SM.upload_id.in_(upload_ids_filter))
                db_direct = SM.query.with_entities(
                    SM.player_id, sqlfunc.max(SM.nickname),
                    sqlfunc.max(SM.role),
                    sqlfunc.sum(SM.pnl), sqlfunc.sum(SM.rake),
                    sqlfunc.sum(SM.hands),
                ).filter(*_dir_filters).group_by(SM.player_id).all()
                for pid, nick, role, pnl, rake, hands in db_direct:
                    cs['direct'].append({
                        'player_id': pid, 'nickname': nick, 'role': role or 'Player',
                        'pnl': round(float(pnl or 0), 2),
                        'rake': round(float(rake or 0), 2),
                        'hands': int(hands or 0),
                    })

        # Add child SA's own game stats as a direct player (if they also play)
        for cs in child_sas:
            sa_id_val = cs.get('sa_id')
            if sa_id_val:
                existing_pids = set(m['player_id'] for m in cs.get('direct', []))
                for ag in cs.get('agents', {}).values():
                    for m in ag.get('members', []):
                        existing_pids.add(m['player_id'])
                if sa_id_val not in existing_pids:
                    _csa_filters = [SM.player_id == sa_id_val, SM.role != 'Name Entry']
                    if managed_club_names_list:
                        _csa_filters.append(SM.club.notin_(managed_club_names_list))
                    if use_archive and archive_period_id:
                        _csa_filters += [SM.period_id == archive_period_id, SM.upload_id.in_(archive_upload_ids)]
                    elif upload_ids_filter:
                        _csa_filters.append(SM.upload_id.in_(upload_ids_filter))
                    sa_own = SM.query.with_entities(
                        sqlfunc.max(SM.nickname),
                        sqlfunc.sum(SM.pnl),
                        sqlfunc.sum(SM.rake),
                        sqlfunc.sum(SM.hands),
                    ).filter(*_csa_filters).first()
                    if sa_own and (float(sa_own[1] or 0) != 0 or float(sa_own[2] or 0) != 0):
                        cs['direct'].insert(0, {
                            'player_id': sa_id_val,
                            'nickname': sa_own[0] or sa_id_val,
                            'role': 'Player',
                            'pnl': round(float(sa_own[1] or 0), 2),
                            'rake': round(float(sa_own[2] or 0), 2),
                            'hands': int(sa_own[3] or 0),
                        })

        # Override ALL child_sas data with cumulative DB data (after missing players added)
        from app.union_data import get_cumulative_stats
        all_child_player_ids = set()
        for cs in child_sas:
            for m in cs.get('direct', []):
                all_child_player_ids.add(m['player_id'])
            for ag in cs.get('agents', {}).values():
                for m in ag.get('members', []):
                    all_child_player_ids.add(m['player_id'])
        for cs in child_sas:
            # Get cumulative stats filtered by THIS SA only (not all SAs a player belongs to)
            cs_sa = cs.get('sa_id')
            cs_player_ids = set()
            for m in cs.get('direct', []):
                cs_player_ids.add(m['player_id'])
            for ag in cs.get('agents', {}).values():
                for m in ag.get('members', []):
                    cs_player_ids.add(m['player_id'])
            # Query cumulative filtered by sa_id
            cumul_cs = {}
            if cs_player_ids and cs_sa:
                _cumul_filters = [or_(SM.sa_id == cs_sa, SM.player_id == cs_sa), SM.player_id.in_(list(cs_player_ids)), SM.role != 'Name Entry']
                if managed_club_names_list:
                    _cumul_filters.append(SM.club.notin_(managed_club_names_list))
                if use_archive and archive_period_id:
                    _cumul_filters += [SM.period_id == archive_period_id, SM.upload_id.in_(archive_upload_ids)]
                elif upload_ids_filter:
                    _cumul_filters.append(SM.upload_id.in_(upload_ids_filter))
                sa_stats = SM.query.with_entities(
                    SM.player_id,
                    sqlfunc.sum(SM.pnl),
                    sqlfunc.sum(SM.rake),
                    sqlfunc.sum(SM.hands),
                ).filter(*_cumul_filters).group_by(SM.player_id).all()
                for pid, pnl, rake, hands in sa_stats:
                    cumul_cs[pid] = {'pnl': round(float(pnl or 0), 2),
                                     'rake': round(float(rake or 0), 2),
                                     'hands': int(hands or 0)}

            # When a date filter is active, players without data in the filtered
            # range are dropped from the display — they didn't play in that window,
            # so they shouldn't appear under the SA at all.
            cs_rake = cs_pnl = cs_hands = 0
            direct_kept = []
            for m in cs.get('direct', []):
                c = cumul_cs.get(m['player_id'])
                if c:
                    m['pnl'] = c['pnl']
                    m['rake'] = c['rake']
                    m['hands'] = c.get('hands', 0)
                    direct_kept.append(m)
                elif not had_date_filter:
                    direct_kept.append(m)
                # else: filtered view and player has no data in range → drop
                if m in direct_kept:
                    cs_rake += m.get('rake', 0)
                    cs_pnl += m.get('pnl', 0)
                    cs_hands += m.get('hands', 0)
            cs['direct'] = direct_kept
            for ag_id_key, ag in list(cs.get('agents', {}).items()):
                ag_r = ag_p = ag_h = 0
                members_kept = []
                for m in ag.get('members', []):
                    c = cumul_cs.get(m['player_id'])
                    if c:
                        m['pnl'] = c['pnl']
                        m['rake'] = c['rake']
                        m['hands'] = c.get('hands', 0)
                        members_kept.append(m)
                    elif not had_date_filter:
                        members_kept.append(m)
                    # else: drop in filtered view
                    if m in members_kept:
                        ag_r += m.get('rake', 0)
                        ag_p += m.get('pnl', 0)
                        ag_h += m.get('hands', 0)
                ag['members'] = members_kept
                ag['total_rake'] = round(ag_r, 2)
                ag['total_pnl'] = round(ag_p, 2)
                ag['total_hands'] = ag_h
                cs_rake += ag_r
                cs_pnl += ag_p
                cs_hands += ag_h
                # Drop empty sub-agents when filtered
                if had_date_filter and not members_kept:
                    cs['agents'].pop(ag_id_key, None)
            cs['total_rake'] = round(cs_rake, 2)
            cs['total_pnl'] = round(cs_pnl, 2)
            cs['total_hands'] = cs_hands

        # Drop child SAs that have no players at all in the filtered range
        if had_date_filter:
            child_sas = [cs for cs in child_sas
                         if cs.get('direct') or cs.get('agents')]

        # Find agent nicknames from Excel + DB
        all_nicks_db = dict(SM.query.with_entities(
            SM.player_id, sqlfunc.max(SM.nickname)
        ).group_by(SM.player_id).all())
        for ag_id in agents_map:
            if agents_map[ag_id]['nick'] == ag_id:
                agents_map[ag_id]['nick'] = all_nicks_db.get(ag_id, ag_id)
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
                # Resolve club name: either via registered club_id, or use
                # the managed_club_id value itself as a literal club name
                # (e.g. "Spc o" has no club_id in the hierarchy).
                club_name = club_id_to_name.get(cfg.managed_club_id) or cfg.managed_club_id
                if not club_name:
                    continue

                # Build ID → nickname map from DB
                all_nicknames = dict(SM.query.with_entities(
                    SM.player_id, sqlfunc.max(SM.nickname)
                ).group_by(SM.player_id).all())

                # Get ALL players in this club from DB
                club_filters = [SM.club == club_name, SM.role != 'Name Entry']
                if use_archive and archive_period_id:
                    club_filters += [SM.period_id == archive_period_id, SM.upload_id.in_(archive_upload_ids)]
                elif upload_ids_filter:
                    club_filters.append(SM.upload_id.in_(upload_ids_filter))
                club_players_db = SM.query.with_entities(
                    SM.player_id,
                    sqlfunc.max(SM.nickname),
                    sqlfunc.max(SM.sa_id),
                    sqlfunc.max(SM.agent_id),
                    sqlfunc.sum(SM.pnl),
                    sqlfunc.sum(SM.rake),
                ).filter(*club_filters
                ).group_by(SM.player_id).all()

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

                # Flat list of all members for simple display
                all_club_members = []
                for pid, nick, sa_id_val, ag_id_val, pnl_val, rake_val in club_players_db:
                    sa_nick = all_nicknames.get(sa_id_val, sa_id_val) if sa_id_val and sa_id_val != '-' else '-'
                    ag_nick = all_nicknames.get(ag_id_val, ag_id_val) if ag_id_val and ag_id_val != '-' else '-'
                    all_club_members.append({
                        'player_id': pid, 'nickname': nick,
                        'sa_nick': sa_nick, 'agent_nick': ag_nick,
                        'pnl_total': round(float(pnl_val or 0), 2),
                        'rake_total': round(float(rake_val or 0), 2),
                    })
                all_club_members.sort(key=lambda m: m['rake_total'], reverse=True)

                from app.routes.admin import MANAGED_CLUB_DISPLAY_NAMES
                display_name = MANAGED_CLUB_DISPLAY_NAMES.get(
                    (sa_id, cfg.managed_club_id), club_name)
                club_obj = {
                    'name': display_name, 'club_id': cfg.managed_club_id,
                    'total_rake': club_rake, 'total_pnl': club_pnl,
                    'super_agents': club_sas, 'no_sa_members': no_sa,
                    'all_members': all_club_members,
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

        # Add child SAs totals to overall totals
        child_sas_rake = round(sum(cs.get('total_rake', 0) for cs in child_sas), 2)
        child_sas_pnl = round(sum(cs.get('total_pnl', 0) for cs in child_sas), 2)
        child_sas_hands = sum(cs.get('total_hands', 0) for cs in child_sas)
        total_rake += child_sas_rake
        total_pnl += child_sas_pnl
        total_hands += child_sas_hands

        personal_rake = round(my_sa_combined['total_rake'] + child_sas_rake, 2)
        clubs_total_rake = round(sum(c.get('total_rake', 0) for c in managed_clubs), 2)
        sa_net_rake = round(personal_rake * rake_pct / 100, 2) if rake_pct else 0
        net_rake = round(sa_net_rake + club_net_rake, 2)

        # Summary totals — override with the unified scope-based calculation
        # that /api/report, /agent/reports and /admin/ overview all use.
        # Each row counted ONCE iff it is in the agent's scope
        # (sa_id/agent_id in hierarchy OR club in managed clubs).
        # Breakdown cards (personal/clubs) keep the display-only values above.
        from app.union_data import get_agent_totals as _unified_agent_totals
        _unified = _unified_agent_totals(
            sa_id,
            upload_ids=upload_ids_filter or None,
            archive_period_id=archive_period_id,
            archive_upload_ids=archive_upload_ids or None,
        )
        total_rake = _unified['total_rake']
        total_pnl = _unified['total_pnl']
        total_hands = _unified['total_hands']
        player_count = _unified['player_count']

        # Sync personal_rake (hier-only bucket) with the unified total so
        # dashboard "רייק אישי" matches the admin overview card. personal_rake
        # = total_rake − clubs_total_rake. This replaces the Excel-derived
        # value (which may double-count agent self-rows via Union Member
        # Statistics). clubs_total_rake stays from managed-club iteration.
        personal_rake = round(total_rake - clubs_total_rake, 2)
        sa_net_rake = round(personal_rake * rake_pct / 100, 2) if rake_pct else 0
        net_rake = round(sa_net_rake + club_net_rake, 2)

        # Sort players by PnL:
        #   1. Positives first  (biggest win at top)
        #   2. Negatives second (biggest loss first)
        #   3. Zeros last
        # Members use either 'pnl' (hierarchy path) or 'pnl_total' (managed-club path).
        def _pnl_key(m):
            v = m.get('pnl')
            if v is None:
                v = m.get('pnl_total', 0)
            v = v or 0
            if v > 0:
                return (0, -v)   # positives first, largest first
            if v < 0:
                return (1, v)    # negatives second, most-negative first
            return (2, 0)         # zeros last

        def _sort_by_pnl(lst):
            if lst:
                lst.sort(key=_pnl_key)

        _sort_by_pnl(my_sa_combined.get('direct'))
        for _ag in my_sa_combined.get('agents', {}).values():
            _sort_by_pnl(_ag.get('members'))
        for _cs in child_sas:
            _sort_by_pnl(_cs.get('direct'))
            for _ag in _cs.get('agents', {}).values():
                _sort_by_pnl(_ag.get('members'))
        for _mc in managed_clubs:
            for _sa in _mc.get('super_agents', {}).values():
                _sort_by_pnl(_sa.get('direct_members'))
                for _ag in _sa.get('agents', {}).values():
                    _sort_by_pnl(_ag.get('members'))
            _sort_by_pnl(_mc.get('no_sa_members'))

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

        hide_personal_breakdown = sa_id in AGENTS_HIDE_PERSONAL_BREAKDOWN

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
                               selected_dates=selected_dates,
                               view_as_username=view_as_username,
                               hide_personal_breakdown=hide_personal_breakdown)

    # Admin preview of any player's dashboard via ?view_player=<player_id>
    _admin_view_player = None
    if (hasattr(current_user, 'role') and current_user.role == 'admin'
            and request.args.get('view_player')):
        _admin_view_player = request.args.get('view_player')

    if (hasattr(current_user, 'role') and current_user.role == 'player' and current_user.player_id) \
            or _admin_view_player:
        from app.union_data import get_cumulative_stats
        from app.models import PlayerSession, MoneyTransfer

        player_id = _admin_view_player if _admin_view_player else current_user.player_id
        cs = get_cumulative_stats([player_id]).get(player_id)
        if cs:
            from app.union_data import get_transfer_adjustments
            xfer_adj = get_transfer_adjustments([player_id])
            cs['pnl'] = round(cs['pnl'] + xfer_adj.get(player_id, 0), 2)
        from app.models import (DailyUpload, DailyPlayerStats,
                                ArchivedPlayerSession, ArchivedUpload,
                                ArchivedPlayerStats)
        from datetime import date, timedelta
        archive_cutoff = date.today() - timedelta(days=90)

        # Active sessions (since last reset — fresh count drives default stats)
        active_rows = (PlayerSession.query
                       .join(DailyUpload, PlayerSession.upload_id == DailyUpload.id)
                       .add_columns(DailyUpload.upload_date)
                       .filter(PlayerSession.player_id == player_id)
                       .order_by(DailyUpload.upload_date.asc())
                       .all())
        active_sessions = [{'table_name': s.table_name, 'game_type': s.game_type,
                            'blinds': s.blinds or '', 'pnl': round(s.pnl, 2),
                            'date': d.strftime('%Y-%m-%d') if d else '',
                            'source': 'active'}
                           for s, d in active_rows]
        active_dates = sorted({s['date'] for s in active_sessions if s['date']})

        # Archived sessions (last 90 days) — available for calendar filtering, not in default stats
        arc_rows = (ArchivedPlayerSession.query
                    .join(ArchivedUpload,
                          db.and_(ArchivedPlayerSession.upload_id == ArchivedUpload.original_id,
                                  ArchivedPlayerSession.period_id == ArchivedUpload.period_id))
                    .add_columns(ArchivedUpload.upload_date)
                    .filter(ArchivedPlayerSession.player_id == player_id,
                            ArchivedUpload.upload_date >= archive_cutoff)
                    .order_by(ArchivedUpload.upload_date.asc())
                    .all())
        archived_sessions = [{'table_name': s.table_name, 'game_type': s.game_type,
                              'blinds': s.blinds or '', 'pnl': round(s.pnl, 2),
                              'date': d.strftime('%Y-%m-%d') if d else '',
                              'source': 'archived'}
                             for s, d in arc_rows
                             # skip archived dates that also exist in active (avoid double-count after re-upload)
                             if not d or d.strftime('%Y-%m-%d') not in set(active_dates)]

        session_list = active_sessions + archived_sessions
        # Ascending by date so the earliest game of the cycle appears first, latest last
        session_list.sort(key=lambda x: x.get('date', ''))

        # Per-date stats (hands, rake) — needed for calendar filtering of top cards
        active_daily = (DailyPlayerStats.query
                        .join(DailyUpload, DailyPlayerStats.upload_id == DailyUpload.id)
                        .add_columns(DailyUpload.upload_date)
                        .filter(DailyPlayerStats.player_id == player_id)
                        .all())
        daily_stats_map = {}
        for ds, d in active_daily:
            key = d.strftime('%Y-%m-%d') if d else ''
            if not key:
                continue
            cur = daily_stats_map.setdefault(key, {'hands': 0, 'rake': 0, 'source': 'active'})
            cur['hands'] += ds.hands or 0
            cur['rake'] += ds.rake or 0

        arc_daily = (ArchivedPlayerStats.query
                     .join(ArchivedUpload,
                           db.and_(ArchivedPlayerStats.upload_id == ArchivedUpload.original_id,
                                   ArchivedPlayerStats.period_id == ArchivedUpload.period_id))
                     .add_columns(ArchivedUpload.upload_date)
                     .filter(ArchivedPlayerStats.player_id == player_id,
                             ArchivedUpload.upload_date >= archive_cutoff)
                     .all())
        for ds, d in arc_daily:
            key = d.strftime('%Y-%m-%d') if d else ''
            if not key or key in daily_stats_map:
                continue
            cur = daily_stats_map.setdefault(key, {'hands': 0, 'rake': 0, 'source': 'archived'})
            cur['hands'] += ds.hands or 0
            cur['rake'] += ds.rake or 0

        # Get transfers for this player
        player_transfers = MoneyTransfer.query.filter(
            db.or_(MoneyTransfer.from_player_id == player_id,
                   MoneyTransfer.to_player_id == player_id)
        ).order_by(MoneyTransfer.created_at.desc()).all()
        transfer_rows = []
        for t in player_transfers:
            if t.from_player_id == player_id:
                transfer_rows.append({'label': f'תשלום ל-{t.to_name}',
                                      'amount': round(t.amount, 2)})
            else:
                transfer_rows.append({'label': f'קבלת תשלום מ-{t.from_name}',
                                      'amount': round(-t.amount, 2)})

        # Check if player has rake refund config
        from app.models import RakeConfig
        player_rc = RakeConfig.query.filter_by(entity_type='player', entity_id=player_id).first()
        rake_refund = None
        if player_rc and cs:
            rake_refund = round(cs['rake'] * player_rc.rake_percent / 100, 2)

        # Build game type stats (default view = active only; calendar filter rebuilds in JS)
        game_stats = {}
        for s in active_sessions:
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

        total_sessions = sum(g['count'] for g in game_stats.values())
        total_wins = sum(g['wins'] for g in game_stats.values())
        total_losses = sum(g['losses'] for g in game_stats.values())

        return render_template('main/player_dashboard.html',
                               player=cs or {'nickname': current_user.username, 'club': '-', 'pnl': 0, 'rake': 0, 'hands': 0},
                               viewing_player_id=player_id,
                               sessions=session_list, transfer_rows=transfer_rows,
                               rake_refund=rake_refund,
                               rake_refund_pct=(player_rc.rake_percent if player_rc else 0),
                               daily_stats_map=daily_stats_map,
                               active_dates=active_dates,
                               game_stats=game_stats,
                               total_sessions=total_sessions,
                               total_wins=total_wins,
                               total_losses=total_losses)

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


@main_bp.route('/top-players')
@login_required
def agent_top_players():
    """Top players page for agent/club users — filtered to their own players."""
    if current_user.role not in ('agent', 'club') or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.models import DailyPlayerStats, SAHierarchy, SARakeConfig
    from app.union_data import get_transfer_adjustments
    from sqlalchemy import func as sqlfunc, or_

    all_players = []

    if current_user.role == 'agent':
        sa_id = current_user.player_id
        known_ids = {sa_id}

        # Resolve actual SA/Agent ID
        is_sa = DailyPlayerStats.query.filter(DailyPlayerStats.sa_id == sa_id).first() is not None
        is_agent = DailyPlayerStats.query.filter(DailyPlayerStats.agent_id == sa_id).first() is not None
        if not is_sa and not is_agent:
            own_row = DailyPlayerStats.query.filter(DailyPlayerStats.player_id == sa_id).first()
            if own_row:
                role_lower = (own_row.role or '').lower()
                if 'super' in role_lower or role_lower in ('sa',):
                    if own_row.sa_id and own_row.sa_id != '-':
                        known_ids.add(own_row.sa_id)
                elif 'agent' in role_lower:
                    if own_row.agent_id and own_row.agent_id != '-':
                        known_ids.add(own_row.agent_id)
        known_ids.discard('')
        known_ids.discard('-')

        # Child SAs
        child_sa_ids = []
        for kid in known_ids:
            child_sa_ids.extend([h.child_sa_id for h in SAHierarchy.query.filter_by(parent_sa_id=kid).all()])
        all_sa_ids = list(known_ids) + child_sa_ids

        # Get ALL players under this SA hierarchy (including agents/SAs who also play)
        # Step 1: Find all player_ids from any upload
        my_pids = [r[0] for r in DailyPlayerStats.query.with_entities(
            DailyPlayerStats.player_id
        ).filter(
            or_(DailyPlayerStats.sa_id.in_(all_sa_ids),
                DailyPlayerStats.agent_id.in_(all_sa_ids)),
            DailyPlayerStats.role != 'Name Entry'
        ).distinct().all()]

        # Also include SA IDs themselves (they may play too)
        for sid in all_sa_ids:
            if sid not in my_pids:
                has_stats = DailyPlayerStats.query.filter(
                    DailyPlayerStats.player_id == sid,
                    DailyPlayerStats.role != 'Name Entry'
                ).first()
                if has_stats:
                    my_pids.append(sid)

        # Step 2: Single unified-scope aggregation — each row counted ONCE
        # iff it's in this agent's scope (sa_id/agent_id in hierarchy OR
        # club in managed clubs). Same logic used by /api/report and
        # agent_dashboard — prevents cross-channel leakage and double counts.
        from app.union_data import get_agent_scope
        _scope_sa_ids, managed_club_names = get_agent_scope(sa_id)
        scope_preds = [DailyPlayerStats.sa_id.in_(_scope_sa_ids),
                       DailyPlayerStats.agent_id.in_(_scope_sa_ids)]
        if managed_club_names:
            scope_preds.append(DailyPlayerStats.club.in_(managed_club_names))
        players_db = DailyPlayerStats.query.with_entities(
            DailyPlayerStats.player_id,
            sqlfunc.max(DailyPlayerStats.nickname),
            sqlfunc.max(DailyPlayerStats.club),
            sqlfunc.max(DailyPlayerStats.agent_id),
            sqlfunc.sum(DailyPlayerStats.pnl),
            sqlfunc.sum(DailyPlayerStats.rake),
            sqlfunc.sum(DailyPlayerStats.hands),
        ).filter(
            or_(*scope_preds),
            DailyPlayerStats.role != 'Name Entry',
        ).group_by(DailyPlayerStats.player_id).all()

        # Nickname lookup for agent names
        all_nicks = dict(DailyPlayerStats.query.with_entities(
            DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname)
        ).group_by(DailyPlayerStats.player_id).all())

        xfer_adj = get_transfer_adjustments([p[0] for p in players_db])

        for p in players_db:
            pid, nick, club, ag_id, pnl, rake, hands = p
            pnl = round(float(pnl or 0) + xfer_adj.get(pid, 0), 2)
            rake = round(float(rake or 0), 2)
            hands = int(hands or 0)
            if hands == 0 and pnl == 0:
                continue
            ag_nick = all_nicks.get(ag_id, ag_id) if ag_id and ag_id != '-' and ag_id not in all_sa_ids else ''
            all_players.append({
                'player_id': pid, 'member_id': pid,
                'nickname': nick, 'club': club or '',
                'agent_nick': ag_nick,
                'pnl': pnl, 'pnl_total': pnl,
                'rake': rake, 'rake_total': rake,
                'hands': hands, 'hands_total': hands,
            })

    elif current_user.role == 'club':
        # Club user — get all players in managed club
        from app.models import SARakeConfig as SRC2
        club_id = current_user.player_id
        from app.union_data import get_members_hierarchy
        clubs_data, _ = get_members_hierarchy()
        club_name = None
        for c in clubs_data:
            if str(c['club_id']) == str(club_id):
                club_name = c['name']
                break
        if club_name:
            all_nicks = dict(DailyPlayerStats.query.with_entities(
                DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname)
            ).group_by(DailyPlayerStats.player_id).all())

            club_players_db = DailyPlayerStats.query.with_entities(
                DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname),
                sqlfunc.max(DailyPlayerStats.club), sqlfunc.max(DailyPlayerStats.agent_id),
                sqlfunc.sum(DailyPlayerStats.pnl), sqlfunc.sum(DailyPlayerStats.rake),
                sqlfunc.sum(DailyPlayerStats.hands),
            ).filter(
                DailyPlayerStats.club == club_name,
                DailyPlayerStats.role != 'Name Entry'
            ).group_by(DailyPlayerStats.player_id).all()

            xfer_adj = get_transfer_adjustments([p[0] for p in club_players_db])

            for p in club_players_db:
                pid, nick, club, ag_id, pnl, rake, hands = p
                pnl = round(float(pnl or 0) + xfer_adj.get(pid, 0), 2)
                rake = round(float(rake or 0), 2)
                hands = int(hands or 0)
                if hands == 0 and pnl == 0:
                    continue
                ag_nick = all_nicks.get(ag_id, ag_id) if ag_id and ag_id != '-' else ''
                all_players.append({
                    'player_id': pid, 'member_id': pid,
                    'nickname': nick, 'club': club or '',
                    'agent_nick': ag_nick,
                    'pnl': pnl, 'pnl_total': pnl,
                    'rake': rake, 'rake_total': rake,
                    'hands': hands, 'hands_total': hands,
                })

    top_winners = [p for p in sorted(all_players, key=lambda x: x['pnl'], reverse=True) if p['pnl'] > 0][:10]
    top_losers = [p for p in sorted(all_players, key=lambda x: x['pnl']) if p['pnl'] < 0][:10]
    top_rake = sorted(all_players, key=lambda x: x['rake'], reverse=True)[:10]
    top_active = sorted(all_players, key=lambda x: x['hands'], reverse=True)[:10]

    biggest_winner = top_winners[0]['pnl'] if top_winners else 0
    biggest_loser = top_losers[0]['pnl'] if top_losers else 0

    return render_template('main/agent_top_players.html',
                           top_winners=top_winners, top_losers=top_losers,
                           top_rake=top_rake, top_active=top_active,
                           total_players=len(all_players),
                           biggest_winner=biggest_winner,
                           biggest_loser=biggest_loser)


@main_bp.route('/agent/reports')
@login_required
def agent_reports():
    if not hasattr(current_user, 'role') or current_user.role != 'agent' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.models import (SAHierarchy, SARakeConfig, DailyPlayerStats,
                            ArchivedPlayerStats, PlayerAssignment)
    from sqlalchemy import func as sqlfunc, or_

    sa_id = current_user.player_id

    # Mirror get_agent_totals() known-IDs resolution so hierarchy breadth
    # matches the dashboard / admin-overview exactly. Use a union of
    # DailyPlayerStats and ArchivedPlayerStats so that archived-only sa_id
    # relationships still resolve (required when the user picks an archived
    # period in the reports page — /api/report switches to archive tables
    # and the client filter must recognise those players).
    known_ids = {sa_id}
    is_sa = (DailyPlayerStats.query.filter(DailyPlayerStats.sa_id == sa_id).first()
             or ArchivedPlayerStats.query.filter(ArchivedPlayerStats.sa_id == sa_id).first()) is not None
    is_ag = (DailyPlayerStats.query.filter(DailyPlayerStats.agent_id == sa_id).first()
             or ArchivedPlayerStats.query.filter(ArchivedPlayerStats.agent_id == sa_id).first()) is not None
    if not is_sa and not is_ag:
        own_row = (DailyPlayerStats.query.filter(DailyPlayerStats.player_id == sa_id).first()
                   or ArchivedPlayerStats.query.filter(ArchivedPlayerStats.player_id == sa_id).first())
        if own_row:
            role_lower = (own_row.role or '').lower()
            if 'super' in role_lower or role_lower in ('sa',):
                if own_row.sa_id and own_row.sa_id != '-':
                    known_ids.add(own_row.sa_id)
            elif 'agent' in role_lower:
                if own_row.agent_id and own_row.agent_id != '-':
                    known_ids.add(own_row.agent_id)
    known_ids.discard('')
    known_ids.discard('-')

    child_sa_ids = []
    for kid in list(known_ids):
        child_sa_ids.extend([h.child_sa_id for h in SAHierarchy.query.filter_by(parent_sa_id=kid).all()])
    all_sa_ids = list(set(list(known_ids) + child_sa_ids))

    # Build player sets PER TABLE so the client can filter correctly whether
    # /api/report used DailyPlayerStats (current period) or
    # ArchivedPlayerStats (archived period). A player whose hierarchy link
    # only exists in Daily shouldn't match archive results and vice versa —
    # otherwise reports over-counts vs the dashboard for that period.
    managed_club_names = []
    from app.union_data import get_members_hierarchy
    rake_cfgs = SARakeConfig.query.filter_by(sa_id=sa_id).filter(
        SARakeConfig.managed_club_id.isnot(None)).all()
    if rake_cfgs:
        clubs_data, _ = get_members_hierarchy()
        club_id_to_name = {c['club_id']: c['name'] for c in clubs_data}
        for cfg in rake_cfgs:
            club_name = club_id_to_name.get(cfg.managed_club_id)
            if club_name and club_name not in managed_club_names:
                managed_club_names.append(club_name)

    def _compute_ids_for(M):
        """Run the dashboard-equivalent player-set logic against a single
        stats model (Daily or Archived). Returns (all_ids, hierarchy_ids)."""
        all_ids = set()
        hier_ids = set()

        # 1) sa_id OR agent_id in hierarchy
        rows = M.query.with_entities(M.player_id).filter(
            or_(M.sa_id.in_(all_sa_ids), M.agent_id.in_(all_sa_ids)),
            M.role != 'Name Entry'
        ).distinct().all()
        for (pid,) in rows:
            all_ids.add(pid); hier_ids.add(pid)

        # 2) PlayerAssignment overrides — these are table-agnostic, so we
        # include them for both models (only if they actually appear in M).
        if override_pids:
            rows = M.query.with_entities(M.player_id).filter(
                M.player_id.in_(override_pids)
            ).distinct().all()
            for (pid,) in rows:
                all_ids.add(pid); hier_ids.add(pid)

        # 3) Missing-agents' players — agents appearing in hierarchy (as sa
        # or agent), then all their members in the same table.
        agent_rows = M.query.with_entities(M.agent_id).filter(
            or_(M.sa_id.in_(all_sa_ids), M.agent_id.in_(all_sa_ids)),
            M.agent_id.isnot(None), M.agent_id != '', M.agent_id != '-',
        ).distinct().all()
        agents_here = [r[0] for r in agent_rows if r[0]]
        if agents_here:
            rows = M.query.with_entities(M.player_id).filter(
                M.agent_id.in_(agents_here), M.role != 'Name Entry'
            ).distinct().all()
            for (pid,) in rows:
                all_ids.add(pid); hier_ids.add(pid)

        # 4) The agent's own game stats (+ sub-SAs' own play)
        rows = M.query.with_entities(M.player_id).filter(
            M.player_id.in_(all_sa_ids), M.role != 'Name Entry'
        ).distinct().all()
        for (pid,) in rows:
            all_ids.add(pid); hier_ids.add(pid)

        # 5) Managed clubs — all members in managed club names.
        # Club players do NOT go into hier_ids (dashboard bucketises them).
        if managed_club_names:
            rows = M.query.with_entities(M.player_id).filter(
                M.club.in_(managed_club_names), M.role != 'Name Entry'
            ).distinct().all()
            for (pid,) in rows:
                all_ids.add(pid)
        return all_ids, hier_ids

    override_rows = PlayerAssignment.query.filter(
        or_(
            PlayerAssignment.assigned_sa_id.in_(all_sa_ids),
            PlayerAssignment.assigned_agent_id.in_(all_sa_ids),
        )
    ).all()
    override_pids = [r.player_id for r in override_rows]

    daily_all_ids,   daily_hier_ids   = _compute_ids_for(DailyPlayerStats)
    archive_all_ids, archive_hier_ids = _compute_ids_for(ArchivedPlayerStats)

    # my_players: union of both for the dropdown. Get nicknames from either table.
    union_ids = daily_all_ids | archive_all_ids
    my_players = []
    my_player_ids = list(union_ids)
    if union_ids:
        nick_map = {}
        for M in (DailyPlayerStats, ArchivedPlayerStats):
            for pid, nick in M.query.with_entities(
                M.player_id, sqlfunc.max(M.nickname)
            ).filter(M.player_id.in_(list(union_ids))).group_by(M.player_id).all():
                nick_map.setdefault(pid, nick)
        my_players = [{'player_id': pid, 'nickname': nick_map.get(pid, pid)}
                      for pid in union_ids]
    # Flatten for downstream compatibility (template vars).
    hierarchy_player_ids = daily_hier_ids | archive_hier_ids

    my_players.sort(key=lambda x: (x['nickname'] or '').lower())

    return render_template('main/agent_reports.html',
                           players=my_players,
                           player_ids=list(my_player_ids),
                           hierarchy_player_ids=list(hierarchy_player_ids),
                           daily_player_ids=list(daily_all_ids),
                           daily_hierarchy_ids=list(daily_hier_ids),
                           archive_player_ids=list(archive_all_ids),
                           archive_hierarchy_ids=list(archive_hier_ids),
                           managed_club_names=managed_club_names)


# ═══════════════════════ EXCEL EXPORTS ═══════════════════════

def _make_excel(sheets_data, filename, period_label=None):
    """Create Excel file from dict of {sheet_name: [{col: val, ...}]}.

    When period_label is given (e.g. "01/04/2026 — 05/04/2026"), a banner row
    is added at the top of every sheet so the reader can see which dates the
    export covers.
    """
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sheet_name, rows in sheets_data.items():
        import re
        safe_name = re.sub(r'[\[\]\*\?:/\\]', '', sheet_name)[:31] or 'Sheet'
        ws = wb.create_sheet(title=safe_name)
        # Optional period banner on row 1
        banner_offset = 0
        if period_label:
            banner = ws.cell(row=1, column=1, value=f'דוח אקסל לתאריכים: {period_label}')
            banner.font = Font(bold=True, color='4361EE', size=12)
            banner.alignment = Alignment(horizontal='right')
            banner_offset = 1
        if not rows:
            continue
        # Headers
        headers = list(rows[0].keys())
        header_row = 1 + banner_offset
        # Merge banner across the header columns for readability
        if banner_offset and len(headers) > 1:
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        header_font = Font(bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color='4361EE', end_color='4361EE', fill_type='solid')
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=header_row, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')
        # Data with color formatting
        green_font = Font(color='2EC4B6', bold=True)
        red_font = Font(color='EF233C', bold=True)
        bold_font = Font(bold=True)
        for row_idx, row_data in enumerate(rows, header_row + 1):
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
                # Green "נטו סוכן" row
                first_val = str(row_data.get(headers[0], ''))
                if first_val.startswith('נטו סוכן'):
                    cell.font = Font(bold=True, color='217346')
        # Auto-width (skip merged banner cells)
        for col in ws.columns:
            lengths = []
            for cell in col:
                if getattr(cell, 'column_letter', None) is None:
                    continue
                lengths.append(len(str(cell.value or '')))
            if lengths:
                # col[0] for a MergedCell may not have column_letter; find first real cell
                first_real = next((c for c in col if getattr(c, 'column_letter', None) is not None), None)
                if first_real is not None:
                    ws.column_dimensions[first_real.column_letter].width = min(max(lengths) + 3, 40)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@main_bp.route('/export/player/<player_id>')
@login_required
def export_player(player_id):
    """Export player personal report - all games, P&L, record.

    Honors ?dates=YYYY-MM-DD,... the same way dashboards do: when set, stats
    and sessions are filtered to those upload dates (active or archived).
    Transfers are only included in the all-time export (no date filter)."""
    from app.models import (PlayerSession, DailyPlayerStats,
                            ArchivedPlayerStats, ArchivedPlayerSession)
    from app.union_data import get_transfer_adjustments
    from sqlalchemy import func as sqlfunc

    # Parse ?dates= filter (shared with dashboards)
    requested_dates = [d.strip() for d in request.args.get('dates', '').split(',') if d.strip()]
    had_date_filter = bool(requested_dates)
    selected_dates = requested_dates
    upload_ids_filter = []
    archive_period_id = None
    archive_upload_ids = []
    use_archive = False
    if selected_dates:
        upload_ids_filter, archive_period_id, archive_upload_ids, selected_dates = _resolve_date_uploads(selected_dates)
        use_archive = bool(archive_upload_ids)

    if use_archive and archive_period_id:
        StatsModel = ArchivedPlayerStats
        SessionModel = ArchivedPlayerSession
        stat_filters = [ArchivedPlayerStats.player_id == player_id,
                        ArchivedPlayerStats.period_id == archive_period_id,
                        ArchivedPlayerStats.upload_id.in_(archive_upload_ids)]
        sess_filters = [ArchivedPlayerSession.player_id == player_id,
                        ArchivedPlayerSession.period_id == archive_period_id,
                        ArchivedPlayerSession.upload_id.in_(archive_upload_ids)]
    else:
        StatsModel = DailyPlayerStats
        SessionModel = PlayerSession
        stat_filters = [DailyPlayerStats.player_id == player_id]
        sess_filters = [PlayerSession.player_id == player_id]
        if upload_ids_filter:
            stat_filters.append(DailyPlayerStats.upload_id.in_(upload_ids_filter))
            sess_filters.append(PlayerSession.upload_id.in_(upload_ids_filter))
        elif had_date_filter:
            # Dates requested but didn't resolve → return empty instead of silent all-time fallback
            stat_filters.append(DailyPlayerStats.upload_id == -1)
            sess_filters.append(PlayerSession.upload_id == -1)

    agg = StatsModel.query.with_entities(
        sqlfunc.sum(StatsModel.pnl), sqlfunc.sum(StatsModel.rake),
        sqlfunc.sum(StatsModel.hands),
        sqlfunc.max(StatsModel.nickname), sqlfunc.max(StatsModel.club),
    ).filter(*stat_filters).first()

    if not agg or agg[3] is None:
        flash('שחקן לא נמצא.', 'danger')
        return redirect(url_for('main.dashboard'))

    cs = {
        'pnl': round(float(agg[0] or 0), 2),
        'rake': round(float(agg[1] or 0), 2),
        'hands': int(agg[2] or 0),
        'nickname': agg[3],
        'club': agg[4],
    }

    # Transfers aren't date-bound; apply them only when exporting the full cumulative view
    if not selected_dates:
        xfer_adj = get_transfer_adjustments([player_id])
        cs['pnl'] = round(cs['pnl'] + xfer_adj.get(player_id, 0), 2)

    sessions = SessionModel.query.filter(*sess_filters).all()
    session_rows = [{'משחק': s.table_name, 'סוג': s.game_type,
                     'בליינדס': s.blinds or '', 'P&L': round(s.pnl, 2)} for s in sessions]

    # Transfer rows — only in the all-time export
    if not selected_dates:
        from app.models import MoneyTransfer
        transfers_out = MoneyTransfer.query.filter_by(from_player_id=player_id).all()
        transfers_in = MoneyTransfer.query.filter_by(to_player_id=player_id).all()
        for t in transfers_out:
            session_rows.append({
                'משחק': f'העברה ל-{t.to_name}',
                'סוג': 'העברה',
                'בליינדס': t.description or '',
                'P&L': round(-t.amount, 2),
            })
        for t in transfers_in:
            session_rows.append({
                'משחק': f'קיבלת מ-{t.from_name}',
                'סוג': 'העברה',
                'בליינדס': t.description or '',
                'P&L': round(t.amount, 2),
            })

    # Add total row at the end
    total_pnl = sum(r['P&L'] for r in session_rows)
    session_rows.append({
        'משחק': 'סה"כ', 'סוג': '', 'בליינדס': '',
        'P&L': round(total_pnl, 2),
    })

    summary = [{'שחקן': cs['nickname'], 'קלאב': cs['club'],
                'P&L': cs['pnl']}]

    suffix = ('_' + '_'.join(selected_dates)) if selected_dates else ''
    period_label = _format_period_label(selected_dates)
    return _make_excel({
        'סיכום': summary,
        'רקורד משחקים': session_rows,
    }, f'{cs["nickname"]}{suffix}_report.xlsx', period_label=period_label)


@main_bp.route('/export/agent/account')
@login_required
def export_agent_account():
    """Export agent account summary - personal rake, club rake, expenses, net.

    Honors ?dates= — limits the personal & club rake to the selected upload
    dates (active or archived). Expenses are all-time because they aren't
    date-bound to uploads. Transfers are applied only in the all-time view."""
    if current_user.role != 'agent' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.models import (SAHierarchy, SARakeConfig, RakeConfig, ExpenseCharge,
                            DailyPlayerStats, ArchivedPlayerStats)
    from app.union_data import get_members_hierarchy, get_transfer_adjustments
    from sqlalchemy import func as sqlfunc, or_

    sa_id = current_user.player_id

    # Date filter
    requested_dates = [d.strip() for d in request.args.get('dates', '').split(',') if d.strip()]
    had_date_filter = bool(requested_dates)
    selected_dates = requested_dates
    upload_ids_filter = []
    archive_period_id = None
    archive_upload_ids = []
    use_archive = False
    if selected_dates:
        upload_ids_filter, archive_period_id, archive_upload_ids, selected_dates = _resolve_date_uploads(selected_dates)
        use_archive = bool(archive_upload_ids)

    if use_archive and archive_period_id:
        SM = ArchivedPlayerStats
        scope = [SM.period_id == archive_period_id, SM.upload_id.in_(archive_upload_ids)]
    else:
        SM = DailyPlayerStats
        scope = []
        if upload_ids_filter:
            scope.append(SM.upload_id.in_(upload_ids_filter))

    # Personal rake — hier channel only (exclude managed-club rows so they
    # aren't double-counted below in the clubs section).
    from app.union_data import get_agent_scope
    _scope_sa_ids, _mc_names = get_agent_scope(sa_id)
    personal_filters = [
        or_(SM.sa_id.in_(_scope_sa_ids), SM.agent_id.in_(_scope_sa_ids)),
        SM.role != 'Name Entry',
    ]
    if _mc_names:
        personal_filters.append(SM.club.notin_(_mc_names))
    personal = SM.query.with_entities(
        sqlfunc.sum(SM.rake), sqlfunc.sum(SM.pnl)
    ).filter(*personal_filters, *scope).first()
    personal_rake = round(float(personal[0] or 0), 2)
    personal_pnl = round(float(personal[1] or 0), 2)
    # Transfers only apply to the unfiltered (all-time) view
    if not had_date_filter:
        all_pids = [r[0] for r in DailyPlayerStats.query.with_entities(
            sqlfunc.distinct(DailyPlayerStats.player_id)
        ).filter(or_(DailyPlayerStats.sa_id.in_(_scope_sa_ids),
                     DailyPlayerStats.agent_id.in_(_scope_sa_ids))).all()]
        if all_pids:
            xfer_adj = get_transfer_adjustments(all_pids)
            personal_pnl = round(personal_pnl + sum(xfer_adj.values()), 2)

    # Club rakes
    rake_cfgs = SARakeConfig.query.filter_by(sa_id=sa_id).filter(SARakeConfig.managed_club_id.isnot(None)).all()
    clubs_data, _ = get_members_hierarchy()
    club_id_to_name = {c['club_id']: c['name'] for c in clubs_data}
    club_rows = []
    total_club_rake = 0
    for cfg in rake_cfgs:
        name = club_id_to_name.get(cfg.managed_club_id, '')
        if name:
            cr = SM.query.with_entities(
                sqlfunc.sum(SM.rake), sqlfunc.sum(SM.pnl)
            ).filter(SM.club == name, *scope).first()
            rake = round(float(cr[0] or 0), 2)
            pnl = round(float(cr[1] or 0), 2)
            club_rc = RakeConfig.query.filter_by(entity_type='club', entity_id=cfg.managed_club_id).first()
            keeps = club_rc.rake_percent if club_rc else 0
            net = round(rake * (100 - keeps) / 100, 2)
            club_rows.append({'מועדון': name, 'Rake': rake, 'P&L': pnl,
                              'מועדון מקבל %': keeps, 'נטו שלי': net})
            total_club_rake += net

    # Expenses — not date-bound to uploads, always included
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

    suffix = ('_' + '_'.join(selected_dates)) if selected_dates else ''
    period_label = _format_period_label(selected_dates)
    return _make_excel(sheets, f'{current_user.username}{suffix}_account.xlsx',
                       period_label=period_label)


@main_bp.route('/export/agent/single/<agent_id>')
@login_required
def export_single_agent(agent_id):
    """Export a single agent's players report."""
    if current_user.role not in ('agent', 'admin', 'club') :
        return redirect(url_for('main.dashboard'))

    from app.models import DailyPlayerStats, ArchivedPlayerStats
    from app.union_data import get_transfer_adjustments
    from sqlalchemy import func as sqlfunc, or_

    # Parse ?dates= filter (shared with dashboards)
    requested_dates = [d.strip() for d in request.args.get('dates', '').split(',') if d.strip()]
    had_date_filter = bool(requested_dates)
    selected_dates = requested_dates
    upload_ids_filter = []
    archive_period_id = None
    archive_upload_ids = []
    use_archive = False
    if selected_dates:
        upload_ids_filter, archive_period_id, archive_upload_ids, selected_dates = _resolve_date_uploads(selected_dates)
        use_archive = bool(archive_upload_ids)

    if use_archive and archive_period_id:
        StatsModel = ArchivedPlayerStats
        base_filters = [ArchivedPlayerStats.period_id == archive_period_id,
                        ArchivedPlayerStats.upload_id.in_(archive_upload_ids),
                        ArchivedPlayerStats.role != 'Name Entry']
    else:
        StatsModel = DailyPlayerStats
        base_filters = [DailyPlayerStats.role != 'Name Entry']
        if upload_ids_filter:
            base_filters.append(DailyPlayerStats.upload_id.in_(upload_ids_filter))
        elif had_date_filter:
            # Dates were requested but didn't resolve to any upload → return empty, don't silently fall back
            base_filters.append(DailyPlayerStats.upload_id == -1)

    # Get all players under this agent/SA (by agent_id or sa_id)
    players = StatsModel.query.with_entities(
        StatsModel.player_id, sqlfunc.max(StatsModel.nickname),
        sqlfunc.max(StatsModel.club), sqlfunc.max(StatsModel.agent_id),
        sqlfunc.sum(StatsModel.pnl), sqlfunc.sum(StatsModel.rake),
        sqlfunc.sum(StatsModel.hands),
    ).filter(
        or_(StatsModel.agent_id == agent_id, StatsModel.sa_id == agent_id),
        *base_filters
    ).group_by(StatsModel.player_id).all()

    # Transfer adjustments only apply to the unfiltered cumulative view
    xfer_adj = get_transfer_adjustments([p[0] for p in players]) if not selected_dates else {}

    # Agent/SA nickname (look in both active and archive)
    agent_nick = StatsModel.query.with_entities(
        sqlfunc.max(StatsModel.nickname)
    ).filter(StatsModel.player_id == agent_id).scalar() or agent_id

    # Nickname lookup
    all_nicks = dict(StatsModel.query.with_entities(
        StatsModel.player_id, sqlfunc.max(StatsModel.nickname)
    ).group_by(StatsModel.player_id).all())

    import re
    full_mode = request.args.get('mode') == 'full'

    all_rows = []
    agent_groups = {}
    direct_rows = []
    for p in players:
        raw_pnl = round(float(p[4] or 0), 2)
        ag = p[3]
        ag_name = all_nicks.get(ag, ag) if ag and ag != '-' and ag != agent_id else ''
        row = {
            'שחקן': p[1], 'ID': p[0], 'קלאב': p[2],
            'סוכן': ag_name,
            'רווח/הפסד': round(raw_pnl + xfer_adj.get(p[0], 0), 2),
            'Rake': round(float(p[5] or 0), 2),
        }
        all_rows.append(row)
        if ag_name:
            if ag_name not in agent_groups:
                agent_groups[ag_name] = []
            agent_groups[ag_name].append(row)
        else:
            direct_rows.append(row)

    sheets = {}

    if full_mode:
        # Single sheet with all players sorted by rake
        all_rows.sort(key=lambda x: x['Rake'], reverse=True)
        all_rows.append({
            'שחקן': 'סה"כ', 'ID': '', 'קלאב': '', 'סוכן': '',
            'רווח/הפסד': round(sum(r['רווח/הפסד'] for r in all_rows), 2),
            'Rake': round(sum(r['Rake'] for r in all_rows), 2),
        })
        sheets[agent_nick[:31]] = all_rows
    else:
        # Sheet per sub-agent
        for ag_name, ag_rows in sorted(agent_groups.items(), key=lambda x: sum(r['Rake'] for r in x[1]), reverse=True):
            ag_rows_clean = [{'שחקן': r['שחקן'], 'ID': r['ID'], 'קלאב': r['קלאב'],
                              'רווח/הפסד': r['רווח/הפסד'], 'Rake': r['Rake']} for r in ag_rows]
            ag_rows_clean.sort(key=lambda x: x['Rake'], reverse=True)
            ag_rows_clean.append({
                'שחקן': 'סה"כ', 'ID': '', 'קלאב': '',
                'רווח/הפסד': round(sum(r['רווח/הפסד'] for r in ag_rows_clean), 2),
                'Rake': round(sum(r['Rake'] for r in ag_rows_clean), 2),
            })
            safe_name = re.sub(r'[\[\]\*\?:/\\]', '', ag_name)[:31] or 'Agent'
            sheets[safe_name] = ag_rows_clean

        if direct_rows:
            dr_clean = [{'שחקן': r['שחקן'], 'ID': r['ID'], 'קלאב': r['קלאב'],
                         'רווח/הפסד': r['רווח/הפסד'], 'Rake': r['Rake']} for r in direct_rows]
            dr_clean.sort(key=lambda x: x['Rake'], reverse=True)
            dr_clean.append({
                'שחקן': 'סה"כ', 'ID': '', 'קלאב': '',
                'רווח/הפסד': round(sum(r['רווח/הפסד'] for r in dr_clean), 2),
                'Rake': round(sum(r['Rake'] for r in dr_clean), 2),
            })
            sheets['שחקנים ישירים'] = dr_clean

    if not sheets:
        sheets[agent_nick[:31]] = []

    suffix = ('_' + '_'.join(selected_dates)) if selected_dates else ''
    period_label = _format_period_label(selected_dates)
    return _make_excel(sheets, f'{agent_nick}{suffix}_players.xlsx', period_label=period_label)


@main_bp.route('/export/agent/players')
@login_required
def export_agent_players():
    """Export all agent's players, agents, SAs, clubs with rake % and totals.

    Honors ?dates= — all stats (players, sub-agents, child SAs, clubs) are
    limited to the selected upload dates. Transfers are only applied in the
    all-time view."""
    if current_user.role != 'agent' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.models import (SAHierarchy, SARakeConfig, DailyPlayerStats,
                            ArchivedPlayerStats, RakeConfig)
    from app.union_data import get_members_hierarchy, get_transfer_adjustments
    from sqlalchemy import func as sqlfunc

    sa_id = current_user.player_id
    all_sa_ids = [sa_id]
    child_sa_ids = [h.child_sa_id for h in SAHierarchy.query.filter_by(parent_sa_id=sa_id).all()]
    all_sa_ids.extend(child_sa_ids)

    # Date filter
    requested_dates = [d.strip() for d in request.args.get('dates', '').split(',') if d.strip()]
    had_date_filter = bool(requested_dates)
    selected_dates = requested_dates
    upload_ids_filter = []
    archive_period_id = None
    archive_upload_ids = []
    use_archive = False
    if selected_dates:
        upload_ids_filter, archive_period_id, archive_upload_ids, selected_dates = _resolve_date_uploads(selected_dates)
        use_archive = bool(archive_upload_ids)

    if use_archive and archive_period_id:
        SM = ArchivedPlayerStats
        scope = [SM.period_id == archive_period_id, SM.upload_id.in_(archive_upload_ids)]
    else:
        SM = DailyPlayerStats
        scope = []
        if upload_ids_filter:
            scope.append(SM.upload_id.in_(upload_ids_filter))

    # Unified scope predicate — row is in scope iff sa_id/agent_id in
    # hierarchy OR club in managed clubs. Every row counted once.
    from app.union_data import get_agent_scope
    from sqlalchemy import or_ as _or
    _scope_sa_ids, _mc_names = get_agent_scope(sa_id)
    _scope_preds = [SM.sa_id.in_(_scope_sa_ids), SM.agent_id.in_(_scope_sa_ids)]
    if _mc_names:
        _scope_preds.append(SM.club.in_(_mc_names))

    # Nickname map (always from active data — needed for resolving names even
    # when archive filter returns no rows for the SA itself)
    all_nicks = dict(DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname)
    ).group_by(DailyPlayerStats.player_id).all())

    sheets = {}

    # ── Sheet 1: My Players (direct) ──
    players = SM.query.with_entities(
        SM.player_id, sqlfunc.max(SM.nickname),
        sqlfunc.max(SM.club), sqlfunc.max(SM.sa_id),
        sqlfunc.max(SM.agent_id),
        sqlfunc.sum(SM.pnl), sqlfunc.sum(SM.rake),
        sqlfunc.sum(SM.hands),
    ).filter(
        _or(*_scope_preds), SM.role != 'Name Entry', *scope,
    ).group_by(SM.player_id).all()

    # Transfers only apply to the unfiltered (all-time) view
    xfer_adj = get_transfer_adjustments([p[0] for p in players]) if not had_date_filter else {}

    # Group players: by agent, by child SA, or direct
    agent_groups = {}  # agent_name -> [players]
    child_sa_groups = {}  # child_sa_name -> [players]
    direct_players = []
    for p in players:
        player_sa = p[3]  # sa_id of this player
        ag_id = p[4] if p[4] and p[4] != '-' else None
        ag_name = all_nicks.get(ag_id, ag_id) if ag_id else None
        raw_pnl = round(float(p[5] or 0), 2)
        row = {
            'שחקן': p[1], 'ID': p[0], 'קלאב': p[2],
            'P&L': round(raw_pnl + xfer_adj.get(p[0], 0), 2),
            'Rake': round(float(p[6] or 0), 2),
        }
        # Check if player belongs to a child SA (not the parent SA)
        if player_sa in child_sa_ids:
            csa_name = all_nicks.get(player_sa, player_sa)
            if csa_name not in child_sa_groups:
                child_sa_groups[csa_name] = []
            child_sa_groups[csa_name].append(row)
        elif ag_name and ag_name != all_nicks.get(sa_id, sa_id):
            if ag_name not in agent_groups:
                agent_groups[ag_name] = []
            agent_groups[ag_name].append(row)
        else:
            direct_players.append(row)

    # Helper: find rake % for an agent/SA by player_id
    def _get_rake_pct(entity_id):
        rc = RakeConfig.query.filter_by(entity_id=entity_id).first()
        return rc.rake_percent if rc else 0

    # Reverse lookup: agent name -> agent player_id
    nicks_to_id = {v: k for k, v in all_nicks.items()}

    # Create sheet per agent
    for ag_name, ag_players in sorted(agent_groups.items(), key=lambda x: sum(r['Rake'] for r in x[1]), reverse=True):
        ag_players.sort(key=lambda x: x['Rake'], reverse=True)
        total_rake = round(sum(r['Rake'] for r in ag_players), 2)
        ag_players.append({
            'שחקן': 'סה"כ', 'ID': '', 'קלאב': '',
            'P&L': round(sum(r['P&L'] for r in ag_players), 2),
            'Rake': total_rake,
        })
        ag_pid = nicks_to_id.get(ag_name, '')
        pct = _get_rake_pct(ag_pid) if ag_pid else 0
        if pct:
            ag_players.append({
                'שחקן': f'נטו סוכן ({pct}%)', 'ID': '', 'קלאב': '',
                'P&L': '', 'Rake': round(total_rake * pct / 100, 2),
            })
        sheets[ag_name[:31]] = ag_players

    # Create sheet per child SA
    import re
    for csa_name, csa_players in sorted(child_sa_groups.items(), key=lambda x: sum(r['Rake'] for r in x[1]), reverse=True):
        csa_players.sort(key=lambda x: x['Rake'], reverse=True)
        total_rake = round(sum(r['Rake'] for r in csa_players), 2)
        csa_players.append({
            'שחקן': 'סה"כ', 'ID': '', 'קלאב': '',
            'P&L': round(sum(r['P&L'] for r in csa_players), 2),
            'Rake': total_rake,
        })
        csa_pid = nicks_to_id.get(csa_name, '')
        pct = _get_rake_pct(csa_pid) if csa_pid else 0
        if pct:
            csa_players.append({
                'שחקן': f'נטו סוכן ({pct}%)', 'ID': '', 'קלאב': '',
                'P&L': '', 'Rake': round(total_rake * pct / 100, 2),
            })
        safe_name = re.sub(r'[\[\]\*\?:/\\]', '', csa_name)[:31] or 'SA'
        sheets[safe_name] = csa_players

    # Direct players sheet
    if direct_players:
        direct_players.sort(key=lambda x: x['Rake'], reverse=True)
        direct_players.append({
            'שחקן': 'סה"כ', 'ID': '', 'קלאב': '',
            'P&L': round(sum(r['P&L'] for r in direct_players), 2),
            'Rake': round(sum(r['Rake'] for r in direct_players), 2),
        })
        sheets['שחקנים ישירים'] = direct_players

    # ── Sheet 2: My Agents ──
    _agent_filters = [
        SM.sa_id.in_(all_sa_ids), SM.role != 'Name Entry',
        SM.agent_id != '', SM.agent_id != '-',
    ]
    if _mc_names:
        _agent_filters.append(SM.club.notin_(_mc_names))
    agent_stats = SM.query.with_entities(
        SM.agent_id,
        sqlfunc.sum(SM.pnl), sqlfunc.sum(SM.rake),
        sqlfunc.sum(SM.hands), sqlfunc.count(sqlfunc.distinct(SM.player_id)),
    ).filter(*_agent_filters, *scope).group_by(SM.agent_id).all()

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
        })
    agent_rows.sort(key=lambda x: x['Rake'], reverse=True)
    if agent_rows:
        agent_rows.append({
            'סוכן': 'סה"כ', 'ID': '', 'שחקנים': sum(r['שחקנים'] for r in agent_rows),
            'P&L': round(sum(r['P&L'] for r in agent_rows), 2),
            'Rake': round(sum(r['Rake'] for r in agent_rows), 2),
            'אחוז רייק %': '',
        })
    sheets['סוכנים'] = agent_rows

    # ── Sheet 3: My Super Agents ──
    sa_rows = []
    for csa_id in child_sa_ids:
        _csa_filters = [SM.sa_id == csa_id, SM.role != 'Name Entry']
        if _mc_names:
            _csa_filters.append(SM.club.notin_(_mc_names))
        sa_data = SM.query.with_entities(
            sqlfunc.sum(SM.pnl), sqlfunc.sum(SM.rake),
            sqlfunc.sum(SM.hands), sqlfunc.count(sqlfunc.distinct(SM.player_id)),
        ).filter(*_csa_filters, *scope).first()
        sa_name = all_nicks.get(csa_id, csa_id)
        rc = RakeConfig.query.filter_by(entity_type='agent', entity_id=csa_id).first()
        rake_pct = rc.rake_percent if rc else 0
        rake = round(float(sa_data[1] or 0), 2)
        sa_rows.append({
            'Super Agent': sa_name, 'ID': csa_id, 'שחקנים': int(sa_data[3] or 0),
            'P&L': round(float(sa_data[0] or 0), 2), 'Rake': rake,
            'אחוז רייק %': rake_pct,
        })
    sa_rows.sort(key=lambda x: x['Rake'], reverse=True)
    if sa_rows:
        sa_rows.append({
            'Super Agent': 'סה"כ', 'ID': '', 'שחקנים': sum(r['שחקנים'] for r in sa_rows),
            'P&L': round(sum(r['P&L'] for r in sa_rows), 2),
            'Rake': round(sum(r['Rake'] for r in sa_rows), 2),
            'אחוז רייק %': '',
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
            cr = SM.query.with_entities(
                sqlfunc.sum(SM.pnl), sqlfunc.sum(SM.rake),
                sqlfunc.sum(SM.hands), sqlfunc.count(sqlfunc.distinct(SM.player_id)),
            ).filter(SM.club == name, SM.role != 'Name Entry', *scope).first()
            club_rc = RakeConfig.query.filter_by(entity_type='club', entity_id=cfg.managed_club_id).first()
            keeps = club_rc.rake_percent if club_rc else 0
            rake = round(float(cr[1] or 0), 2)
            net = round(rake * (100 - keeps) / 100, 2)
            club_rows.append({
                'מועדון': name, 'שחקנים': int(cr[3] or 0),
                'P&L': round(float(cr[0] or 0), 2), 'Rake': rake,
                'מועדון מקבל %': keeps, 'נטו שלי': net,
            })
        club_rows.sort(key=lambda x: x['Rake'], reverse=True)
        if club_rows:
            club_rows.append({
                'מועדון': 'סה"כ', 'שחקנים': sum(r['שחקנים'] for r in club_rows),
                'P&L': round(sum(r['P&L'] for r in club_rows), 2),
                'Rake': round(sum(r['Rake'] for r in club_rows), 2),
                'מועדון מקבל %': '', 'נטו שלי': round(sum(r['נטו שלי'] for r in club_rows), 2),
            })
        sheets['מועדונים'] = club_rows

    suffix = ('_' + '_'.join(selected_dates)) if selected_dates else ''
    period_label = _format_period_label(selected_dates)
    return _make_excel(sheets, f'{current_user.username}{suffix}_players.xlsx',
                       period_label=period_label)


@main_bp.route('/export/agent/full_box')
@login_required
def export_agent_full_box():
    """Full-box report — every player under this agent's scope in ONE flat sheet.

    No per-agent / per-SA grouping. Honors ?dates= like the other agent
    exports, and uses the same scope predicate to avoid cross-channel leakage.
    """
    if current_user.role != 'agent' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.models import DailyPlayerStats, ArchivedPlayerStats
    from app.union_data import get_agent_scope, get_transfer_adjustments
    from sqlalchemy import func as sqlfunc, or_ as _or

    sa_id = current_user.player_id

    # Date filter (same logic as export_agent_players)
    requested_dates = [d.strip() for d in request.args.get('dates', '').split(',') if d.strip()]
    had_date_filter = bool(requested_dates)
    selected_dates = requested_dates
    upload_ids_filter = []
    archive_period_id = None
    archive_upload_ids = []
    use_archive = False
    if selected_dates:
        upload_ids_filter, archive_period_id, archive_upload_ids, selected_dates = _resolve_date_uploads(selected_dates)
        use_archive = bool(archive_upload_ids)

    if use_archive and archive_period_id:
        SM = ArchivedPlayerStats
        scope = [SM.period_id == archive_period_id, SM.upload_id.in_(archive_upload_ids)]
    else:
        SM = DailyPlayerStats
        scope = []
        if upload_ids_filter:
            scope.append(SM.upload_id.in_(upload_ids_filter))

    # Agent scope — zero-leakage rule
    _scope_sa_ids, _mc_names = get_agent_scope(sa_id)
    _scope_preds = [SM.sa_id.in_(_scope_sa_ids), SM.agent_id.in_(_scope_sa_ids)]
    if _mc_names:
        _scope_preds.append(SM.club.in_(_mc_names))

    # Nickname lookup (from active data — names are stable across archives)
    all_nicks = dict(DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname)
    ).group_by(DailyPlayerStats.player_id).all())

    players = SM.query.with_entities(
        SM.player_id, sqlfunc.max(SM.nickname),
        sqlfunc.max(SM.club), sqlfunc.max(SM.agent_id),
        sqlfunc.max(SM.sa_id),
        sqlfunc.sum(SM.pnl), sqlfunc.sum(SM.rake),
        sqlfunc.sum(SM.hands),
    ).filter(
        _or(*_scope_preds), SM.role != 'Name Entry', *scope,
    ).group_by(SM.player_id).all()

    xfer_adj = get_transfer_adjustments([p[0] for p in players]) if not had_date_filter else {}

    # Group players by club so the sheet isn't interleaved
    # (spc t block, then spc o block, then rafi ginat block, etc.)
    clubs = {}  # club_name -> [row dicts]
    for p in players:
        ag_id = p[3] if p[3] and p[3] != '-' else None
        ag_name = all_nicks.get(ag_id, ag_id) if ag_id else ''
        sa_pid = p[4] if p[4] and p[4] != '-' else None
        sa_name = all_nicks.get(sa_pid, sa_pid) if sa_pid else ''
        raw_pnl = round(float(p[5] or 0), 2)
        row = {
            'שחקן': p[1],
            'ID': p[0],
            'קלאב': p[2] or '',
            'Super Agent': sa_name,
            'סוכן': ag_name,
            'P&L': round(raw_pnl + xfer_adj.get(p[0], 0), 2),
            'Rake': round(float(p[6] or 0), 2),
            'ידיים': int(p[7] or 0),
        }
        clubs.setdefault(row['קלאב'], []).append(row)

    # Sort clubs by total rake desc; within each club sort players by rake desc.
    club_order = sorted(
        clubs.items(),
        key=lambda kv: sum(r['Rake'] for r in kv[1]),
        reverse=True,
    )

    rows = []
    for club_name, club_rows in club_order:
        club_rows.sort(key=lambda r: r['Rake'], reverse=True)
        rows.extend(club_rows)
        # Club subtotal — visually separates each group
        rows.append({
            'שחקן': f'סה"כ {club_name}' if club_name else 'סה"כ',
            'ID': '', 'קלאב': club_name, 'Super Agent': '', 'סוכן': '',
            'P&L': round(sum(r['P&L'] for r in club_rows), 2),
            'Rake': round(sum(r['Rake'] for r in club_rows), 2),
            'ידיים': sum(r['ידיים'] for r in club_rows),
        })

    if rows:
        # Grand total row — label matches the exact 'סה"כ' string so
        # _make_excel applies its bold-total formatting to it.
        data_rows = [r for r in rows if not str(r['שחקן']).startswith('סה"כ')]
        rows.append({
            'שחקן': 'סה"כ', 'ID': '', 'קלאב': '', 'Super Agent': '', 'סוכן': '',
            'P&L': round(sum(r['P&L'] for r in data_rows), 2),
            'Rake': round(sum(r['Rake'] for r in data_rows), 2),
            'ידיים': sum(r['ידיים'] for r in data_rows),
        })

    suffix = ('_' + '_'.join(selected_dates)) if selected_dates else ''
    period_label = _format_period_label(selected_dates)
    return _make_excel({'קופסא מלאה': rows},
                       f'{current_user.username}{suffix}_full_box.xlsx',
                       period_label=period_label)


@main_bp.route('/export/agent/club/<club_id>')
@login_required
def export_agent_club(club_id):
    """Export specific club details - SAs, Agents, Players.

    Honors ?dates= — limits to the selected upload dates."""
    if current_user.role not in ('agent', 'admin') or (current_user.role == 'agent' and not current_user.player_id):
        return redirect(url_for('main.dashboard'))

    from app.models import DailyPlayerStats, ArchivedPlayerStats
    from app.union_data import get_members_hierarchy, get_transfer_adjustments
    from sqlalchemy import func as sqlfunc
    import re

    from app.union_data import resolve_club_name
    club_name = resolve_club_name(club_id)
    if not club_name:
        flash('מועדון לא נמצא.', 'danger')
        return redirect(url_for('main.dashboard'))

    # Date filter
    requested_dates = [d.strip() for d in request.args.get('dates', '').split(',') if d.strip()]
    had_date_filter = bool(requested_dates)
    selected_dates = requested_dates
    upload_ids_filter = []
    archive_period_id = None
    archive_upload_ids = []
    use_archive = False
    if selected_dates:
        upload_ids_filter, archive_period_id, archive_upload_ids, selected_dates = _resolve_date_uploads(selected_dates)
        use_archive = bool(archive_upload_ids)

    if use_archive and archive_period_id:
        SM = ArchivedPlayerStats
        scope = [SM.period_id == archive_period_id, SM.upload_id.in_(archive_upload_ids)]
    else:
        SM = DailyPlayerStats
        scope = []
        if upload_ids_filter:
            scope.append(SM.upload_id.in_(upload_ids_filter))

    players = SM.query.with_entities(
        SM.player_id, sqlfunc.max(SM.nickname),
        sqlfunc.max(SM.sa_id), sqlfunc.max(SM.agent_id),
        sqlfunc.max(SM.role), sqlfunc.sum(SM.pnl),
        sqlfunc.sum(SM.rake), sqlfunc.sum(SM.hands),
    ).filter(SM.club == club_name, SM.role != 'Name Entry', *scope
    ).group_by(SM.player_id).all()

    # Nickname map (always from active data so names resolve even if the SA
    # itself has no rows in the filtered range)
    all_nicks = dict(DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname)
    ).group_by(DailyPlayerStats.player_id).all())

    xfer_adj = get_transfer_adjustments([p[0] for p in players]) if not had_date_filter else {}

    full_mode = request.args.get('mode') == 'full'

    all_rows = []
    sa_groups = {}   # sa_name -> [rows]
    no_sa_rows = []
    for p in players:
        sa_name = all_nicks.get(p[2], p[2]) if p[2] and p[2] != '-' else ''
        ag_name = all_nicks.get(p[3], p[3]) if p[3] and p[3] != '-' else ''
        row = {
            'שחקן': p[1], 'ID': p[0],
            'Super Agent': sa_name,
            'סוכן': ag_name,
            'רווח/הפסד': round(float(p[5] or 0) + xfer_adj.get(p[0], 0), 2),
            'Rake': round(float(p[6] or 0), 2),
        }
        all_rows.append(row)
        if sa_name:
            if sa_name not in sa_groups:
                sa_groups[sa_name] = []
            sa_groups[sa_name].append(row)
        else:
            no_sa_rows.append(row)

    sheets = {}

    if full_mode:
        all_rows.sort(key=lambda x: x['Rake'], reverse=True)
        all_rows.append({
            'שחקן': 'סה"כ', 'ID': '', 'Super Agent': '', 'סוכן': '',
            'רווח/הפסד': round(sum(r['רווח/הפסד'] for r in all_rows), 2),
            'Rake': round(sum(r['Rake'] for r in all_rows), 2),
        })
        sheets[club_name[:31]] = all_rows
    else:
        # Sheet per SA
        for sa_name, sa_rows in sorted(sa_groups.items(), key=lambda x: sum(r['Rake'] for r in x[1]), reverse=True):
            sa_rows_clean = [{'שחקן': r['שחקן'], 'ID': r['ID'], 'סוכן': r['סוכן'],
                              'רווח/הפסד': r['רווח/הפסד'], 'Rake': r['Rake']} for r in sa_rows]
            sa_rows_clean.sort(key=lambda x: x['Rake'], reverse=True)
            sa_rows_clean.append({
                'שחקן': 'סה"כ', 'ID': '', 'סוכן': '',
                'רווח/הפסד': round(sum(r['רווח/הפסד'] for r in sa_rows_clean), 2),
                'Rake': round(sum(r['Rake'] for r in sa_rows_clean), 2),
            })
            safe_name = re.sub(r'[\[\]\*\?:/\\]', '', sa_name)[:31] or 'SA'
            sheets[safe_name] = sa_rows_clean

        if no_sa_rows:
            no_sa_clean = [{'שחקן': r['שחקן'], 'ID': r['ID'], 'סוכן': r['סוכן'],
                            'רווח/הפסד': r['רווח/הפסד'], 'Rake': r['Rake']} for r in no_sa_rows]
            no_sa_clean.sort(key=lambda x: x['Rake'], reverse=True)
            no_sa_clean.append({
                'שחקן': 'סה"כ', 'ID': '', 'סוכן': '',
                'רווח/הפסד': round(sum(r['רווח/הפסד'] for r in no_sa_clean), 2),
                'Rake': round(sum(r['Rake'] for r in no_sa_clean), 2),
            })
            sheets['ללא סוכן'] = no_sa_clean

    if not sheets:
        sheets[club_name[:31]] = []

    suffix = ('_' + '_'.join(selected_dates)) if selected_dates else ''
    period_label = _format_period_label(selected_dates)
    return _make_excel(sheets, f'{club_name}{suffix}_report.xlsx',
                       period_label=period_label)


@main_bp.route('/export/agent/period')
@login_required
def export_agent_period():
    """Export agent data for specific date range."""
    if current_user.role != 'agent' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.models import SAHierarchy, SARakeConfig, DailyPlayerStats, DailyUpload, PlayerSession
    from app.models import ArchivedUpload, ArchivedPlayerStats, ArchivedPlayerSession
    from app.union_data import get_members_hierarchy
    from sqlalchemy import func as sqlfunc
    from datetime import datetime

    from_date = request.args.get('from', '')
    to_date = request.args.get('to', '')
    player_id_filter = request.args.get('player_id', '')
    period_id = request.args.get('period_id', '')
    if not from_date or not to_date:
        flash('יש לבחור תאריכים.', 'danger')
        return redirect(url_for('main.agent_reports'))

    fd = datetime.strptime(from_date, '%Y-%m-%d').date()
    td = datetime.strptime(to_date, '%Y-%m-%d').date()

    sa_id = current_user.player_id
    all_sa_ids = [sa_id]
    child_sa_ids = [h.child_sa_id for h in SAHierarchy.query.filter_by(parent_sa_id=sa_id).all()]
    all_sa_ids.extend(child_sa_ids)

    # Unified scope: sa_id/agent_id in hierarchy OR club in managed clubs.
    from app.union_data import get_agent_scope
    from sqlalchemy import or_ as _or
    _scope_sa_ids, _mc_names = get_agent_scope(current_user.player_id)

    def _scope_preds(M):
        preds = [M.sa_id.in_(_scope_sa_ids), M.agent_id.in_(_scope_sa_ids)]
        if _mc_names:
            preds.append(M.club.in_(_mc_names))
        return _or(*preds)

    if period_id:
        # Query from archive tables
        uploads = ArchivedUpload.query.filter(
            ArchivedUpload.period_id == int(period_id),
            ArchivedUpload.upload_date >= fd, ArchivedUpload.upload_date <= td
        ).all()
        upload_ids = [u.original_id for u in uploads]
        if not upload_ids:
            flash('אין נתונים בטווח התאריכים.', 'warning')
            return redirect(url_for('main.agent_reports'))

        base_filters = [
            ArchivedPlayerStats.period_id == int(period_id),
            ArchivedPlayerStats.upload_id.in_(upload_ids),
            _scope_preds(ArchivedPlayerStats),
            ArchivedPlayerStats.role != 'Name Entry',
        ]
        if player_id_filter:
            base_filters.append(ArchivedPlayerStats.player_id == player_id_filter)

        players = ArchivedPlayerStats.query.with_entities(
            ArchivedPlayerStats.player_id, sqlfunc.max(ArchivedPlayerStats.nickname),
            sqlfunc.max(ArchivedPlayerStats.club), sqlfunc.sum(ArchivedPlayerStats.pnl),
            sqlfunc.sum(ArchivedPlayerStats.rake), sqlfunc.sum(ArchivedPlayerStats.hands),
        ).filter(*base_filters).group_by(ArchivedPlayerStats.player_id).all()

        SessionModel = ArchivedPlayerSession
        session_period_filter = [ArchivedPlayerSession.period_id == int(period_id)]
    else:
        # Query from active tables (existing behavior)
        uploads = DailyUpload.query.filter(DailyUpload.upload_date >= fd, DailyUpload.upload_date <= td).all()
        upload_ids = [u.id for u in uploads]
        if not upload_ids:
            flash('אין נתונים בטווח התאריכים.', 'warning')
            return redirect(url_for('main.agent_reports'))

        base_filters = [
            DailyPlayerStats.upload_id.in_(upload_ids),
            _scope_preds(DailyPlayerStats),
            DailyPlayerStats.role != 'Name Entry',
        ]
        if player_id_filter:
            base_filters.append(DailyPlayerStats.player_id == player_id_filter)

        players = DailyPlayerStats.query.with_entities(
            DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname),
            sqlfunc.max(DailyPlayerStats.club), sqlfunc.sum(DailyPlayerStats.pnl),
            sqlfunc.sum(DailyPlayerStats.rake), sqlfunc.sum(DailyPlayerStats.hands),
        ).filter(*base_filters).group_by(DailyPlayerStats.player_id).all()

        SessionModel = PlayerSession
        session_period_filter = []

    # Transfer adjustments
    from app.union_data import get_transfer_adjustments
    xfer_adj = get_transfer_adjustments([p[0] for p in players])

    rows = []
    for p in players:
        raw_pnl = round(float(p[3] or 0), 2)
        rows.append({'שחקן': p[1], 'ID': p[0], 'קלאב': p[2],
                     'P&L': round(raw_pnl + xfer_adj.get(p[0], 0), 2),
                     'Rake': round(float(p[4] or 0), 2),
                     })
    rows.sort(key=lambda x: x['Rake'], reverse=True)

    sheets = {f'{from_date} - {to_date}': rows}

    # If single player selected, add game sessions sheet
    if player_id_filter and rows:
        sessions = SessionModel.query.filter(
            *session_period_filter,
            SessionModel.upload_id.in_(upload_ids),
            SessionModel.player_id == player_id_filter
        ).all()
        if sessions:
            sess_rows = [{'משחק': s.table_name, 'סוג': s.game_type,
                          'בליינדס': s.blinds or '', 'רווח/הפסד': round(s.pnl, 2)} for s in sessions]
            sess_rows.sort(key=lambda x: x['רווח/הפסד'])
            total_pnl = round(sum(s['רווח/הפסד'] for s in sess_rows), 2)
            sess_rows.append({'משחק': 'סה"כ', 'סוג': '', 'בליינדס': '', 'רווח/הפסד': total_pnl})
            sheets['משחקים'] = sess_rows

    player_nick = rows[0]['שחקן'] if len(rows) == 1 else current_user.username
    return _make_excel(sheets, f'{player_nick}_{from_date}_{to_date}.xlsx')


@main_bp.route('/export/club/report')
@login_required
def export_club_report():
    """Export club report - all SAs, Agents, Players with balances.

    Honors ?dates= — supports both active and archived uploads, with a
    banner on each sheet showing the period."""
    if current_user.role != 'club' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.models import DailyPlayerStats, ArchivedPlayerStats
    from app.union_data import get_members_hierarchy, get_transfer_adjustments
    from sqlalchemy import func as sqlfunc

    club_id = current_user.player_id
    from app.union_data import resolve_club_name
    club_name = resolve_club_name(club_id)
    if not club_name:
        flash('מועדון לא נמצא.', 'danger')
        return redirect(url_for('main.dashboard'))

    # Date filter (shared helper — supports active + archive)
    requested_dates = [d.strip() for d in request.args.get('dates', '').split(',') if d.strip()]
    had_date_filter = bool(requested_dates)
    selected_dates = requested_dates
    upload_ids_filter = []
    archive_period_id = None
    archive_upload_ids = []
    use_archive = False
    if selected_dates:
        upload_ids_filter, archive_period_id, archive_upload_ids, selected_dates = _resolve_date_uploads(selected_dates)
        use_archive = bool(archive_upload_ids)

    if use_archive and archive_period_id:
        SM = ArchivedPlayerStats
        base_filters = [SM.club == club_name, SM.role != 'Name Entry',
                        SM.period_id == archive_period_id,
                        SM.upload_id.in_(archive_upload_ids)]
    else:
        SM = DailyPlayerStats
        base_filters = [SM.club == club_name, SM.role != 'Name Entry']
        if upload_ids_filter:
            base_filters.append(SM.upload_id.in_(upload_ids_filter))

    # All players in this club
    players = SM.query.with_entities(
        SM.player_id, sqlfunc.max(SM.nickname),
        sqlfunc.max(SM.sa_id), sqlfunc.max(SM.agent_id),
        sqlfunc.sum(SM.pnl), sqlfunc.sum(SM.rake),
        sqlfunc.sum(SM.hands),
    ).filter(*base_filters).group_by(SM.player_id).all()

    # Nickname map (always from active data so names resolve)
    all_nicks = dict(DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname)
    ).group_by(DailyPlayerStats.player_id).all())

    import re
    xfer_adj = get_transfer_adjustments([p[0] for p in players]) if not had_date_filter else {}
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
            'P&L': round(float(p[4] or 0) + xfer_adj.get(p[0], 0), 2),
            'Rake': round(float(p[5] or 0), 2),
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
        })
    sa_rows.sort(key=lambda x: x['Rake'], reverse=True)
    if sa_rows:
        sa_rows.append({
            'Super Agent': 'סה"כ', 'ID': '', 'שחקנים': sum(r['שחקנים'] for r in sa_rows),
            'P&L': round(sum(r['P&L'] for r in sa_rows), 2),
            'Rake': round(sum(r['Rake'] for r in sa_rows), 2),
        })
        sheets['Super Agents'] = sa_rows

    filename_suffix = ('_' + '_'.join(selected_dates)) if selected_dates else ''
    period_label = _format_period_label(selected_dates)
    return _make_excel(sheets, f'{club_name}_report{filename_suffix}.xlsx',
                       period_label=period_label)


@main_bp.route('/club/reports')
@login_required
def club_reports():
    if not hasattr(current_user, 'role') or current_user.role != 'club' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.models import DailyPlayerStats
    from app.union_data import get_members_hierarchy
    from sqlalchemy import func as sqlfunc

    club_id = current_user.player_id
    from app.union_data import resolve_club_name
    club_name = resolve_club_name(club_id)
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

    from app.models import DailyPlayerStats, DailyUpload, PlayerSession
    from app.union_data import get_members_hierarchy
    from sqlalchemy import func as sqlfunc
    from datetime import datetime

    from_date = request.args.get('from', '')
    to_date = request.args.get('to', '')
    player_id_filter = request.args.get('player_id', '')
    if not from_date or not to_date:
        flash('יש לבחור תאריכים.', 'danger')
        return redirect(url_for('main.club_reports'))

    fd = datetime.strptime(from_date, '%Y-%m-%d').date()
    td = datetime.strptime(to_date, '%Y-%m-%d').date()

    club_id = current_user.player_id
    from app.union_data import resolve_club_name
    club_name = resolve_club_name(club_id)
    if not club_name:
        flash('מועדון לא נמצא.', 'danger')
        return redirect(url_for('main.club_reports'))

    uploads = DailyUpload.query.filter(DailyUpload.upload_date >= fd, DailyUpload.upload_date <= td).all()
    upload_ids = [u.id for u in uploads]
    if not upload_ids:
        flash('אין נתונים בטווח התאריכים.', 'warning')
        return redirect(url_for('main.club_reports'))

    base_filters = [
        DailyPlayerStats.upload_id.in_(upload_ids),
        DailyPlayerStats.club == club_name,
        DailyPlayerStats.role != 'Name Entry',
    ]
    if player_id_filter:
        base_filters.append(DailyPlayerStats.player_id == player_id_filter)

    players = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname),
        sqlfunc.max(DailyPlayerStats.club), sqlfunc.sum(DailyPlayerStats.pnl),
        sqlfunc.sum(DailyPlayerStats.rake), sqlfunc.sum(DailyPlayerStats.hands),
    ).filter(
        *base_filters
    ).group_by(DailyPlayerStats.player_id).all()

    from app.union_data import get_transfer_adjustments
    xfer_adj = get_transfer_adjustments([p[0] for p in players])
    rows = [{'שחקן': p[1], 'ID': p[0], 'קלאב': p[2],
             'P&L': round(float(p[3] or 0) + xfer_adj.get(p[0], 0), 2),
             'Rake': round(float(p[4] or 0), 2),
             } for p in players]
    rows.sort(key=lambda x: x['Rake'], reverse=True)

    sheets = {f'{from_date} - {to_date}': rows}

    if player_id_filter and rows:
        sessions = PlayerSession.query.filter(
            PlayerSession.upload_id.in_(upload_ids),
            PlayerSession.player_id == player_id_filter
        ).all()
        if sessions:
            sess_rows = [{'משחק': s.table_name, 'סוג': s.game_type,
                          'בליינדס': s.blinds or '', 'רווח/הפסד': round(s.pnl, 2)} for s in sessions]
            sess_rows.sort(key=lambda x: x['רווח/הפסד'])
            total_pnl = round(sum(s['רווח/הפסד'] for s in sess_rows), 2)
            sess_rows.append({'משחק': 'סה"כ', 'סוג': '', 'בליינדס': '', 'רווח/הפסד': total_pnl})
            sheets['משחקים'] = sess_rows

    player_nick = rows[0]['שחקן'] if len(rows) == 1 else club_name
    return _make_excel(sheets, f'{player_nick}_{from_date}_{to_date}.xlsx')


@main_bp.route('/club/transfers', methods=['GET', 'POST'])
@login_required
def club_transfers():
    if not hasattr(current_user, 'role') or current_user.role != 'club' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.union_data import get_player_balance, get_all_balances, get_members_hierarchy
    from app.models import MoneyTransfer, DailyPlayerStats
    from sqlalchemy import func as sqlfunc

    club_id = current_user.player_id
    from app.union_data import resolve_club_name
    club_name = resolve_club_name(club_id)
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


@main_bp.route('/export/admin/period')
@login_required
def export_admin_period():
    """Export all players data for specific date range (admin)."""
    if current_user.role != 'admin':
        return redirect(url_for('main.dashboard'))

    from app.models import DailyPlayerStats, DailyUpload, PlayerSession
    from sqlalchemy import func as sqlfunc
    from datetime import datetime

    from_date = request.args.get('from', '')
    to_date = request.args.get('to', '')
    player_id_filter = request.args.get('player_id', '')
    if not from_date or not to_date:
        flash('יש לבחור תאריכים.', 'danger')
        return redirect(url_for('admin.reports'))

    fd = datetime.strptime(from_date, '%Y-%m-%d').date()
    td = datetime.strptime(to_date, '%Y-%m-%d').date()

    uploads = DailyUpload.query.filter(DailyUpload.upload_date >= fd, DailyUpload.upload_date <= td).all()
    upload_ids = [u.id for u in uploads]
    if not upload_ids:
        flash('אין נתונים בטווח התאריכים.', 'warning')
        return redirect(url_for('admin.reports'))

    base_filters = [
        DailyPlayerStats.upload_id.in_(upload_ids),
        DailyPlayerStats.role != 'Name Entry',
    ]
    if player_id_filter:
        base_filters.append(DailyPlayerStats.player_id == player_id_filter)

    players = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname),
        sqlfunc.max(DailyPlayerStats.club), sqlfunc.sum(DailyPlayerStats.pnl),
        sqlfunc.sum(DailyPlayerStats.rake), sqlfunc.sum(DailyPlayerStats.hands),
    ).filter(*base_filters).group_by(DailyPlayerStats.player_id).all()

    from app.union_data import get_transfer_adjustments
    xfer_adj = get_transfer_adjustments([p[0] for p in players])

    rows = []
    for p in players:
        raw_pnl = round(float(p[3] or 0), 2)
        rows.append({'שחקן': p[1], 'ID': p[0], 'קלאב': p[2],
                     'P&L': round(raw_pnl + xfer_adj.get(p[0], 0), 2),
                     'Rake': round(float(p[4] or 0), 2),
                     })
    rows.sort(key=lambda x: x['Rake'], reverse=True)

    sheets = {f'{from_date} - {to_date}': rows}

    if player_id_filter and rows:
        sessions = PlayerSession.query.filter(
            PlayerSession.upload_id.in_(upload_ids),
            PlayerSession.player_id == player_id_filter
        ).all()
        if sessions:
            sess_rows = [{'משחק': s.table_name, 'סוג': s.game_type,
                          'בליינדס': s.blinds or '', 'רווח/הפסד': round(s.pnl, 2)} for s in sessions]
            sess_rows.sort(key=lambda x: x['רווח/הפסד'])
            total_pnl = round(sum(s['רווח/הפסד'] for s in sess_rows), 2)
            sess_rows.append({'משחק': 'סה"כ', 'סוג': '', 'בליינדס': '', 'רווח/הפסד': total_pnl})
            sheets['משחקים'] = sess_rows

    player_nick = rows[0]['שחקן'] if len(rows) == 1 else 'all'
    return _make_excel(sheets, f'{player_nick}_{from_date}_{to_date}.xlsx')


@main_bp.route('/reports/periodic')
@login_required
def periodic_report():
    """Periodic report page — pick date range, download Excel."""
    if current_user.role not in ('admin', 'agent'):
        return redirect(url_for('main.dashboard'))
    from app.models import PlayerSession
    from sqlalchemy import func as sqlfunc
    game_types = [r[0] for r in PlayerSession.query.with_entities(
        sqlfunc.distinct(PlayerSession.game_type)
    ).filter(PlayerSession.game_type.isnot(None)).all() if r[0]]
    return render_template('main/periodic_report.html', game_types=sorted(game_types))


@main_bp.route('/export/periodic')
@login_required
def export_periodic():
    """Generate periodic Excel report for date range."""
    if current_user.role not in ('admin', 'agent'):
        return redirect(url_for('main.dashboard'))

    from app.models import DailyPlayerStats, DailyUpload, PlayerSession, MoneyTransfer, SAHierarchy
    from app.union_data import get_transfer_adjustments
    from sqlalchemy import func as sqlfunc, or_
    from datetime import datetime, timedelta

    from_date = request.args.get('from', '')
    to_date = request.args.get('to', '')
    game_type_filter = request.args.get('game_type', '')
    if not from_date or not to_date:
        flash('יש לבחור תאריכים.', 'danger')
        return redirect(url_for('main.periodic_report'))

    fd = datetime.strptime(from_date, '%Y-%m-%d').date()
    td = datetime.strptime(to_date, '%Y-%m-%d').date()

    # Get uploads in range
    uploads = DailyUpload.query.filter(DailyUpload.upload_date >= fd, DailyUpload.upload_date <= td).all()
    upload_ids = [u.id for u in uploads]
    if not upload_ids:
        flash('אין נתונים בטווח התאריכים.', 'warning')
        return redirect(url_for('main.periodic_report'))

    # Filter by role: agent sees only their players
    base_filters = [
        DailyPlayerStats.upload_id.in_(upload_ids),
        DailyPlayerStats.role != 'Name Entry',
    ]
    if current_user.role == 'agent' and current_user.player_id:
        # Unified scope — hierarchy + managed clubs (no leakage).
        from app.union_data import get_agent_scope
        _scope_sa_ids, _mc_names = get_agent_scope(current_user.player_id)
        _scope_preds = [DailyPlayerStats.sa_id.in_(_scope_sa_ids),
                        DailyPlayerStats.agent_id.in_(_scope_sa_ids)]
        if _mc_names:
            _scope_preds.append(DailyPlayerStats.club.in_(_mc_names))
        base_filters.append(or_(*_scope_preds))

    # Sheet 1: Player summary
    players = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id, sqlfunc.max(DailyPlayerStats.nickname),
        sqlfunc.max(DailyPlayerStats.club), sqlfunc.sum(DailyPlayerStats.pnl),
        sqlfunc.sum(DailyPlayerStats.rake), sqlfunc.sum(DailyPlayerStats.hands),
    ).filter(*base_filters).group_by(DailyPlayerStats.player_id).all()

    player_ids = [p[0] for p in players]
    xfer_adj = get_transfer_adjustments(player_ids)

    summary_rows = []
    for p in players:
        summary_rows.append({
            'שחקן': p[1], 'ID': p[0], 'קלאב': p[2],
            'P&L': round(float(p[3] or 0) + xfer_adj.get(p[0], 0), 2),
            'Rake': round(float(p[4] or 0), 2),
        })
    summary_rows.sort(key=lambda x: x['Rake'], reverse=True)
    if summary_rows:
        summary_rows.append({
            'שחקן': 'סה"כ', 'ID': '', 'קלאב': '',
            'P&L': round(sum(r['P&L'] for r in summary_rows), 2),
            'Rake': round(sum(r['Rake'] for r in summary_rows), 2),
        })

    # Sheet 2: Sessions
    sess_filters = [PlayerSession.upload_id.in_(upload_ids), PlayerSession.player_id.in_(player_ids)]
    if game_type_filter:
        sess_filters.append(PlayerSession.game_type == game_type_filter)
    sessions = (PlayerSession.query
                .join(DailyUpload, PlayerSession.upload_id == DailyUpload.id)
                .add_columns(DailyUpload.upload_date)
                .filter(*sess_filters)
                .order_by(DailyUpload.upload_date.asc())
                .all())

    sess_rows = []
    for s, upload_date in sessions:
        sess_rows.append({
            'תאריך': upload_date.strftime('%d/%m/%Y') if upload_date else '',
            'שחקן': s.player_id,
            'משחק': s.table_name, 'סוג': s.game_type,
            'בליינדס': s.blinds or '',
            'P&L': round(s.pnl, 2),
        })
    if sess_rows:
        sess_rows.append({
            'תאריך': '', 'שחקן': '', 'משחק': 'סה"כ', 'סוג': '', 'בליינדס': '',
            'P&L': round(sum(r['P&L'] for r in sess_rows), 2),
        })

    # Sheet 3: Transfers in period
    transfer_filters = [MoneyTransfer.created_at >= datetime.combine(fd, datetime.min.time()),
                        MoneyTransfer.created_at <= datetime.combine(td, datetime.max.time())]
    if current_user.role == 'agent' and player_ids:
        transfer_filters.append(or_(
            MoneyTransfer.from_player_id.in_(player_ids),
            MoneyTransfer.to_player_id.in_(player_ids),
        ))
    transfers = MoneyTransfer.query.filter(*transfer_filters).order_by(MoneyTransfer.created_at.asc()).all()

    xfer_rows = []
    for t in transfers:
        il_time = t.created_at + timedelta(hours=3) if t.created_at else None
        xfer_rows.append({
            'תאריך': il_time.strftime('%d/%m/%Y %H:%M') if il_time else '',
            'משלם': t.from_name, 'מקבל': t.to_name,
            'סכום': round(t.amount, 2),
            'תיאור': t.description or '',
        })
    if xfer_rows:
        xfer_rows.append({
            'תאריך': '', 'משלם': '', 'מקבל': 'סה"כ',
            'סכום': round(sum(r['סכום'] for r in xfer_rows), 2),
            'תיאור': '',
        })

    sheets = {'סיכום שחקנים': summary_rows or []}
    if sess_rows:
        sheets['רקורד משחקים'] = sess_rows
    if xfer_rows:
        sheets['העברות'] = xfer_rows

    return _make_excel(sheets, f'periodic_{from_date}_{to_date}.xlsx')


@main_bp.route('/api/periodic-report')
@login_required
def periodic_report_api():
    """Return periodic report data as JSON for preview."""
    if current_user.role not in ('admin', 'agent'):
        return jsonify({'error': 'unauthorized'}), 403

    from app.models import DailyPlayerStats, DailyUpload, MoneyTransfer, SAHierarchy
    from app.union_data import get_transfer_adjustments
    from sqlalchemy import func as sqlfunc, or_
    from datetime import datetime, timedelta

    from_date = request.args.get('from', '')
    to_date = request.args.get('to', '')
    if not from_date or not to_date:
        return jsonify({'error': 'missing dates'}), 400

    fd = datetime.strptime(from_date, '%Y-%m-%d').date()
    td = datetime.strptime(to_date, '%Y-%m-%d').date()

    # Check active + archive
    active_ids, arc_pid, arc_ids, _ = _resolve_date_uploads(
        [(fd + timedelta(days=i)).strftime('%Y-%m-%d') for i in range((td - fd).days + 1)]
    )
    all_upload_ids = active_ids + arc_ids

    if arc_ids and arc_pid:
        from app.models import ArchivedPlayerStats
        SM = ArchivedPlayerStats
        base_filters = [SM.period_id == arc_pid, SM.upload_id.in_(arc_ids), SM.role != 'Name Entry']
    else:
        SM = DailyPlayerStats
        base_filters = [SM.role != 'Name Entry']
        if active_ids:
            base_filters.append(SM.upload_id.in_(active_ids))

    game_type_filter = request.args.get('game_type', '')

    if current_user.role == 'agent' and current_user.player_id:
        # Unified scope — hierarchy + managed clubs (no leakage across channels).
        from app.union_data import get_agent_scope
        _scope_sa_ids, _mc_names = get_agent_scope(current_user.player_id)
        _scope_preds = [SM.sa_id.in_(_scope_sa_ids), SM.agent_id.in_(_scope_sa_ids)]
        if _mc_names:
            _scope_preds.append(SM.club.in_(_mc_names))
        base_filters.append(or_(*_scope_preds))

    if game_type_filter:
        # Filter by game type — use PlayerSession for P&L per game type
        from app.models import PlayerSession
        all_upload_ids = active_ids + arc_ids
        sess_filters = [PlayerSession.game_type == game_type_filter]
        if all_upload_ids:
            sess_filters.append(PlayerSession.upload_id.in_(all_upload_ids))

        # Get player_ids from the base SM query first (for permission filtering)
        allowed_pids = [r[0] for r in SM.query.with_entities(SM.player_id).filter(*base_filters).distinct().all()]
        if allowed_pids:
            sess_filters.append(PlayerSession.player_id.in_(allowed_pids))

        players = db.session.query(
            PlayerSession.player_id,
            sqlfunc.sum(PlayerSession.pnl),
        ).filter(*sess_filters).group_by(PlayerSession.player_id).all()

        # Get nicknames/clubs from SM
        nick_map = dict(SM.query.with_entities(SM.player_id, sqlfunc.max(SM.nickname)).filter(
            SM.player_id.in_([p[0] for p in players])
        ).group_by(SM.player_id).all())
        club_map = dict(SM.query.with_entities(SM.player_id, sqlfunc.max(SM.club)).filter(
            SM.player_id.in_([p[0] for p in players])
        ).group_by(SM.player_id).all())

        player_ids = [p[0] for p in players]
        xfer_adj = get_transfer_adjustments(player_ids)

        summary = []
        tot_pnl = tot_rake = tot_hands = 0
        for p in players:
            pnl = round(float(p[1] or 0) + xfer_adj.get(p[0], 0), 2)
            summary.append({'name': nick_map.get(p[0], p[0]), 'id': p[0],
                            'club': club_map.get(p[0], ''), 'pnl': pnl, 'rake': 0, 'hands': 0})
            tot_pnl += pnl
        summary.sort(key=lambda x: x['pnl'])
    else:
        players = SM.query.with_entities(
            SM.player_id, sqlfunc.max(SM.nickname),
            sqlfunc.max(SM.club), sqlfunc.sum(SM.pnl),
            sqlfunc.sum(SM.rake), sqlfunc.sum(SM.hands),
        ).filter(*base_filters).group_by(SM.player_id).all()

        player_ids = [p[0] for p in players]
        xfer_adj = get_transfer_adjustments(player_ids)

        summary = []
        tot_pnl = tot_rake = tot_hands = 0
        for p in players:
            pnl = round(float(p[3] or 0) + xfer_adj.get(p[0], 0), 2)
            rake = round(float(p[4] or 0), 2)
            hands = int(p[5] or 0)
            summary.append({'name': p[1], 'id': p[0], 'club': p[2], 'pnl': pnl, 'rake': rake, 'hands': hands})
            tot_pnl += pnl
            tot_rake += rake
            tot_hands += hands
        summary.sort(key=lambda x: x['rake'], reverse=True)

    # Transfers
    transfer_filters = [MoneyTransfer.created_at >= datetime.combine(fd, datetime.min.time()),
                        MoneyTransfer.created_at <= datetime.combine(td, datetime.max.time())]
    if current_user.role == 'agent' and player_ids:
        transfer_filters.append(or_(MoneyTransfer.from_player_id.in_(player_ids), MoneyTransfer.to_player_id.in_(player_ids)))
    transfers = MoneyTransfer.query.filter(*transfer_filters).order_by(MoneyTransfer.created_at.asc()).all()

    xfer_list = []
    for t in transfers:
        il_time = t.created_at + timedelta(hours=3) if t.created_at else None
        xfer_list.append({
            'date': il_time.strftime('%d/%m/%Y %H:%M') if il_time else '',
            'from': t.from_name, 'to': t.to_name,
            'amount': round(t.amount, 2), 'desc': t.description or '',
        })

    return jsonify({
        'summary': summary,
        'totals': {'pnl': round(tot_pnl, 2), 'rake': round(tot_rake, 2), 'hands': tot_hands},
        'transfers': xfer_list,
    })


@main_bp.route('/api/report')
@login_required
def report_api():
    from app.models import DailyPlayerStats, DailyUpload, ArchivedUpload, ArchivedPlayerStats
    from datetime import datetime
    from sqlalchemy import func, or_
    from app.models import SAHierarchy, SARakeConfig

    from_date = request.args.get('from')
    to_date = request.args.get('to')
    player_id = request.args.get('player_id', '')
    period_id = request.args.get('period_id', '')
    club_names_raw = request.args.get('club_names', '')
    club_names = [c for c in club_names_raw.split(',') if c.strip()]

    if not from_date or not to_date:
        return jsonify({'error': 'missing dates'}), 400

    try:
        fd = datetime.strptime(from_date, '%Y-%m-%d').date()
        td = datetime.strptime(to_date, '%Y-%m-%d').date()
    except ValueError:
        return jsonify({'error': 'invalid date format'}), 400

    # Compute the agent's hierarchy for row-level channel filtering. When a
    # player appears in multiple channels (e.g. rows under sa_id=Hatofer AND
    # rows under club=AnDenDino), only the rows that belong to the agent's
    # own channels should be aggregated into their totals — otherwise reports
    # mixes in activity from outside the agent's scope.
    agent_sa_ids = []
    agent_club_names = list(club_names)
    if getattr(current_user, 'role', None) == 'agent' and current_user.player_id:
        _sa = current_user.player_id
        _all = {_sa}
        _all.update(h.child_sa_id for h in SAHierarchy.query.filter_by(parent_sa_id=_sa).all())
        _all.discard(''); _all.discard('-')
        agent_sa_ids = list(_all)
        if not agent_club_names:
            _rake_cfgs = SARakeConfig.query.filter_by(sa_id=_sa).filter(
                SARakeConfig.managed_club_id.isnot(None)).all()
            if _rake_cfgs:
                from app.union_data import get_members_hierarchy
                _clubs_data, _ = get_members_hierarchy()
                _cid_to_name = {c['club_id']: c['name'] for c in _clubs_data}
                # Fall back to raw managed_club_id as literal club name when
                # not registered in clubs_data (e.g. "Spc o").
                agent_club_names = [_cid_to_name.get(c.managed_club_id) or c.managed_club_id
                                    for c in _rake_cfgs]

    def _hierarchy_row_filter(M):
        """Row-level filter: keep only rows whose sa_id/agent_id is in our
        hierarchy, or whose club is one of our managed clubs. Returns None
        if the user is not an agent (no filtering applied — admin case)."""
        if not agent_sa_ids and not agent_club_names:
            return None
        preds = []
        if agent_sa_ids:
            preds.append(M.sa_id.in_(agent_sa_ids))
            preds.append(M.agent_id.in_(agent_sa_ids))
        if agent_club_names:
            preds.append(M.club.in_(agent_club_names))
        return or_(*preds)

    if period_id:
        # Query from archive tables
        uploads = ArchivedUpload.query.filter(
            ArchivedUpload.period_id == int(period_id),
            ArchivedUpload.upload_date >= fd,
            ArchivedUpload.upload_date <= td
        ).all()
        upload_ids = [u.original_id for u in uploads]

        if not upload_ids:
            return jsonify({'players': [], 'totals': {'pnl': 0, 'rake': 0, 'hands': 0}, 'days': 0,
                            'managed_clubs_totals': None})

        base_filters = [
            ArchivedPlayerStats.period_id == int(period_id),
            ArchivedPlayerStats.upload_id.in_(upload_ids),
            ArchivedPlayerStats.role != 'Name Entry',
        ]
        row_filter = _hierarchy_row_filter(ArchivedPlayerStats)
        if row_filter is not None:
            base_filters.append(row_filter)
        query = ArchivedPlayerStats.query.with_entities(
            ArchivedPlayerStats.player_id,
            func.max(ArchivedPlayerStats.nickname),
            func.max(ArchivedPlayerStats.club),
            func.sum(ArchivedPlayerStats.pnl),
            func.sum(ArchivedPlayerStats.rake),
            func.sum(ArchivedPlayerStats.hands),
        ).filter(*base_filters)
        if player_id:
            query = query.filter(ArchivedPlayerStats.player_id == player_id)
        query = query.group_by(ArchivedPlayerStats.player_id)
    else:
        # Query from active tables (existing behavior)
        uploads = DailyUpload.query.filter(
            DailyUpload.upload_date >= fd,
            DailyUpload.upload_date <= td
        ).all()
        upload_ids = [u.id for u in uploads]

        if not upload_ids:
            return jsonify({'players': [], 'totals': {'pnl': 0, 'rake': 0, 'hands': 0}, 'days': 0,
                            'managed_clubs_totals': None})

        base_filters = [
            DailyPlayerStats.upload_id.in_(upload_ids),
            DailyPlayerStats.role != 'Name Entry',
        ]
        row_filter = _hierarchy_row_filter(DailyPlayerStats)
        if row_filter is not None:
            base_filters.append(row_filter)
        query = DailyPlayerStats.query.with_entities(
            DailyPlayerStats.player_id,
            func.max(DailyPlayerStats.nickname),
            func.max(DailyPlayerStats.club),
            func.sum(DailyPlayerStats.pnl),
            func.sum(DailyPlayerStats.rake),
            func.sum(DailyPlayerStats.hands),
        ).filter(*base_filters)
        if player_id:
            query = query.filter(DailyPlayerStats.player_id == player_id)
        query = query.group_by(DailyPlayerStats.player_id)

    results = query.all()

    from app.union_data import get_transfer_adjustments
    xfer_adj = get_transfer_adjustments([r[0] for r in results])

    players = []
    total_pnl = 0
    total_rake = 0
    total_hands = 0
    for pid, nick, club, pnl, rake, hands in results:
        p = round(float(pnl or 0) + xfer_adj.get(pid, 0), 2)
        r = round(float(rake or 0), 2)
        h = int(hands or 0)
        players.append({'player_id': pid, 'nickname': nick, 'club': club,
                        'pnl': p, 'rake': r, 'hands': h})
        total_pnl += p
        total_rake += r
        total_hands += h

    players.sort(key=lambda x: x['pnl'], reverse=True)

    # Managed-clubs totals — sum rake/pnl over ALL players in the given clubs
    # in the same date range. Mirrors the dashboard's "רייק מועדונים" bucket,
    # which is added on top of the hierarchy total (overlap is counted twice —
    # this is intentional, to match dashboard arithmetic).
    managed_clubs_totals = None
    if club_names and upload_ids:
        if period_id:
            mc_q = ArchivedPlayerStats.query.with_entities(
                func.sum(ArchivedPlayerStats.pnl),
                func.sum(ArchivedPlayerStats.rake),
                func.sum(ArchivedPlayerStats.hands),
            ).filter(
                ArchivedPlayerStats.period_id == int(period_id),
                ArchivedPlayerStats.upload_id.in_(upload_ids),
                ArchivedPlayerStats.club.in_(club_names),
                ArchivedPlayerStats.role != 'Name Entry',
            ).first()
        else:
            mc_q = DailyPlayerStats.query.with_entities(
                func.sum(DailyPlayerStats.pnl),
                func.sum(DailyPlayerStats.rake),
                func.sum(DailyPlayerStats.hands),
            ).filter(
                DailyPlayerStats.upload_id.in_(upload_ids),
                DailyPlayerStats.club.in_(club_names),
                DailyPlayerStats.role != 'Name Entry',
            ).first()
        mc_pnl, mc_rake, mc_hands = mc_q if mc_q else (0, 0, 0)
        managed_clubs_totals = {
            'pnl': round(float(mc_pnl or 0), 2),
            'rake': round(float(mc_rake or 0), 2),
            'hands': int(mc_hands or 0),
        }

    return jsonify({
        'players': players,
        'totals': {'pnl': round(total_pnl, 2), 'rake': round(total_rake, 2), 'hands': total_hands},
        'days': len(upload_ids),
        'managed_clubs_totals': managed_clubs_totals,
    })


@main_bp.route('/api/report-dates')
@login_required
def report_dates_api():
    """Return list of dates that have upload data, plus archived periods."""
    from app.models import DailyUpload, ArchivedUpload, ArchivePeriod

    period_id = request.args.get('period_id', '')

    if period_id:
        # Return dates for specific archived period
        archived = ArchivedUpload.query.with_entities(ArchivedUpload.upload_date).filter(
            ArchivedUpload.period_id == int(period_id)
        ).distinct().all()
        dates = [u[0].strftime('%Y-%m-%d') for u in archived]
    else:
        # Return active dates
        uploads = DailyUpload.query.with_entities(DailyUpload.upload_date).distinct().all()
        dates = [u[0].strftime('%Y-%m-%d') for u in uploads]

    # Always return periods list
    periods = ArchivePeriod.query.order_by(ArchivePeriod.last_date.desc()).all()
    periods_list = [{'id': p.id, 'label': p.label} for p in periods]

    # Current period label from active uploads
    from sqlalchemy import func as sqlfunc
    current_range = db.session.query(
        sqlfunc.min(DailyUpload.upload_date),
        sqlfunc.max(DailyUpload.upload_date)
    ).first()
    current_label = ''
    if current_range and current_range[0] is not None:
        f, l = current_range
        current_label = f"{f.strftime('%d/%m/%Y')} — {l.strftime('%d/%m/%Y')}"

    return jsonify({'dates': dates, 'periods': periods_list, 'current_label': current_label})


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
    from app.models import PlayerSession, DailyUpload
    # Read ALL sessions from cumulative DB with upload date
    db_sessions = (PlayerSession.query
                   .join(DailyUpload, PlayerSession.upload_id == DailyUpload.id)
                   .add_columns(DailyUpload.upload_date)
                   .filter(PlayerSession.player_id == player_id)
                   .order_by(DailyUpload.upload_date.asc())
                   .all())
    sessions = []
    for s, upload_date in db_sessions:
        date_fmt = upload_date.strftime('%d/%m/%Y') if upload_date else ''
        sessions.append({
            'table': s.table_name,
            'game': s.game_type,
            'blinds': s.blinds or '',
            'date': date_fmt,
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
