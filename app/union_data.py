import pandas as pd
import os

# Mutable config so reset can clear it
_config = {
    'excel_path': os.environ.get(
        'UNION_EXCEL_PATH',
        r'C:\Users\Administrator\Downloads\קבצים\18.xlsx'
    )
}


def set_excel_path(path):
    _config['excel_path'] = path


def get_excel_path():
    return _config['excel_path']


def _load_sa_hierarchy():
    """Load parent→children SA mapping from DB."""
    from app.models import SAHierarchy
    hierarchy = {}
    for row in SAHierarchy.query.all():
        hierarchy.setdefault(row.parent_sa_id, []).append(row.child_sa_id)
    return hierarchy


def get_all_super_agents():
    """Returns list of all unique super agents from the Excel: [{id, nick, club}]."""
    sheets = _read_sheets()
    if 'Union Member Statistics' not in sheets:
        return []
    df = sheets['Union Member Statistics']
    sa_map = {}
    current_club = ''
    for i in range(6, len(df)):
        row = df.iloc[i]
        if '(ID:' in str(row.iloc[0]):
            current_club = str(row.iloc[0]).split(' (ID:')[0]
        sa_id = str(row.iloc[2])
        sa_nick = str(row.iloc[3])
        if sa_id not in ('-', 'nan') and sa_id not in sa_map:
            sa_map[sa_id] = {'id': sa_id, 'nick': sa_nick, 'club': current_club}
    return sorted(sa_map.values(), key=lambda x: x['nick'].lower())


def get_all_clubs():
    """Returns list of all clubs: [{club_id, name}]."""
    sheets = _read_sheets()
    if 'Union Member Statistics' not in sheets:
        return []
    df = sheets['Union Member Statistics']
    clubs = []
    for i in range(6, len(df)):
        cell = str(df.iloc[i, 0])
        if '(ID:' in cell:
            clubs.append({
                'club_id': cell.split('(ID:')[1].rstrip(')'),
                'name': cell.split(' (ID:')[0],
            })
    return clubs


def get_all_members():
    """Returns list of all unique members: [{player_id, nickname, role, club, sa_nick, agent_nick}]."""
    sheets = _read_sheets()
    if 'Union Member Statistics' not in sheets:
        return []
    df = sheets['Union Member Statistics']
    members = []
    seen = set()
    current_club = ''
    for i in range(6, len(df)):
        row = df.iloc[i]
        if '(ID:' in str(row.iloc[0]):
            current_club = str(row.iloc[0]).split(' (ID:')[0]

        # Add Super Agent if not seen
        sa_id = str(row.iloc[2])
        sa_nick = str(row.iloc[3])
        if sa_id not in ('-', 'nan') and sa_id not in seen:
            seen.add(sa_id)
            members.append({
                'player_id': sa_id, 'nickname': sa_nick,
                'role': 'Super Agent', 'club': current_club,
                'sa_nick': '-', 'agent_nick': '-',
            })

        # Add Agent if not seen
        ag_id = str(row.iloc[4])
        ag_nick = str(row.iloc[5])
        if ag_id not in ('-', 'nan') and ag_id not in seen:
            seen.add(ag_id)
            members.append({
                'player_id': ag_id, 'nickname': ag_nick,
                'role': 'Agent', 'club': current_club,
                'sa_nick': sa_nick, 'agent_nick': '-',
            })

        # Add Player
        pid = str(row.iloc[8])
        nickname = str(row.iloc[9])
        if nickname in ('nan', '-') or pid in seen:
            continue
        seen.add(pid)
        members.append({
            'player_id': pid,
            'nickname': nickname,
            'role': str(row.iloc[7]),
            'club': current_club,
            'sa_nick': sa_nick if sa_nick not in ('nan', '-') else '-',
            'agent_nick': ag_nick if ag_nick not in ('nan', '-') else '-',
        })
    return sorted(members, key=lambda x: x['nickname'].lower())


def _read_sheets():
    import io
    # Try local file first
    path = get_excel_path()
    if path and os.path.exists(path):
        return pd.read_excel(path, sheet_name=None, header=None)

    # Fallback: read from DB (for Vercel/cloud)
    try:
        from app.models import ActiveExcelData
        active = ActiveExcelData.query.order_by(ActiveExcelData.id.desc()).first()
        if active and active.file_data:
            return pd.read_excel(io.BytesIO(active.file_data), sheet_name=None, header=None)
    except Exception:
        pass

    return {}


def get_union_overview():
    sheets = _read_sheets()
    if 'Union Overview' not in sheets:
        meta = {'union_name': '-', 'union_id': '-', 'period': '-'}
        return meta, [], {'active_players': 0, 'total_hands': 0, 'total_fee': 0, 'pnl': 0}
    df = sheets['Union Overview']

    meta = {
        'union_name': str(df.iloc[0, 0]).replace('Union Name : ', ''),
        'union_id': str(df.iloc[1, 0]).replace('Union ID : ', ''),
        'period': str(df.iloc[2, 0]).replace('Period : ', ''),
    }

    # Col layout: 0=No, 1=NaN(Country), 2=ClubID, 3=ClubName, 4=MasterID,
    #             5=MasterNick, 6=ActivePlayers, 7=TotalHands, 8=TotalFee,
    #             9-16=sub-fees, 17=P&L
    clubs = []
    for i in range(5, len(df)):
        row = df.iloc[i]
        no = str(row.iloc[0])
        if no == 'TOTAL':
            # TOTAL row: 1-5=NaN, 6=ActivePlayers, 7=Hands, 8=Fee, 17=P&L
            total = {
                'active_players': _num(row.iloc[6]),
                'total_hands': _num(row.iloc[7]),
                'total_fee': _num(row.iloc[8]),
                'pnl': _num(row.iloc[17]),
            }
            break
        clubs.append({
            'no': _num(row.iloc[0]),
            'club_id': str(row.iloc[2]),
            'club_name': str(row.iloc[3]),
            'master_id': str(row.iloc[4]),
            'master_nickname': str(row.iloc[5]),
            'active_players': _num(row.iloc[6]),
            'total_hands': _num(row.iloc[7]),
            'total_fee': _num(row.iloc[8]),
            'pnl': _num(row.iloc[17]),
        })
    else:
        total = {}

    return meta, clubs, total


def get_ring_games():
    sheets = _read_sheets()
    if 'Union Ring Game Statistics' not in sheets:
        return [], {'hands': 0, 'buy_in': 0, 'rake': 0}
    df = sheets['Union Ring Game Statistics']

    # Col layout: 0=ClubID, 1=ClubName, 2=TableName, 3=GameType, 4=Role,
    #             5=CreatorID, 6=CreatorNick, 7=BuyinMin, 8=BuyinMax,
    #             9=Rake%, 10=RakeCap, 11=SB, 12=BB, 13=Ante,
    #             14=Insurance, 15=EVCashout, 16=Start, 17=End, 18=Duration,
    #             19=Players, 20=Hands, 21=BuyinAmt, 22=Rake
    games = []
    total_hands = 0
    for i in range(5, len(df)):
        row = df.iloc[i]
        club_name = str(row.iloc[1])
        if club_name in ('nan', 'TOTAL'):
            continue
        hands = _num(row.iloc[20])
        total_hands += hands
        games.append({
            'club_name': club_name,
            'table_name': str(row.iloc[2]),
            'game_type': str(row.iloc[3]),
            'creator': str(row.iloc[6]),
            'blinds': f"{row.iloc[11]}/{row.iloc[12]}",
            'players': _num(row.iloc[19]),
            'hands': hands,
            'buy_in': _num(row.iloc[21]),
            'rake': _num(row.iloc[22]),
            'start': str(row.iloc[16])[:16],
            'duration': str(row.iloc[18]),
        })

    last = df.iloc[-1]
    totals = {
        'hands': total_hands,
        'buy_in': _num(last.iloc[21]),
        'rake': _num(last.iloc[22]),
    }
    return games, totals


def get_mtts():
    sheets = _read_sheets()
    if 'Union MTT Statistics' not in sheets:
        return [], {'entries': 0, 'total_buyin': 0, 'prize_pool': 0}
    df = sheets['Union MTT Statistics']

    # Col layout: 0=ClubID, 1=ClubName, 2=Title, 3=Status, 4=GameType,
    #             5=Role, 6=CreatorID, 7=CreatorNick, 8=Buyin, 9=Fee,
    #             10=ReEntry, 11=GTD, 12=Structure, 13=Rake%,
    #             14=Start, 15=End, 16=Duration, 17=Entries,
    #             18=TotalBuyinChips, 19=TotalBuyinTickets,
    #             20=TotalReEntryChips, 21=TotalReEntryTickets,
    #             22=PayoutPrizeChips, 23=PayoutPrizeTickets,
    #             24=BountyPrize, 25=..., 26=TotalPrizePool, 27=..., 28=Overlay
    mtts = []
    total_entries = 0
    for i in range(6, len(df)):
        row = df.iloc[i]
        club_name = str(row.iloc[1])
        if club_name in ('nan', 'TOTAL'):
            continue
        title = str(row.iloc[2])
        if title == 'TOTAL':
            break
        entries = _num(row.iloc[17])
        total_entries += int(entries)
        mtts.append({
            'club_name': club_name,
            'title': title,
            'status': str(row.iloc[3]),
            'game_type': str(row.iloc[4]),
            'creator': str(row.iloc[7]),
            'buyin': _num(row.iloc[8]),
            'fee': _num(row.iloc[9]),
            'reentry': str(row.iloc[10]),
            'gtd': _num(row.iloc[11]),
            'entries': entries,
            'prize_pool': _num(row.iloc[26]),
            'start': str(row.iloc[14])[:16] if str(row.iloc[14]) != 'nan' else '-',
            'duration': str(row.iloc[16]),
        })

    last = df.iloc[-1]
    total_rake = round(sum(m['fee'] * m['entries'] for m in mtts if m['fee'] > 0), 2)
    totals = {
        'entries': total_entries,
        'total_buyin': _num(last.iloc[18]),
        'prize_pool': _num(last.iloc[26]),
        'total_rake': total_rake,
    }
    return mtts, totals


def get_top_members(limit=20):
    sheets = _read_sheets()
    if 'Union Member Statistics' not in sheets:
        return [], []
    df = sheets['Union Member Statistics']

    # Col layout (row 5 sub-headers): 0=ClubGroup, 1=No, 2=SuperAgentID,
    # 3=SuperAgentNick, 4=AgentID, 5=AgentNick, 6=Country, 7=Role,
    # 8=MemberID, 9=Nickname, 10=P&L_Ring_NLH, 11=P&L_Ring_PLO,
    # 12-17=other game P&L, 37=P&L_Total, 64=Rake_Total, 151=Hands_Total
    members = []
    current_club = ''
    for i in range(6, len(df)):
        row = df.iloc[i]
        club_cell = str(row.iloc[0])
        if '(ID:' in club_cell:
            current_club = club_cell.split(' (ID:')[0]

        nickname = str(row.iloc[9])
        if nickname in ('nan', '-'):
            continue

        members.append({
            'club': current_club,
            'member_id': str(row.iloc[8]),
            'nickname': nickname,
            'country': str(row.iloc[6]),
            'pnl_total': _num(row.iloc[37]),
            'rake_total': _num(row.iloc[64]),
            'hands_total': _num(row.iloc[151]),
        })

    members.sort(key=lambda x: x['pnl_total'], reverse=True)
    return members[:limit], members[-limit:][::-1]


def get_members_hierarchy():
    """Returns list of clubs, each with super_agents → agents → players hierarchy."""
    sheets = _read_sheets()
    if 'Union Member Statistics' not in sheets:
        return [], {'rake': 0, 'pnl': 0}
    df = sheets['Union Member Statistics']

    clubs = []
    current_club = None

    for i in range(6, len(df)):
        row = df.iloc[i]
        club_cell = str(row.iloc[0])

        # New club block
        if '(ID:' in club_cell:
            current_club = {
                'name': club_cell.split(' (ID:')[0],
                'club_id': club_cell.split('(ID:')[1].rstrip(')'),
                'super_agents': {},   # sa_id → {id, nick, agents: {}, direct_members: []}
                'no_sa_members': [],  # members without super agent
            }
            clubs.append(current_club)

        nickname = str(row.iloc[9])
        if nickname in ('nan', '-') or current_club is None:
            continue

        member = {
            'player_id': str(row.iloc[8]),
            'nickname': nickname,
            'role': str(row.iloc[7]),
            'country': str(row.iloc[6]),
            'sa_id': str(row.iloc[2]),
            'sa_nick': str(row.iloc[3]),
            'agent_id': str(row.iloc[4]),
            'agent_nick': str(row.iloc[5]),
            'pnl_total': _num(row.iloc[37]),
            'rake_total': _num(row.iloc[64]),
            'hands_total': _num(row.iloc[151]),
        }

        sa_id = member['sa_id']
        agent_id = member['agent_id']

        # Accumulate club totals
        current_club['total_rake'] = round(current_club.get('total_rake', 0) + member['rake_total'], 2)
        current_club['total_pnl']  = round(current_club.get('total_pnl', 0)  + member['pnl_total'],  2)

        if sa_id == '-':
            current_club['no_sa_members'].append(member)
            continue

        # Ensure super agent entry exists
        if sa_id not in current_club['super_agents']:
            current_club['super_agents'][sa_id] = {
                'id': sa_id,
                'nick': member['sa_nick'],
                'agents': {},
                'direct_members': [],
            }
        sa = current_club['super_agents'][sa_id]

        if agent_id == '-':
            sa['direct_members'].append(member)
        else:
            if agent_id not in sa['agents']:
                sa['agents'][agent_id] = {
                    'id': agent_id,
                    'nick': member['agent_nick'],
                    'members': [],
                }
            sa['agents'][agent_id]['members'].append(member)

    # Apply SA hierarchy: nest child SAs under parent SAs
    hierarchy = _load_sa_hierarchy()
    # Build reverse map: child_sa_id → parent_sa_id
    child_to_parent = {}
    for parent_id, children in hierarchy.items():
        for child_id in children:
            child_to_parent[child_id] = parent_id

    for club in clubs:
        sa_dict = club['super_agents']
        # For each child SA that exists in this club, move it under its parent
        for child_id, parent_id in child_to_parent.items():
            if child_id in sa_dict and parent_id in sa_dict:
                parent_sa = sa_dict[parent_id]
                if 'child_super_agents' not in parent_sa:
                    parent_sa['child_super_agents'] = {}
                parent_sa['child_super_agents'][child_id] = sa_dict[child_id]
                del sa_dict[child_id]

    # Grand totals across all clubs
    grand_rake = round(sum(c.get('total_rake', 0) for c in clubs), 2)
    grand_pnl  = round(sum(c.get('total_pnl',  0) for c in clubs), 2)
    grand = {'rake': grand_rake, 'pnl': grand_pnl}

    return clubs, grand


def get_ring_game_detail():
    """Returns list of table sessions with player results."""
    sheets = _read_sheets()
    if 'Union Ring Game Detail' not in sheets:
        return []
    df = sheets['Union Ring Game Detail']

    tables = []
    current = None

    for i in range(len(df)):
        col0 = str(df.iloc[i, 0])

        if col0.startswith('Table Name :'):
            # "Table Name : PLO6 1/2 D.B , Creator : SHIFKA(1478-4435) , Club : SPC Un(970996)"
            try:
                table_name = col0.split('Table Name : ')[1].split(' , Creator')[0].strip()
                club = col0.split('Club : ')[1].rsplit('(', 1)[0].strip()
            except Exception:
                table_name, club = col0, ''
            current = {'table_name': table_name, 'club': club,
                       'game_type': '', 'blinds': '', 'players': []}
            tables.append(current)

        elif col0.startswith('Table Information :'):
            if current:
                try:
                    current['game_type'] = col0.split('Game : ')[1].split(' ,')[0].strip()
                    current['blinds'] = col0.split('Blinds : ')[1].split(' ,')[0].strip()
                except Exception:
                    pass

        elif current is not None and col0 not in ('Club', 'ID', 'Total', 'nan'):
            # Player data row: col0 is club ID (numeric string), col2 is player ID (XXXX-XXXX)
            player_id = str(df.iloc[i, 2])
            if '-' in player_id and len(player_id) == 9:
                current['players'].append({
                    'club_id': col0,
                    'club_name': str(df.iloc[i, 1]),
                    'player_id': player_id,
                    'nickname': str(df.iloc[i, 3]),
                    'buyin': _num(df.iloc[i, 4]),
                    'cashout': _num(df.iloc[i, 5]),
                    'hands': _num(df.iloc[i, 6]),
                    'rake': _num(df.iloc[i, 12]),
                    'pnl': _num(df.iloc[i, 13]),
                })

    return tables


def get_super_agent_tables():
    """Returns list of super agents, each with their agents and members + stats."""
    sheets = _read_sheets()
    if 'Union Member Statistics' not in sheets:
        return []
    df = sheets['Union Member Statistics']

    sa_map = {}   # (club, sa_id) → sa dict
    order = []    # preserve insertion order

    current_club = ''
    for i in range(6, len(df)):
        row = df.iloc[i]
        if '(ID:' in str(row.iloc[0]):
            current_club = str(row.iloc[0]).split(' (ID:')[0]

        nickname = str(row.iloc[9])
        if nickname in ('nan', '-'):
            continue

        sa_id   = str(row.iloc[2])
        sa_nick = str(row.iloc[3])
        if sa_id == '-':
            continue

        ag_id   = str(row.iloc[4])
        ag_nick = str(row.iloc[5])
        role    = str(row.iloc[7])
        pnl     = _num(row.iloc[37])
        rake    = _num(row.iloc[64])
        hands   = _num(row.iloc[151])
        pid     = str(row.iloc[8])

        key = (current_club, sa_id)
        if key not in sa_map:
            sa_map[key] = {
                'sa_id': sa_id,
                'sa_nick': sa_nick,
                'club': current_club,
                'agents': {},   # ag_id → {id, nick, members:[]}
                'direct': [],
                'total_pnl': 0, 'total_rake': 0, 'total_hands': 0,
            }
            order.append(key)

        sa = sa_map[key]
        sa['total_pnl']   = round(sa['total_pnl']   + pnl,   2)
        sa['total_rake']  = round(sa['total_rake']  + rake,  2)
        sa['total_hands'] = round(sa['total_hands'] + hands, 0)

        member = {'player_id': pid, 'nickname': nickname, 'role': role,
                  'pnl': pnl, 'rake': rake, 'hands': hands}

        if ag_id == '-':
            sa['direct'].append(member)
        else:
            if ag_id not in sa['agents']:
                sa['agents'][ag_id] = {'id': ag_id, 'nick': ag_nick,
                                       'members': [],
                                       'total_pnl': 0, 'total_rake': 0, 'total_hands': 0}
            ag = sa['agents'][ag_id]
            ag['members'].append(member)
            ag['total_pnl']   = round(ag['total_pnl']   + pnl,   2)
            ag['total_rake']  = round(ag['total_rake']  + rake,  2)
            ag['total_hands'] = round(ag['total_hands'] + hands, 0)

    return [sa_map[k] for k in order]


def get_player_detail(player_id):
    """Returns all club entries + ring game sessions for a given player_id."""
    sheets = _read_sheets()
    if 'Union Member Statistics' not in sheets:
        return {}, [], []
    df = sheets['Union Member Statistics']

    # Collect ALL club entries for this player
    club_entries = []
    current_club = ''
    for i in range(6, len(df)):
        row = df.iloc[i]
        if '(ID:' in str(row.iloc[0]):
            current_club = str(row.iloc[0]).split(' (ID:')[0]
        if str(row.iloc[8]) == player_id:
            club_entries.append({
                'club': current_club,
                'role': str(row.iloc[7]),
                'country': str(row.iloc[6]),
                'sa_nick': str(row.iloc[3]),
                'agent_nick': str(row.iloc[5]),
                'pnl_total': _num(row.iloc[37]),
                'rake_total': _num(row.iloc[64]),
                'hands_total': _num(row.iloc[151]),
            })

    # Build member_info from first entry (for basic info)
    member_info = {}
    if club_entries:
        first = club_entries[0]
        member_info = {
            'player_id': player_id,
            'nickname': '',
            'role': first['role'],
            'country': first['country'],
            'club': first['club'],
            'sa_nick': first['sa_nick'],
            'agent_nick': first['agent_nick'],
            'pnl_total': first['pnl_total'],
            'rake_total': first['rake_total'],
            'hands_total': first['hands_total'],
        }
        # Get nickname from first entry
        for i in range(6, len(df)):
            if str(df.iloc[i, 8]) == player_id:
                member_info['nickname'] = str(df.iloc[i, 9])
                break

    # Ring game sessions
    sessions = []
    for table in get_ring_game_detail():
        for p in table['players']:
            if p['player_id'] == player_id:
                sessions.append({
                    'table_name': table['table_name'],
                    'game_type': table['game_type'],
                    'blinds': table['blinds'],
                    'club': table['club'],
                    'club_name': p['club_name'],
                    'buyin': p['buyin'],
                    'cashout': p['cashout'],
                    'hands': p['hands'],
                    'rake': p['rake'],
                    'pnl': p['pnl'],
                })

    return member_info, sessions, club_entries


def get_cumulative_stats(player_ids=None):
    """Returns cumulative stats from all uploads: {player_id: {pnl, rake, hands, nickname, club}}."""
    from app.models import DailyPlayerStats
    from sqlalchemy import func

    query = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id,
        func.sum(DailyPlayerStats.pnl),
        func.sum(DailyPlayerStats.rake),
        func.sum(DailyPlayerStats.hands),
        func.max(DailyPlayerStats.nickname),
        func.max(DailyPlayerStats.club),
    ).group_by(DailyPlayerStats.player_id)

    if player_ids:
        query = query.filter(DailyPlayerStats.player_id.in_(player_ids))

    result = {}
    for pid, pnl, rake, hands, nick, club in query.all():
        result[pid] = {
            'pnl': round(float(pnl or 0), 2),
            'rake': round(float(rake or 0), 2),
            'hands': int(hands or 0),
            'nickname': nick,
            'club': club,
        }
    return result


def get_cumulative_totals():
    """Returns cumulative totals across all uploads for dashboard."""
    from app.models import DailyPlayerStats, DailyUpload
    from sqlalchemy import func

    stats = DailyPlayerStats.query.with_entities(
        func.sum(DailyPlayerStats.pnl),
        func.sum(DailyPlayerStats.rake),
        func.sum(DailyPlayerStats.hands),
        func.count(func.distinct(DailyPlayerStats.player_id)),
    ).first()

    uploads_count = DailyUpload.query.count()

    # Date range from uploads
    date_range = DailyUpload.query.with_entities(
        func.min(DailyUpload.upload_date),
        func.max(DailyUpload.upload_date),
    ).first()
    first_date = date_range[0].strftime('%d/%m/%Y') if date_range[0] else '-'
    last_date = date_range[1].strftime('%d/%m/%Y') if date_range[1] else '-'

    # Per-club totals
    club_stats = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.club,
        func.sum(DailyPlayerStats.pnl),
        func.sum(DailyPlayerStats.rake),
        func.sum(DailyPlayerStats.hands),
        func.count(func.distinct(DailyPlayerStats.player_id)),
    ).group_by(DailyPlayerStats.club).all()

    clubs = []
    for club_name, pnl, rake, hands, players in club_stats:
        if club_name:
            clubs.append({
                'club_name': club_name,
                'pnl': round(float(pnl or 0), 2),
                'total_fee': round(float(rake or 0), 2),
                'total_hands': int(hands or 0),
                'active_players': int(players or 0),
            })

    # Rake breakdown from sessions (Ring vs MTT)
    from app.models import PlayerSession
    ring_rake_total = 0
    mtt_rake_total = 0
    try:
        ring_sessions = PlayerSession.query.with_entities(
            func.count(PlayerSession.id)
        ).filter(PlayerSession.game_type != 'MTT').first()
        mtt_sessions = PlayerSession.query.with_entities(
            func.count(PlayerSession.id)
        ).filter(PlayerSession.game_type == 'MTT').first()

        from app.models import TournamentStats
        mtt_rake_data = TournamentStats.query.with_entities(
            func.sum(TournamentStats.fee * TournamentStats.entries)
        ).first()
        mtt_rake_total = round(float(mtt_rake_data[0] or 0), 2)
    except Exception:
        pass

    total_rake = round(float(stats[1] or 0), 2)
    ring_rake_total = round(total_rake - mtt_rake_total, 2)

    return {
        'total_pnl': round(float(stats[0] or 0), 2),
        'total_rake': total_rake,
        'ring_rake': ring_rake_total,
        'mtt_rake': mtt_rake_total,
        'total_hands': int(stats[2] or 0),
        'total_players': int(stats[3] or 0),
        'uploads_count': uploads_count,
        'clubs': clubs,
        'period': f'{first_date} — {last_date}' if first_date != last_date else first_date,
    }


def get_player_balance(player_id):
    """Returns current balance: cumulative P&L + incoming transfers - outgoing transfers."""
    from app.models import MoneyTransfer
    from sqlalchemy import func

    # Cumulative P&L from all uploads
    stats = get_cumulative_stats([player_id])
    pnl = stats.get(player_id, {}).get('pnl', 0)

    # Transfers
    incoming = MoneyTransfer.query.with_entities(
        func.coalesce(func.sum(MoneyTransfer.amount), 0)
    ).filter_by(to_player_id=player_id).scalar()

    outgoing = MoneyTransfer.query.with_entities(
        func.coalesce(func.sum(MoneyTransfer.amount), 0)
    ).filter_by(from_player_id=player_id).scalar()

    return round(pnl + float(incoming) - float(outgoing), 2)


def get_all_balances(player_ids=None):
    """Returns dict of player_id → balance for all or specific players."""
    from app.models import MoneyTransfer
    from sqlalchemy import func

    # Cumulative P&Ls from all uploads
    pnl_map = {pid: s['pnl'] for pid, s in get_cumulative_stats(player_ids).items()}

    # Get all transfer sums
    transfers_in = dict(MoneyTransfer.query.with_entities(
        MoneyTransfer.to_player_id, func.sum(MoneyTransfer.amount)
    ).group_by(MoneyTransfer.to_player_id).all())

    transfers_out = dict(MoneyTransfer.query.with_entities(
        MoneyTransfer.from_player_id, func.sum(MoneyTransfer.amount)
    ).group_by(MoneyTransfer.from_player_id).all())

    all_ids = set(pnl_map.keys()) | set(transfers_in.keys()) | set(transfers_out.keys())
    if player_ids:
        all_ids &= set(player_ids)

    balances = {}
    for pid in all_ids:
        pnl = pnl_map.get(pid, 0)
        inc = float(transfers_in.get(pid, 0))
        out = float(transfers_out.get(pid, 0))
        balances[pid] = round(pnl + inc - out, 2)

    return balances


def _num(val):
    try:
        f = float(val)
        return round(f, 2)
    except (ValueError, TypeError):
        return 0
