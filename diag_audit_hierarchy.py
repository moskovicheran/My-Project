"""Audit: every SAHierarchy / SARakeConfig entry should be reachable
from the owner's dashboard.

Background — repeated bugs ("I assigned X in the control panel but it
disappeared from the dashboard") all stem from code that intersects
DB-stored hierarchy with Excel-derived lists. This script checks every
agent user and reports anything in the DB that wouldn't render.

Run:
    python diag_audit_hierarchy.py

Exits 0 when everything lines up, 1 if anything is unreachable.
Safe to run anytime — read-only.
"""
import sys

from app import create_app
from app.models import User, SAHierarchy, SARakeConfig, DailyPlayerStats
from app.union_data import (
    get_child_sa_entries,
    get_super_agent_tables,
    get_members_hierarchy,
)
from sqlalchemy import func as sqlfunc


def _nick_lookup():
    return dict(DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id,
        sqlfunc.max(DailyPlayerStats.nickname),
    ).group_by(DailyPlayerStats.player_id).all())


def _resolve_known_ids(user):
    """Minimal mirror of the dashboard's known-ids logic — enough for an
    audit. Uses the user's own player_id and any ids used as sa_id /
    agent_id in player stats.
    """
    ids = {user.player_id}
    hits = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.sa_id, DailyPlayerStats.agent_id,
    ).filter(
        (DailyPlayerStats.sa_id == user.player_id) |
        (DailyPlayerStats.agent_id == user.player_id)
    ).first()
    if hits:
        ids.update(v for v in hits if v and v != '-')
    ids.discard('')
    ids.discard('-')
    return list(ids)


def audit():
    nicks = _nick_lookup()
    sa_tables = get_super_agent_tables()
    excel_sa_ids = {sa.get('sa_id') for sa in sa_tables}

    clubs_data, _ = get_members_hierarchy()
    club_id_to_name = {c['club_id']: c['name'] for c in clubs_data}

    total_issues = 0
    users = (User.query.filter_by(role='agent')
             .filter(User.player_id.isnot(None))
             .order_by(User.username).all())

    print(f'=== Hierarchy audit — {len(users)} agent users ===')
    for u in users:
        known = _resolve_known_ids(u)

        # DB truth
        db_children = set()
        for kid in known:
            for h in SAHierarchy.query.filter_by(parent_sa_id=kid).all():
                if h.child_sa_id:
                    db_children.add(h.child_sa_id)

        managed_club_names = set()
        cfgs = (SARakeConfig.query.filter_by(sa_id=u.player_id)
                .filter(SARakeConfig.managed_club_id.isnot(None)).all())
        for cfg in cfgs:
            managed_club_names.add(
                club_id_to_name.get(cfg.managed_club_id, cfg.managed_club_id))

        # What the dashboard would actually render
        rendered = get_child_sa_entries(known, managed_club_names)
        rendered_ids = {cs['sa_id'] for cs in rendered}

        # Children present in DB but missing from rendered list (should be 0)
        unreachable = db_children - rendered_ids

        # Children that exist ONLY in the DB (informational — still rendered
        # thanks to the helper, but worth knowing so the user can confirm
        # they're intentional)
        db_only = {cid for cs in rendered
                   for cid in [cs['sa_id']]
                   if cs.get('_source') == 'db_only'}

        if unreachable or db_only or managed_club_names:
            print(f'\n  {u.username} (pid={u.player_id})')
            if unreachable:
                total_issues += len(unreachable)
                for cid in sorted(unreachable):
                    print(f'    UNREACHABLE child SA: {cid} ({nicks.get(cid, "?")})')
            if db_only:
                for cid in sorted(db_only):
                    print(f'    db-only child SA (rendered via helper): '
                          f'{cid} ({nicks.get(cid, "?")})')
            if managed_club_names:
                for mc in sorted(managed_club_names):
                    print(f'    managed club: {mc}')

    print()
    if total_issues:
        print(f'❌ FAIL — {total_issues} unreachable entries')
        return 1
    print('✅ OK — every DB hierarchy entry is reachable from its owner dashboard')
    return 0


if __name__ == '__main__':
    app = create_app()
    with app.app_context():
        sys.exit(audit())
