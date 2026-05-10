"""One-shot: remove the 'robin hood 777' managed-club entry from
Mangisto San (sa_id=4406-1298) and register 'SPC Un' + 'SPC T' as
managed clubs of Robin Hood 777's own SA (sa_id=3849-4104), so his
play in those clubs surfaces under "תוספות" on his agent dashboard.

Pattern: same as add_marmalades_to_mangisto.py / add_mang0_to_mangisto.py
— literal club name stored as managed_club_id, matches rows where
DailyPlayerStats.club == <name>.

The matching code change in app/routes/main.py opts Robin Hood 777
into the "include SA in his own managed-club card" path
(MANAGED_CLUB_SHOW_SELF), so his own SPC Un rows appear inside the
"תוספת / SPC Un" card instead of being rolled into "שחקנים ישירים".

Usage:
  $env:DATABASE_URL="<neon-url>"
  python setup_robinhood777_managed_clubs.py
  Remove-Item env:DATABASE_URL
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from app import create_app
from app.models import db, SARakeConfig, DailyPlayerStats

MANGISTO_SA   = '4406-1298'
ROBINHOOD_SA  = '3849-4104'

# Canonical casings — match how they appear in DailyPlayerStats.club
# and in admin.py (MANAGED_CLUB_DISPLAY_NAMES uses 'SPC Un').
ROBINHOOD_CLUB = 'robin hood 777'
ADD_TO_ROBINHOOD = ['SPC Un', 'SPC T']
RAKE_PCT = 0


def _print_state(sa_id, label):
    rows = SARakeConfig.query.filter_by(sa_id=sa_id).filter(
        SARakeConfig.managed_club_id.isnot(None)).all()
    print(f'\n{label} ({sa_id}) managed clubs ({len(rows)}):')
    for r in rows:
        print(f'  id={r.id} managed_club_id={r.managed_club_id!r} rake%={r.rake_percent}')


app = create_app()
with app.app_context():
    print('=== BEFORE ===')
    _print_state(MANGISTO_SA, 'Mangisto San')
    _print_state(ROBINHOOD_SA, 'Robin Hood 777')

    # 1. Remove "robin hood 777" entry from Mangisto San (case-insensitive
    #    match — could have been stored as "robin hood 777", "Robin Hood 777",
    #    etc.).
    print('\n=== STEP 1: remove "robin hood 777" entry from Mangisto San ===')
    bad = [c for c in SARakeConfig.query.filter_by(sa_id=MANGISTO_SA).all()
           if (c.managed_club_id or '').strip().lower() == ROBINHOOD_CLUB.lower()]
    if not bad:
        print(f'  No SARakeConfig found under sa={MANGISTO_SA} matching {ROBINHOOD_CLUB!r} — nothing to remove.')
    else:
        for c in bad:
            print(f'  Deleting id={c.id} managed_club_id={c.managed_club_id!r}')
            db.session.delete(c)

    # 2. Ensure Robin Hood 777 has his three managed-club entries.
    print('\n=== STEP 2: add managed clubs under Robin Hood 777 ===')
    for cname in ADD_TO_ROBINHOOD:
        # Sanity: confirm rows actually exist with this exact spelling so
        # the new card won't render as empty.
        row_count = DailyPlayerStats.query.filter(DailyPlayerStats.club == cname).count()
        print(f'\n  club={cname!r}  DailyPlayerStats rows: {row_count}')
        if row_count == 0:
            print('    WARN: no rows found with this exact spelling — card will be empty.')

        existing = SARakeConfig.query.filter_by(
            sa_id=ROBINHOOD_SA, managed_club_id=cname).first()
        if existing:
            print(f'    Already registered under sa={ROBINHOOD_SA} (id={existing.id}) — skipping.')
        else:
            print(f'    Adding SARakeConfig(sa_id={ROBINHOOD_SA}, managed_club_id={cname!r}, rake_percent={RAKE_PCT})')
            db.session.add(SARakeConfig(
                sa_id=ROBINHOOD_SA, managed_club_id=cname, rake_percent=RAKE_PCT))

    db.session.commit()
    print('\nCommitted.')

    print('\n=== AFTER ===')
    _print_state(MANGISTO_SA, 'Mangisto San')
    _print_state(ROBINHOOD_SA, 'Robin Hood 777')
