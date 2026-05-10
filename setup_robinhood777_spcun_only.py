"""One-shot: add SARakeConfig for Robin Hood 777 to surface his own
SPC Un play under an "SPC Un" card on his agent dashboard. No deletes
— Mangisto's existing rows (including 'robin hood 777') stay as-is.

Pairs with MANAGED_CLUB_PLAYER_ONLY in app/routes/admin.py: Robin Hood
is registered there, so his SPC Un card filters to just his own
player_id (no downline), and his row is excluded from Mangisto's
"תוספת" / SPC Un card to prevent double-count.

Usage:
  $env:DATABASE_URL="<neon-url>"
  python setup_robinhood777_spcun_only.py
  Remove-Item Env:DATABASE_URL
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from app import create_app
from app.models import db, SARakeConfig, DailyPlayerStats

ROBINHOOD_SA = '3849-4104'
CLUB_NAME    = 'SPC Un'
RAKE_PCT     = 0


def _print_state(sa_id, label):
    rows = SARakeConfig.query.filter_by(sa_id=sa_id).filter(
        SARakeConfig.managed_club_id.isnot(None)).all()
    print(f'\n{label} ({sa_id}) managed clubs ({len(rows)}):')
    for r in rows:
        print(f'  id={r.id} managed_club_id={r.managed_club_id!r} rake%={r.rake_percent}')


app = create_app()
with app.app_context():
    print('=== BEFORE ===')
    _print_state(ROBINHOOD_SA, 'Robin Hood 777')

    print(f'\n=== Adding SARakeConfig under Robin Hood 777 for {CLUB_NAME!r} ===')
    row_count = DailyPlayerStats.query.filter(DailyPlayerStats.club == CLUB_NAME).count()
    print(f'  DailyPlayerStats rows for club={CLUB_NAME!r}: {row_count}')
    if row_count == 0:
        print('  WARN: no rows with this exact spelling — card will be empty.')

    existing = SARakeConfig.query.filter_by(
        sa_id=ROBINHOOD_SA, managed_club_id=CLUB_NAME).first()
    if existing:
        print(f'  Already registered (id={existing.id}) — skipping.')
    else:
        print(f'  Adding SARakeConfig(sa_id={ROBINHOOD_SA}, managed_club_id={CLUB_NAME!r}, rake_percent={RAKE_PCT})')
        db.session.add(SARakeConfig(
            sa_id=ROBINHOOD_SA, managed_club_id=CLUB_NAME, rake_percent=RAKE_PCT))
        db.session.commit()
        print('  Committed.')

    print('\n=== AFTER ===')
    _print_state(ROBINHOOD_SA, 'Robin Hood 777')
