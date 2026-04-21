"""One-shot: register 'Marmalades' as a managed club of Mangisto San
(sa_id=4406-1298) so Marmalades rows are attributed to his card.

Marmalades has no Excel club_id in the members hierarchy, so we store
the literal club name as managed_club_id (matches rows where
DailyPlayerStats.club == 'Marmalades'). Same pattern as 'Spc o' / 'SPC Un'.

Usage:
  $env:DATABASE_URL="<neon-url>"
  python add_marmalades_to_mangisto.py
  Remove-Item env:DATABASE_URL
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from app import create_app
from app.models import db, SARakeConfig, DailyPlayerStats

SA_ID = '4406-1298'       # Mangisto San
CLUB = 'Marmalades'       # literal club name as it appears in DailyPlayerStats.club
RAKE_PCT = 0              # keep-percent (adjust later via admin panel if needed)

app = create_app()
with app.app_context():
    # Sanity: do Marmalades rows actually exist with this exact spelling?
    row_count = DailyPlayerStats.query.filter(DailyPlayerStats.club == CLUB).count()
    print(f'DailyPlayerStats rows with club="{CLUB}": {row_count}')
    if row_count == 0:
        print('WARN: no rows found — check spelling before proceeding.')

    existing = SARakeConfig.query.filter_by(
        sa_id=SA_ID, managed_club_id=CLUB).first()
    if existing:
        print(f'Already registered: sa={SA_ID} club={CLUB!r} — no change.')
    else:
        print(f'Adding SARakeConfig(sa_id={SA_ID}, managed_club_id={CLUB!r}, rake_percent={RAKE_PCT})')
        db.session.add(SARakeConfig(
            sa_id=SA_ID, managed_club_id=CLUB, rake_percent=RAKE_PCT))
        db.session.commit()
        print('Done.')

    print('\nAll Mangisto managed clubs now:')
    for c in SARakeConfig.query.filter_by(sa_id=SA_ID).filter(
            SARakeConfig.managed_club_id.isnot(None)).all():
        print(f'  {c.managed_club_id!r}')
