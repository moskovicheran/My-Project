# -*- coding: utf-8 -*-
"""Club 596339 was renamed by PPPoker: "SPC TP" -> "SPC PK".

The app matches a managed club to its rows BY NAME
(`DailyPlayerStats.club == club_name`), and the name comes from the uploaded
file. So the moment a file arrives calling 596339 "SPC PK", every row already
stored under "SPC TP" stops matching: the SA's card loses the whole cycle and
the 85% club rake config stops being found.

This renames the literal, everywhere it is stored:

  rake_config.entity_name        'SPC TP' -> 'SPC PK'   (85% would be lost:
                                 the managed-club lookup matches entity_id OR
                                 entity_name against managed_club_id, and
                                 entity_id is '596339', not the name)
  sa_rake_config.managed_club_id 'SPC TP' -> 'SPC PK'
  daily_player_stats.club        'SPC TP' -> 'SPC PK'   (current cycle)
  archived_player_stats.club     'SPC TP' -> 'SPC PK'   (closed cycles)
  player_payments.club           'SPC TP' -> 'SPC PK'   (collection records)

Run AFTER the first file that uses the new name is uploaded, so the active
Excel and the DB agree. Default is a DRY RUN; pass --apply to write.

    python normalize_spc_pk.py            # count only, writes nothing
    python normalize_spc_pk.py --apply    # perform the rename

Same shape as normalize_gorilla.py (GORILLA ISRAELu, 2026-09).
"""
import os
import sys

OLD = 'SPC TP'
NEW = 'SPC PK'
CLUB_ID = '596339'

TARGETS = [
    ('rake_config', 'entity_name'),
    ('sa_rake_config', 'managed_club_id'),
    ('daily_player_stats', 'club'),
    ('archived_player_stats', 'club'),
    ('player_payments', 'club'),
]

HERE = os.path.abspath(os.path.dirname(__file__))
os.chdir(HERE)
sys.path.insert(0, HERE)

APPLY = '--apply' in sys.argv

url = None
for line in open(r'C:\Project\.env', encoding='utf-8', errors='ignore'):
    if line.startswith('PROD_DATABASE_URL='):
        url = line.split('=', 1)[1].strip().strip('"').strip("'")
if not url:
    sys.exit(r'PROD_DATABASE_URL not found in C:\Project\.env')
if url.startswith('postgres://'):
    url = url.replace('postgres://', 'postgresql://', 1)
os.environ['DATABASE_URL'] = url
os.environ.pop('DATABASE_URL_POOLER', None)
os.environ['SECRET_KEY'] = 'normalize-only'

from app import create_app                      # noqa: E402
from app.models import db                       # noqa: E402
from sqlalchemy import text                     # noqa: E402


def p(s):
    sys.stdout.buffer.write((str(s) + '\n').encode('utf-8'))


app = create_app()
with app.app_context():
    p('club %s:  %r -> %r     mode: %s'
      % (CLUB_ID, OLD, NEW, 'APPLY' if APPLY else 'DRY RUN'))

    # What the file now calls this club — a guard against renaming the DB to
    # a name the uploads do not actually use.
    live = db.session.execute(text(
        "SELECT DISTINCT s.club FROM daily_player_stats s "
        "JOIN daily_uploads d ON d.id = s.upload_id "
        "WHERE d.upload_date = (SELECT MAX(upload_date) FROM daily_uploads) "
        "AND s.club IN (:old, :new)"), {'old': OLD, 'new': NEW}).scalars().all()
    p('newest upload calls it: %s' % (live or '(absent from the newest file)'))

    total = 0
    for table, col in TARGETS:
        n_old = db.session.execute(text(
            'SELECT COUNT(*) FROM %s WHERE %s = :v' % (table, col)),
            {'v': OLD}).scalar()
        n_new = db.session.execute(text(
            'SELECT COUNT(*) FROM %s WHERE %s = :v' % (table, col)),
            {'v': NEW}).scalar()
        p('  %-24s %-16s %r=%-6d %r=%d' % (table, col, OLD, n_old, NEW, n_new))
        total += n_old
        if APPLY and n_old:
            db.session.execute(text(
                'UPDATE %s SET %s = :new WHERE %s = :old' % (table, col, col)),
                {'new': NEW, 'old': OLD})

    if not APPLY:
        p('\nDRY RUN — %d rows would be renamed. Re-run with --apply.' % total)
        sys.exit(0)

    db.session.commit()
    p('\ncommitted. verifying:')
    left = 0
    for table, col in TARGETS:
        n_old = db.session.execute(text(
            'SELECT COUNT(*) FROM %s WHERE %s = :v' % (table, col)),
            {'v': OLD}).scalar()
        n_new = db.session.execute(text(
            'SELECT COUNT(*) FROM %s WHERE %s = :v' % (table, col)),
            {'v': NEW}).scalar()
        p('  %-24s %-16s %r=%-6d %r=%d' % (table, col, OLD, n_old, NEW, n_new))
        left += n_old
    cfg = db.session.execute(text(
        "SELECT entity_id, entity_name, rake_percent, mtt_rake_percent "
        "FROM rake_config WHERE entity_type='club' AND entity_id=:c"),
        {'c': CLUB_ID}).first()
    p('  rake_config for %s: %s' % (CLUB_ID, tuple(cfg) if cfg else None))
    p('\n%s  (%d rows renamed, %d stragglers)'
      % ('OK' if left == 0 else 'CHECK THE STRAGGLERS', total, left))
