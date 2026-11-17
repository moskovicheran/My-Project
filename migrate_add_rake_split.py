"""One-off migration: add the ring/MTT rake split columns.

Adds (idempotent, safe to re-run):
  daily_player_stats.ring_rake, .mtt_rake
  archived_player_stats.ring_rake, .mtt_rake
  rake_config.mtt_rake_percent          (NULL = no split, today's behaviour)
  player_payments.total_mtt_rake        (tournament share of a frozen snapshot)

Existing rows are backfilled with ring_rake = rake, mtt_rake = 0, so the
invariant  ring_rake + mtt_rake == rake  holds for EVERY row from day one.
Days imported before this migration therefore read as "all ring" — their
real tournament share cannot be recovered, because only the currently
active Excel file is kept on the server.

Usage (prod):
    PROD_DATABASE_URL=... python migrate_add_rake_split.py
or, from the project root with the url in .env:
    python migrate_add_rake_split.py
"""
import os
import sys


def _database_url():
    url = os.environ.get('PROD_DATABASE_URL') or os.environ.get('DATABASE_URL')
    if url:
        return url
    for path in ('.env', os.path.join(os.path.dirname(__file__), '.env')):
        if not os.path.exists(path):
            continue
        with open(path, encoding='utf-8', errors='ignore') as fh:
            for line in fh:
                for key in ('PROD_DATABASE_URL', 'DATABASE_URL'):
                    if line.startswith(key + '='):
                        return line.split('=', 1)[1].strip().strip('"').strip("'")
    return None


DDL = [
    "ALTER TABLE daily_player_stats    ADD COLUMN IF NOT EXISTS ring_rake DOUBLE PRECISION DEFAULT 0",
    "ALTER TABLE daily_player_stats    ADD COLUMN IF NOT EXISTS mtt_rake  DOUBLE PRECISION DEFAULT 0",
    "ALTER TABLE archived_player_stats ADD COLUMN IF NOT EXISTS ring_rake DOUBLE PRECISION DEFAULT 0",
    "ALTER TABLE archived_player_stats ADD COLUMN IF NOT EXISTS mtt_rake  DOUBLE PRECISION DEFAULT 0",
    "ALTER TABLE rake_config           ADD COLUMN IF NOT EXISTS mtt_rake_percent DOUBLE PRECISION",
    "ALTER TABLE player_payments      ADD COLUMN IF NOT EXISTS total_mtt_rake DOUBLE PRECISION DEFAULT 0",
]

BACKFILL = [
    ("daily_player_stats",
     "UPDATE daily_player_stats SET ring_rake = rake, mtt_rake = 0 "
     "WHERE ring_rake IS NULL OR (ring_rake = 0 AND mtt_rake = 0 AND rake <> 0)"),
    ("archived_player_stats",
     "UPDATE archived_player_stats SET ring_rake = rake, mtt_rake = 0 "
     "WHERE ring_rake IS NULL OR (ring_rake = 0 AND mtt_rake = 0 AND rake <> 0)"),
]


def main():
    import sqlalchemy as sa

    url = _database_url()
    if not url:
        print('ERROR: no PROD_DATABASE_URL / DATABASE_URL found.')
        return 1

    engine = sa.create_engine(url)
    with engine.begin() as conn:
        for stmt in DDL:
            conn.execute(sa.text(stmt))
            print('OK  ', stmt.split('ADD COLUMN IF NOT EXISTS')[0].strip(),
                  '->', stmt.split('IF NOT EXISTS')[1].strip())
        for table, stmt in BACKFILL:
            res = conn.execute(sa.text(stmt))
            print(f'BACKFILL {table}: {res.rowcount} rows set to ring_rake = rake')

    # Verify the invariant straight away.
    with engine.connect() as conn:
        for table in ('daily_player_stats', 'archived_player_stats'):
            row = conn.execute(sa.text(
                f"SELECT COALESCE(SUM(rake),0), COALESCE(SUM(ring_rake),0), "
                f"COALESCE(SUM(mtt_rake),0), COUNT(*) FROM {table}")).first()
            rake, ring, mtt, n = (float(row[0]), float(row[1]),
                                  float(row[2]), int(row[3]))
            ok = abs((ring + mtt) - rake) < 0.01
            print(f'{table}: rows={n} rake={rake:.2f} ring={ring:.2f} '
                  f'mtt={mtt:.2f} invariant={"OK" if ok else "MISMATCH"}')
            if not ok:
                return 2
    print('\nMigration complete.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
