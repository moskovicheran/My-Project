"""Verification tool for the ring/MTT rake split.

Answers "is the split correct?" with four independent checks:

  1. INVARIANT   ring_rake + mtt_rake == rake, per row and in total.
                 If this ever breaks, the import is wrong — nothing else
                 matters.
  2. CROSS-CHECK total mtt_rake  vs  SUM(fee x entries) from tournament_stats.
                 Two completely different sheets of the same Excel file; they
                 must agree for every day that was imported WITH the split.
  3. COVERAGE    which upload days carry a split and which are pre-feature
                 (mtt_rake = 0 on every row of that day).
  4. ENTITY      for one entity (default: Mangisto San), the refund under the
                 old single-% rule vs the split rule, per sub-agent.

Usage:
    python diag_rake_split.py                  # all checks, Mangisto San
    python diag_rake_split.py 4406-1298        # another SA id
"""
import os
import sys

MANGISTO = '4406-1298'


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


def main():
    import sqlalchemy as sa

    sa_id = sys.argv[1] if len(sys.argv) > 1 else MANGISTO
    url = _database_url()
    if not url:
        print('ERROR: no PROD_DATABASE_URL / DATABASE_URL found.')
        return 1
    engine = sa.create_engine(url)
    failures = []

    with engine.connect() as c:
        # ---------- 1. invariant ----------
        print('=' * 72)
        print('1. INVARIANT  ring_rake + mtt_rake == rake')
        print('=' * 72)
        for table in ('daily_player_stats', 'archived_player_stats'):
            row = c.execute(sa.text(
                f"SELECT COALESCE(SUM(rake),0), COALESCE(SUM(ring_rake),0), "
                f"COALESCE(SUM(mtt_rake),0), COUNT(*) FROM {table}")).first()
            rake, ring, mtt, n = float(row[0]), float(row[1]), float(row[2]), int(row[3])
            bad = c.execute(sa.text(
                f"SELECT COUNT(*) FROM {table} "
                f"WHERE ABS(COALESCE(ring_rake,0) + COALESCE(mtt_rake,0) - rake) > 0.01"
            )).first()[0]
            ok = bad == 0 and abs((ring + mtt) - rake) < 0.01
            print(f'  {table:24s} rows={n:<7d} rake={rake:12,.2f} '
                  f'ring={ring:12,.2f} mtt={mtt:10,.2f}  bad_rows={bad}  '
                  f'{"OK" if ok else "FAIL"}')
            if not ok:
                failures.append(f'invariant broken in {table} ({bad} rows)')

        # ---------- 2. cross-check against the tournament sheet ----------
        print()
        print('=' * 72)
        print('2. CROSS-CHECK  SUM(mtt_rake)  vs  SUM(fee x entries)')
        print('=' * 72)
        per_day = c.execute(sa.text(
            "SELECT u.id, u.upload_date, "
            "       COALESCE((SELECT SUM(s.mtt_rake) FROM daily_player_stats s "
            "                 WHERE s.upload_id = u.id), 0) AS mtt, "
            "       COALESCE((SELECT SUM(t.fee * t.entries) FROM tournament_stats t "
            "                 WHERE t.upload_id = u.id), 0) AS fees, "
            "       COALESCE((SELECT SUM(s.rake) FROM daily_player_stats s "
            "                 WHERE s.upload_id = u.id), 0) AS rake "
            "FROM daily_uploads u ORDER BY u.upload_date")).all()
        print(f'  {"date":12s} {"rake":>12s} {"mtt_rake":>11s} {"fee*entries":>12s} '
              f'{"diff":>9s}  status')
        for uid, d, mtt, fees, rake in per_day:
            mtt, fees, rake = float(mtt), float(fees), float(rake)
            if mtt == 0 and rake != 0:
                status = 'pre-split day (no data)'
                diff = 0.0
            else:
                diff = mtt - fees
                status = 'OK' if abs(diff) < 0.5 else 'MISMATCH'
                if status == 'MISMATCH':
                    failures.append(f'{d}: mtt_rake {mtt:.2f} != fee*entries {fees:.2f}')
            print(f'  {str(d):12s} {rake:12,.2f} {mtt:11,.2f} {fees:12,.2f} '
                  f'{diff:9,.2f}  {status}')

        # ---------- 3. coverage ----------
        print()
        print('=' * 72)
        print('3. COVERAGE  which days carry a split')
        print('=' * 72)
        split_days = [str(d) for _, d, mtt, _, rake in per_day if float(mtt) > 0]
        plain_days = [str(d) for _, d, mtt, _, rake in per_day
                      if float(mtt) == 0 and float(rake) != 0]
        print(f'  with split : {len(split_days)} day(s) '
              f'{"  " + ", ".join(split_days) if split_days else ""}')
        print(f'  all-ring   : {len(plain_days)} day(s) '
              f'{"  " + ", ".join(plain_days) if plain_days else ""}')
        if plain_days:
            print('  NOTE: all-ring days are counted entirely as ring rake. Their real')
            print('        tournament share cannot be recovered (only the active Excel')
            print('        file is stored), so cycle totals mixing both are partial.')

        # ---------- 4. one entity, old rule vs split rule ----------
        print()
        print('=' * 72)
        print(f'4. ENTITY  {sa_id} — refund today vs with the split')
        print('=' * 72)
        cfgs = {r[0]: (float(r[1] or 0), None if r[2] is None else float(r[2]))
                for r in c.execute(sa.text(
                    "SELECT entity_id, rake_percent, mtt_rake_percent FROM rake_config "
                    "WHERE entity_type IN ('agent','sub_agent')")).all()}
        rows = c.execute(sa.text(
            "SELECT agent_id, COALESCE(SUM(rake),0), COALESCE(SUM(mtt_rake),0), "
            "       COUNT(DISTINCT player_id) "
            "FROM daily_player_stats WHERE sa_id = :s AND agent_id <> '' "
            "GROUP BY agent_id ORDER BY SUM(rake) DESC"), {'s': sa_id}).all()
        if not rows:
            print('  no sub-agent rows for this SA in the active cycle.')
        print(f'  {"agent":12s} {"pl":>3s} {"rake":>11s} {"mtt":>9s} {"ring%":>6s} '
              f'{"mtt%":>6s} {"refund now":>11s} {"refund split":>13s} {"diff":>9s}')
        t_old = t_new = 0.0
        for agent_id, rake, mtt, n in rows:
            rake, mtt = float(rake), float(mtt)
            pct, mtt_pct = cfgs.get(agent_id, (0.0, None))
            old = rake * pct / 100
            if mtt_pct is None:
                new = old
            else:
                m = max(0.0, min(mtt, rake))
                new = (rake - m) * pct / 100 + m * mtt_pct / 100
            t_old += old
            t_new += new
            print(f'  {agent_id:12s} {n:3d} {rake:11,.2f} {mtt:9,.2f} {pct:6.1f} '
                  f'{("-" if mtt_pct is None else f"{mtt_pct:.1f}"):>6s} '
                  f'{old:11,.2f} {new:13,.2f} {old - new:9,.2f}')
        print(f'  {"TOTAL":12s} {"":3s} {"":11s} {"":9s} {"":6s} {"":6s} '
              f'{t_old:11,.2f} {t_new:13,.2f} {t_old - t_new:9,.2f}')

    print()
    print('=' * 72)
    if failures:
        print('RESULT: PROBLEMS FOUND')
        for f in failures:
            print('  - ' + f)
        return 2
    print('RESULT: all checks passed')
    return 0


if __name__ == '__main__':
    sys.exit(main())
