"""Register a club as managed by a Super Agent (SARakeConfig).
Generalized version of `add_marmalades_to_mangisto.py` /
`add_mang0_to_mangisto.py`.

`--club` can be either an Excel club_id (e.g. '630307') or a literal
club name that appears in DailyPlayerStats.club (e.g. 'MANG0',
'Marmalades'). The code that consumes SARakeConfig already handles
both via `cid_to_name.get(...) or managed_club_id` fallback.

Idempotent: won't re-add a duplicate for the same (sa_id, club).
Pass --delete to remove a registration.

Usage examples (PowerShell):
  # Register a literal club name:
  $env:DATABASE_URL="<neon-url>"
  python tools_add_managed_club.py --sa-id 4406-1298 --club "Marmalades"
  Remove-Item env:DATABASE_URL

  # Register via Excel club_id:
  python tools_add_managed_club.py --sa-id 4447-3687 --club "630307"

  # Remove a registration:
  python tools_add_managed_club.py --sa-id 4406-1298 --club "MANG0" --delete
"""
import sys, io, argparse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from app import create_app
from app.models import db, SARakeConfig, DailyPlayerStats


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--sa-id', required=True)
    ap.add_argument('--club', required=True,
                    help="Excel club_id or literal club name")
    ap.add_argument('--rake-percent', type=float, default=0.0)
    ap.add_argument('--delete', action='store_true')
    args = ap.parse_args()

    app = create_app()
    with app.app_context():
        literal_rows = DailyPlayerStats.query.filter(DailyPlayerStats.club == args.club).count()
        print(f'DailyPlayerStats rows with club={args.club!r} (literal): {literal_rows}')
        existing = SARakeConfig.query.filter_by(
            sa_id=args.sa_id, managed_club_id=args.club).first()

        if args.delete:
            if existing is None:
                print(f'No SARakeConfig(sa={args.sa_id}, club={args.club!r}) — nothing to delete.')
                return
            print(f'Deleting SARakeConfig id={existing.id}')
            db.session.delete(existing)
            db.session.commit()
        else:
            if existing is not None:
                print(f'Already registered (id={existing.id}). Leaving rake_percent={existing.rake_percent}.')
            else:
                print(f'Adding SARakeConfig(sa_id={args.sa_id}, managed_club_id={args.club!r}, rake_percent={args.rake_percent})')
                db.session.add(SARakeConfig(
                    sa_id=args.sa_id, managed_club_id=args.club,
                    rake_percent=args.rake_percent))
                db.session.commit()

        print('\nManaged clubs for this SA now:')
        for c in SARakeConfig.query.filter_by(sa_id=args.sa_id).filter(
                SARakeConfig.managed_club_id.isnot(None)).all():
            print(f'  id={c.id}  {c.managed_club_id!r}  rake_pct={c.rake_percent}')


if __name__ == '__main__':
    main()
