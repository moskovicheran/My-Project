"""Claim a player's play in a HOUSE club (admin OVERVIEW_CLUBS) into one
agent's box — permanently, for every cycle and every future upload.

A house club has its own card on the admin overview, so every scope in the
app deliberately keeps house-club rows on that card: a normal
`PlayerAssignment` (tools_assign_player.py) answers "who owns this player"
and is explicitly carved OUT of house clubs. This tool writes the narrower
statement "this player's play IN THIS HOUSE CLUB belongs to that box", which
is the one thing allowed to move those rows.

The rows leave the club card and the club dashboard in the same breath as
they join the agent's box, so no total is counted twice.

Idempotent: re-running for the same (player, club) UPDATES the claim.

Usage examples (PowerShell):
  $env:DATABASE_URL="<neon-url>"
  # ScooPokeR (875349) play of F@Taxi Driver counts in Mangisto San's box:
  python tools_claim_house_club.py --player-id 6018-5756 --club-id 875349 `
      --sa-id 4406-1298 --agent-id 6018-5756 --note "..."

  # List every claim:
  python tools_claim_house_club.py --list

  # Remove a claim:
  python tools_claim_house_club.py --player-id 6018-5756 --club-id 875349 --delete
  Remove-Item env:DATABASE_URL
"""
import sys, io, argparse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from app import create_app
from app.models import db, HouseClubClaim


def _show(c):
    return (f'{c.player_id} @ club {c.club_id} -> sa={c.assigned_sa_id!r} '
            f'ag={c.assigned_agent_id!r} note={c.note!r}')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--player-id')
    ap.add_argument('--club-id', help="OVERVIEW_CLUBS club id, e.g. 875349 (ScooPokeR)")
    ap.add_argument('--sa-id', default='', help="SA whose box claims these rows")
    ap.add_argument('--agent-id', default='', help="Agent whose box claims these rows")
    ap.add_argument('--note', default='')
    ap.add_argument('--delete', action='store_true', help="Remove this claim")
    ap.add_argument('--list', action='store_true', help="List all claims and exit")
    args = ap.parse_args()

    app = create_app()
    with app.app_context():
        if args.list:
            rows = HouseClubClaim.query.order_by(HouseClubClaim.club_id,
                                                 HouseClubClaim.player_id).all()
            print(f'{len(rows)} house-club claim(s):')
            for c in rows:
                print('  ' + _show(c))
            return

        if not args.player_id or not args.club_id:
            ap.error('--player-id and --club-id are required (or use --list)')
        if not args.delete and not (args.sa_id or args.agent_id):
            ap.error('Must provide --sa-id or --agent-id, or use --delete')

        claim = HouseClubClaim.query.filter_by(
            player_id=args.player_id, club_id=args.club_id).first()

        if args.delete:
            if claim is None:
                print(f'No claim for {args.player_id} @ {args.club_id} — nothing to delete.')
                return
            print('Deleting: ' + _show(claim))
            db.session.delete(claim)
            db.session.commit()
            print('Done.')
            return

        if claim is None:
            print(f'Creating claim: {args.player_id} @ {args.club_id} -> '
                  f'sa={args.sa_id!r} ag={args.agent_id!r}')
            db.session.add(HouseClubClaim(
                player_id=args.player_id,
                club_id=args.club_id,
                assigned_sa_id=args.sa_id,
                assigned_agent_id=args.agent_id,
                note=args.note,
            ))
        else:
            print('Updating: ' + _show(claim))
            claim.assigned_sa_id = args.sa_id
            claim.assigned_agent_id = args.agent_id
            if args.note:
                claim.note = args.note
        db.session.commit()

        claim = HouseClubClaim.query.filter_by(
            player_id=args.player_id, club_id=args.club_id).first()
        print('After commit: ' + _show(claim))


if __name__ == '__main__':
    main()
