"""Assign or move a player to a specific SA (or sub-agent) via
PlayerAssignment — generalized version of `move_areyoufold_to_niroha.py`.

Idempotent: if the player already has a PlayerAssignment it gets
UPDATED (not duplicated). Pass --delete to remove an override.

Usage examples (PowerShell):
  # Attach areyoufold to niroha27:
  $env:DATABASE_URL="<neon-url>"
  python tools_assign_player.py --player-id 1443-8481 --sa-id 8040-6815 --note "SPC T under Dolar 10"
  Remove-Item env:DATABASE_URL

  # Attach to a specific agent:
  python tools_assign_player.py --player-id 1234-5678 --agent-id 6670-6318

  # Remove override:
  python tools_assign_player.py --player-id 1443-8481 --delete
"""
import sys, io, argparse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from app import create_app
from app.models import db, PlayerAssignment


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--player-id', required=True)
    ap.add_argument('--sa-id', default='',
                    help="SA to assign the player to (leave empty when using --agent-id)")
    ap.add_argument('--agent-id', default='',
                    help="Agent to assign the player to (optional)")
    ap.add_argument('--note', default='')
    ap.add_argument('--delete', action='store_true',
                    help="Remove the PlayerAssignment for this player")
    args = ap.parse_args()

    if not args.delete and not (args.sa_id or args.agent_id):
        ap.error('Must provide --sa-id or --agent-id, or use --delete')

    app = create_app()
    with app.app_context():
        pa = PlayerAssignment.query.filter_by(player_id=args.player_id).first()
        if args.delete:
            if pa is None:
                print(f'No PlayerAssignment for {args.player_id} — nothing to delete.')
                return
            print(f'Deleting PlayerAssignment: {args.player_id} → sa={pa.assigned_sa_id!r} ag={pa.assigned_agent_id!r}')
            db.session.delete(pa)
            db.session.commit()
            print('Done.')
            return

        if pa is None:
            print(f'Creating new PlayerAssignment: {args.player_id} → sa={args.sa_id!r} ag={args.agent_id!r}')
            db.session.add(PlayerAssignment(
                player_id=args.player_id,
                assigned_sa_id=args.sa_id,
                assigned_agent_id=args.agent_id,
                note=args.note,
            ))
        else:
            print(f'Updating PlayerAssignment {args.player_id}:')
            print(f'  sa: {pa.assigned_sa_id!r} → {args.sa_id!r}')
            print(f'  ag: {pa.assigned_agent_id!r} → {args.agent_id!r}')
            pa.assigned_sa_id = args.sa_id
            pa.assigned_agent_id = args.agent_id
            if args.note:
                pa.note = args.note
        db.session.commit()

        pa = PlayerAssignment.query.filter_by(player_id=args.player_id).first()
        print(f'\nAfter commit: sa={pa.assigned_sa_id!r} ag={pa.assigned_agent_id!r} note={pa.note!r}')


if __name__ == '__main__':
    main()
