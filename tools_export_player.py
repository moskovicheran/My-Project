"""Export all data about one player to Excel — generalized version
of `export_areyoufold.py`. Breaks down rake/PnL per club, per
(SA/Agent), and per game type + blinds.

Usage (PowerShell):
  $env:DATABASE_URL="<neon-url>"
  python tools_export_player.py --player-id 1443-8481
  python tools_export_player.py --player-id 1443-8481 --out custom_name.xlsx
  Remove-Item env:DATABASE_URL

The output workbook has sheets:
  - סיכום שחקנים  /  סיכום לפי מועדון
  - פירוט לפי סוג משחק (+ בליינדס)
  - סוג משחק × דרך מי
  - סשנים (מפורט)
  - פעיל לפי העלאה
  - ארכיון / העברות כספים / שיוך ידני
"""
import sys, io, os, argparse
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from datetime import timedelta
from collections import defaultdict
import pandas as pd
from sqlalchemy import func
from app import create_app
from app.models import (DailyPlayerStats, DailyUpload, PlayerSession,
                         MoneyTransfer, PlayerAssignment, ArchivedPlayerStats,
                         ArchivedUpload)


def _nick(pid):
    if not pid or pid in ('-',):
        return ''
    return DailyPlayerStats.query.with_entities(
        func.max(DailyPlayerStats.nickname)
    ).filter(DailyPlayerStats.player_id == pid,
             DailyPlayerStats.nickname.isnot(None),
             DailyPlayerStats.nickname != '').scalar() or ''


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--player-id', required=True)
    ap.add_argument('--out', default=None,
                    help="Output xlsx path (default: player_<id>_export.xlsx)")
    args = ap.parse_args()

    PID = args.player_id
    out_path = args.out or f'player_{PID}_export.xlsx'

    app = create_app()
    with app.app_context():
        # SA/Agent nick resolution across all player's rows
        sa_ag_ids = set()
        for M in (DailyPlayerStats, ArchivedPlayerStats):
            for r in M.query.filter(M.player_id == PID).all():
                if r.sa_id and r.sa_id != '-': sa_ag_ids.add(r.sa_id)
                if r.agent_id and r.agent_id != '-': sa_ag_ids.add(r.agent_id)
        nick_map = {pid: (_nick(pid) or pid) for pid in sa_ag_ids}

        def label(pid):
            if not pid or pid == '-': return '(ללא)'
            return f'{nick_map.get(pid, pid)} ({pid})'

        # 1) Active rows per upload
        rows = (DailyPlayerStats.query
                .join(DailyUpload, DailyPlayerStats.upload_id == DailyUpload.id)
                .add_columns(DailyUpload.upload_date)
                .filter(DailyPlayerStats.player_id == PID)
                .order_by(DailyUpload.upload_date.asc()).all())
        active_df = pd.DataFrame([{
            'תאריך': d.strftime('%Y-%m-%d') if d else '',
            'upload_id': r.upload_id, 'מועדון': r.club,
            'כינוי': r.nickname, 'תפקיד': r.role,
            'דרך SA': label(r.sa_id), 'דרך Agent': label(r.agent_id),
            'Rake': round(float(r.rake or 0), 2),
            'PnL': round(float(r.pnl or 0), 2),
        } for r, d in rows])

        # 2) Through-who breakdown (per club × SA × Agent)
        by_through = defaultdict(lambda: {'rake': 0.0, 'pnl': 0.0})
        up2ctx = {}
        for r in DailyPlayerStats.query.filter(
                DailyPlayerStats.player_id == PID,
                DailyPlayerStats.role != 'Name Entry').all():
            up2ctx[r.upload_id] = {'sa': r.sa_id or '', 'ag': r.agent_id or '',
                                    'club': r.club or ''}
            k = (r.club or '(ללא)', label(r.sa_id), label(r.agent_id))
            by_through[k]['rake'] += float(r.rake or 0)
            by_through[k]['pnl'] += float(r.pnl or 0)
        through_df = pd.DataFrame([{
            'מועדון': c, 'דרך SA': s, 'דרך Agent': a,
            'Rake': round(v['rake'], 2), 'PnL': round(v['pnl'], 2),
        } for (c, s, a), v in sorted(
            by_through.items(), key=lambda kv: -abs(kv[1]['rake']) - abs(kv[1]['pnl']))])

        # 3) Sessions enriched with context
        sess = (PlayerSession.query
                .join(DailyUpload, PlayerSession.upload_id == DailyUpload.id)
                .add_columns(DailyUpload.upload_date)
                .filter(PlayerSession.player_id == PID)
                .order_by(DailyUpload.upload_date.asc()).all())
        sess_rows = []
        for s, d in sess:
            ctx = up2ctx.get(s.upload_id, {})
            sess_rows.append({
                'תאריך': d.strftime('%Y-%m-%d') if d else '',
                'סוג משחק': s.game_type or '',
                'בליינדס': s.blinds or '',
                'שם שולחן': s.table_name or '',
                'מועדון': ctx.get('club', ''),
                'דרך SA': label(ctx.get('sa', '')),
                'דרך Agent': label(ctx.get('ag', '')),
                'PnL': round(float(s.pnl or 0), 2),
            })
        sess_df = pd.DataFrame(sess_rows)

        # 4) Game type summary + game type × blinds + game type × through
        def _summ(key_fn):
            d = defaultdict(lambda: {'sessions': 0, 'pnl': 0.0,
                                      'wins': 0, 'losses': 0})
            for row in sess_rows:
                k = key_fn(row)
                e = d[k]
                e['sessions'] += 1
                e['pnl'] += row['PnL']
                if row['PnL'] > 0: e['wins'] += 1
                elif row['PnL'] < 0: e['losses'] += 1
            return d

        gt = _summ(lambda r: r['סוג משחק'] or 'Other')
        gt_df = pd.DataFrame([{
            'סוג משחק': k, 'סשנים': v['sessions'],
            'זכיות': v['wins'], 'הפסדים': v['losses'],
            'PnL': round(v['pnl'], 2),
        } for k, v in sorted(gt.items(), key=lambda kv: -abs(kv[1]['pnl']))])

        gtb = _summ(lambda r: (r['סוג משחק'] or 'Other', r['בליינדס'] or '-'))
        gtb_df = pd.DataFrame([{
            'סוג משחק': k[0], 'בליינדס': k[1], 'סשנים': v['sessions'],
            'זכיות': v['wins'], 'הפסדים': v['losses'],
            'PnL': round(v['pnl'], 2),
        } for k, v in sorted(gtb.items(), key=lambda kv: -abs(kv[1]['pnl']))])

        gtth = _summ(lambda r: (r['סוג משחק'] or 'Other',
                                r['מועדון'] or '(ללא)',
                                r['דרך SA'], r['דרך Agent']))
        gtth_df = pd.DataFrame([{
            'סוג משחק': k[0], 'מועדון': k[1], 'דרך SA': k[2], 'דרך Agent': k[3],
            'סשנים': v['sessions'], 'PnL': round(v['pnl'], 2),
        } for k, v in sorted(gtth.items(), key=lambda kv: -abs(kv[1]['pnl']))])

        # 5) Transfers
        xfers = []
        for t in MoneyTransfer.query.filter_by(from_player_id=PID).all():
            il = (t.created_at + timedelta(hours=3)) if t.created_at else None
            xfers.append({
                'תאריך': il.strftime('%Y-%m-%d %H:%M') if il else '',
                'כיוון': 'החוצה',
                'צד שני': f'{t.to_name} ({t.to_player_id})',
                'סכום (לשחקן)': round(-float(t.amount or 0), 2),
                'תיאור': t.description or '',
            })
        for t in MoneyTransfer.query.filter_by(to_player_id=PID).all():
            il = (t.created_at + timedelta(hours=3)) if t.created_at else None
            xfers.append({
                'תאריך': il.strftime('%Y-%m-%d %H:%M') if il else '',
                'כיוון': 'פנימה',
                'צד שני': f'{t.from_name} ({t.from_player_id})',
                'סכום (לשחקן)': round(float(t.amount or 0), 2),
                'תיאור': t.description or '',
            })
        xfers_df = pd.DataFrame(sorted(xfers, key=lambda r: r['תאריך']))

        # 6) PlayerAssignment
        pa = PlayerAssignment.query.filter_by(player_id=PID).first()
        pa_df = pd.DataFrame([{
            'player_id': PID,
            'assigned_sa': label(pa.assigned_sa_id) if pa else '(אין)',
            'assigned_agent': label(pa.assigned_agent_id) if pa else '',
            'note': pa.note if pa else '',
            'updated_at': pa.updated_at.strftime('%Y-%m-%d %H:%M')
                           if (pa and pa.updated_at) else '',
        }])

        def _w(wr, df, name):
            (df if not df.empty else pd.DataFrame([{'info': 'אין נתונים'}])
             ).to_excel(wr, sheet_name=name, index=False)

        with pd.ExcelWriter(out_path, engine='openpyxl') as wr:
            _w(wr, gt_df,       'סיכום לפי סוג משחק')
            _w(wr, gtb_df,      'סוג משחק × בליינדס')
            _w(wr, gtth_df,     'סוג משחק × דרך מי')
            _w(wr, through_df,  'סיכום דרך מי (לפי מועדון)')
            _w(wr, sess_df,     'סשנים (מפורט)')
            _w(wr, active_df,   'פעיל לפי העלאה')
            _w(wr, xfers_df,    'העברות כספים')
            _w(wr, pa_df,       'שיוך ידני')

        print(f'Written: {os.path.abspath(out_path)}')
        print(f'Sessions: {len(sess_rows)}  Active rows: {len(rows)}')


if __name__ == '__main__':
    main()
