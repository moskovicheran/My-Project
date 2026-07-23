import os
import io
from datetime import date
from flask import Blueprint, render_template, request, redirect, url_for, flash, session
from flask_login import login_required, current_user
from werkzeug.utils import secure_filename

upload_bp = Blueprint('upload', __name__, url_prefix='/upload')

UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'uploads')
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
ACTIVE_FILE_PATH = os.path.join(UPLOAD_FOLDER, '_active.txt')


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def _parse_and_store_stats_from_bytes(file_bytes, filename):
    """Parse player stats from Excel bytes and store as cumulative daily data."""
    import pandas as pd
    from app.models import db, DailyUpload, DailyPlayerStats
    from sqlalchemy import text

    try:
        sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None)
    except Exception:
        return 0

    if 'Union Member Statistics' not in sheets:
        return 0

    df = sheets['Union Member Statistics']
    # Hands Total is the last column; layout shifted from 152→156 cols on 2026-04-16
    hands_col = df.shape[1] - 1

    # Try to extract date from Excel (Union Overview period)
    upload_date = date.today()
    try:
        if 'Union Overview' in sheets:
            period_str = str(sheets['Union Overview'].iloc[2, 0])  # "Period : 2026-03-31 ~ ..."
            date_part = period_str.replace('Period : ', '').strip().split(' ')[0].split('~')[0].strip()
            if len(date_part) == 10:
                from datetime import datetime
                upload_date = datetime.strptime(date_part, '%Y-%m-%d').date()
    except Exception:
        pass

    upload = DailyUpload(filename=filename, upload_date=upload_date)
    db.session.add(upload)
    db.session.flush()
    upload_id = upload.id

    # Collect all rows first
    rows = []
    sa_agent_names = {}  # id -> nickname for SAs and Agents
    current_club = ''
    for i in range(6, len(df)):
        row = df.iloc[i]
        if '(ID:' in str(row.iloc[0]):
            current_club = str(row.iloc[0]).split(' (ID:')[0]
        nickname = str(row.iloc[9])
        if nickname in ('nan', '-'):
            continue
        player_id = str(row.iloc[8])
        try:
            pnl = float(row.iloc[37]) if str(row.iloc[37]) != 'nan' else 0
        except (ValueError, TypeError):
            pnl = 0
        try:
            rake = float(row.iloc[64]) if str(row.iloc[64]) != 'nan' else 0
        except (ValueError, TypeError):
            rake = 0
        try:
            hands = float(row.iloc[hands_col]) if str(row.iloc[hands_col]) != 'nan' else 0
        except (ValueError, TypeError):
            hands = 0

        sa_id_val = str(row.iloc[2]) if str(row.iloc[2]) != 'nan' else ''
        sa_nick_val = str(row.iloc[3]) if str(row.iloc[3]) != 'nan' else ''
        agent_id_val = str(row.iloc[4]) if str(row.iloc[4]) != 'nan' else ''
        agent_nick_val = str(row.iloc[5]) if str(row.iloc[5]) != 'nan' else ''
        role_val = str(row.iloc[7]) if str(row.iloc[7]) != 'nan' else ''
        # Track SA/Agent names
        if sa_id_val and sa_id_val != '-' and sa_nick_val and sa_nick_val != '-':
            sa_agent_names[sa_id_val] = sa_nick_val
        if agent_id_val and agent_id_val != '-' and agent_nick_val and agent_nick_val != '-':
            sa_agent_names[agent_id_val] = agent_nick_val
        rows.append({
            'upload_id': upload_id, 'player_id': player_id,
            'nickname': nickname, 'club': current_club,
            'sa_id': sa_id_val, 'agent_id': agent_id_val, 'role': role_val,
            'pnl': round(pnl, 2), 'rake': round(rake, 2),
            'hands': round(hands, 0),
        })

    # Add SA/Agent name entries (0 stats, just for name resolution)
    existing_pids = {r['player_id'] for r in rows}
    for eid, enick in sa_agent_names.items():
        if eid not in existing_pids:
            rows.append({
                'upload_id': upload_id, 'player_id': eid,
                'nickname': enick, 'club': '',
                'sa_id': '', 'agent_id': '', 'role': 'Name Entry',
                'pnl': 0, 'rake': 0, 'hands': 0,
            })

    # Bulk insert all at once (much faster for cloud DB)
    if rows:
        db.session.execute(
            DailyPlayerStats.__table__.insert(),
            rows
        )

    # Parse and store game sessions (Ring + MTT)
    from app.models import PlayerSession
    sessions = []

    # Ring Game sessions
    if 'Union Ring Game Detail' in sheets:
        rdf = sheets['Union Ring Game Detail']
        current_table = ''
        current_game = ''
        current_blinds = ''
        for i in range(len(rdf)):
            col0 = str(rdf.iloc[i, 0])
            if col0.startswith('Table Name :'):
                try:
                    current_table = col0.split('Table Name : ')[1].split(' , Creator')[0].strip()
                except Exception:
                    current_table = col0
            elif col0.startswith('Table Information :'):
                try:
                    current_game = col0.split('Game : ')[1].split(' ,')[0].strip()
                    current_blinds = col0.split('Blinds : ')[1].split(' ,')[0].strip()
                except Exception:
                    pass
            else:
                pid = str(rdf.iloc[i, 2]) if rdf.shape[1] > 2 else ''
                if '-' in pid and len(pid) == 9:
                    try:
                        pnl_val = float(rdf.iloc[i, 13]) if rdf.shape[1] > 13 and str(rdf.iloc[i, 13]) != 'nan' else 0
                    except (ValueError, TypeError):
                        pnl_val = 0
                    sessions.append({
                        'upload_id': upload_id, 'player_id': pid,
                        'game_type': current_game or 'Ring', 'table_name': current_table,
                        'blinds': current_blinds, 'pnl': round(pnl_val, 2),
                    })

    # MTT sessions
    if 'Union MTT Detail' in sheets:
        mdf = sheets['Union MTT Detail']
        current_tournament = ''
        for i in range(len(mdf)):
            col0 = str(mdf.iloc[i, 0])
            if col0.startswith('Table Name :'):
                try:
                    current_tournament = col0.split('Table Name : ')[1].split(' , Creator')[0].strip()
                except Exception:
                    current_tournament = col0
            pid = str(mdf.iloc[i, 2]) if mdf.shape[1] > 2 else ''
            if '-' in pid and len(pid) == 9:
                try:
                    pnl_val = float(mdf.iloc[i, 16]) if mdf.shape[1] > 16 and str(mdf.iloc[i, 16]) != 'nan' else 0
                except (ValueError, TypeError):
                    pnl_val = 0
                sessions.append({
                    'upload_id': upload_id, 'player_id': pid,
                    'game_type': 'MTT', 'table_name': current_tournament[:200],
                    'blinds': '', 'pnl': round(pnl_val, 2),
                })

    if sessions:
        db.session.execute(PlayerSession.__table__.insert(), sessions)

    # Parse MTT Statistics (tournament list)
    from app.models import TournamentStats
    if 'Union MTT Statistics' in sheets:
        mtt_df = sheets['Union MTT Statistics']
        tournaments = []
        for i in range(6, len(mtt_df)):
            row = mtt_df.iloc[i]
            title = str(row.iloc[2])
            if title in ('nan', 'TOTAL'):
                continue
            club_name = str(row.iloc[1])
            if club_name in ('nan', 'TOTAL'):
                continue
            try:
                buyin = float(row.iloc[8]) if str(row.iloc[8]) != 'nan' else 0
                fee = float(row.iloc[9]) if str(row.iloc[9]) != 'nan' else 0
                entries = float(row.iloc[17]) if str(row.iloc[17]) != 'nan' else 0
                gtd = float(row.iloc[11]) if str(row.iloc[11]) != 'nan' else 0
                prize_pool = float(row.iloc[26]) if str(row.iloc[26]) != 'nan' else 0
            except (ValueError, TypeError):
                continue
            game_type = str(row.iloc[4]) if str(row.iloc[4]) != 'nan' else ''
            # col 3 = Status ('Ended' / 'In Progress')
            status = str(row.iloc[3]) if str(row.iloc[3]) != 'nan' else ''
            reentry = str(row.iloc[10]) if str(row.iloc[10]) != 'nan' else ''
            start = str(row.iloc[14])[:16] if str(row.iloc[14]) != 'nan' else ''
            duration = str(row.iloc[16]) if str(row.iloc[16]) != 'nan' else ''
            tournaments.append({
                'upload_id': upload_id, 'title': title[:200],
                'status': status[:30],
                'game_type': game_type, 'buyin': round(buyin, 2),
                'fee': round(fee, 2), 'reentry': reentry,
                'gtd': round(gtd, 2), 'entries': round(entries, 0),
                'prize_pool': round(prize_pool, 2), 'start': start,
                'duration': duration,
            })
        if tournaments:
            db.session.execute(TournamentStats.__table__.insert(), tournaments)

    db.session.commit()

    return len(rows)


@upload_bp.route('/', methods=['GET', 'POST'])
@login_required
def index():
    if current_user.role != 'admin':
        flash('אין לך הרשאה לדף זה.', 'danger')
        return redirect(url_for('main.dashboard'))
    from app.models import DailyUpload

    if request.method == 'POST':
        # Accept either single ('file') or multiple ('file' multiple) — same
        # field name keeps backwards compatibility with the old form.
        files = request.files.getlist('file')
        files = [f for f in files if f and f.filename]
        if not files:
            flash('לא נבחר קובץ.', 'danger')
            return redirect(request.url)

        from app.models import db as _db
        import pandas as pd
        from datetime import datetime as _dt

        # Pre-read each file: validate type, peek the Period date so we can
        # process oldest-first (matches normal upload chronology so any
        # duplicate-period block fires deterministically).
        prepared = []
        for f in files:
            if not allowed_file(f.filename):
                flash(f'סוג קובץ לא נתמך: {f.filename}', 'danger')
                continue
            fname = secure_filename(f.filename)
            fbytes = f.read()
            try:
                overview = pd.read_excel(io.BytesIO(fbytes), sheet_name='Union Overview', header=None)
                period_str = str(overview.iloc[2, 0])
                date_part = period_str.replace('Period : ', '').strip().split(' ')[0].split('~')[0].strip()
                peek_date = _dt.strptime(date_part, '%Y-%m-%d').date() if len(date_part) == 10 else None
            except Exception:
                peek_date = None
            prepared.append({'filename': fname, 'bytes': fbytes, 'period': peek_date})
        # Oldest period first; files without a parseable date go last.
        prepared.sort(key=lambda x: (x['period'] is None, x['period']))

        ok_count = 0; skip_count = 0; fail_count = 0
        last_ok = None  # (filename, bytes) of the most recent successful upload
        for p in prepared:
            fname, fbytes, peek_date = p['filename'], p['bytes'], p['period']

            # Block 1: same filename.
            if DailyUpload.query.filter_by(filename=fname).first():
                flash(f'דילוג — "{fname}" כבר קיים במערכת.', 'warning')
                skip_count += 1
                continue

            # Block 2: same period date under any filename.
            if peek_date is not None and DailyUpload.query.filter_by(upload_date=peek_date).first():
                flash(
                    f'דילוג — "{fname}": כבר קיימת העלאה לתאריך {peek_date.strftime("%d/%m/%Y")}.',
                    'warning',
                )
                skip_count += 1
                continue

            # Parse + store DPS/sessions/tournaments. The function commits at
            # the end; on exception we roll back so the next file isn't
            # poisoned by half-staged rows.
            try:
                _parse_and_store_stats_from_bytes(fbytes, fname)
                ok_count += 1
                last_ok = (fname, fbytes, peek_date)
            except Exception as e:
                _db.session.rollback()
                import logging
                logging.getLogger(__name__).error(f'Parse error on {fname}: {e}')
                flash(f'שגיאה בפרסינג "{fname}".', 'danger')
                fail_count += 1

        # The file with the latest Period date wins as "active" (used for
        # structure reading by union_data._read_sheets). prepared is sorted
        # ascending so the last successful upload is the newest.
        if last_ok:
            fname, fbytes, _ = last_ok
            try:
                from app.models import ActiveExcelData
                ActiveExcelData.query.delete()
                _db.session.add(ActiveExcelData(filename=fname, file_data=fbytes))
                _db.session.commit()
            except Exception:
                _db.session.rollback()

            # Save locally too (best-effort, dev-only).
            try:
                os.makedirs(UPLOAD_FOLDER, exist_ok=True)
                filepath = os.path.join(UPLOAD_FOLDER, fname)
                with open(filepath, 'wb') as f:
                    f.write(fbytes)
                session['uploaded_file'] = filepath
                with open(ACTIVE_FILE_PATH, 'w', encoding='utf-8') as f:
                    f.write(filepath)
                from app.union_data import set_excel_path
                set_excel_path(filepath)
            except Exception:
                pass

        # Summary flash. If exactly one OK, keep it specific; otherwise roll up.
        if ok_count == 1 and skip_count == 0 and fail_count == 0:
            flash(f'הקובץ "{last_ok[0]}" הועלה.', 'success')
        elif ok_count > 0:
            flash(f'הועלו {ok_count} קבצים בהצלחה. דילוגים: {skip_count}, שגיאות: {fail_count}.', 'success')
        elif skip_count > 0 and fail_count == 0:
            flash(f'כל {skip_count} הקבצים כבר קיימים — לא נוספו עליות חדשות.', 'warning')
        return redirect(url_for('upload.index'))

    uploaded = session.get('uploaded_file')
    active_file = None
    try:
        if os.path.exists(ACTIVE_FILE_PATH):
            with open(ACTIVE_FILE_PATH, 'r', encoding='utf-8') as f:
                active_path = f.read().strip()
            if active_path and os.path.exists(active_path):
                active_file = os.path.basename(active_path)
    except Exception:
        pass

    uploads = DailyUpload.query.order_by(DailyUpload.created_at.desc()).all()
    return render_template('upload/index.html',
                           current_file=os.path.basename(uploaded) if uploaded else None,
                           active_file=active_file, uploads=uploads)


@upload_bp.route('/preview')
@login_required
def preview():
    filepath = session.get('uploaded_file')
    if not filepath or not os.path.exists(filepath):
        flash('תצוגה מקדימה זמינה רק בהרצה מקומית.', 'warning')
        return redirect(url_for('upload.index'))

    import pandas as pd
    sheets_info = []
    try:
        all_sheets = pd.read_excel(filepath, sheet_name=None, header=None)
        for name, df in all_sheets.items():
            preview_rows = []
            for i, row in df.iterrows():
                cols = [str(v) if str(v) != 'nan' else '' for v in row]
                preview_rows.append(cols)
                if i >= 7:
                    break
            sheets_info.append({
                'name': name,
                'rows': df.shape[0],
                'cols': df.shape[1],
                'preview': preview_rows,
            })
    except Exception as e:
        import logging
        logging.getLogger(__name__).error(f'File read error: {e}')
        flash('שגיאה בקריאת הקובץ.', 'danger')
        return redirect(url_for('upload.index'))

    return render_template('upload/preview.html',
                           filename=os.path.basename(filepath),
                           sheets=sheets_info)


@upload_bp.route('/reset-all', methods=['POST'])
@login_required
def reset_all():
    """Archive the current active cycle, then clear active tables so the next
    upload starts a fresh cycle. Historical data remains accessible via the
    archive; dashboards re-count from the next file upload onward."""
    if not hasattr(current_user, 'role') or current_user.role != 'admin':
        flash('אין הרשאה.', 'danger')
        return redirect(url_for('upload.index'))

    return _archive_and_clear_active()


def _archive_and_clear_active():
    from app.models import (db, DailyUpload, DailyPlayerStats, ActiveExcelData,
                            PlayerSession, TournamentStats, ArchivePeriod,
                            MoneyTransfer)
    from sqlalchemy import func as sqlfunc, text

    # Archive data before deleting
    period_label = None
    try:
        date_range = db.session.query(
            sqlfunc.min(DailyUpload.upload_date),
            sqlfunc.max(DailyUpload.upload_date)
        ).first()

        if date_range and date_range[0] is not None:
            first_date, last_date = date_range
            period_label = f"{first_date.strftime('%d/%m/%Y')} — {last_date.strftime('%d/%m/%Y')}"

            period = ArchivePeriod(label=period_label, first_date=first_date, last_date=last_date)
            db.session.add(period)
            db.session.flush()
            pid = period.id

            # Bulk copy to archive tables using INSERT...SELECT
            db.session.execute(text(
                "INSERT INTO archived_uploads (period_id, original_id, filename, upload_date, created_at) "
                "SELECT :pid, id, filename, upload_date, created_at FROM daily_uploads"
            ), {'pid': pid})

            db.session.execute(text(
                "INSERT INTO archived_player_stats (period_id, upload_id, player_id, nickname, club, sa_id, agent_id, role, pnl, rake, hands) "
                "SELECT :pid, upload_id, player_id, nickname, club, sa_id, agent_id, role, pnl, rake, hands FROM daily_player_stats"
            ), {'pid': pid})

            db.session.execute(text(
                "INSERT INTO archived_player_sessions (period_id, upload_id, player_id, game_type, table_name, blinds, pnl) "
                "SELECT :pid, upload_id, player_id, game_type, table_name, blinds, pnl FROM player_sessions"
            ), {'pid': pid})

            db.session.execute(text(
                "INSERT INTO archived_tournament_stats (period_id, upload_id, title, status, game_type, buyin, fee, reentry, gtd, entries, prize_pool, start, duration) "
                "SELECT :pid, upload_id, title, status, game_type, buyin, fee, reentry, gtd, entries, prize_pool, start, duration FROM tournament_stats"
            ), {'pid': pid})
    except Exception as e:
        db.session.rollback()
        import logging
        logging.getLogger(__name__).error(f'Archive error: {e}')
        flash('שגיאה בארכוב הנתונים.', 'danger')
        return redirect(url_for('upload.index'))

    # Make the archive durable BEFORE the cycle-summary block. The archive
    # INSERTs above were only flushed; a rollback in the next try (e.g. an
    # error inside _build_cycle_summary_xlsx, or any DB hiccup) would otherwise
    # take the archive down with it — and the delete block that follows would
    # then wipe active data with nothing recoverable. See the 2026-04 incident
    # where exactly this happened.
    try:
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        import logging
        logging.getLogger(__name__).error(f'Archive commit error: {e}')
        flash('שגיאה בארכוב הנתונים.', 'danger')
        return redirect(url_for('upload.index'))

    # Snapshot the cycle summary Excel before clearing active data — the
    # report is built from live tables (DailyPlayerStats, TournamentStats),
    # so this must run BEFORE the delete block. Saved as a historical
    # (is_current=False) row so /admin/cycle-summary/ can list it; kept
    # 180 days independent of the 90-day archive retention.
    if period_label:
        try:
            from app.routes.admin import _build_cycle_summary_xlsx
            from app.models import CycleSummaryReport
            content, filename, _plabel = _build_cycle_summary_xlsx()
            db.session.add(CycleSummaryReport(
                is_current=False, period_label=period_label,
                filename=filename, content=content,
            ))
            db.session.commit()
        except Exception as e:
            db.session.rollback()
            import logging
            logging.getLogger(__name__).error(f'Cycle summary snapshot error: {e}')
            # Non-fatal — continue with reset.

    # Freeze open collection cycles — snapshot their live amounts before the
    # active data is cleared, so each cycle survives the reset as a standalone
    # record ('previous cycle'). Non-fatal: a failure here must not block the
    # reset itself.
    try:
        from app.models import CollectionCycle, PlayerPayment
        from app.routes.main import _collection_live_rows
        for cyc in CollectionCycle.query.filter_by(frozen=False, is_closed=False).all():
            live = _collection_live_rows(cyc.owner_id, None, None)
            existing = {p.player_id: p for p in cyc.payments}
            for pid, d in live.items():
                pay = existing.get(pid)
                if not pay:
                    pay = PlayerPayment(cycle_id=cyc.id, player_id=pid)
                    db.session.add(pay)
                pay.nickname = d['nickname']
                pay.club = d['club']
                pay.base_amount = d['base']
                pay.total_rake = d['rake']
            cyc.frozen = True
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        import logging
        logging.getLogger(__name__).error(f'Collection freeze error: {e}')
        # Non-fatal — continue with reset.

    # Delete active data
    TournamentStats.query.delete()
    PlayerSession.query.delete()
    ActiveExcelData.query.delete()
    DailyPlayerStats.query.delete()
    DailyUpload.query.delete()
    # Money transfers belong to the cycle that's being reset: they adjust
    # per-card PnL but not the top box, so leaving them behind after a reset
    # produces a phantom DELTA on the health page (one-sided settlements from
    # the old cycle keep counting). Clear them with the rest of the active data
    # — the frozen collection-cycle snapshots already preserve the settlement
    # record for history.
    MoneyTransfer.query.delete()
    db.session.commit()

    session.pop('uploaded_file', None)

    try:
        with open(ACTIVE_FILE_PATH, 'w', encoding='utf-8') as f:
            f.write('')
        from app.union_data import set_excel_path
        set_excel_path('')
        if os.path.exists(UPLOAD_FOLDER):
            for fname in os.listdir(UPLOAD_FOLDER):
                fpath = os.path.join(UPLOAD_FOLDER, fname)
                if os.path.isfile(fpath) and fname != '_active.txt':
                    os.remove(fpath)
    except Exception:
        pass

    if period_label:
        flash(f'הנתונים אורכבו לתקופה: {period_label}. כל הנתונים הפעילים אופסו.', 'success')
    else:
        flash('כל הנתונים המצטברים אופסו לגמרי.', 'success')
    return redirect(url_for('upload.index'))
