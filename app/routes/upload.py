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
    upload = DailyUpload(filename=filename, upload_date=date.today())
    db.session.add(upload)
    db.session.flush()
    upload_id = upload.id

    # Collect all rows first
    rows = []
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
            hands = float(row.iloc[151]) if str(row.iloc[151]) != 'nan' else 0
        except (ValueError, TypeError):
            hands = 0

        rows.append({
            'upload_id': upload_id, 'player_id': player_id,
            'nickname': nickname, 'club': current_club,
            'pnl': round(pnl, 2), 'rake': round(rake, 2),
            'hands': round(hands, 0),
        })

    # Bulk insert all at once (much faster for cloud DB)
    if rows:
        db.session.execute(
            DailyPlayerStats.__table__.insert(),
            rows
        )

    db.session.commit()
    return len(rows)


@upload_bp.route('/', methods=['GET', 'POST'])
@login_required
def index():
    from app.models import DailyUpload

    if request.method == 'POST':
        if 'file' not in request.files:
            flash('לא נבחר קובץ.', 'danger')
            return redirect(request.url)

        file = request.files['file']
        if file.filename == '':
            flash('לא נבחר קובץ.', 'danger')
            return redirect(request.url)

        if not allowed_file(file.filename):
            flash('סוג קובץ לא נתמך. נא להעלות קובץ .xlsx או .xls', 'danger')
            return redirect(request.url)

        filename = secure_filename(file.filename)
        file_bytes = file.read()

        try:
            # Save Excel as active file in DB (for structure/hierarchy reading)
            from app.models import db as _db, ActiveExcelData
            ActiveExcelData.query.delete()
            _db.session.add(ActiveExcelData(filename=filename, file_data=file_bytes))
            _db.session.commit()

            # Also try to save locally (works on local dev, fails silently on Vercel)
            try:
                os.makedirs(UPLOAD_FOLDER, exist_ok=True)
                filepath = os.path.join(UPLOAD_FOLDER, filename)
                with open(filepath, 'wb') as f:
                    f.write(file_bytes)
                session['uploaded_file'] = filepath
                with open(ACTIVE_FILE_PATH, 'w', encoding='utf-8') as f:
                    f.write(filepath)
                from app.union_data import set_excel_path
                set_excel_path(filepath)
            except Exception:
                pass

            # Parse and store CUMULATIVE stats (adds to existing, never deletes)
            player_count = _parse_and_store_stats_from_bytes(file_bytes, filename)

            flash(f'הקובץ "{filename}" הועלה — {player_count} שחקנים נוספו למצטבר.', 'success')
        except Exception as e:
            try:
                _db.session.rollback()
            except Exception:
                pass
            flash(f'שגיאה בהעלאה: {str(e)[:150]}', 'danger')
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
        flash(f'שגיאה בקריאת הקובץ: {e}', 'danger')
        return redirect(url_for('upload.index'))

    return render_template('upload/preview.html',
                           filename=os.path.basename(filepath),
                           sheets=sheets_info)


@upload_bp.route('/reset', methods=['POST'])
@login_required
def reset():
    if not hasattr(current_user, 'role') or current_user.role != 'admin':
        flash('אין הרשאה.', 'danger')
        return redirect(url_for('upload.index'))

    from app.models import db, DailyUpload, DailyPlayerStats, ActiveExcelData

    # Clear all data from DB
    ActiveExcelData.query.delete()
    # Clear cumulative data from DB
    DailyPlayerStats.query.delete()
    DailyUpload.query.delete()
    db.session.commit()

    session.pop('uploaded_file', None)

    # Try to clean local files (works locally, fails silently on Vercel)
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

    flash('כל הנתונים המצטברים אופסו.', 'success')
    return redirect(url_for('upload.index'))
