from functools import wraps
from datetime import datetime, timedelta
from flask import Blueprint, render_template, redirect, url_for, flash, request, session, jsonify
from flask_login import login_user, logout_user, login_required, current_user
from app.models import db, User

auth_bp = Blueprint('auth', __name__, url_prefix='/auth')


# ── Per-agent extra-password gate (2FA on agent cards) ──────────────────
# Unlock state lives in the Flask session so it auto-expires on logout.
# Window is 10 minutes from the last successful unlock for that sa_id.

_UNLOCK_WINDOW = timedelta(minutes=10)


def _now_ts():
    return datetime.utcnow().timestamp()


def is_agent_unlocked(sa_id):
    """True iff the current session has a fresh unlock for this sa_id."""
    pa = session.get('protected_unlocked', {})
    ts = pa.get(sa_id)
    if not ts:
        return False
    if (_now_ts() - ts) > _UNLOCK_WINDOW.total_seconds():
        # Stale — drop it so the dict doesn't grow forever.
        pa.pop(sa_id, None)
        session['protected_unlocked'] = pa
        session.modified = True
        return False
    return True


def mark_agent_unlocked(sa_id):
    pa = session.get('protected_unlocked', {})
    pa[sa_id] = _now_ts()
    session['protected_unlocked'] = pa
    session.modified = True


def _purge_unlock(sa_id):
    """Remove this sa_id from the current session's unlock dict (only —
    other sessions are unaffected, but they'll be forced to re-enter the
    new password the next time they hit the gate since the password hash
    changed)."""
    pa = session.get('protected_unlocked', {})
    if sa_id in pa:
        pa.pop(sa_id, None)
        session['protected_unlocked'] = pa
        session.modified = True


def admin_required(f):
    @wraps(f)
    @login_required
    def decorated(*args, **kwargs):
        if not hasattr(current_user, 'role') or current_user.role != 'admin':
            flash('אין לך הרשאה לדף זה.', 'danger')
            return redirect(url_for('main.dashboard'))
        return f(*args, **kwargs)
    return decorated


@auth_bp.route('/register', methods=['GET', 'POST'])
def register():
    if current_user.is_authenticated:
        return redirect(url_for('main.dashboard'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        email = request.form.get('email', '').strip()
        password = request.form.get('password', '')
        confirm_password = request.form.get('confirm_password', '')

        error = None
        if not username or len(username) < 3:
            error = 'שם המשתמש חייב להכיל לפחות 3 תווים.'
        elif not email or '@' not in email:
            error = 'כתובת אימייל לא תקינה.'
        elif not password or len(password) < 6:
            error = 'הסיסמה חייבת להכיל לפחות 6 תווים.'
        elif password != confirm_password:
            error = 'הסיסמאות אינן תואמות.'
        elif User.query.filter_by(username=username).first():
            error = 'שם המשתמש כבר קיים.'
        elif User.query.filter_by(email=email).first():
            error = 'כתובת האימייל כבר רשומה.'

        if error:
            flash(error, 'danger')
        else:
            user = User(username=username, email=email)
            user.set_password(password)
            db.session.add(user)
            db.session.commit()
            flash('ההרשמה הצליחה! ניתן להתחבר עכשיו.', 'success')
            return redirect(url_for('auth.login'))

    return render_template('auth/register.html')


@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('main.dashboard'))

    if request.method == 'POST':
        login_id = request.form.get('email', '').strip()
        password = request.form.get('password', '')
        remember = request.form.get('remember') == 'on'

        user = User.query.filter_by(email=login_id).first()
        if not user:
            user = User.query.filter_by(username=login_id).first()
        if not user:
            user = User.query.filter_by(player_id=login_id).first()

        if not user or not user.check_password(password):
            flash('שם משתמש או סיסמה שגויים.', 'danger')
        else:
            login_user(user, remember=remember)
            session.permanent = True
            # Log the login
            try:
                from app.models import LoginLog
                log = LoginLog(
                    user_id=user.id,
                    username=user.username,
                    role=user.role,
                    ip_address=request.headers.get('X-Forwarded-For', request.remote_addr),
                    user_agent=str(request.user_agent)[:300],
                )
                db.session.add(log)
                db.session.commit()
            except Exception:
                db.session.rollback()
            # admin1 always lands on the dashboard, never the control panel,
            # even if ?next= pointed to /admin/ from a prior deep link.
            if user.username == 'admin1':
                return redirect(url_for('main.dashboard'))
            next_page = request.args.get('next')
            if not next_page or not next_page.startswith('/') or next_page.startswith('//'):
                next_page = url_for('main.dashboard')
            return redirect(next_page)

    return render_template('auth/login.html')


@auth_bp.route('/logout')
@login_required
def logout():
    logout_user()
    flash('התנתקת בהצלחה.', 'info')
    return redirect(url_for('auth.login'))


@auth_bp.route('/users', methods=['GET', 'POST'])
@admin_required
def users():
    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'add':
            user_type = request.form.get('user_type', 'member')
            password = request.form.get('password', '')
            role = request.form.get('role', 'player')

            if role not in ('admin', 'agent', 'player', 'club'):
                role = 'player'

            error = None
            player_id = ''
            username = ''
            lookup_mode = request.form.get('lookup_mode', 'list')

            if user_type == 'club':
                role = 'club'
                if lookup_mode == 'list':
                    club_key = request.form.get('club_key', '').strip()
                    if '|' in club_key:
                        player_id, username = club_key.split('|', 1)
                    else:
                        error = 'יש לבחור מועדון מהרשימה.'
                elif lookup_mode == 'name':
                    username = request.form.get('manual_username', '').strip()
                    player_id = request.form.get('manual_id', '').strip()
                    if not username:
                        error = 'יש להזין שם מועדון.'
                elif lookup_mode == 'id':
                    player_id = request.form.get('manual_id', '').strip()
                    username = request.form.get('manual_username', '').strip() or player_id
                    if not player_id:
                        error = 'יש להזין Club ID.'
            else:
                if lookup_mode == 'list':
                    member_key = request.form.get('member_key', '').strip()
                    if '|' in member_key:
                        player_id, username = member_key.split('|', 1)
                    else:
                        error = 'יש לבחור שחקן מהרשימה.'
                elif lookup_mode == 'user':
                    username = request.form.get('manual_username', '').strip()
                    player_id = request.form.get('manual_id', '').strip()
                    if not username:
                        error = 'יש להזין שם משתמש.'
                elif lookup_mode == 'id':
                    player_id = request.form.get('manual_id', '').strip()
                    username = request.form.get('manual_username', '').strip() or player_id
                    if not player_id:
                        error = 'יש להזין Player ID.'

            if not error and (not password or len(password) < 6):
                error = 'הסיסמה חייבת להכיל לפחות 6 תווים.'
            elif not error and User.query.filter_by(username=username).first():
                error = f'משתמש {username} כבר קיים במערכת.'

            if error:
                flash(error, 'danger')
            else:
                try:
                    import uuid
                    unique_email = f'{player_id}-{uuid.uuid4().hex[:6]}@player.local'
                    user = User(username=username, email=unique_email,
                               player_id=player_id, role=role)
                    user.set_password(password)
                    db.session.add(user)
                    db.session.commit()
                    role_name = {'admin': 'מנהל', 'agent': 'סוכן', 'player': 'שחקן', 'club': 'מועדון'}[role]
                    flash(f'משתמש {username} ({role_name}) נוצר בהצלחה.', 'success')
                except Exception as e:
                    db.session.rollback()
                    import logging
                    logging.getLogger(__name__).error(f'User creation error: {e}')
                    flash('שגיאה ביצירת משתמש.', 'danger')

        elif action == 'delete':
            user_id = request.form.get('user_id')
            user = User.query.get(user_id)
            if user and user.id != current_user.id:
                db.session.delete(user)
                db.session.commit()
                flash(f'משתמש {user.username} נמחק.', 'success')
            elif user and user.id == current_user.id:
                flash('לא ניתן למחוק את עצמך.', 'warning')

        elif action == 'change_password':
            user_id = request.form.get('user_id')
            new_password = request.form.get('new_password', '')
            user = User.query.get(user_id)
            if not user:
                flash('משתמש לא נמצא.', 'danger')
            elif len(new_password) < 6:
                flash('הסיסמה חייבת להכיל לפחות 6 תווים.', 'danger')
            else:
                user.set_password(new_password)
                db.session.commit()
                flash(f'סיסמה עודכנה עבור {user.username}.', 'success')

        elif action == 'update_role':
            user_id = request.form.get('user_id')
            new_role = request.form.get('role')
            user = User.query.get(user_id)
            if user and new_role in ('admin', 'agent', 'player', 'club'):
                if user.id == current_user.id:
                    flash('לא ניתן לשנות את התפקיד של עצמך.', 'warning')
                else:
                    user.role = new_role
                    db.session.commit()
                    flash(f'תפקיד {user.username} עודכן.', 'success')

        elif action == 'set_agent_protection':
            from app.models import ProtectedAgent
            sa_id = request.form.get('sa_id', '').strip()
            password = request.form.get('protection_password', '')
            if not sa_id:
                flash('יש לבחור סוכן.', 'danger')
            elif len(password) < 4:
                flash('הסיסמה חייבת להכיל לפחות 4 תווים.', 'danger')
            else:
                row = ProtectedAgent.query.filter_by(sa_id=sa_id).first()
                if row:
                    row.set_password(password)
                    flash(f'סיסמת ההגנה לסוכן {sa_id} עודכנה.', 'success')
                else:
                    row = ProtectedAgent(sa_id=sa_id,
                                         created_by_user_id=current_user.id)
                    row.set_password(password)
                    db.session.add(row)
                    flash(f'הגנת אימות-כפול הופעלה על סוכן {sa_id}.', 'success')
                db.session.commit()
                # Clear any stale unlock for this sa_id so the new password
                # takes effect for everyone who had unlocked previously.
                _purge_unlock(sa_id)

        elif action == 'remove_agent_protection':
            from app.models import ProtectedAgent
            sa_id = request.form.get('sa_id', '').strip()
            row = ProtectedAgent.query.filter_by(sa_id=sa_id).first()
            if row:
                db.session.delete(row)
                db.session.commit()
                _purge_unlock(sa_id)
                flash(f'הגנת אימות-כפול הוסרה מסוכן {sa_id}.', 'success')

        elif action == 'set_panel_protection':
            from app.models import ProtectedAgent
            from app.routes.admin import ADMIN_PANEL_GATE_KEY
            password = request.form.get('panel_password', '')
            if len(password) < 4:
                flash('הסיסמה חייבת להכיל לפחות 4 תווים.', 'danger')
            else:
                row = ProtectedAgent.query.filter_by(
                    sa_id=ADMIN_PANEL_GATE_KEY).first()
                if row:
                    row.set_password(password)
                    flash('סיסמת הגנת לוח הבקרה עודכנה.', 'success')
                else:
                    row = ProtectedAgent(sa_id=ADMIN_PANEL_GATE_KEY,
                                         created_by_user_id=current_user.id)
                    row.set_password(password)
                    db.session.add(row)
                    flash('הגנת לוח הבקרה הופעלה.', 'success')
                db.session.commit()
                _purge_unlock(ADMIN_PANEL_GATE_KEY)

        elif action == 'remove_panel_protection':
            from app.models import ProtectedAgent
            from app.routes.admin import ADMIN_PANEL_GATE_KEY
            row = ProtectedAgent.query.filter_by(
                sa_id=ADMIN_PANEL_GATE_KEY).first()
            if row:
                db.session.delete(row)
                db.session.commit()
                _purge_unlock(ADMIN_PANEL_GATE_KEY)
                flash('הגנת לוח הבקרה הוסרה.', 'success')

        return redirect(url_for('auth.users'))

    from app.union_data import get_all_members, get_all_clubs
    from app.models import ProtectedAgent
    from app.routes.admin import (OVERVIEW_MANAGERS, OVERVIEW_EXTERNAL_AGENTS,
                                   ADMIN_PANEL_GATE_KEY)
    all_users = User.query.order_by(User.created_at.desc()).all()
    members = get_all_members()
    clubs = get_all_clubs()
    existing_pids = {u.player_id for u in all_users if u.player_id}
    available_members = [m for m in members if m['player_id'] not in existing_pids]
    available_clubs = [c for c in clubs if c['club_id'] not in existing_pids]

    # Agents shown on the admin overview — these are the ones eligible for
    # extra-password protection. Pre-build the dropdown options including
    # tracked external agents so the admin can protect those too.
    protectable_agents = [{'sa_id': pid, 'name': name}
                          for name, pid in OVERVIEW_MANAGERS + OVERVIEW_EXTERNAL_AGENTS]
    protectable_agents.sort(key=lambda a: a['name'].lower())

    protected_rows = ProtectedAgent.query.order_by(
        ProtectedAgent.updated_at.desc()).all()
    name_by_sa = {pid: name for name, pid in
                  OVERVIEW_MANAGERS + OVERVIEW_EXTERNAL_AGENTS}
    # Split off the panel-gate sentinel so the per-agent table doesn't show
    # "__admin_panel__" as a row — it gets its own dedicated section.
    panel_row = None
    protected_list = []
    for r in protected_rows:
        meta = {
            'sa_id': r.sa_id,
            'name': name_by_sa.get(r.sa_id, r.sa_id),
            'updated_at': r.updated_at.strftime('%d/%m/%Y %H:%M') if r.updated_at else '',
            'created_by': r.created_by.username if r.created_by else '—',
        }
        if r.sa_id == ADMIN_PANEL_GATE_KEY:
            panel_row = meta
        else:
            protected_list.append(meta)

    return render_template('auth/users.html', users=all_users,
                           members=available_members, clubs=available_clubs,
                           protectable_agents=protectable_agents,
                           protected_list=protected_list,
                           panel_protection=panel_row)


def _safe_next(next_url, fallback):
    """Only allow same-origin relative paths in `next` to prevent
    open-redirect into a phishing page."""
    if next_url and next_url.startswith('/') and not next_url.startswith('//'):
        return next_url
    return fallback


@auth_bp.route('/agent-protection/<sa_id>/unlock', methods=['POST'])
@login_required
def unlock_protected_agent(sa_id):
    """Verify the per-agent extra password and stamp the session unlock.
    Returns JSON for the modal flow on the overview page; also accepts
    form-encoded POST from the fallback gate page (in which case it
    redirects on success)."""
    from app.models import ProtectedAgent
    if request.is_json:
        password = (request.get_json(silent=True) or {}).get('password', '')
    else:
        password = request.form.get('password', '')
    next_url = request.form.get('next') or request.args.get('next', '')
    row = ProtectedAgent.query.filter_by(sa_id=sa_id).first()

    wants_json = request.is_json or (
        request.headers.get('Accept', '').startswith('application/json'))

    if not row:
        # No protection set — nothing to unlock. Treat as success so the
        # caller can proceed (defensive — the UI shouldn't reach here).
        if wants_json:
            return jsonify({'ok': True}), 200
        return redirect(_safe_next(next_url,
                                   url_for('main.dashboard', view_as=sa_id)))

    if not password or not row.check_password(password):
        if wants_json:
            return jsonify({'ok': False, 'error': 'wrong_password'}), 403
        flash('סיסמה שגויה.', 'danger')
        return redirect(url_for('auth.agent_gate', sa_id=sa_id, next=next_url))

    mark_agent_unlocked(sa_id)
    if wants_json:
        return jsonify({'ok': True})
    return redirect(_safe_next(next_url,
                               url_for('main.dashboard', view_as=sa_id)))


@auth_bp.route('/agent-protection/<sa_id>/gate')
@login_required
def agent_gate(sa_id):
    """Fallback gate page — shown when admin lands on /dashboard?view_as=
    or /admin/agent-view/<sa_id> for a protected agent without an active
    unlock. Also handles direct-URL access AND the whole-panel sentinel
    used by admin_bp.before_request. The overview page normally pops a
    modal instead of redirecting here, but the server still gates the
    actual data routes so URL-typing can't bypass it."""
    from app.models import ProtectedAgent
    from app.routes.admin import (OVERVIEW_MANAGERS, OVERVIEW_EXTERNAL_AGENTS,
                                   ADMIN_PANEL_GATE_KEY)
    next_url = request.args.get('next', '')
    row = ProtectedAgent.query.filter_by(sa_id=sa_id).first()
    if not row:
        # No protection set for this id — let the caller through.
        # Default fallback differs per kind: panel goes to /admin/, agent
        # goes to its dashboard.
        fallback = (url_for('admin.overview') if sa_id == ADMIN_PANEL_GATE_KEY
                    else url_for('main.dashboard', view_as=sa_id))
        return redirect(_safe_next(next_url, fallback))

    if sa_id == ADMIN_PANEL_GATE_KEY:
        agent_name = 'לוח בקרה'
        is_panel = True
    else:
        name_by_sa = {pid: name for name, pid in
                      OVERVIEW_MANAGERS + OVERVIEW_EXTERNAL_AGENTS}
        agent_name = name_by_sa.get(sa_id, sa_id)
        is_panel = False

    return render_template('auth/agent_gate.html',
                           sa_id=sa_id,
                           agent_name=agent_name,
                           is_panel=is_panel,
                           next_url=next_url)
