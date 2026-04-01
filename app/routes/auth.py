from functools import wraps
from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_user, logout_user, login_required, current_user
from app.models import db, User

auth_bp = Blueprint('auth', __name__, url_prefix='/auth')


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

        if not user or not user.check_password(password):
            flash('שם משתמש או סיסמה שגויים.', 'danger')
        else:
            login_user(user, remember=remember)
            next_page = request.args.get('next')
            return redirect(next_page or url_for('main.dashboard'))

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

            if user_type == 'club':
                club_key = request.form.get('club_key', '').strip()
                role = 'club'
                if '|' in club_key:
                    player_id, username = club_key.split('|', 1)
                else:
                    error = 'יש לבחור מועדון מהרשימה.'
            else:
                member_key = request.form.get('member_key', '').strip()
                if '|' in member_key:
                    player_id, username = member_key.split('|', 1)
                else:
                    error = 'יש לבחור שחקן מהרשימה.'

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
                    flash(f'שגיאה ביצירת משתמש: {str(e)[:100]}', 'danger')

        elif action == 'delete':
            user_id = request.form.get('user_id')
            user = User.query.get(user_id)
            if user and user.id != current_user.id:
                db.session.delete(user)
                db.session.commit()
                flash(f'משתמש {user.username} נמחק.', 'success')
            elif user and user.id == current_user.id:
                flash('לא ניתן למחוק את עצמך.', 'warning')

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

        return redirect(url_for('auth.users'))

    from app.union_data import get_all_members, get_all_clubs, get_all_super_agents
    all_users = User.query.order_by(User.created_at.desc()).all()
    members = get_all_members()
    super_agents = get_all_super_agents()
    clubs = get_all_clubs()
    # Merge SA into members list (they might not appear as players)
    member_ids = {m['player_id'] for m in members}
    for sa in super_agents:
        if sa['id'] not in member_ids:
            members.append({
                'player_id': sa['id'], 'nickname': sa['nick'],
                'role': 'Super Agent', 'club': sa['club'],
                'sa_nick': '-', 'agent_nick': '-',
            })
    members.sort(key=lambda x: x['nickname'].lower())
    existing_pids = {u.player_id for u in all_users if u.player_id}
    available_members = [m for m in members if m['player_id'] not in existing_pids]
    available_clubs = [c for c in clubs if c['club_id'] not in existing_pids]
    return render_template('auth/users.html', users=all_users,
                           members=available_members, clubs=available_clubs)
