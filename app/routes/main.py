from datetime import date
from flask import Blueprint, render_template, redirect, url_for, flash, request, jsonify
from flask_login import login_required, current_user
from sqlalchemy import func
from app.models import db, Transaction

main_bp = Blueprint('main', __name__)

INCOME_CATEGORIES = ['משכורת', 'פרילנס', 'השקעות', 'מתנה', 'אחר']
EXPENSE_CATEGORIES = ['מזון', 'דיור', 'תחבורה', 'בריאות', 'בידור', 'קניות', 'חינוך', 'חשבונות', 'אחר']


@main_bp.route('/')
def index():
    if current_user.is_authenticated:
        return redirect(url_for('main.dashboard'))
    return redirect(url_for('auth.login'))


@main_bp.route('/dashboard')
@login_required
def dashboard():
    if hasattr(current_user, 'role') and current_user.role == 'admin':
        from app.union_data import get_union_overview, get_cumulative_totals
        meta, _, _ = get_union_overview()
        ct = get_cumulative_totals()
        meta['period'] = ct['period']
        return render_template('main/admin_dashboard.html',
                               meta=meta, clubs=ct['clubs'],
                               total={'active_players': ct['total_players'],
                                      'total_hands': ct['total_hands'],
                                      'total_fee': ct['total_rake'], 'pnl': ct['total_pnl']},
                               tables_count=ct['uploads_count'],
                               total_rake=ct['total_rake'], total_pnl=ct['total_pnl'],
                               total_hands=ct['total_hands'])

    if hasattr(current_user, 'role') and current_user.role == 'agent' and current_user.player_id:
        from app.union_data import get_super_agent_tables, get_members_hierarchy
        from app.models import SAHierarchy, SARakeConfig, RakeConfig, ExpenseCharge
        sa_id = current_user.player_id
        sa_tables = get_super_agent_tables()

        # Find this SA's data (may appear in multiple clubs)
        my_sas = [sa for sa in sa_tables if sa['sa_id'] == sa_id]

        # Find child SAs (SAs managed by this agent)
        child_sa_ids = [h.child_sa_id for h in SAHierarchy.query.filter_by(parent_sa_id=sa_id).all()]
        child_sas = [sa for sa in sa_tables if sa['sa_id'] in child_sa_ids]

        # Managed club - get all SAs in that club
        rake_cfg = SARakeConfig.query.filter_by(sa_id=sa_id).first()
        rake_pct = rake_cfg.rake_percent if rake_cfg else 0
        managed_club_id = rake_cfg.managed_club_id if rake_cfg else None
        managed_club_data = None
        if managed_club_id:
            clubs_data, _ = get_members_hierarchy()
            for club in clubs_data:
                if club['club_id'] == managed_club_id:
                    managed_club_data = club
                    break

        # Totals from direct SA data
        total_rake = sum(s['total_rake'] for s in my_sas) + sum(s['total_rake'] for s in child_sas)
        total_pnl = sum(s['total_pnl'] for s in my_sas) + sum(s['total_pnl'] for s in child_sas)
        total_hands = sum(s['total_hands'] for s in my_sas) + sum(s['total_hands'] for s in child_sas)

        # Add managed club totals + calculate net rake per entity
        club_net_rake = 0
        if managed_club_data:
            club_rake = managed_club_data.get('total_rake', 0)
            total_rake += club_rake
            total_pnl += managed_club_data.get('total_pnl', 0)
            # Check if club has a RakeConfig (what % the club keeps)
            club_rc = RakeConfig.query.filter_by(entity_type='club',
                                                  entity_id=managed_club_id).first()
            club_keeps_pct = club_rc.rake_percent if club_rc else 0
            club_net_rake = round(club_rake * (100 - club_keeps_pct) / 100, 2)

        # Net rake = what's left after all entities take their cut
        # For SA's own players: use rake_pct (SA keeps this %)
        sa_net_rake = round(
            (sum(s['total_rake'] for s in my_sas) + sum(s['total_rake'] for s in child_sas))
            * rake_pct / 100, 2) if rake_pct else 0
        net_rake = round(sa_net_rake + club_net_rake, 2)

        # Count players
        player_count = 0
        for sa in my_sas:
            player_count += len(sa['direct'])
            for ag in sa['agents'].values():
                player_count += len(ag['members'])
        for cs in child_sas:
            player_count += len(cs['direct'])
            for ag in cs['agents'].values():
                player_count += len(ag['members'])

        # Expense charges for this agent
        expense_charges = ExpenseCharge.query.filter_by(agent_player_id=sa_id).all()
        total_expenses = round(sum(c.charge_amount for c in expense_charges), 2)
        net_rake_after_expenses = round(net_rake - total_expenses, 2)

        return render_template('main/agent_dashboard.html',
                               my_sas=my_sas, child_sas=child_sas,
                               managed_club=managed_club_data,
                               total_rake=total_rake, total_pnl=total_pnl,
                               total_hands=int(total_hands), net_rake=net_rake,
                               net_rake_after_expenses=net_rake_after_expenses,
                               total_expenses=total_expenses,
                               expense_charges=expense_charges,
                               rake_pct=rake_pct, player_count=player_count,
                               club_net_rake=club_net_rake,
                               club_keeps_pct=club_keeps_pct if managed_club_data else 0)

    transactions = (Transaction.query
                    .filter_by(user_id=current_user.id)
                    .order_by(Transaction.date.desc())
                    .limit(5)
                    .all())

    total_income = db.session.query(func.sum(Transaction.amount)).filter_by(
        user_id=current_user.id, type='income').scalar() or 0

    total_expense = db.session.query(func.sum(Transaction.amount)).filter_by(
        user_id=current_user.id, type='expense').scalar() or 0

    balance = total_income - total_expense

    return render_template('main/dashboard.html',
                           transactions=transactions,
                           total_income=total_income,
                           total_expense=total_expense,
                           balance=balance)


@main_bp.route('/agent/reports')
@login_required
def agent_reports():
    if not hasattr(current_user, 'role') or current_user.role != 'agent' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.union_data import get_super_agent_tables
    from app.models import SAHierarchy

    sa_id = current_user.player_id
    sa_tables = get_super_agent_tables()
    my_sas = [sa for sa in sa_tables if sa['sa_id'] == sa_id]
    child_sa_ids = [h.child_sa_id for h in SAHierarchy.query.filter_by(parent_sa_id=sa_id).all()]
    child_sas = [sa for sa in sa_tables if sa['sa_id'] in child_sa_ids]

    my_players = []
    my_player_ids = []
    for sa in my_sas + child_sas:
        for m in sa['direct']:
            my_players.append({'player_id': m['player_id'], 'nickname': m['nickname']})
            my_player_ids.append(m['player_id'])
        for ag in sa['agents'].values():
            for m in ag['members']:
                my_players.append({'player_id': m['player_id'], 'nickname': m['nickname']})
                my_player_ids.append(m['player_id'])

    return render_template('main/agent_reports.html',
                           players=my_players, player_ids=my_player_ids)


@main_bp.route('/agent/transfers', methods=['GET', 'POST'])
@login_required
def agent_transfers():
    if not hasattr(current_user, 'role') or current_user.role != 'agent' or not current_user.player_id:
        return redirect(url_for('main.dashboard'))

    from app.union_data import get_super_agent_tables, get_player_balance, get_all_balances
    from app.models import MoneyTransfer, SAHierarchy

    sa_id = current_user.player_id
    sa_tables = get_super_agent_tables()

    # Collect all player IDs under this agent
    my_player_ids = set()
    my_players = []  # [{player_id, nickname, club}]
    my_sas = [sa for sa in sa_tables if sa['sa_id'] == sa_id]
    child_sa_ids = [h.child_sa_id for h in SAHierarchy.query.filter_by(parent_sa_id=sa_id).all()]
    child_sas = [sa for sa in sa_tables if sa['sa_id'] in child_sa_ids]

    for sa in my_sas + child_sas:
        for m in sa['direct']:
            my_player_ids.add(m['player_id'])
            my_players.append({'player_id': m['player_id'], 'nickname': m['nickname'], 'club': sa['club']})
        for ag in sa['agents'].values():
            for m in ag['members']:
                my_player_ids.add(m['player_id'])
                my_players.append({'player_id': m['player_id'], 'nickname': m['nickname'], 'club': sa['club']})

    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add':
            from_key = request.form.get('from_key', '').strip()
            to_key = request.form.get('to_key', '').strip()
            description = request.form.get('description', '').strip()
            try:
                amount = float(request.form.get('amount', 0))
            except ValueError:
                flash('סכום לא תקין.', 'danger')
                return redirect(url_for('main.agent_transfers'))

            if not from_key or not to_key or '|' not in from_key or '|' not in to_key:
                flash('יש לבחור שולח ומקבל.', 'danger')
            elif from_key == to_key:
                flash('לא ניתן להעביר לאותו שחקן.', 'warning')
            elif amount <= 0:
                flash('הסכום חייב להיות חיובי.', 'danger')
            else:
                from_pid = from_key.split('|', 1)[0]
                to_pid = to_key.split('|', 1)[0]
                from_name = from_key.split('|', 1)[1]
                to_name = to_key.split('|', 1)[1]
                # Verify both players belong to this agent
                if from_pid not in my_player_ids or to_pid not in my_player_ids:
                    flash('אין הרשאה להעביר לשחקן שלא שייך אליך.', 'danger')
                else:
                    from_balance = get_player_balance(from_pid)
                    to_balance = get_player_balance(to_pid)
                    max_transfer = min(abs(from_balance), to_balance)
                    if from_balance >= 0:
                        flash(f'{from_name} לא במינוס.', 'danger')
                    elif to_balance <= 0:
                        flash(f'{to_name} לא בפלוס.', 'danger')
                    elif amount > max_transfer:
                        flash(f'חריגה! מקסימום: {max_transfer:.2f}', 'danger')
                    else:
                        t = MoneyTransfer(user_id=current_user.id,
                                          from_player_id=from_pid, from_name=from_name,
                                          to_player_id=to_pid, to_name=to_name,
                                          amount=amount, description=description)
                        db.session.add(t)
                        db.session.commit()
                        flash(f'העברה של {amount} מ-{from_name} ל-{to_name} בוצעה.', 'success')
        elif action == 'delete':
            tid = request.form.get('transfer_id')
            t = MoneyTransfer.query.get(tid)
            if t and (t.from_player_id in my_player_ids or t.to_player_id in my_player_ids):
                db.session.delete(t)
                db.session.commit()
                flash('העברה נמחקה.', 'success')
        return redirect(url_for('main.agent_transfers'))

    balances = get_all_balances(my_player_ids)
    # Get transfers for my players only
    my_transfers = MoneyTransfer.query.filter(
        db.or_(
            MoneyTransfer.from_player_id.in_(my_player_ids),
            MoneyTransfer.to_player_id.in_(my_player_ids)
        )
    ).order_by(MoneyTransfer.created_at.desc()).all()

    return render_template('main/agent_transfers.html',
                           players=my_players, balances=balances,
                           transfers=my_transfers)


@main_bp.route('/api/report')
@login_required
def report_api():
    from app.models import DailyPlayerStats, DailyUpload
    from datetime import datetime
    from sqlalchemy import func

    from_date = request.args.get('from')
    to_date = request.args.get('to')
    player_id = request.args.get('player_id', '')

    if not from_date or not to_date:
        return jsonify({'error': 'missing dates'}), 400

    try:
        fd = datetime.strptime(from_date, '%Y-%m-%d').date()
        td = datetime.strptime(to_date, '%Y-%m-%d').date()
    except ValueError:
        return jsonify({'error': 'invalid date format'}), 400

    # Get upload IDs in range
    uploads = DailyUpload.query.filter(
        DailyUpload.upload_date >= fd,
        DailyUpload.upload_date <= td
    ).all()
    upload_ids = [u.id for u in uploads]

    if not upload_ids:
        return jsonify({'players': [], 'totals': {'pnl': 0, 'rake': 0, 'hands': 0}, 'days': 0})

    query = DailyPlayerStats.query.with_entities(
        DailyPlayerStats.player_id,
        func.max(DailyPlayerStats.nickname),
        func.max(DailyPlayerStats.club),
        func.sum(DailyPlayerStats.pnl),
        func.sum(DailyPlayerStats.rake),
        func.sum(DailyPlayerStats.hands),
    ).filter(DailyPlayerStats.upload_id.in_(upload_ids))

    if player_id:
        query = query.filter(DailyPlayerStats.player_id == player_id)

    query = query.group_by(DailyPlayerStats.player_id)
    results = query.all()

    players = []
    total_pnl = 0
    total_rake = 0
    total_hands = 0
    for pid, nick, club, pnl, rake, hands in results:
        p = round(float(pnl or 0), 2)
        r = round(float(rake or 0), 2)
        h = int(hands or 0)
        players.append({'player_id': pid, 'nickname': nick, 'club': club,
                        'pnl': p, 'rake': r, 'hands': h})
        total_pnl += p
        total_rake += r
        total_hands += h

    players.sort(key=lambda x: x['pnl'], reverse=True)

    return jsonify({
        'players': players,
        'totals': {'pnl': round(total_pnl, 2), 'rake': round(total_rake, 2), 'hands': total_hands},
        'days': len(upload_ids)
    })


@main_bp.route('/api/report-dates')
@login_required
def report_dates_api():
    """Return list of dates that have upload data."""
    from app.models import DailyUpload
    uploads = DailyUpload.query.with_entities(DailyUpload.upload_date).distinct().all()
    dates = [u[0].strftime('%Y-%m-%d') for u in uploads]
    return jsonify({'dates': dates})


@main_bp.route('/api/player-record/<player_id>')
@login_required
def player_record_api(player_id):
    from app.union_data import get_ring_game_detail, _read_sheets, _num
    sessions = []

    # Ring games
    for table in get_ring_game_detail():
        for p in table['players']:
            if p['player_id'] == player_id:
                sessions.append({
                    'table': table['table_name'],
                    'game': table['game_type'],
                    'blinds': table['blinds'],
                    'pnl': p['pnl'],
                })

    # MTT tournaments
    sheets = _read_sheets()
    if 'Union MTT Detail' in sheets:
        df = sheets['Union MTT Detail']
        current_tournament = ''
        for i in range(len(df)):
            col0 = str(df.iloc[i, 0])
            if col0.startswith('Table Name :'):
                try:
                    current_tournament = col0.split('Table Name : ')[1].split(' , Creator')[0].strip()
                except Exception:
                    current_tournament = col0
            if str(df.iloc[i, 2]) == player_id:
                pnl = _num(df.iloc[i, 16])  # P&L column
                sessions.append({
                    'table': current_tournament[:40],
                    'game': 'MTT',
                    'blinds': '',
                    'pnl': pnl,
                })

    total_pnl = round(sum(s['pnl'] for s in sessions), 2)
    return jsonify({'sessions': sessions, 'total_pnl': total_pnl})


@main_bp.route('/transactions')
@login_required
def transactions():
    tx_type = request.args.get('type', '')
    category = request.args.get('category', '')

    query = Transaction.query.filter_by(user_id=current_user.id)

    if tx_type in ('income', 'expense'):
        query = query.filter_by(type=tx_type)
    if category:
        query = query.filter_by(category=category)

    all_transactions = query.order_by(Transaction.date.desc()).all()

    return render_template('main/transactions.html',
                           transactions=all_transactions,
                           income_categories=INCOME_CATEGORIES,
                           expense_categories=EXPENSE_CATEGORIES,
                           selected_type=tx_type,
                           selected_category=category)


@main_bp.route('/transactions/add', methods=['GET', 'POST'])
@login_required
def add_transaction():
    if request.method == 'POST':
        tx_type = request.form.get('type')
        amount_str = request.form.get('amount', '')
        category = request.form.get('category', '').strip()
        description = request.form.get('description', '').strip()
        date_str = request.form.get('date', '')

        error = None
        try:
            amount = float(amount_str)
            if amount <= 0:
                error = 'הסכום חייב להיות חיובי.'
        except ValueError:
            error = 'סכום לא תקין.'

        if not error:
            if tx_type not in ('income', 'expense'):
                error = 'סוג עסקה לא תקין.'
            elif not category:
                error = 'יש לבחור קטגוריה.'

        if not error:
            try:
                tx_date = date.fromisoformat(date_str)
            except ValueError:
                tx_date = date.today()

            transaction = Transaction(
                user_id=current_user.id,
                type=tx_type,
                amount=amount,
                category=category,
                description=description,
                date=tx_date
            )
            db.session.add(transaction)
            db.session.commit()
            flash('העסקה נוספה בהצלחה.', 'success')
            return redirect(url_for('main.transactions'))

        flash(error, 'danger')

    return render_template('main/add_transaction.html',
                           income_categories=INCOME_CATEGORIES,
                           expense_categories=EXPENSE_CATEGORIES,
                           today=date.today().isoformat())


@main_bp.route('/transactions/delete/<int:tx_id>', methods=['POST'])
@login_required
def delete_transaction(tx_id):
    transaction = Transaction.query.filter_by(id=tx_id, user_id=current_user.id).first_or_404()
    db.session.delete(transaction)
    db.session.commit()
    flash('העסקה נמחקה.', 'info')
    return redirect(url_for('main.transactions'))
