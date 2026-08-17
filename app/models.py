from datetime import datetime
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()


class User(UserMixin, db.Model):
    __tablename__ = 'users'

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(64), unique=True, nullable=False)
    email = db.Column(db.String(120), nullable=True)
    password_hash = db.Column(db.String(256), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='admin')  # admin, agent, player
    player_id = db.Column(db.String(20), nullable=True)  # links to Excel player ID (e.g. 2197-2365)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    transactions = db.relationship('Transaction', backref='user', lazy='dynamic')

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def __repr__(self):
        return f'<User {self.username}>'


class Transaction(db.Model):
    __tablename__ = 'transactions'

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    type = db.Column(db.String(10), nullable=False)  # 'income' or 'expense'
    category = db.Column(db.String(50), nullable=False)
    description = db.Column(db.String(200))
    date = db.Column(db.Date, nullable=False, default=datetime.utcnow)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return f'<Transaction {self.type} {self.amount}>'


class AdminNote(db.Model):
    __tablename__ = 'admin_notes'

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    content = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    user = db.relationship('User', backref='notes')


# RETIRED synthetic "house" account. The return-to-house / distribute-from-house
# forms were removed from every dashboard: a house row has only one real side,
# so unlike a player-to-player transfer it is NOT zero-sum inside a box — it
# quietly shrank the card totals while the cash stayed with the agent, and the
# settlement stopped matching the dashboard.
# The constants remain because rows written before the removal still exist and
# must keep rendering with their name.
HOUSE_PLAYER_ID = '__house__'
HOUSE_PLAYER_NAME = 'הבית'


class MoneyTransfer(db.Model):
    __tablename__ = 'money_transfers'

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    # Wide enough for synthetic counterparty ids like '__club__<club name>'
    # (a club wallet), alongside real player/agent ids and '__house__<sa>'.
    from_player_id = db.Column(db.String(120), nullable=False)
    from_name = db.Column(db.String(100), nullable=False)
    to_player_id = db.Column(db.String(120), nullable=False)
    to_name = db.Column(db.String(100), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    description = db.Column(db.String(200))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    user = db.relationship('User', backref='transfers')


class PlayerCross(db.Model):
    """A per-club "balance" (הצלבה) for a SINGLE player whose P&L is split
    across two clubs — e.g. +602.95 in GORILLA ISRAELu and −600 in SPC T. It
    shifts `amount` from the player's winning club (from_club) to his losing
    club (to_club) so both sides net toward zero on the dashboard, leaving
    only the true remainder.

    Deliberately NOT a MoneyTransfer: it never touches the global wallet
    balance and is zero-sum for the player. Read by the dashboard's per-club
    aggregations (managed-club cards + the main list) to redistribute a
    player's P&L between his clubs."""
    __tablename__ = 'player_crosses'

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    player_id = db.Column(db.String(20), nullable=False, index=True)
    player_name = db.Column(db.String(100))
    from_club = db.Column(db.String(200), nullable=False)  # the +side (reduced)
    to_club = db.Column(db.String(200), nullable=False)    # the −side (raised)
    amount = db.Column(db.Float, nullable=False)           # always positive
    description = db.Column(db.String(200))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    user = db.relationship('User', backref='player_crosses')


class SAHierarchy(db.Model):
    __tablename__ = 'sa_hierarchy'

    id = db.Column(db.Integer, primary_key=True)
    parent_sa_id = db.Column(db.String(20), nullable=False)
    child_sa_id = db.Column(db.String(20), unique=True, nullable=False)

    def __repr__(self):
        return f'<SAHierarchy {self.parent_sa_id} -> {self.child_sa_id}>'


class SARakeConfig(db.Model):
    __tablename__ = 'sa_rake_config'

    id = db.Column(db.Integer, primary_key=True)
    sa_id = db.Column(db.String(20), nullable=False)
    rake_percent = db.Column(db.Float, nullable=False, default=0)
    managed_club_id = db.Column(db.String(20), nullable=True)

    def __repr__(self):
        return f'<SARakeConfig {self.sa_id} {self.rake_percent}%>'


class RakeConfig(db.Model):
    __tablename__ = 'rake_config'

    id = db.Column(db.Integer, primary_key=True)
    entity_type = db.Column(db.String(20), nullable=False)  # 'club', 'agent', 'player'
    entity_id = db.Column(db.String(20), nullable=False)     # club_id, agent SA ID, or player_id
    entity_name = db.Column(db.String(100), nullable=False)
    rake_percent = db.Column(db.Float, nullable=False, default=0)

    def __repr__(self):
        return f'<RakeConfig {self.entity_type}:{self.entity_name} {self.rake_percent}%>'


class CollectionCycle(db.Model):
    """A two-week settlement round owned by one agent. A new cycle can be
    opened while a previous one is still open ("not closed") — the old cycle's
    table stays visible until the agent explicitly closes it."""
    __tablename__ = 'collection_cycles'

    id = db.Column(db.Integer, primary_key=True)
    owner_id = db.Column(db.String(20), nullable=False, index=True)  # agent player_id
    label = db.Column(db.String(100), nullable=False)
    is_closed = db.Column(db.Boolean, default=False)   # agent finished settling it
    frozen = db.Column(db.Boolean, default=False)      # files were reset — now a snapshot
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    closed_at = db.Column(db.DateTime, nullable=True)

    payments = db.relationship('PlayerPayment', backref='cycle',
                               cascade='all, delete-orphan')

    def __repr__(self):
        return f'<CollectionCycle {self.label} owner={self.owner_id}>'


class PlayerPayment(db.Model):
    """One player's settlement row inside a collection cycle. Amounts are a
    frozen snapshot taken when the cycle is opened, so an old cycle does not
    change when new daily data is uploaded.

    settlement = base_amount + manual_rake
      base_amount > 0  -> agent owes the player (needs to receive)
      base_amount < 0  -> player owes the agent (in minus)
    manual_rake is rakeback the agent grants the player; it always benefits
    the player and is capped at the agent's configured rake percentage."""
    __tablename__ = 'player_payments'

    id = db.Column(db.Integer, primary_key=True)
    cycle_id = db.Column(db.Integer, db.ForeignKey('collection_cycles.id'),
                         nullable=False, index=True)
    player_id = db.Column(db.String(20), nullable=False, index=True)
    nickname = db.Column(db.String(100), default='')
    club = db.Column(db.String(100), default='')
    base_amount = db.Column(db.Float, default=0)   # snapshot PnL (+ transfer adj)
    total_rake = db.Column(db.Float, default=0)    # snapshot rake — caps manual_rake
    manual_rake = db.Column(db.Float, default=0)   # rakeback granted by the agent
    is_paid = db.Column(db.Boolean, default=False)  # manual "fully paid" mark
    paid_at = db.Column(db.DateTime, nullable=True)
    note = db.Column(db.String(200), default='')
    # How much the player has paid so far this cycle (supports partial/split
    # payments — owes 300, paid 200, remaining 100). Editable running figure;
    # settled once it covers the debt. A settlement is cash only — it never
    # touches poker PnL/rake or the admin reconciliation.
    paid_so_far = db.Column(db.Float, default=0)

    def __repr__(self):
        return f'<PlayerPayment {self.player_id} cycle={self.cycle_id} paid={self.is_paid}>'


class SharedExpense(db.Model):
    __tablename__ = 'shared_expenses'

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    description = db.Column(db.String(200), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    charged = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    user = db.relationship('User')
    charges = db.relationship('ExpenseCharge', backref='expense', cascade='all, delete-orphan')


class ExpenseCharge(db.Model):
    __tablename__ = 'expense_charges'

    id = db.Column(db.Integer, primary_key=True)
    expense_id = db.Column(db.Integer, db.ForeignKey('shared_expenses.id'), nullable=False)
    agent_player_id = db.Column(db.String(20), nullable=False)
    agent_name = db.Column(db.String(100), nullable=False)
    charge_amount = db.Column(db.Float, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class LoginLog(db.Model):
    __tablename__ = 'login_logs'

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    username = db.Column(db.String(64), nullable=False)
    role = db.Column(db.String(20))
    ip_address = db.Column(db.String(50))
    user_agent = db.Column(db.String(300))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class ActiveExcelData(db.Model):
    __tablename__ = 'active_excel_data'

    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(200), nullable=False)
    file_data = db.Column(db.LargeBinary, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class DailyUpload(db.Model):
    __tablename__ = 'daily_uploads'

    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(200), nullable=False)
    upload_date = db.Column(db.Date, nullable=False, default=datetime.utcnow)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    stats = db.relationship('DailyPlayerStats', backref='upload', cascade='all, delete-orphan')

    def __repr__(self):
        return f'<DailyUpload {self.filename} {self.upload_date}>'


class DailyPlayerStats(db.Model):
    __tablename__ = 'daily_player_stats'

    id = db.Column(db.Integer, primary_key=True)
    upload_id = db.Column(db.Integer, db.ForeignKey('daily_uploads.id'), nullable=False, index=True)
    player_id = db.Column(db.String(20), nullable=False, index=True)
    nickname = db.Column(db.String(100), nullable=False)
    club = db.Column(db.String(100), nullable=False, index=True)
    sa_id = db.Column(db.String(20), default='', index=True)
    agent_id = db.Column(db.String(20), default='', index=True)
    role = db.Column(db.String(30), default='')
    pnl = db.Column(db.Float, default=0)
    rake = db.Column(db.Float, default=0)
    hands = db.Column(db.Float, default=0)


class TournamentStats(db.Model):
    __tablename__ = 'tournament_stats'

    id = db.Column(db.Integer, primary_key=True)
    upload_id = db.Column(db.Integer, db.ForeignKey('daily_uploads.id'), nullable=False, index=True)
    title = db.Column(db.String(200), nullable=False)
    # 'Ended' / 'In Progress' straight from the report. A tournament still
    # in progress when the file was pulled has incomplete entries and no
    # prize pool, and PPPoker binds it to its START date — so late
    # registrations never reach any later file. Surfacing this is what
    # turns that from a silent shortfall into something visible.
    status = db.Column(db.String(30), default='')
    game_type = db.Column(db.String(30), default='')
    buyin = db.Column(db.Float, default=0)
    fee = db.Column(db.Float, default=0)
    reentry = db.Column(db.String(20), default='')
    gtd = db.Column(db.Float, default=0)
    entries = db.Column(db.Float, default=0)
    prize_pool = db.Column(db.Float, default=0)
    start = db.Column(db.String(20), default='')
    duration = db.Column(db.String(20), default='')


class PlayerSession(db.Model):
    __tablename__ = 'player_sessions'

    id = db.Column(db.Integer, primary_key=True)
    upload_id = db.Column(db.Integer, db.ForeignKey('daily_uploads.id'), nullable=False)
    player_id = db.Column(db.String(20), nullable=False)
    game_type = db.Column(db.String(20), nullable=False)  # 'Ring', 'MTT'
    table_name = db.Column(db.String(200), nullable=False)
    blinds = db.Column(db.String(20), default='')
    pnl = db.Column(db.Float, default=0)

    def __repr__(self):
        return f'<DailyPlayerStats {self.nickname} pnl={self.pnl}>'


# ── Archive Models ──

class ArchivePeriod(db.Model):
    __tablename__ = 'archive_periods'

    id = db.Column(db.Integer, primary_key=True)
    label = db.Column(db.String(100), nullable=False)  # e.g. "30/03/2026 — 12/04/2026"
    first_date = db.Column(db.Date, nullable=False)
    last_date = db.Column(db.Date, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return f'<ArchivePeriod {self.label}>'


class ArchivedUpload(db.Model):
    __tablename__ = 'archived_uploads'

    id = db.Column(db.Integer, primary_key=True)
    period_id = db.Column(db.Integer, db.ForeignKey('archive_periods.id'), nullable=False)
    original_id = db.Column(db.Integer, nullable=False)
    filename = db.Column(db.String(200), nullable=False)
    upload_date = db.Column(db.Date, nullable=False)
    created_at = db.Column(db.DateTime)

    period = db.relationship('ArchivePeriod', backref='uploads')


class ArchivedPlayerStats(db.Model):
    __tablename__ = 'archived_player_stats'

    id = db.Column(db.Integer, primary_key=True)
    period_id = db.Column(db.Integer, db.ForeignKey('archive_periods.id'), nullable=False, index=True)
    upload_id = db.Column(db.Integer, nullable=False)
    player_id = db.Column(db.String(20), nullable=False, index=True)
    nickname = db.Column(db.String(100), nullable=False)
    club = db.Column(db.String(100), nullable=False, index=True)
    sa_id = db.Column(db.String(20), default='', index=True)
    agent_id = db.Column(db.String(20), default='', index=True)
    role = db.Column(db.String(30), default='')
    pnl = db.Column(db.Float, default=0)
    rake = db.Column(db.Float, default=0)
    hands = db.Column(db.Float, default=0)

    period = db.relationship('ArchivePeriod', backref='stats')


class ArchivedPlayerSession(db.Model):
    __tablename__ = 'archived_player_sessions'

    id = db.Column(db.Integer, primary_key=True)
    period_id = db.Column(db.Integer, db.ForeignKey('archive_periods.id'), nullable=False)
    upload_id = db.Column(db.Integer, nullable=False)
    player_id = db.Column(db.String(20), nullable=False)
    game_type = db.Column(db.String(20), nullable=False)
    table_name = db.Column(db.String(200), nullable=False)
    blinds = db.Column(db.String(20), default='')
    pnl = db.Column(db.Float, default=0)


class ArchivedTournamentStats(db.Model):
    __tablename__ = 'archived_tournament_stats'

    id = db.Column(db.Integer, primary_key=True)
    period_id = db.Column(db.Integer, db.ForeignKey('archive_periods.id'), nullable=False)
    upload_id = db.Column(db.Integer, nullable=False)
    title = db.Column(db.String(200), nullable=False)
    status = db.Column(db.String(30), default='')
    game_type = db.Column(db.String(30), default='')
    buyin = db.Column(db.Float, default=0)
    fee = db.Column(db.Float, default=0)
    reentry = db.Column(db.String(20), default='')
    gtd = db.Column(db.Float, default=0)
    entries = db.Column(db.Float, default=0)
    prize_pool = db.Column(db.Float, default=0)
    start = db.Column(db.String(20), default='')
    duration = db.Column(db.String(20), default='')


class CycleSummaryReport(db.Model):
    """Persisted Excel snapshot of a cycle summary, created ONLY at cycle
    reset (in _archive_and_clear_active). Never regenerated or overwritten —
    this is the historical record of a closed cycle. Retention: 180 days
    (see app/__init__).

    No FK to ArchivePeriod: archive data is cleaned up at 90 days while
    cycle summaries live 180 days, so the report must survive its period's
    deletion. `period_label` carries the display string.

    The "current cycle" is NOT stored here — /admin/cycle-summary.xlsx
    builds it on demand from live tables."""
    __tablename__ = 'cycle_summary_reports'

    id = db.Column(db.Integer, primary_key=True)
    period_label = db.Column(db.String(100), default='')
    filename = db.Column(db.String(200), nullable=False)
    content = db.Column(db.LargeBinary, nullable=False)
    generated_at = db.Column(db.DateTime, default=datetime.utcnow, index=True)
    # Distinguishes the "live" snapshot (regenerated on demand) from a
    # closed-cycle snapshot saved at reset. Historical rows are is_current=False
    # and subject to the 180-day cleanup; the True case is reserved.
    is_current = db.Column(db.Boolean, nullable=False, default=False)


class PlayerAssignment(db.Model):
    """Manual override of a player's sa_id / agent_id.

    The Excel/PPPoker file is treated as read-only. When an admin wants to
    attach a player to a specific SA or agent, we store the override here.
    Dashboards and exports apply it at display time via
    `union_data.apply_player_overrides`."""
    __tablename__ = 'player_assignments'

    id = db.Column(db.Integer, primary_key=True)
    player_id = db.Column(db.String(20), unique=True, nullable=False, index=True)
    assigned_sa_id = db.Column(db.String(20), default='')
    assigned_agent_id = db.Column(db.String(20), default='')
    assigned_by_user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    note = db.Column(db.String(200), default='')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    assigned_by = db.relationship('User', backref='player_assignments')

    def __repr__(self):
        return f'<PlayerAssignment {self.player_id} → sa={self.assigned_sa_id} ag={self.assigned_agent_id}>'


class ProtectedAgent(db.Model):
    """Per-agent extra-password gate. When a row exists for an sa_id, the
    agent's card on the admin overview is blurred and clicking through to
    the agent dashboard requires re-entering this password. Removing the
    row removes the protection (no migration of existing data).

    The password is stored as a werkzeug hash, identical to user passwords.
    Unlock state lives in the Flask session (10-minute window) — see
    `auth.unlock_protected_agent`.
    """
    __tablename__ = 'protected_agents'

    sa_id = db.Column(db.String(20), primary_key=True)
    password_hash = db.Column(db.String(256), nullable=False)
    created_by_user_id = db.Column(db.Integer, db.ForeignKey('users.id'),
                                   nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow,
                           onupdate=datetime.utcnow)

    created_by = db.relationship('User', backref='protected_agents')

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def __repr__(self):
        return f'<ProtectedAgent {self.sa_id}>'


class BotSuspectDismissal(db.Model):
    """A player the admin has already reviewed for bot suspicion and
    decided is fine. Used to filter the player out of the /admin/bot-suspects
    list on subsequent visits.

    Restoring (un-dismissing) is supported — admin can change their mind
    later and put the player back on the list."""
    __tablename__ = 'bot_suspect_dismissals'

    id = db.Column(db.Integer, primary_key=True)
    player_id = db.Column(db.String(20), unique=True, nullable=False, index=True)
    dismissed_by_user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    note = db.Column(db.String(200), default='')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    dismissed_by = db.relationship('User', backref='bot_dismissals')

    def __repr__(self):
        return f'<BotSuspectDismissal {self.player_id}>'
