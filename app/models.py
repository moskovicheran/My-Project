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


class MoneyTransfer(db.Model):
    __tablename__ = 'money_transfers'

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    from_player_id = db.Column(db.String(20), nullable=False)
    from_name = db.Column(db.String(100), nullable=False)
    to_player_id = db.Column(db.String(20), nullable=False)
    to_name = db.Column(db.String(100), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    description = db.Column(db.String(200))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    user = db.relationship('User', backref='transfers')


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
    upload_id = db.Column(db.Integer, db.ForeignKey('daily_uploads.id'), nullable=False)
    player_id = db.Column(db.String(20), nullable=False)
    nickname = db.Column(db.String(100), nullable=False)
    club = db.Column(db.String(100), nullable=False)
    sa_id = db.Column(db.String(20), default='')
    agent_id = db.Column(db.String(20), default='')
    role = db.Column(db.String(30), default='')
    pnl = db.Column(db.Float, default=0)
    rake = db.Column(db.Float, default=0)
    hands = db.Column(db.Float, default=0)

    def __repr__(self):
        return f'<DailyPlayerStats {self.nickname} pnl={self.pnl}>'
