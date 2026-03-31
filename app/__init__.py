import os
from flask import Flask
from flask_login import LoginManager
from config import Config
from app.models import db, User


def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    db.init_app(app)

    login_manager = LoginManager(app)
    login_manager.login_view = 'auth.login'
    login_manager.login_message = 'יש להתחבר כדי לגשת לדף זה.'
    login_manager.login_message_category = 'warning'

    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))

    from app.routes.auth import auth_bp
    from app.routes.main import main_bp
    from app.routes.union import union_bp
    from app.routes.upload import upload_bp
    from app.routes.admin import admin_bp
    app.register_blueprint(auth_bp)
    app.register_blueprint(main_bp)
    app.register_blueprint(union_bp)
    app.register_blueprint(upload_bp)
    app.register_blueprint(admin_bp)

    app.jinja_env.filters['enumerate'] = enumerate

    with app.app_context():
        db.create_all()
        # Migrations
        import sqlalchemy
        def _try_add_col(col, ddl):
            try:
                db.session.execute(sqlalchemy.text(f"SELECT {col} FROM users LIMIT 1"))
            except Exception:
                db.session.rollback()
                db.session.execute(sqlalchemy.text(ddl))
                db.session.commit()
        _try_add_col('role', "ALTER TABLE users ADD COLUMN role VARCHAR(20) NOT NULL DEFAULT 'admin'")
        _try_add_col('player_id', "ALTER TABLE users ADD COLUMN player_id VARCHAR(20)")
        # MoneyTransfer migrations
        def _try_add_col_table(table, col, ddl):
            try:
                db.session.execute(sqlalchemy.text(f"SELECT {col} FROM {table} LIMIT 1"))
            except Exception:
                db.session.rollback()
                db.session.execute(sqlalchemy.text(ddl))
                db.session.commit()
        _try_add_col_table('money_transfers', 'from_player_id',
                           "ALTER TABLE money_transfers ADD COLUMN from_player_id VARCHAR(20) NOT NULL DEFAULT ''")
        _try_add_col_table('money_transfers', 'to_player_id',
                           "ALTER TABLE money_transfers ADD COLUMN to_player_id VARCHAR(20) NOT NULL DEFAULT ''")

        # Cleanup data older than 60 days
        from datetime import datetime, timedelta
        from app.models import DailyUpload, DailyPlayerStats
        cutoff = datetime.utcnow().date() - timedelta(days=60)
        old_uploads = DailyUpload.query.filter(DailyUpload.upload_date < cutoff).all()
        if old_uploads:
            for u in old_uploads:
                db.session.delete(u)  # cascade deletes DailyPlayerStats
            db.session.commit()

        # Load active excel file - if _active.txt exists use it, otherwise keep default
        active_file = os.path.join(os.path.dirname(__file__), '..', 'uploads', '_active.txt')
        from app.union_data import set_excel_path
        if os.path.exists(active_file):
            with open(active_file, 'r', encoding='utf-8') as f:
                path = f.read().strip()
            if path and os.path.exists(path):
                set_excel_path(path)
            else:
                set_excel_path('')  # Reset state - no file

    return app
