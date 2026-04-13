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

    def comma_format(value):
        """Format number with commas: 29344.97 -> 29,344.97"""
        try:
            val = float(value)
            if val == int(val) and '.' not in str(value):
                return f'{int(val):,}'
            return f'{val:,.2f}'
        except (ValueError, TypeError):
            return value
    app.jinja_env.filters['comma'] = comma_format

    @app.context_processor
    def inject_last_upload():
        from datetime import timedelta
        from app.models import DailyUpload
        try:
            last = DailyUpload.query.order_by(DailyUpload.created_at.desc()).first()
            if last and last.created_at:
                il_time = last.created_at + timedelta(hours=3)
                return {'last_upload_time': il_time.strftime('%d/%m/%Y %H:%M')}
        except Exception:
            pass
        return {'last_upload_time': None}

    @app.context_processor
    def inject_archive_warnings():
        from datetime import datetime, timedelta
        from app.models import ArchivePeriod
        try:
            cutoff_85 = datetime.utcnow().date() - timedelta(days=85)
            expiring = ArchivePeriod.query.filter(ArchivePeriod.last_date < cutoff_85).all()
            if expiring:
                warnings = []
                for p in expiring:
                    days_left = 90 - (datetime.utcnow().date() - p.last_date).days
                    if days_left > 0:
                        warnings.append({'label': p.label, 'days_left': days_left})
                return {'archive_warnings': warnings}
        except Exception:
            pass
        return {'archive_warnings': []}

    with app.app_context():
        try:
            db.create_all()
        except Exception as e:
            print(f"DB create_all warning: {e}")

        # Cleanup archived data older than 90 days
        try:
            from datetime import datetime, timedelta
            from app.models import ArchivePeriod, ArchivedUpload, ArchivedPlayerStats, ArchivedPlayerSession, ArchivedTournamentStats
            cutoff = datetime.utcnow().date() - timedelta(days=90)
            old_periods = ArchivePeriod.query.filter(ArchivePeriod.last_date < cutoff).all()
            if old_periods:
                old_ids = [p.id for p in old_periods]
                ArchivedTournamentStats.query.filter(ArchivedTournamentStats.period_id.in_(old_ids)).delete(synchronize_session=False)
                ArchivedPlayerSession.query.filter(ArchivedPlayerSession.period_id.in_(old_ids)).delete(synchronize_session=False)
                ArchivedPlayerStats.query.filter(ArchivedPlayerStats.period_id.in_(old_ids)).delete(synchronize_session=False)
                ArchivedUpload.query.filter(ArchivedUpload.period_id.in_(old_ids)).delete(synchronize_session=False)
                ArchivePeriod.query.filter(ArchivePeriod.id.in_(old_ids)).delete(synchronize_session=False)
                db.session.commit()
        except Exception:
            pass

        # Load active excel file if exists (local only)
        try:
            active_file = os.path.join(os.path.dirname(__file__), '..', 'uploads', '_active.txt')
            from app.union_data import set_excel_path
            if os.path.exists(active_file):
                with open(active_file, 'r', encoding='utf-8') as f:
                    path = f.read().strip()
                if path and os.path.exists(path):
                    set_excel_path(path)
                else:
                    set_excel_path('')
        except Exception:
            pass

    return app
