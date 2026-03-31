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
        try:
            db.create_all()
        except Exception as e:
            print(f"DB create_all warning: {e}")

        # Cleanup data older than 60 days
        try:
            from datetime import datetime, timedelta
            from app.models import DailyUpload
            cutoff = datetime.utcnow().date() - timedelta(days=60)
            old_uploads = DailyUpload.query.filter(DailyUpload.upload_date < cutoff).all()
            if old_uploads:
                for u in old_uploads:
                    db.session.delete(u)
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
