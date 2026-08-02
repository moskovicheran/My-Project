import os
from flask import Flask
from flask_login import LoginManager
from flask_wtf.csrf import CSRFProtect
from config import Config
from app.models import db, User

csrf = CSRFProtect()


def _init_sentry():
    """Activate Sentry only when SENTRY_DSN is set — keeps local dev silent
    and optional on prod (won't crash if the package isn't installed)."""
    dsn = os.environ.get('SENTRY_DSN')
    if not dsn:
        return
    try:
        import sentry_sdk
        from sentry_sdk.integrations.flask import FlaskIntegration
        from sentry_sdk.integrations.sqlalchemy import SqlalchemyIntegration
        sentry_sdk.init(
            dsn=dsn,
            integrations=[FlaskIntegration(), SqlalchemyIntegration()],
            # Performance monitoring sample rate — 10% of requests traced
            # is enough signal without bloating the free-tier event budget.
            traces_sample_rate=float(os.environ.get('SENTRY_TRACES_RATE', '0.1')),
            send_default_pii=False,  # don't send cookies / auth headers
            environment=os.environ.get('FLASK_ENV', 'production'),
        )
    except ImportError:
        pass


def create_app():
    _init_sentry()
    app = Flask(__name__)
    app.config.from_object(Config)

    db.init_app(app)
    csrf.init_app(app)

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
    def inject_asset_version():
        # Cache-busting token for static assets — changes whenever the file is
        # rewritten (every deploy), so browsers fetch the fresh JS instead of a
        # stale cached copy.
        try:
            _p = os.path.join(os.path.dirname(__file__), 'static', 'js', 'transfer_form.js')
            return {'asset_v': str(int(os.path.getmtime(_p)))}
        except Exception:
            return {'asset_v': '1'}

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
            # Warn when archive is 85+ days old (created_at based), deleted at 90
            cutoff_85 = datetime.utcnow() - timedelta(days=85)
            expiring = ArchivePeriod.query.filter(ArchivePeriod.created_at < cutoff_85).all()
            if expiring:
                warnings = []
                for p in expiring:
                    age_days = (datetime.utcnow() - p.created_at).days
                    days_left = 90 - age_days
                    if days_left > 0:
                        warnings.append({'label': p.label, 'days_left': days_left})
                return {'archive_warnings': warnings}
        except Exception:
            pass
        return {'archive_warnings': []}

    # Slow startup DB work (create_all + archive cleanup) can be skipped
    # locally when Neon is slow to respond over a home connection — set
    # SKIP_STARTUP_DB_WORK=1 in the env. The two operations are only needed
    # on first-ever startup (create_all) and as maintenance (cleanup);
    # skipping them on every dev restart is safe.
    skip_startup_db = os.environ.get('SKIP_STARTUP_DB_WORK', '').lower() in ('1', 'true', 'yes')

    with app.app_context():
        # Tables added AFTER the original schema must be created even when
        # SKIP_STARTUP_DB_WORK is set, since the deployed DB doesn't have
        # them yet. Idempotent (checkfirst=True) and very fast on warm DBs.
        try:
            from app.models import (BotSuspectDismissal, ProtectedAgent,
                                     CollectionCycle, PlayerPayment, PlayerCross)
            for _newtbl in (BotSuspectDismissal, ProtectedAgent,
                            CollectionCycle, PlayerPayment, PlayerCross):
                _newtbl.__table__.create(db.engine, checkfirst=True)
            # `frozen` was added to collection_cycles after it first shipped.
            # Ensure it exists here too — the main column-migration block
            # below is skipped when SKIP_STARTUP_DB_WORK is set (as in prod).
            from sqlalchemy import text as _text
            try:
                db.session.execute(_text(
                    'ALTER TABLE collection_cycles ADD COLUMN IF NOT EXISTS '
                    'frozen BOOLEAN NOT NULL DEFAULT FALSE'))
                db.session.commit()
            except Exception:
                db.session.rollback()
            # Tournament status ('Ended' / 'In Progress') — added after the
            # tables shipped, so prod needs the explicit ALTER here too.
            for _tbl in ('tournament_stats', 'archived_tournament_stats'):
                try:
                    db.session.execute(_text(
                        f'ALTER TABLE {_tbl} ADD COLUMN IF NOT EXISTS '
                        "status VARCHAR(30) DEFAULT ''"))
                    db.session.commit()
                except Exception:
                    db.session.rollback()
            # Widen money_transfers id columns so synthetic counterparty ids
            # like '__club__<club name>' fit (were VARCHAR(20)). Postgres
            # widens a varchar as a fast metadata-only change; runs on every
            # boot (incl. prod, where create_all is skipped) and is a no-op
            # once applied. Guarded so SQLite (no such ALTER) just passes.
            try:
                db.session.execute(_text(
                    'ALTER TABLE money_transfers '
                    'ALTER COLUMN from_player_id TYPE VARCHAR(120)'))
                db.session.execute(_text(
                    'ALTER TABLE money_transfers '
                    'ALTER COLUMN to_player_id TYPE VARCHAR(120)'))
                db.session.commit()
            except Exception:
                db.session.rollback()
        except Exception as e:
            print(f"New-table create warning: {e}")

        if not skip_startup_db:
            try:
                db.create_all()
            except Exception as e:
                print(f"DB create_all warning: {e}")

            # Ensure indexes exist on hot columns for tables that were
            # created BEFORE index=True was added to the model. create_all()
            # only adds indexes on fresh tables; existing tables keep the
            # old schema until we explicitly CREATE INDEX. Safe to run on
            # every boot — CREATE INDEX IF NOT EXISTS is idempotent.
            try:
                from sqlalchemy import text
                hot_indexes = [
                    ('ix_daily_player_stats_player_id', 'daily_player_stats', 'player_id'),
                    ('ix_daily_player_stats_club',      'daily_player_stats', 'club'),
                    ('ix_daily_player_stats_sa_id',     'daily_player_stats', 'sa_id'),
                    ('ix_daily_player_stats_agent_id',  'daily_player_stats', 'agent_id'),
                    ('ix_daily_player_stats_upload_id', 'daily_player_stats', 'upload_id'),
                    ('ix_archived_player_stats_period_id', 'archived_player_stats', 'period_id'),
                    ('ix_archived_player_stats_player_id', 'archived_player_stats', 'player_id'),
                    ('ix_archived_player_stats_club',      'archived_player_stats', 'club'),
                    ('ix_archived_player_stats_sa_id',     'archived_player_stats', 'sa_id'),
                    ('ix_archived_player_stats_agent_id',  'archived_player_stats', 'agent_id'),
                    ('ix_tournament_stats_upload_id',      'tournament_stats',      'upload_id'),
                ]
                for idx_name, tbl, col in hot_indexes:
                    try:
                        db.session.execute(text(
                            f'CREATE INDEX IF NOT EXISTS {idx_name} ON {tbl} ({col})'
                        ))
                    except Exception as _e:
                        # Per-index failure shouldn't block startup (e.g.
                        # table doesn't exist yet on a clean install).
                        pass
                db.session.commit()

                # Add columns that were introduced AFTER the table already
                # existed in prod. create_all() never alters existing tables,
                # so columns added later need an explicit idempotent ALTER.
                # Wrapped per-statement so SQLite (no IF NOT EXISTS for ADD
                # COLUMN) just falls through the except.
                column_adds = [
                    ('cycle_summary_reports', 'is_current',
                     'BOOLEAN NOT NULL DEFAULT FALSE'),
                    ('collection_cycles', 'frozen',
                     'BOOLEAN NOT NULL DEFAULT FALSE'),
                    ('tournament_stats', 'status', "VARCHAR(30) DEFAULT ''"),
                    ('archived_tournament_stats', 'status',
                     "VARCHAR(30) DEFAULT ''"),
                ]
                for tbl, col, coldef in column_adds:
                    try:
                        db.session.execute(text(
                            f'ALTER TABLE {tbl} ADD COLUMN IF NOT EXISTS {col} {coldef}'
                        ))
                        db.session.commit()
                    except Exception:
                        db.session.rollback()
                        try:
                            db.session.execute(text(
                                f'ALTER TABLE {tbl} ADD COLUMN {col} {coldef}'
                            ))
                            db.session.commit()
                        except Exception:
                            db.session.rollback()
            except Exception as e:
                db.session.rollback()
                print(f"Index creation warning: {e}")

            # Cleanup archived data older than 90 days (from archive creation date)
            try:
                from datetime import datetime, timedelta
                from app.models import ArchivePeriod, ArchivedUpload, ArchivedPlayerStats, ArchivedPlayerSession, ArchivedTournamentStats
                cutoff = datetime.utcnow() - timedelta(days=90)
                old_periods = ArchivePeriod.query.filter(ArchivePeriod.created_at < cutoff).all()
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

            # Cleanup cycle summary reports older than 180 days. Kept twice
            # as long as raw archive data so admins can still pull an older
            # settlement Excel even after the underlying rows are gone.
            try:
                from datetime import datetime, timedelta
                from app.models import CycleSummaryReport
                cs_cutoff = datetime.utcnow() - timedelta(days=180)
                CycleSummaryReport.query.filter(
                    CycleSummaryReport.is_current == False,  # noqa: E712
                    CycleSummaryReport.generated_at < cs_cutoff,
                ).delete(synchronize_session=False)
                db.session.commit()
            except Exception:
                pass

        # Load active excel file if exists (local only) — fast, keep always.
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

    @app.after_request
    def set_security_headers(response):
        response.headers['X-Content-Type-Options'] = 'nosniff'
        response.headers['X-Frame-Options'] = 'SAMEORIGIN'
        response.headers['X-XSS-Protection'] = '1; mode=block'
        response.headers['Referrer-Policy'] = 'strict-origin-when-cross-origin'
        return response

    return app
