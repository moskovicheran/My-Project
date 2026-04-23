import os
import warnings
from datetime import timedelta

BASE_DIR = os.path.abspath(os.path.dirname(__file__))


class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY')
    if not SECRET_KEY:
        SECRET_KEY = os.urandom(24).hex()
        warnings.warn('SECRET_KEY not set — using random key. Sessions will reset on restart.')

    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = 'Lax'
    SESSION_COOKIE_SECURE = os.environ.get('SESSION_COOKIE_SECURE', 'false').lower() == 'true'

    # CSRF token tied to session lifetime (no fixed expiry) — avoids
    # "CSRF token has expired" on login forms left open for a while.
    WTF_CSRF_TIME_LIMIT = None

    # Keep users logged in for 30 days when session.permanent is set.
    PERMANENT_SESSION_LIFETIME = timedelta(days=30)

    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16 MB upload limit

    database_url = os.environ.get('DATABASE_URL_POOLER') or \
        os.environ.get('DATABASE_URL') or \
        'sqlite:///' + os.path.join(BASE_DIR, 'finance.db')
    if database_url.startswith('postgres://'):
        database_url = database_url.replace('postgres://', 'postgresql://', 1)
    SQLALCHEMY_DATABASE_URI = database_url

    SQLALCHEMY_TRACK_MODIFICATIONS = False
    TEMPLATES_AUTO_RELOAD = True

    # Connection-pool tuning for Neon / Postgres. Not applied to SQLite
    # (local dev) because StaticPool doesn't support these kwargs.
    #   pool_pre_ping — ping the connection before checking it out so dead
    #                   connections from Neon's scale-to-zero are replaced
    #                   transparently instead of raising "server closed
    #                   the connection unexpectedly".
    #   pool_recycle — force recycling of connections older than 30 min
    #                  to avoid stale-connection errors.
    #   pool_timeout=10 — fail fast (10s) when the pool is exhausted
    #                     instead of hanging for 30s while the user waits.
    if not database_url.startswith('sqlite'):
        SQLALCHEMY_ENGINE_OPTIONS = {
            'pool_pre_ping': True,
            'pool_recycle': 1800,
            'pool_timeout': 10,
        }
