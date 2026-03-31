import sqlite3
import sqlalchemy
from sqlalchemy import text

LOCAL_DB = 'C:/Project/finance.db'
NEON_URL = "postgresql://neondb_owner:npg_H78sMvjYPLWu@ep-sparkling-thunder-am3ip16r-pooler.c-5.us-east-1.aws.neon.tech/neondb?sslmode=require"

TABLES = [
    {
        'name': 'users',
        'columns': 'id, username, email, password_hash, created_at, role, player_id',
        'conflict': 'username',
    },
    {
        'name': 'sa_hierarchy',
        'columns': 'id, parent_sa_id, child_sa_id',
        'conflict': 'child_sa_id',
    },
    {
        'name': 'sa_rake_config',
        'columns': 'id, sa_id, rake_percent, managed_club_id',
        'conflict': 'sa_id',
    },
    {
        'name': 'rake_config',
        'columns': 'id, entity_type, entity_id, entity_name, rake_percent',
        'conflict': None,
    },
    {
        'name': 'shared_expenses',
        'columns': 'id, user_id, description, amount, charged, created_at',
        'conflict': None,
    },
    {
        'name': 'expense_charges',
        'columns': 'id, expense_id, agent_player_id, agent_name, charge_amount, created_at',
        'conflict': None,
    },
    {
        'name': 'money_transfers',
        'columns': 'id, user_id, from_player_id, from_name, to_player_id, to_name, amount, description, created_at',
        'conflict': None,
    },
    {
        'name': 'admin_notes',
        'columns': 'id, user_id, content, created_at',
        'conflict': None,
    },
]


def migrate():
    print("--- Reading from local SQLite ---")
    sqlite_conn = sqlite3.connect(LOCAL_DB)
    sqlite_cursor = sqlite_conn.cursor()

    print("--- Connecting to Neon ---")
    engine = sqlalchemy.create_engine(NEON_URL)

    with engine.begin() as neon_conn:
        for table in TABLES:
            name = table['name']
            cols = table['columns']
            col_list = [c.strip() for c in cols.split(',')]

            try:
                sqlite_cursor.execute(f"SELECT {cols} FROM {name}")
                rows = sqlite_cursor.fetchall()
            except Exception as e:
                print(f"  {name}: skip ({e})")
                continue

            if not rows:
                print(f"  {name}: 0 rows (empty)")
                continue

            params = ', '.join([f':{c}' for c in col_list])
            conflict = ''
            if table['conflict']:
                conflict = f" ON CONFLICT ({table['conflict']}) DO NOTHING"
            query = text(f'INSERT INTO {name} ({cols}) VALUES ({params}){conflict}')

            count = 0
            for row in rows:
                data = {}
                for i in range(len(col_list)):
                    val = row[i]
                    # Fix SQLite boolean (1/0) → Python bool for PostgreSQL
                    if col_list[i] == 'charged' and val in (0, 1):
                        val = bool(val)
                    data[col_list[i]] = val
                try:
                    neon_conn.execute(query, data)
                    count += 1
                except Exception as e:
                    print(f"    skip row: {str(e)[:80]}")

            print(f"  {name}: {count}/{len(rows)} rows migrated")

    sqlite_conn.close()
    print("\n--- Migration complete! ---")


if __name__ == "__main__":
    migrate()
