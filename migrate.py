"""
Runs db/schema.sql against DATABASE_URL (read from .env, or the environment).

Usage:
    python migrate.py
"""

import os
import re
import sys
from pathlib import Path

import psycopg2


def load_database_url() -> str:
    if os.environ.get("DATABASE_URL"):
        return os.environ["DATABASE_URL"]

    env_path = Path(".env")
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8").splitlines():
            m = re.match(r"^\s*DATABASE_URL\s*=\s*(.+?)\s*$", line)
            if m:
                return m.group(1).strip().strip('"').strip("'")

    print("DATABASE_URL not found in the environment or .env", file=sys.stderr)
    sys.exit(1)


def main():
    db_url = load_database_url()
    schema_sql = Path("db/schema.sql").read_text(encoding="utf-8")

    conn = psycopg2.connect(db_url)
    try:
        with conn.cursor() as cur:
            cur.execute(schema_sql)
        conn.commit()
        print("Migration applied: schema.sql executed successfully.")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
