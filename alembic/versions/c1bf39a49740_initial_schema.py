"""initial schema

Revision ID: c1bf39a49740
Revises:
Create Date: 2026-09-17 15:40:09.640458

Runs db/schema.sql verbatim -- that file stays the single source of truth
for the schema (its CREATE TABLE/INDEX ... IF NOT EXISTS guards also make it
safe to run against a database that already has these tables). Later schema
changes get their own revision on top of this one rather than edits here.
"""
from pathlib import Path
from typing import Sequence, Union

from alembic import op

# revision identifiers, used by Alembic.
revision: str = 'c1bf39a49740'
down_revision: Union[str, Sequence[str], None] = None
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None

SCHEMA_SQL_PATH = Path(__file__).resolve().parents[2] / "db" / "schema.sql"


def upgrade() -> None:
    op.execute(SCHEMA_SQL_PATH.read_text(encoding="utf-8"))


def downgrade() -> None:
    op.execute("""
        DROP TABLE IF EXISTS idiom_quiz_options;
        DROP TABLE IF EXISTS idiom_quizzes;
        DROP TABLE IF EXISTS idiom_images;
        DROP TABLE IF EXISTS idioms;
        DROP TABLE IF EXISTS quiz_options;
        DROP TABLE IF EXISTS quizzes;
        DROP TABLE IF EXISTS word_images;
        DROP TABLE IF EXISTS word_synonyms;
        DROP TABLE IF EXISTS word_examples;
        DROP TABLE IF EXISTS words;
        DROP TABLE IF EXISTS source_files;
    """)
