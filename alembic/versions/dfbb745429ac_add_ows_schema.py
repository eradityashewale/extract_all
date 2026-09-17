"""add ows schema

Revision ID: dfbb745429ac
Revises: c1bf39a49740
Create Date: 2026-09-17 22:24:53.090719

Adds the OWS (one-word-substitution) tables from db/schema.sql -- source_files
and the vocab/idioms tables from the initial migration are untouched. Mirrors
how idioms was structured (one row per word, one quiz per word, one image per
word), plus a per-quiz-option `description` column like vocab's quiz_options.
"""
from typing import Sequence, Union

from alembic import op

# revision identifiers, used by Alembic.
revision: str = 'dfbb745429ac'
down_revision: Union[str, Sequence[str], None] = 'c1bf39a49740'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.execute("""
        CREATE TABLE IF NOT EXISTS ows_words (
            id                   SERIAL PRIMARY KEY,
            source_file_id       INTEGER NOT NULL REFERENCES source_files(id) ON DELETE CASCADE,
            word_number          INTEGER NOT NULL,
            word                 TEXT NOT NULL,
            part_of_speech       TEXT,
            meaning              TEXT NOT NULL,
            meaning_translation  TEXT,
            example_text         TEXT,
            created_at           TIMESTAMPTZ NOT NULL DEFAULT now(),
            UNIQUE (source_file_id, word_number)
        );

        CREATE TABLE IF NOT EXISTS ows_quizzes (
            id           SERIAL PRIMARY KEY,
            ows_word_id  INTEGER NOT NULL UNIQUE REFERENCES ows_words(id) ON DELETE CASCADE,
            question     TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS ows_quiz_options (
            id            SERIAL PRIMARY KEY,
            quiz_id       INTEGER NOT NULL REFERENCES ows_quizzes(id) ON DELETE CASCADE,
            option_label  CHAR(1) NOT NULL CHECK (option_label IN ('A', 'B', 'C', 'D')),
            option_text   TEXT NOT NULL,
            is_correct    BOOLEAN NOT NULL DEFAULT FALSE,
            description   TEXT,
            UNIQUE (quiz_id, option_label)
        );

        CREATE TABLE IF NOT EXISTS ows_images (
            id                SERIAL PRIMARY KEY,
            ows_word_id       INTEGER NOT NULL UNIQUE REFERENCES ows_words(id) ON DELETE CASCADE,
            file_path         TEXT NOT NULL,
            content_type      TEXT,
            file_size_bytes   INTEGER
        );

        CREATE UNIQUE INDEX IF NOT EXISTS idx_ows_words_word_unique ON ows_words (lower(word));
        CREATE INDEX IF NOT EXISTS idx_ows_quiz_options_quiz_id ON ows_quiz_options (quiz_id);
    """)


def downgrade() -> None:
    op.execute("""
        DROP TABLE IF EXISTS ows_quiz_options;
        DROP TABLE IF EXISTS ows_quizzes;
        DROP TABLE IF EXISTS ows_images;
        DROP TABLE IF EXISTS ows_words;
    """)
