-- Vocab database schema
-- One row in source_files per imported .docx. Everything else fans out from words.

CREATE TABLE IF NOT EXISTS source_files (
    id          SERIAL PRIMARY KEY,
    file_name   TEXT NOT NULL UNIQUE,
    imported_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- Re-run of the same source file updates imported_at instead of erroring.

CREATE TABLE IF NOT EXISTS words (
    id                   SERIAL PRIMARY KEY,
    source_file_id       INTEGER NOT NULL REFERENCES source_files(id) ON DELETE CASCADE,
    word_number          INTEGER NOT NULL,          -- position within the file (1-10)
    word                 TEXT NOT NULL,
    part_of_speech       TEXT,                       -- n, v, adj, ph.v, ...
    meaning              TEXT NOT NULL,
    meaning_translation  TEXT,                       -- e.g. the Hindi gloss in parentheses
    hint                 TEXT,                       -- mnemonic, when present
    created_at           TIMESTAMPTZ NOT NULL DEFAULT now(),
    UNIQUE (source_file_id, word_number)
);

CREATE TABLE IF NOT EXISTS word_examples (
    id            SERIAL PRIMARY KEY,
    word_id       INTEGER NOT NULL REFERENCES words(id) ON DELETE CASCADE,
    example_text  TEXT NOT NULL,
    sort_order    INTEGER NOT NULL
);

CREATE TABLE IF NOT EXISTS word_synonyms (
    id          SERIAL PRIMARY KEY,
    word_id     INTEGER NOT NULL REFERENCES words(id) ON DELETE CASCADE,
    synonym     TEXT NOT NULL,
    sort_order  INTEGER NOT NULL
);

-- One quiz per word (the doc only ever has one "Quiz -" question per entry).
CREATE TABLE IF NOT EXISTS quizzes (
    id        SERIAL PRIMARY KEY,
    word_id   INTEGER NOT NULL UNIQUE REFERENCES words(id) ON DELETE CASCADE,
    question  TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS quiz_options (
    id            SERIAL PRIMARY KEY,
    quiz_id       INTEGER NOT NULL REFERENCES quizzes(id) ON DELETE CASCADE,
    option_label  CHAR(1) NOT NULL CHECK (option_label IN ('A', 'B', 'C', 'D')),
    option_text   TEXT NOT NULL,
    is_correct    BOOLEAN NOT NULL DEFAULT FALSE,
    description   TEXT,                              -- definition of the option, when the doc gives one (usually the wrong answers)
    UNIQUE (quiz_id, option_label)
);

-- Image stays on disk; this row is just the pointer + metadata.
CREATE TABLE IF NOT EXISTS word_images (
    id                SERIAL PRIMARY KEY,
    word_id           INTEGER NOT NULL UNIQUE REFERENCES words(id) ON DELETE CASCADE,
    file_path         TEXT NOT NULL,                 -- relative to MEDIA_ROOT, e.g. vocab-1/faction.jpeg
    content_type      TEXT,
    file_size_bytes   INTEGER
);

-- Global upsert key: importing a word that already exists (by name, case-insensitive)
-- updates that row instead of inserting a duplicate. Drop the old plain index first
-- in case this schema was created before this became unique.
DROP INDEX IF EXISTS idx_words_word;
CREATE UNIQUE INDEX IF NOT EXISTS idx_words_word_unique ON words (lower(word));
CREATE INDEX IF NOT EXISTS idx_word_examples_word_id ON word_examples (word_id);
CREATE INDEX IF NOT EXISTS idx_word_synonyms_word_id ON word_synonyms (word_id);
CREATE INDEX IF NOT EXISTS idx_quiz_options_quiz_id ON quiz_options (quiz_id);

-- Idioms database schema
-- One row in source_files (shared with vocab) per imported "Idioms - N (...).docx".
-- Unlike vocab, these files carry no part of speech, synonyms, or hint, and
-- exactly one example per idiom -- so idioms keeps that as a plain column
-- instead of a child table.

CREATE TABLE IF NOT EXISTS idioms (
    id                   SERIAL PRIMARY KEY,
    source_file_id       INTEGER NOT NULL REFERENCES source_files(id) ON DELETE CASCADE,
    idiom_number         INTEGER NOT NULL,          -- position within the file (1-10)
    idiom                TEXT NOT NULL,
    meaning              TEXT NOT NULL,
    meaning_translation  TEXT,                       -- e.g. the Hindi gloss in parentheses
    example_text         TEXT,
    created_at           TIMESTAMPTZ NOT NULL DEFAULT now(),
    UNIQUE (source_file_id, idiom_number)
);

-- One quiz per idiom (the doc only ever has one "Quiz -" question per entry).
CREATE TABLE IF NOT EXISTS idiom_quizzes (
    id        SERIAL PRIMARY KEY,
    idiom_id  INTEGER NOT NULL UNIQUE REFERENCES idioms(id) ON DELETE CASCADE,
    question  TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS idiom_quiz_options (
    id            SERIAL PRIMARY KEY,
    quiz_id       INTEGER NOT NULL REFERENCES idiom_quizzes(id) ON DELETE CASCADE,
    option_label  CHAR(1) NOT NULL CHECK (option_label IN ('A', 'B', 'C', 'D')),
    option_text   TEXT NOT NULL,
    is_correct    BOOLEAN NOT NULL DEFAULT FALSE,
    UNIQUE (quiz_id, option_label)
);

-- Image stays on disk; this row is just the pointer + metadata.
CREATE TABLE IF NOT EXISTS idiom_images (
    id                SERIAL PRIMARY KEY,
    idiom_id          INTEGER NOT NULL UNIQUE REFERENCES idioms(id) ON DELETE CASCADE,
    file_path         TEXT NOT NULL,                 -- relative to MEDIA_ROOT, e.g. idioms-41/live-high-on-the-hog.jpeg
    content_type      TEXT,
    file_size_bytes   INTEGER
);

-- Global upsert key: importing an idiom that already exists (by name, case-insensitive)
-- updates that row instead of inserting a duplicate.
CREATE UNIQUE INDEX IF NOT EXISTS idx_idioms_idiom_unique ON idioms (lower(idiom));
CREATE INDEX IF NOT EXISTS idx_idiom_quiz_options_quiz_id ON idiom_quiz_options (quiz_id);

-- OWS (one-word-substitution) database schema
-- One row in source_files (shared with vocab/idioms) per imported
-- "Substitution(s) - N (...).docx". Like vocab, these files carry a part of
-- speech and per-quiz-option definitions; like idioms, exactly one example
-- per word, so it's a plain column instead of a child table.

CREATE TABLE IF NOT EXISTS ows_words (
    id                   SERIAL PRIMARY KEY,
    source_file_id       INTEGER NOT NULL REFERENCES source_files(id) ON DELETE CASCADE,
    word_number          INTEGER NOT NULL,          -- position within the file (1-10)
    word                 TEXT NOT NULL,
    part_of_speech       TEXT,                       -- n, v, adj, ...
    meaning              TEXT NOT NULL,
    meaning_translation  TEXT,                       -- e.g. the Hindi gloss in parentheses
    example_text         TEXT,
    created_at           TIMESTAMPTZ NOT NULL DEFAULT now(),
    UNIQUE (source_file_id, word_number)
);

-- One quiz per word (the doc only ever has one "Quiz -" question per entry).
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
    description   TEXT,                              -- definition of the option, when the doc gives one (usually the wrong answers)
    UNIQUE (quiz_id, option_label)
);

-- Image stays on disk; this row is just the pointer + metadata.
CREATE TABLE IF NOT EXISTS ows_images (
    id                SERIAL PRIMARY KEY,
    ows_word_id       INTEGER NOT NULL UNIQUE REFERENCES ows_words(id) ON DELETE CASCADE,
    file_path         TEXT NOT NULL,                 -- relative to MEDIA_ROOT, e.g. substitution-1-28-may/warrant.jpeg
    content_type      TEXT,
    file_size_bytes   INTEGER
);

-- Global upsert key: importing an OWS word that already exists (by word, case-insensitive)
-- updates that row instead of inserting a duplicate.
CREATE UNIQUE INDEX IF NOT EXISTS idx_ows_words_word_unique ON ows_words (lower(word));
CREATE INDEX IF NOT EXISTS idx_ows_quiz_options_quiz_id ON ows_quiz_options (quiz_id);
