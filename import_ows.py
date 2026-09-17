"""
Imports OWS (one-word-substitution) .docx files (see ows_parser.py for the
expected shape) into Postgres.

Images are extracted to disk under --media-dir; only their relative path is
stored in the database (ows_images.file_path). Point a static file
server / CDN at --media-dir and file_path is the URL suffix.

OWS words are upserted by word (case-insensitive): if a word already exists
in the DB it's updated in place (along with its quiz and image), never
duplicated. This makes the import safe to re-run on the same file, and safe
to run across files that happen to repeat a word.

Usage:
    # 1) create the schema once (ows tables live in the same db/schema.sql as vocab/idioms)
    psql "$DATABASE_URL" -f db/schema.sql

    # 2) see what would happen, without touching the DB or writing files
    python import_ows.py --input-dir ows --dry-run

    # 3) actually import everything
    python import_ows.py --input-dir ows --db-url "$DATABASE_URL"

DATABASE_URL can also just be set as an env var instead of passed with --db-url.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import docx
import psycopg2

from import_vocab import slugify
from migrate import load_database_url
from ows_parser import parse_docx


def extract_images(doc_path: str, entries, media_dir: Path) -> dict[str, dict]:
    """rId -> {relative_path, content_type, size_bytes}, and writes the files."""
    doc = docx.Document(doc_path)
    rels = doc.part.rels
    file_slug = slugify(Path(doc_path).stem)
    out_dir = media_dir / file_slug
    out_dir.mkdir(parents=True, exist_ok=True)

    result = {}
    for e in entries:
        if not e.image_rid:
            continue
        rel = rels.get(e.image_rid)
        if rel is None or "image" not in rel.reltype:
            continue
        blob = rel.target_part.blob
        ext = Path(rel.target_ref).suffix or ".bin"
        fname = f"{e.number:02d}-{slugify(e.word)}{ext}"
        (out_dir / fname).write_bytes(blob)
        result[e.image_rid] = {
            "file_path": f"{file_slug}/{fname}",
            "content_type": rel.target_part.content_type,
            "size_bytes": len(blob),
        }
    return result


def import_file(conn, doc_path: str, media_dir: Path) -> list[str]:
    """
    Upserts every OWS word by word (case-insensitive): if the word already
    exists anywhere in the DB, its row and all child rows (quiz, image) are
    replaced with what's in this file; otherwise it's inserted new. Safe to
    re-run on the same or overlapping files.
    """
    entries, file_warnings = parse_docx(doc_path)
    file_name = Path(doc_path).name
    report = [f"[{file_name}] {w}" for w in file_warnings]

    with conn.cursor() as cur:
        cur.execute(
            """
            INSERT INTO source_files (file_name) VALUES (%s)
            ON CONFLICT (file_name) DO UPDATE SET imported_at = now()
            RETURNING id
            """,
            (file_name,),
        )
        source_file_id = cur.fetchone()[0]

        images = extract_images(doc_path, entries, media_dir)

        for e in entries:
            for w in e.warnings:
                report.append(f"[{file_name}] word {e.number} ({e.word}): {w}")

            cur.execute(
                """
                INSERT INTO ows_words (source_file_id, word_number, word, part_of_speech,
                                        meaning, meaning_translation, example_text)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (lower(word)) DO UPDATE SET
                    source_file_id = EXCLUDED.source_file_id,
                    word_number = EXCLUDED.word_number,
                    part_of_speech = EXCLUDED.part_of_speech,
                    meaning = EXCLUDED.meaning,
                    meaning_translation = EXCLUDED.meaning_translation,
                    example_text = EXCLUDED.example_text
                RETURNING id
                """,
                (source_file_id, e.number, e.word, e.part_of_speech, e.meaning, e.meaning_translation, e.example),
            )
            ows_word_id = cur.fetchone()[0]

            # Replace child rows wholesale rather than merging old + new.
            cur.execute("DELETE FROM ows_quizzes WHERE ows_word_id = %s", (ows_word_id,))  # cascades to ows_quiz_options
            cur.execute("DELETE FROM ows_images WHERE ows_word_id = %s", (ows_word_id,))

            if e.quiz_question:
                cur.execute(
                    "INSERT INTO ows_quizzes (ows_word_id, question) VALUES (%s, %s) RETURNING id",
                    (ows_word_id, e.quiz_question),
                )
                quiz_id = cur.fetchone()[0]
                for opt in e.quiz_options:
                    cur.execute(
                        """
                        INSERT INTO ows_quiz_options (quiz_id, option_label, option_text, is_correct, description)
                        VALUES (%s, %s, %s, %s, %s)
                        """,
                        (quiz_id, opt.label, opt.text, opt.is_correct, opt.description),
                    )

            img = images.get(e.image_rid) if e.image_rid else None
            if img:
                cur.execute(
                    """
                    INSERT INTO ows_images (ows_word_id, file_path, content_type, file_size_bytes)
                    VALUES (%s, %s, %s, %s)
                    """,
                    (ows_word_id, img["file_path"], img["content_type"], img["size_bytes"]),
                )

    conn.commit()
    return report


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input-dir", default="ows", help="Directory of .docx files to import")
    ap.add_argument("--media-dir", default="media", help="Directory to extract images into")
    ap.add_argument("--db-url", default=None, help="Postgres connection string (defaults to DATABASE_URL from the env or .env)")
    ap.add_argument("--dry-run", action="store_true", help="Parse and report only; no DB writes, no image extraction")
    args = ap.parse_args()

    input_dir = Path(args.input_dir)
    docx_files = sorted(input_dir.glob("*.docx"))
    if not docx_files:
        print(f"No .docx files found in {input_dir}", file=sys.stderr)
        sys.exit(1)

    if args.dry_run:
        total_warnings = 0
        for path in docx_files:
            entries, file_warnings = parse_docx(str(path))
            for w in file_warnings:
                print(f"[{path.name}] {w}")
                total_warnings += 1
            for e in entries:
                for w in e.warnings:
                    print(f"[{path.name}] word {e.number} ({e.word}): {w}")
                    total_warnings += 1
            print(f"[{path.name}] parsed {len(entries)} words")
        print(f"\n{len(docx_files)} file(s), {total_warnings} warning(s). Dry run only, nothing written.")
        return

    db_url = args.db_url or load_database_url()

    media_dir = Path(args.media_dir)
    media_dir.mkdir(parents=True, exist_ok=True)

    conn = psycopg2.connect(db_url)
    try:
        all_warnings = []
        for path in docx_files:
            print(f"Importing {path.name} ...")
            try:
                all_warnings += import_file(conn, str(path), media_dir)
            except Exception as exc:
                conn.rollback()
                print(f"FAILED on {path.name}: {exc}. Rolled back that file, continuing with the rest.", file=sys.stderr)
        print(f"\nDone. {len(docx_files)} file(s) processed, {len(all_warnings)} warning(s):")
        for w in all_warnings:
            print(f"  {w}")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
