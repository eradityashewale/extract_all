"""
Imports idiom .docx files (see idiom_parser.py for the expected shape) into
Postgres.

Images are extracted to disk under --media-dir; only their relative path is
stored in the database (idiom_images.file_path). Point a static file
server / CDN at --media-dir and file_path is the URL suffix.

Idioms are upserted by name (case-insensitive): if an idiom already exists in
the DB it's updated in place (along with its quiz and image), never
duplicated. This makes the import safe to re-run on the same file, and safe
to run across files that happen to repeat an idiom.

Usage:
    # 1) create the schema once (idioms tables live in the same db/schema.sql as vocab)
    psql "$DATABASE_URL" -f db/schema.sql

    # 2) see what would happen, without touching the DB or writing files
    python import_idioms.py --input-dir idioms --dry-run

    # 3) actually import everything
    python import_idioms.py --input-dir idioms --db-url "$DATABASE_URL"

DATABASE_URL can also just be set as an env var instead of passed with --db-url.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import docx
import psycopg2

from idiom_parser import parse_docx
from import_vocab import slugify
from migrate import load_database_url


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
        fname = f"{e.idiom_number:02d}-{slugify(e.idiom)}{ext}"
        (out_dir / fname).write_bytes(blob)
        result[e.image_rid] = {
            "file_path": f"{file_slug}/{fname}",
            "content_type": rel.target_part.content_type,
            "size_bytes": len(blob),
        }
    return result


def import_file(conn, doc_path: str, media_dir: Path) -> list[str]:
    """
    Upserts every idiom by name (case-insensitive): if the idiom already
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
                report.append(f"[{file_name}] idiom {e.idiom_number} ({e.idiom}): {w}")

            cur.execute(
                """
                INSERT INTO idioms (source_file_id, idiom_number, idiom,
                                     meaning, meaning_translation, example_text)
                VALUES (%s, %s, %s, %s, %s, %s)
                ON CONFLICT (lower(idiom)) DO UPDATE SET
                    source_file_id = EXCLUDED.source_file_id,
                    idiom_number = EXCLUDED.idiom_number,
                    meaning = EXCLUDED.meaning,
                    meaning_translation = EXCLUDED.meaning_translation,
                    example_text = EXCLUDED.example_text
                RETURNING id
                """,
                (source_file_id, e.idiom_number, e.idiom, e.meaning, e.meaning_translation, e.example),
            )
            idiom_id = cur.fetchone()[0]

            # Replace child rows wholesale rather than merging old + new.
            cur.execute("DELETE FROM idiom_quizzes WHERE idiom_id = %s", (idiom_id,))  # cascades to idiom_quiz_options
            cur.execute("DELETE FROM idiom_images WHERE idiom_id = %s", (idiom_id,))

            if e.quiz_question:
                cur.execute(
                    "INSERT INTO idiom_quizzes (idiom_id, question) VALUES (%s, %s) RETURNING id",
                    (idiom_id, e.quiz_question),
                )
                quiz_id = cur.fetchone()[0]
                for opt in e.quiz_options:
                    cur.execute(
                        """
                        INSERT INTO idiom_quiz_options (quiz_id, option_label, option_text, is_correct)
                        VALUES (%s, %s, %s, %s)
                        """,
                        (quiz_id, opt.label, opt.text, opt.is_correct),
                    )

            img = images.get(e.image_rid) if e.image_rid else None
            if img:
                cur.execute(
                    """
                    INSERT INTO idiom_images (idiom_id, file_path, content_type, file_size_bytes)
                    VALUES (%s, %s, %s, %s)
                    """,
                    (idiom_id, img["file_path"], img["content_type"], img["size_bytes"]),
                )

    conn.commit()
    return report


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--input-dir", default="idioms", help="Directory of .docx files to import")
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
                    print(f"[{path.name}] idiom {e.idiom_number} ({e.idiom}): {w}")
                    total_warnings += 1
            print(f"[{path.name}] parsed {len(entries)} idioms")
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
