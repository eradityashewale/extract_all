"""
Uploads idioms from the "Idioms - N ( D Month ).docx" files into the Tarun
Grover English admin panel (https://admin.tarungroverenglish.com), as
"idioms" type entries.

Input is a list of "fileNumber(idiomNumber)" tokens: which file to pull from,
and which numbered idiom within that file to take. By default each file's
date is parsed from its own filename (e.g. "Idioms - 41 ( 8 July ).docx" ->
8 July), paired with --year. Pass --date to override and use that date for
every item instead (--year is then not needed).

Usage:
    # preview what would be sent, without touching the site
    python upload_idioms.py --items "41(1),42(1),43(1),44(1),45(1),46(1),47(1),48(1),49(1),50(1)" --year 2026 --dry-run

    # actually create the entries on the site (text fields only)
    python upload_idioms.py --items "41(1),42(1),...,50(1)" --year 2026

    # also upload each idiom's image and attach it
    python upload_idioms.py --items "41(1),42(1),...,50(1)" --year 2026 --with-images

    # upload all 10 idioms from a single file as one day's batch
    python upload_idioms.py --items "41(1),41(2),41(3),41(4),41(5),41(6),41(7),41(8),41(9),41(10)" --year 2026 --with-images

    # force a specific date instead of using the one in the filename
    python upload_idioms.py --items "41(1)" --date 18/9/2026 --with-images

Auth: reuses the same ADMIN_SESSION_COOKIE / ADMIN_DEVICE_ID env vars (see
upload_vocab.py) via .env, or --cookie / --device-id.
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import requests

from idiom_parser import parse_docx
from import_idioms import extract_images
from upload_vocab import (
    ADMIN_API_BASE,
    CREATE_VOCAB_ENDPOINT,
    build_session,
    load_device_id,
    load_session_cookie,
    parse_date,
    upload_image,
)

ITEM_RE = re.compile(r"(\d+)\s*\(\s*(\d+)\s*\)")
FILENAME_DATE_RE = re.compile(r"\(\s*(\d{1,2})\s+([A-Za-z]+)\s*\)")


def parse_items(text: str) -> list[tuple[int, int]]:
    """Parses "41(1), 42(1), ..." into [(41, 1), (42, 1), ...]."""
    matches = ITEM_RE.findall(text)
    if not matches:
        print(f"Couldn't parse any 'fileNumber(idiomNumber)' items from '{text}'", file=sys.stderr)
        sys.exit(1)
    return [(int(file_num), int(idiom_num)) for file_num, idiom_num in matches]


def find_idiom_file(number: int, input_dir: Path) -> Path:
    # Names vary: "Idioms - 1.docx", "Idioms - 7 - (3 June).docx", "Idioms - 41 ( 8 July ).docx"
    matches = sorted(p for p in input_dir.glob("Idioms - *.docx") if re.match(rf"Idioms - {number}\b", p.name))
    if not matches:
        print(f"No file found in {input_dir} for idiom file number {number}", file=sys.stderr)
        sys.exit(1)
    if len(matches) > 1:
        print(f"Multiple files matched idiom file number {number}: {[m.name for m in matches]}", file=sys.stderr)
        sys.exit(1)
    return matches[0]


def parse_date_from_filename(path: Path, year: int) -> str:
    """Extracts "( 8 July )" from the filename and combines it with --year -> ISO date."""
    m = FILENAME_DATE_RE.search(path.name)
    if not m:
        print(f"Couldn't find a '( D Month )' date in filename '{path.name}'", file=sys.stderr)
        sys.exit(1)
    day, month = m.groups()
    try:
        try:
            dt = datetime.strptime(f"{day} {month} {year}", "%d %B %Y")
        except ValueError:
            dt = datetime.strptime(f"{day} {month} {year}", "%d %b %Y")  # "1 Aug"
    except ValueError as exc:
        print(f"Couldn't parse date '{day} {month} {year}' from filename '{path.name}': {exc}", file=sys.stderr)
        sys.exit(1)
    return dt.strftime("%Y-%m-%d")


def build_definition(meaning: str, translation: str | None) -> str:
    cleaned = re.sub(r"\s{2,}", " ", meaning.strip())
    if translation:
        cleaned = f"{cleaned} ({translation})"
    return cleaned


def build_payload(entry, iso_date: str, media_id: str | None = None) -> dict:
    payload = {
        "type": "idioms",
        "phrase": entry.idiom,
        "definition": build_definition(entry.meaning, entry.meaning_translation),
        "example": entry.example or "",
        "synonyms": [],
        "antonyms": [],
        "related": [],
        "forDate": iso_date,
    }
    if media_id:
        payload["mediaId"] = media_id
    return payload


class FileContext:
    """Caches the parsed entries, date, and doc path for one idiom file number."""

    def __init__(self, file_number: int, input_dir: Path, year: int | None, date_override: str | None):
        self.file_number = file_number
        self.path = find_idiom_file(file_number, input_dir)
        self.iso_date = date_override if date_override else parse_date_from_filename(self.path, year)
        self.entries, self.warnings = parse_docx(str(self.path))
        self.entries_by_number = {e.idiom_number: e for e in self.entries}


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except AttributeError:
        pass

    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--items", "-i", required=True, help='Comma-separated "fileNumber(idiomNumber)" list, e.g. "41(1),42(1),...,50(1)"')
    ap.add_argument("--year", type=int, default=None, help="Year to combine with each file's embedded 'D Month' date, e.g. 2026 (not needed if --date is given)")
    ap.add_argument("--date", "-d", default=None, help="Override forDate for every item, DD/MM/YYYY, e.g. 18/9/2026 (default: date parsed from each file's name)")
    ap.add_argument("--input-dir", default="idioms", help="Directory containing the Idioms - N ( D Month ).docx files")
    ap.add_argument("--cookie", default=None, help="Session cookie header value (overrides ADMIN_SESSION_COOKIE)")
    ap.add_argument("--device-id", default=None, help="x-device-id header value (overrides ADMIN_DEVICE_ID)")
    ap.add_argument("--with-images", action="store_true", help="Also extract and upload each idiom's image")
    ap.add_argument("--media-dir", default="media", help="Directory to extract images into (for --with-images)")
    ap.add_argument("--dry-run", action="store_true", help="Only print what would be sent; no requests made")
    args = ap.parse_args()

    if not args.date and args.year is None:
        print("Either --date or --year is required.", file=sys.stderr)
        sys.exit(1)

    date_override = parse_date(args.date) if args.date else None

    input_dir = Path(args.input_dir)
    items = parse_items(args.items)

    file_ctx_cache: dict[int, FileContext] = {}
    selected: list[tuple[FileContext, object]] = []  # (ctx, IdiomEntry)
    for file_num, idiom_num in items:
        ctx = file_ctx_cache.get(file_num)
        if ctx is None:
            ctx = FileContext(file_num, input_dir, args.year, date_override)
            file_ctx_cache[file_num] = ctx
            for w in ctx.warnings:
                print(f"[warning] [{ctx.path.name}] {w}", file=sys.stderr)

        entry = ctx.entries_by_number.get(idiom_num)
        if entry is None:
            print(f"[warning] File {file_num} has no idiom numbered {idiom_num}; skipping.", file=sys.stderr)
            continue
        for w in entry.warnings:
            print(f"[warning] [{ctx.path.name}] idiom {entry.idiom_number} ({entry.idiom}): {w}", file=sys.stderr)
        selected.append((ctx, entry))

    if not selected:
        print("No idioms resolved from --items; nothing to do.", file=sys.stderr)
        sys.exit(1)

    print(f"Items: {len(selected)}")
    print(f"Mode: {'DRY RUN (nothing sent)' if args.dry_run else 'LIVE upload'}")
    if args.with_images:
        print("Images: enabled")
    print()

    media_dir = Path(args.media_dir)
    images_by_file: dict[int, dict] = {}
    if args.with_images:
        media_dir.mkdir(parents=True, exist_ok=True)
        for ctx, entry in selected:
            if ctx.file_number not in images_by_file:
                images_by_file[ctx.file_number] = {}
            if entry.image_rid and entry.image_rid not in images_by_file[ctx.file_number]:
                images_by_file[ctx.file_number].update(extract_images(str(ctx.path), [entry], media_dir))

    if args.dry_run:
        for ctx, entry in selected:
            p = build_payload(entry, ctx.iso_date)
            if args.with_images:
                img = images_by_file.get(ctx.file_number, {}).get(entry.image_rid) if entry.image_rid else None
                p["media_file"] = img["file_path"] if img else "(no image found)"
            print(f"[file {ctx.file_number}] {p}")
        return

    cookie = load_session_cookie(args.cookie)
    if not cookie:
        print(
            "No session cookie found. Pass --cookie \"...\" or set ADMIN_SESSION_COOKIE in .env "
            "(copy the Cookie header value from an authenticated request in DevTools).",
            file=sys.stderr,
        )
        sys.exit(1)

    device_id = load_device_id(args.device_id)
    if not device_id:
        print(
            "No device id found. Pass --device-id \"...\" or set ADMIN_DEVICE_ID in .env "
            "(copy the x-device-id header value from an authenticated request in DevTools).",
            file=sys.stderr,
        )
        sys.exit(1)

    session = build_session(cookie, device_id)

    for ctx, entry in selected:
        media_id = None
        if args.with_images:
            img = images_by_file.get(ctx.file_number, {}).get(entry.image_rid) if entry.image_rid else None
            if img:
                try:
                    media_id = upload_image(session, media_dir / img["file_path"], img["content_type"])
                    print(f"  uploaded image for {entry.idiom} -> mediaId {media_id}")
                except (requests.HTTPError, RuntimeError) as exc:
                    print(f"  IMAGE FAILED for {entry.idiom}: {exc}", file=sys.stderr)
            else:
                print(f"  [warning] no image found for {entry.idiom}; creating without one", file=sys.stderr)

        payload = build_payload(entry, ctx.iso_date, media_id)
        resp = session.post(ADMIN_API_BASE + CREATE_VOCAB_ENDPOINT, json=payload)
        if resp.status_code >= 400:
            print(f"FAILED: {payload['phrase']} ({ctx.iso_date}) -> {resp.status_code} {resp.text}", file=sys.stderr)
        else:
            print(f"OK: {payload['phrase']} ({ctx.iso_date})")


if __name__ == "__main__":
    main()
