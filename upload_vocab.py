"""
Uploads a "Vocab - N with photos.docx" file's words into the Tarun Grover
English admin panel (https://admin.tarungroverenglish.com), as "Normal" type
vocab entries for a given date.

Usage:
    # preview what would be sent, without touching the site
    python upload_vocab.py --file 32 --date 18/9/2026 --dry-run

    # actually create the entries on the site (text fields only)
    python upload_vocab.py --file 32 --date 18/9/2026

    # also upload each word's image and attach it
    python upload_vocab.py --file 32 --date 18/9/2026 --with-images

Auth: reuses your logged-in browser session cookie and device id rather than
doing the email+OTP flow itself.
  - ADMIN_SESSION_COOKIE: full Cookie header value (DevTools > Application >
    Storage > Cookies > prod-api.tarungroverenglish.com > combine every
    Name=Value row with "; ").
  - ADMIN_DEVICE_ID: the x-device-id header value (DevTools > Network > any
    authenticated request > Request Headers > x-device-id).
Set both in .env, or pass --cookie / --device-id.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from pathlib import Path

import requests

from import_vocab import extract_images
from vocab_parser import parse_docx

# Strips any "(...)" span containing Devanagari script (Hindi/Marathi) out of
# a definition. vocab_parser.py already splits most of these into
# meaning_translation, but a few source files have trailing punctuation or
# reordered text that slips past that split, leaving the translation in the
# meaning text itself.
DEVANAGARI_PAREN_RE = re.compile(r"\s*\([^)]*[ऀ-ॿ][^)]*\)\.?")

ADMIN_API_BASE = "https://prod-api.tarungroverenglish.com"
CREATE_VOCAB_ENDPOINT = "/admin/vocabs"
PRESIGNED_UPLOAD_ENDPOINT = "/admin/medias/presigned-upload-url"
MEDIAS_LIST_ENDPOINT = "/admin/medias"
MEDIA_BUCKET_ID = "cdn.tarungrover.in"

# The backend ties the session to a specific device id (matches the
# account's `activeDeviceId`), sent by the browser on every request.
# Without it, a valid session cookie still gets 401 Unauthorized.
BASE_HEADERS = {
    "rid": "anti-csrf",
    "referer": "https://admin.tarungroverenglish.com/",
}

POS_MAP = {
    "n": "noun",
    "v": "verb",
    "adj": "adjective",
    "adv": "adverb",
    "prep": "preposition",
    "conj": "conjunction",
    "pron": "pronoun",
    "interj": "interjection",
}


def _load_env_value(name: str, cli_value: str | None) -> str | None:
    if cli_value:
        return cli_value
    if os.environ.get(name):
        return os.environ[name]
    env_path = Path(".env")
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8").splitlines():
            m = re.match(rf"^\s*{name}\s*=\s*(.+?)\s*$", line)
            if m:
                return m.group(1).strip().strip('"').strip("'")
    return None


def _load_cookie_parts() -> str | None:
    """Joins every `COOKIE_<name>=value` line in .env into one Cookie header value."""
    env_path = Path(".env")
    if not env_path.exists():
        return None
    parts = []
    for line in env_path.read_text(encoding="utf-8").splitlines():
        m = re.match(r"^\s*COOKIE_(?P<name>[^=\s]+)\s*=\s*(?P<value>.*?)\s*$", line)
        if m and m.group("value"):
            parts.append(f"{m.group('name')}={m.group('value').strip().strip(chr(34)).strip(chr(39))}")
    return "; ".join(parts) or None


def load_session_cookie(cli_cookie: str | None) -> str | None:
    return cli_cookie or _load_cookie_parts() or _load_env_value("ADMIN_SESSION_COOKIE", None)


def load_device_id(cli_device_id: str | None) -> str | None:
    return _load_env_value("ADMIN_DEVICE_ID", cli_device_id)


def build_session(cookie: str, device_id: str) -> requests.Session:
    session = requests.Session()
    session.headers.update({
        "Cookie": cookie,
        "Content-Type": "application/json",
        "x-device-id": device_id,
        **BASE_HEADERS,
    })
    return session


def find_vocab_file(number: int, input_dir: Path) -> Path:
    matches = sorted(input_dir.glob(f"Vocab - {number} with photo*.docx"))
    if not matches:
        matches = sorted(input_dir.glob(f"Vocab - {number} *.docx"))
    if not matches:
        print(f"No file found in {input_dir} for vocab number {number}", file=sys.stderr)
        sys.exit(1)
    if len(matches) > 1:
        print(f"Multiple files matched vocab number {number}: {[m.name for m in matches]}", file=sys.stderr)
        sys.exit(1)
    return matches[0]


def parse_date(text: str) -> str:
    """Accepts D/M/YYYY (or D-M-YYYY) and returns ISO YYYY-MM-DD."""
    m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$", text.strip())
    if not m:
        print(f"Couldn't parse date '{text}'; expected DD/MM/YYYY", file=sys.stderr)
        sys.exit(1)
    day, month, year = m.groups()
    return f"{year}-{int(month):02d}-{int(day):02d}"


def map_part_of_speech(pos: str | None) -> str:
    if not pos:
        return ""
    mapped = POS_MAP.get(pos.strip().lower())
    if mapped:
        return mapped
    print(f"[warning] Unrecognized part of speech '{pos}'; sending as-is.", file=sys.stderr)
    return pos.strip().lower()


def build_definition(meaning: str, translation: str | None) -> str:
    cleaned = DEVANAGARI_PAREN_RE.sub("", meaning).strip()
    cleaned = re.sub(r"\s{2,}", " ", cleaned)
    if translation:
        cleaned = f"{cleaned} ({translation})"
    return cleaned


def build_payload(entry, iso_date: str, media_id: str | None = None) -> dict:
    """Maps a parsed WordEntry onto the confirmed POST /admin/vocabs body shape."""
    shortest_example = min(entry.examples, key=len) if entry.examples else ""
    payload = {
        "type": "normal",
        "word": entry.word,
        "partOfSpeech": map_part_of_speech(entry.part_of_speech),
        "definition": build_definition(entry.meaning, entry.meaning_translation),
        "example": shortest_example,
        "synonyms": entry.synonyms,
        "antonyms": [],
        "related": [],
        "trick": entry.hint or "",
        "forDate": iso_date,
    }
    if media_id:
        payload["mediaId"] = media_id
    return payload


def find_media_id(session: requests.Session, expected_url: str, max_attempts: int = 10, delay: float = 1.0) -> str | None:
    """
    Polls the media library list for an item matching expected_url. The
    library's "medias" records (which is what a vocab's mediaId actually
    references) show up a moment after the raw S3 upload finishes -- the
    fileObject id from presigned-upload-url is NOT the same id and isn't
    valid as a vocab's mediaId.
    """
    params = {
        "page": 1,
        "limit": 10,
        "filters": json.dumps({"filters": [], "sorts": [{"field": "createdAt", "sort": "descending"}]}),
        "type": "image",
    }
    for attempt in range(max_attempts):
        resp = session.get(ADMIN_API_BASE + MEDIAS_LIST_ENDPOINT, params=params)
        resp.raise_for_status()
        for item in resp.json()["body"]["items"]:
            if item["downloadUrl"] == expected_url:
                return item["id"]
        if attempt < max_attempts - 1:
            time.sleep(delay)
    return None


def upload_image(session: requests.Session, file_path: Path, content_type: str) -> str:
    """
    Uploads one image to the admin panel's media library:
      1. POST presigned-upload-url -> a signed S3 PUT url + a storage key
         (the returned fileObject id is NOT usable as a vocab's mediaId).
      2. PUT the raw bytes straight to S3 (no app auth needed, the signature
         is in the URL's query string).
      3. Poll the medias list until the matching library entry appears, and
         use its id as the vocab's mediaId.
    """
    resp = session.post(
        ADMIN_API_BASE + PRESIGNED_UPLOAD_ENDPOINT,
        json={
            "bucketId": MEDIA_BUCKET_ID,
            "type": content_type,
            "size": file_path.stat().st_size,
            "originalFileName": file_path.name,
            "service": "media",
        },
    )
    resp.raise_for_status()
    body = resp.json()["body"]

    put_resp = requests.put(body["url"], data=file_path.read_bytes(), headers={"Content-Type": content_type})
    put_resp.raise_for_status()

    expected_url = f"https://{MEDIA_BUCKET_ID}/{body['key']}"
    media_id = find_media_id(session, expected_url)
    if not media_id:
        raise RuntimeError(f"Uploaded {file_path.name} but it never appeared in the media library list")
    return media_id


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except AttributeError:
        pass

    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--file", "-f", type=int, required=True, help="Vocab file number, e.g. 32")
    ap.add_argument("--date", "-d", type=str, required=True, help="Date to assign, DD/MM/YYYY, e.g. 18/9/2026")
    ap.add_argument("--input-dir", default="vocab_files", help="Directory containing the Vocab - N ... .docx files")
    ap.add_argument("--count", type=int, default=None, help="Max number of words to upload (default: all in the file)")
    ap.add_argument("--cookie", default=None, help="Session cookie header value (overrides ADMIN_SESSION_COOKIE)")
    ap.add_argument("--device-id", default=None, help="x-device-id header value (overrides ADMIN_DEVICE_ID)")
    ap.add_argument("--with-images", action="store_true", help="Also extract and upload each word's image")
    ap.add_argument("--media-dir", default="media", help="Directory to extract images into (for --with-images)")
    ap.add_argument("--dry-run", action="store_true", help="Only print what would be sent; no requests made")
    args = ap.parse_args()

    input_dir = Path(args.input_dir)
    doc_path = find_vocab_file(args.file, input_dir)
    iso_date = parse_date(args.date)

    entries, file_warnings = parse_docx(str(doc_path))
    for w in file_warnings:
        print(f"[warning] {w}", file=sys.stderr)

    if args.count is not None:
        entries = entries[: args.count]

    if not entries:
        print(f"No words parsed from {doc_path.name}", file=sys.stderr)
        sys.exit(1)

    images_by_rid = {}
    media_dir = Path(args.media_dir)
    if args.with_images:
        media_dir.mkdir(parents=True, exist_ok=True)
        images_by_rid = extract_images(str(doc_path), entries, media_dir)

    print(f"File: {doc_path.name}")
    print(f"Date: {iso_date}")
    print(f"Words: {len(entries)}")
    print(f"Mode: {'DRY RUN (nothing sent)' if args.dry_run else 'LIVE upload'}")
    if args.with_images:
        print("Images: enabled")
    print()

    if args.dry_run:
        for e in entries:
            p = build_payload(e, iso_date)
            if args.with_images:
                img = images_by_rid.get(e.image_rid) if e.image_rid else None
                p["media_file"] = img["file_path"] if img else "(no image found)"
            print(p)
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

    for e in entries:
        media_id = None
        if args.with_images:
            img = images_by_rid.get(e.image_rid) if e.image_rid else None
            if img:
                try:
                    media_id = upload_image(session, media_dir / img["file_path"], img["content_type"])
                    print(f"  uploaded image for {e.word} -> mediaId {media_id}")
                except (requests.HTTPError, RuntimeError) as exc:
                    print(f"  IMAGE FAILED for {e.word}: {exc}", file=sys.stderr)
            else:
                print(f"  [warning] no image found for {e.word}; creating without one", file=sys.stderr)

        payload = build_payload(e, iso_date, media_id)
        resp = session.post(ADMIN_API_BASE + CREATE_VOCAB_ENDPOINT, json=payload)
        if resp.status_code >= 400:
            print(f"FAILED: {payload['word']} -> {resp.status_code} {resp.text}", file=sys.stderr)
        else:
            print(f"OK: {payload['word']}")


if __name__ == "__main__":
    main()
