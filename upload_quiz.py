"""
Creates a 30-question quiz (10 vocab + 10 idioms + 10 OWS) for a given date
on the Tarun Grover English admin panel (https://admin.tarungroverenglish.com),
reusing each source item's own "Quiz -" section that's already embedded in
its .docx file (vocab_parser.py / idiom_parser.py / ows_parser.py parse
these out already).

This mirrors the two admin-panel API calls the browser makes:
    1. POST /admin/quizzes
           -> creates an (empty) quiz for a date, returns its id.
    2. POST /admin/quizzes/update-questions
           -> attaches all 30 questions + their 4 choices each to that id.

Item selection uses the same "fileNumber(itemNumber)" syntax as
upload_idioms.py / upload_ows.py (offered for vocab here too, for symmetry,
even though upload_vocab.py itself takes a whole file at a time).

Usage:
    # preview the built quiz without touching the site
    python upload_quiz.py \
        --vocab-items "55(1),55(2),55(3),55(4),55(5),55(6),55(7),55(8),55(9),55(10)" \
        --idiom-items "41(1),42(1),43(1),44(1),45(1),46(1),47(1),48(1),49(1),50(1)" \
        --ows-items "1(1),2(1),3(1),4(1),5(1),6(1),7(1),8(1),9(1),10(1)" \
        --date 18/9/2026 --dry-run

    # create it for real
    python upload_quiz.py --vocab-items "..." --idiom-items "..." --ows-items "..." --date 18/9/2026

Auth: reuses the same ADMIN_SESSION_COOKIE / ADMIN_DEVICE_ID env vars as the
other upload_*.py scripts, via .env, or --cookie / --device-id.
"""

from __future__ import annotations

import argparse
import random
import re
import string
import sys
from pathlib import Path

import idiom_parser
import ows_parser
import vocab_parser
from upload_idioms import find_idiom_file
from upload_ows import find_ows_file
from upload_vocab import (
    ADMIN_API_BASE,
    build_session,
    find_vocab_file,
    load_device_id,
    load_session_cookie,
    parse_date,
)

CREATE_QUIZ_ENDPOINT = "/admin/quizzes"
UPDATE_QUESTIONS_ENDPOINT = "/admin/quizzes/update-questions"

ITEM_RE = re.compile(r"(\d+)\s*\(\s*(\d+)\s*\)")
NANOID_ALPHABET = string.ascii_letters + string.digits + "_-"


def gen_id(size: int = 21) -> str:
    """Client-side question/choice id, in the same shape the admin frontend generates (nanoid)."""
    return "".join(random.choices(NANOID_ALPHABET, k=size))


def parse_items(text: str, label: str) -> list[tuple[int, int]]:
    """Parses "41(1), 42(1), ..." into [(41, 1), (42, 1), ...]."""
    matches = ITEM_RE.findall(text)
    if not matches:
        print(f"Couldn't parse any 'fileNumber(itemNumber)' {label} items from '{text}'", file=sys.stderr)
        sys.exit(1)
    return [(int(a), int(b)) for a, b in matches]


class Question:
    def __init__(self, source: str, text: str, options: list[tuple[str, bool, str | None]]):
        self.source = source
        self.text = text
        self.options = options  # (option_text, is_correct, description|None)


def _quiz_ok(entry) -> bool:
    return bool(entry.quiz_question) and len(entry.quiz_options) == 4 and any(o.is_correct for o in entry.quiz_options)


def collect_vocab_questions(items: list[tuple[int, int]], input_dir: Path) -> list[Question]:
    cache: dict[int, dict] = {}
    questions = []
    for file_num, word_num in items:
        if file_num not in cache:
            path = find_vocab_file(file_num, input_dir)
            entries, warnings = vocab_parser.parse_docx(str(path))
            for w in warnings:
                print(f"[warning] [{path.name}] {w}", file=sys.stderr)
            cache[file_num] = {e.word_number: e for e in entries}
        entry = cache[file_num].get(word_num)
        if entry is None:
            print(f"[warning] Vocab file {file_num} has no word numbered {word_num}; skipping.", file=sys.stderr)
            continue
        for w in entry.warnings:
            print(f"[warning] [vocab {file_num}] word {word_num} ({entry.word}): {w}", file=sys.stderr)
        if not _quiz_ok(entry):
            print(f"[warning] Vocab {file_num}({word_num}) {entry.word}: incomplete quiz data; skipping.", file=sys.stderr)
            continue
        questions.append(Question(
            source=f"vocab {file_num}({word_num}) {entry.word}",
            text=entry.quiz_question,
            options=[(o.text.strip(), o.is_correct, o.description) for o in entry.quiz_options],
        ))
    return questions


def collect_idiom_questions(items: list[tuple[int, int]], input_dir: Path) -> list[Question]:
    cache: dict[int, dict] = {}
    questions = []
    for file_num, idiom_num in items:
        if file_num not in cache:
            path = find_idiom_file(file_num, input_dir)
            entries, warnings = idiom_parser.parse_docx(str(path))
            for w in warnings:
                print(f"[warning] [{path.name}] {w}", file=sys.stderr)
            cache[file_num] = {e.idiom_number: e for e in entries}
        entry = cache[file_num].get(idiom_num)
        if entry is None:
            print(f"[warning] Idiom file {file_num} has no idiom numbered {idiom_num}; skipping.", file=sys.stderr)
            continue
        for w in entry.warnings:
            print(f"[warning] [idioms {file_num}] idiom {idiom_num} ({entry.idiom}): {w}", file=sys.stderr)
        if not _quiz_ok(entry):
            print(f"[warning] Idiom {file_num}({idiom_num}) {entry.idiom}: incomplete quiz data; skipping.", file=sys.stderr)
            continue
        # Idioms carry no per-option definitions in the source docs.
        questions.append(Question(
            source=f"idiom {file_num}({idiom_num}) {entry.idiom}",
            text=entry.quiz_question,
            options=[(o.text.strip(), o.is_correct, None) for o in entry.quiz_options],
        ))
    return questions


def collect_ows_questions(items: list[tuple[int, int]], input_dir: Path) -> list[Question]:
    cache: dict[int, dict] = {}
    questions = []
    for file_num, entry_num in items:
        if file_num not in cache:
            path = find_ows_file(file_num, input_dir)
            entries, warnings = ows_parser.parse_docx(str(path))
            for w in warnings:
                print(f"[warning] [{path.name}] {w}", file=sys.stderr)
            cache[file_num] = {e.number: e for e in entries}
        entry = cache[file_num].get(entry_num)
        if entry is None:
            print(f"[warning] OWS file {file_num} has no entry numbered {entry_num}; skipping.", file=sys.stderr)
            continue
        for w in entry.warnings:
            print(f"[warning] [ows {file_num}] entry {entry_num} ({entry.word}): {w}", file=sys.stderr)
        if not _quiz_ok(entry):
            print(f"[warning] OWS {file_num}({entry_num}) {entry.word}: incomplete quiz data; skipping.", file=sys.stderr)
            continue
        questions.append(Question(
            source=f"ows {file_num}({entry_num}) {entry.word}",
            text=entry.quiz_question,
            options=[(o.text.strip(), o.is_correct, o.description) for o in entry.quiz_options],
        ))
    return questions


def build_payload(quiz_id: str, questions: list[Question]) -> dict:
    question_payloads = []
    choice_payloads = []
    for q_order, q in enumerate(questions, start=1):
        q_id = gen_id()
        question_payloads.append({
            "id": q_id,
            "order": q_order,
            "explanation": "",
            "description": q.text,
        })
        for c_order, (text, is_correct, description) in enumerate(q.options, start=1):
            choice = {
                "id": gen_id(),
                "order": c_order,
                "quizQuestionId": q_id,
                "description": text,
                "isCorrect": is_correct,
            }
            if description:
                choice["explanation"] = description
            choice_payloads.append(choice)
    return {"id": quiz_id, "questions": question_payloads, "choices": choice_payloads}


def extract_quiz_id(resp_json: dict) -> str:
    body = resp_json.get("body", resp_json) if isinstance(resp_json, dict) else resp_json
    if isinstance(body, dict):
        if isinstance(body.get("id"), str):
            return body["id"]
        quiz = body.get("quiz")
        if isinstance(quiz, dict) and isinstance(quiz.get("id"), str):
            return quiz["id"]
    raise RuntimeError(f"Couldn't find a quiz id in the create-quiz response: {resp_json}")


def print_preview(all_questions: list[Question], payload: dict) -> None:
    choices_by_qid: dict[str, list[dict]] = {}
    for c in payload["choices"]:
        choices_by_qid.setdefault(c["quizQuestionId"], []).append(c)

    for src_q, q in zip(all_questions, payload["questions"]):
        print(f"[{q['order']:2}] {q['description']}   ({src_q.source})")
        for c in choices_by_qid[q["id"]]:
            mark = "*" if c["isCorrect"] else " "
            expl = f"  -- {c['explanation']}" if c.get("explanation") else ""
            print(f"      {mark} {c['description']}{expl}")


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except AttributeError:
        pass

    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--vocab-items", required=True, help='Comma-separated "fileNumber(wordNumber)" list, e.g. "55(1),...,55(10)"')
    ap.add_argument("--idiom-items", required=True, help='Comma-separated "fileNumber(idiomNumber)" list, e.g. "41(1),...,50(1)"')
    ap.add_argument("--ows-items", required=True, help='Comma-separated "fileNumber(entryNumber)" list, e.g. "1(1),...,10(1)"')
    ap.add_argument("--date", "-d", required=True, help="Quiz date, DD/MM/YYYY, e.g. 18/9/2026")
    ap.add_argument("--type", default="daily_free", help="Quiz type sent to the create-quiz call (default: daily_free)")
    ap.add_argument("--points", default="10", help="Points value sent to the create-quiz call (default: 10)")
    ap.add_argument("--vocab-dir", default="vocab_files", help="Directory containing the Vocab - N ... .docx files")
    ap.add_argument("--idiom-dir", default="idioms", help="Directory containing the Idioms - N ( D Month ).docx files")
    ap.add_argument("--ows-dir", default="ows", help="Directory containing the Substitution(s) - N ( D Month ).docx files")
    ap.add_argument("--cookie", default=None, help="Session cookie header value (overrides ADMIN_SESSION_COOKIE)")
    ap.add_argument("--device-id", default=None, help="x-device-id header value (overrides ADMIN_DEVICE_ID)")
    ap.add_argument("--dry-run", action="store_true", help="Only print what would be sent; no requests made")
    args = ap.parse_args()

    iso_date = parse_date(args.date)

    vocab_items = parse_items(args.vocab_items, "vocab")
    idiom_items = parse_items(args.idiom_items, "idiom")
    ows_items = parse_items(args.ows_items, "ows")

    vocab_qs = collect_vocab_questions(vocab_items, Path(args.vocab_dir))
    idiom_qs = collect_idiom_questions(idiom_items, Path(args.idiom_dir))
    ows_qs = collect_ows_questions(ows_items, Path(args.ows_dir))

    for label, qs, expected in (
        ("vocab", vocab_qs, len(vocab_items)),
        ("idiom", idiom_qs, len(idiom_items)),
        ("ows", ows_qs, len(ows_items)),
    ):
        if len(qs) != expected:
            print(f"[warning] Resolved {len(qs)}/{expected} {label} questions.", file=sys.stderr)

    all_questions = vocab_qs + idiom_qs + ows_qs
    if not all_questions:
        print("No questions resolved; nothing to do.", file=sys.stderr)
        sys.exit(1)

    print(f"Date: {iso_date}")
    print(f"Questions: {len(all_questions)} (vocab {len(vocab_qs)}, idioms {len(idiom_qs)}, ows {len(ows_qs)})")
    print(f"Mode: {'DRY RUN (nothing sent)' if args.dry_run else 'LIVE upload'}")
    print()

    if args.dry_run:
        print(f"Create-quiz payload: {{'type': {args.type!r}, 'points': {args.points!r}, 'forDate': {iso_date!r}}}")
        print()
        payload = build_payload("DRY-RUN-QUIZ-ID", all_questions)
        print_preview(all_questions, payload)
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

    create_payload = {"type": args.type, "points": str(args.points), "forDate": iso_date}
    resp = session.post(ADMIN_API_BASE + CREATE_QUIZ_ENDPOINT, json=create_payload)
    if resp.status_code >= 400:
        print(f"FAILED to create quiz -> {resp.status_code} {resp.text}", file=sys.stderr)
        sys.exit(1)
    quiz_id = extract_quiz_id(resp.json())
    print(f"Created quiz {quiz_id} for {iso_date}")

    payload = build_payload(quiz_id, all_questions)
    resp = session.post(ADMIN_API_BASE + UPDATE_QUESTIONS_ENDPOINT, json=payload)
    if resp.status_code >= 400:
        print(f"FAILED to upload questions -> {resp.status_code} {resp.text}", file=sys.stderr)
        sys.exit(1)
    print(f"OK: uploaded {len(all_questions)} questions to quiz {quiz_id}")


if __name__ == "__main__":
    main()
