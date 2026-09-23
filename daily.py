"""
One command for the whole daily upload: vocab + idioms + OWS + quiz.

Builds the long "--items" strings for upload_vocab.py / upload_idioms.py /
upload_ows.py / upload_quiz.py from a few short numbers, then runs them in order.

Usage:
    # vocab file 38, idiom files 41-50 item 8, ows files 1-10 item 8, for 24/9/2026
    python daily.py --vocab 38 --idioms 41-50 --ows 1-10 --item 8 --date 24/9/2026

    # --date omitted -> tomorrow
    python daily.py --vocab 38 --idioms 41-50 --ows 1-10 --item 8

    # only the quiz (content already uploaded)
    python daily.py --vocab 38 --idioms 41-50 --ows 1-10 --item 8 --date 24/9/2026 --only quiz

    # preview every command without running anything
    python daily.py --vocab 38 --idioms 41-50 --ows 1-10 --item 8 --print
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from datetime import date, timedelta

STEPS = ["vocab", "idioms", "ows", "quiz"]


def parse_range(text: str) -> list[int]:
    """"41-50" -> [41..50], "41" -> [41], "41,43,45" -> [41, 43, 45]."""
    nums: list[int] = []
    for part in text.split(","):
        part = part.strip()
        if "-" in part:
            lo, hi = part.split("-", 1)
            nums.extend(range(int(lo), int(hi) + 1))
        elif part:
            nums.append(int(part))
    return nums


def items(files: list[int], item: int) -> str:
    return ",".join(f"{f}({item})" for f in files)


def main() -> None:
    ap = argparse.ArgumentParser(description="Run the daily vocab/idioms/OWS/quiz uploads")
    ap.add_argument("--vocab", "-v", type=int, required=True, help="Vocab file number, e.g. 38")
    ap.add_argument("--idioms", "-i", required=True, help='Idiom file range, e.g. "41-50"')
    ap.add_argument("--ows", "-o", required=True, help='OWS file range, e.g. "1-10"')
    ap.add_argument("--item", "-n", type=int, required=True, help="Item number inside each idiom/OWS file, e.g. 8")
    ap.add_argument("--idiom-item", type=int, default=None, help="Override --item for idioms only")
    ap.add_argument("--ows-item", type=int, default=None, help="Override --item for OWS only")
    ap.add_argument("--vocab-quiz", default="1-10", help='Vocab word numbers used for the quiz (default "1-10")')
    ap.add_argument("--date", "-d", default=None, help="DD/MM/YYYY (default: tomorrow)")
    ap.add_argument("--only", nargs="+", choices=STEPS, default=STEPS, help="Run only these steps, e.g. --only quiz")
    ap.add_argument("--no-images", action="store_true", help="Don't pass --with-images")
    ap.add_argument("--dry-run", action="store_true", help="Pass --dry-run to every script")
    ap.add_argument("--print", dest="print_only", action="store_true", help="Just print the commands, don't run them")
    args = ap.parse_args()

    if args.date is None:
        t = date.today() + timedelta(days=1)
        args.date = f"{t.day}/{t.month}/{t.year}"

    idiom_items = items(parse_range(args.idioms), args.idiom_item or args.item)
    ows_items = items(parse_range(args.ows), args.ows_item or args.item)
    vocab_quiz_items = ",".join(f"{args.vocab}({w})" for w in parse_range(args.vocab_quiz))

    py = sys.executable
    img = [] if args.no_images else ["--with-images"]
    dry = ["--dry-run"] if args.dry_run else []
    commands = {
        "vocab": [py, "upload_vocab.py", "--file", str(args.vocab), "--date", args.date, *img, *dry],
        "idioms": [py, "upload_idioms.py", "--items", idiom_items, "--date", args.date, *img, *dry],
        "ows": [py, "upload_ows.py", "--items", ows_items, "--date", args.date, *img, *dry],
        "quiz": [py, "upload_quiz.py", "--vocab-items", vocab_quiz_items, "--idiom-items", idiom_items,
                 "--ows-items", ows_items, "--date", args.date, *dry],
    }

    print(f"Date: {args.date}")
    for step in STEPS:
        if step not in args.only:
            continue
        cmd = commands[step]
        print(f"\n=== {step} ===\n" + subprocess.list2cmdline(["python", *cmd[1:]]))
        if args.print_only:
            continue
        if subprocess.run(cmd).returncode != 0:
            print(f"\n'{step}' failed - stopping. Fix it and re-run with --only {' '.join(STEPS[STEPS.index(step):])}",
                  file=sys.stderr)
            sys.exit(1)
    print("\nDone.")


if __name__ == "__main__":
    main()
