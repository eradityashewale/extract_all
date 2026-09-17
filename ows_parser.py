"""
Parses a "Substitution(s) - N ( D Month ).docx" one-word-substitution file
into structured entries.

Expected shape per entry (repeated ~10x per file):

    1.Word (pos) - definition. (translation)
    Example -
    example sentence
    Quiz - question text
        option A
        option B      <- the correct option is written in bold
        option C
        option D
    Vocab of/for quiz -
    OptionWord (pos) - meaning. (translation)   <- definition for each WRONG
    OptionWord (pos) - meaning. (translation)      option (the correct one
                                                    has no definition here,
                                                    since it's already
                                                    defined as the headword)

    ... then at the end of the file, a "<label> with photos" / "Photos:"
    marker, followed by:

    1.Word
    <image>
    2.Word
    <image>
    ...

Source files are inconsistent: some use an en-dash after the part of speech,
some use ":"; some put the example sentence on its own line, some inline
after "Example:"; the photo-section marker text itself varies
("Substitutions with photos", "Vocab with photos", "Photos:"). This parser
tolerates all of that.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import docx
from docx.oxml.ns import qn

HEADER_RE = re.compile(r"^(\d+)\.\s*(?P<word>.+?)\s*\((?P<pos>[^)]{1,15})\)\s*[:\-–—]\s*(?P<rest>.*)$")
PHOTO_HEADER_RE = re.compile(r"^(\d+)\.\s*(?P<word>.+?):?\s*$")
PHOTO_MARKER_RE = re.compile(r"^(?:[a-z]+\s+with\s+photos|photos:?)$", re.IGNORECASE)
TRANSLATION_RE = re.compile(r"^(?P<meaning>.*?)\s*\((?P<translation>[^)]*[ऀ-ॿ][^)]*)\)\s*$")
EXAMPLE_LABEL_RE = re.compile(r"^examples?\s*[:\-–—]?\s*", re.IGNORECASE)
DEVANAGARI_RE = re.compile(r"[ऀ-ॿ]")
DEF_RE = re.compile(r"^(?P<word>[^(]+?)\s*\((?P<pos>[^)]+)\)\s*[-–]?\s*(?P<rest>.*)$")
VOCAB_OF_QUIZ_RE = re.compile(r"^vocab\s+(?:of|for)\s+quiz\b", re.IGNORECASE)
# Requires the dash, so an option word that happens to start with "quiz"
# (e.g. "Quizzical") isn't mistaken for the "Quiz - ..." marker line.
QUIZ_LINE_RE = re.compile(r"^quiz\s*[-–:]", re.IGNORECASE)

LABELS = ["A", "B", "C", "D"]


@dataclass
class QuizOption:
    label: str
    text: str
    is_correct: bool = False
    description: str | None = None  # definition of the option, when the doc gives one (the wrong answers)
    is_bold: bool = False  # scratch flag used while parsing; not persisted


@dataclass
class OwsEntry:
    number: int
    word: str
    part_of_speech: str | None
    meaning: str
    meaning_translation: str | None = None
    example: str | None = None
    quiz_question: str | None = None
    quiz_options: list[QuizOption] = field(default_factory=list)
    image_rid: str | None = None
    warnings: list[str] = field(default_factory=list)


def _split_translation(text: str) -> tuple[str, str | None]:
    m = TRANSLATION_RE.match(text.strip())
    if m:
        return m.group("meaning").strip(" .,"), m.group("translation").strip()
    return text.strip(), None


def _get_paragraphs(doc: docx.document.Document) -> list[dict]:
    """Flatten to significant paragraphs (has text or an inline image), in order."""
    out = []
    for p in doc.paragraphs:
        text = p.text.strip()
        blips = p._p.findall(".//" + qn("a:blip"))
        rids = [b.get(qn("r:embed")) for b in blips]
        bold = any(r.bold for r in p.runs if r.text.strip())
        if text or rids:
            out.append({"text": text, "rids": rids, "bold": bool(bold)})
    return out


def parse_docx(path: str) -> tuple[list[OwsEntry], list[str]]:
    doc = docx.Document(path)
    paras = _get_paragraphs(doc)

    file_warnings: list[str] = []

    # Split into the definitions section and the "... with photos" / "Photos:" section.
    split_idx = next(
        (i for i, p in enumerate(paras) if PHOTO_MARKER_RE.match(p["text"])),
        len(paras),
    )
    body, photos = paras[:split_idx], paras[split_idx + 1 :]

    header_idxs = [i for i, p in enumerate(body) if HEADER_RE.match(p["text"])]
    if not header_idxs:
        file_warnings.append("No numbered entries found at all.")
        return [], file_warnings

    entries: list[OwsEntry] = []
    for n, start in enumerate(header_idxs):
        end = header_idxs[n + 1] if n + 1 < len(header_idxs) else len(body)
        entries.append(_parse_block(body[start:end]))

    # --- match images from the photos section (same layout as idiom_parser.py) ---
    current_idx = -1
    photo_words_seen: list[str] = []
    for p in photos:
        m = PHOTO_HEADER_RE.match(p["text"])
        if m:
            photo_words_seen.append(m.group("word").strip())
            current_idx += 1
            continue
        if p["rids"]:
            if current_idx < 0 or current_idx >= len(entries):
                file_warnings.append(f"Image found with no matching entry slot (rid={p['rids'][0]}).")
                continue
            target = entries[current_idx]
            expected = photo_words_seen[current_idx] if current_idx < len(photo_words_seen) else None
            if expected and expected.lower() != target.word.lower():
                file_warnings.append(
                    f"Photo section word '{expected}' at position {current_idx + 1} "
                    f"doesn't match parsed word '{target.word}' at the same position; "
                    "assigned by position anyway."
                )
            target.image_rid = p["rids"][0]

    for e in entries:
        if e.image_rid is None:
            e.warnings.append("No image found for this entry.")

    return entries, file_warnings


def _parse_block(block: list[dict]) -> OwsEntry:
    header = HEADER_RE.match(block[0]["text"])
    number = int(header.group(1))
    word = header.group("word").strip()
    pos = header.group("pos").strip()
    meaning, translation = _split_translation(header.group("rest").strip())

    entry = OwsEntry(number=number, word=word, part_of_speech=pos, meaning=meaning, meaning_translation=translation)

    example = None
    i = 1
    while i < len(block):
        text = block[i]["text"]
        tl = text.lower()
        if tl.startswith("example"):
            inline = EXAMPLE_LABEL_RE.sub("", text).strip()
            if inline:
                example = inline
                i += 1
            elif i + 1 < len(block):
                example = block[i + 1]["text"].strip()
                i += 2
            else:
                i += 1
            continue
        if QUIZ_LINE_RE.match(text):
            entry.quiz_question = re.sub(r"^quiz\s*[-–]?\s*", "", text, flags=re.IGNORECASE).strip()
            i += 1
            break
        i += 1

    if not example:
        entry.warnings.append("No 'Example' line found.")
    entry.example = example

    if entry.quiz_question is None:
        entry.warnings.append("No 'Quiz' question found.")
        return entry

    # --- quiz options: collect every line up to the "Vocab of/for quiz"
    # marker, or (some entries have no marker at all) up to the first
    # definition-shaped line, whichever comes first. Some entries give an
    # alternate phrasing of the question on an extra "or ..." line before
    # the real 4 options -- when there are more than 4 candidates, treat
    # the leading overflow as part of the question.
    candidates = []
    while (
        i < len(block)
        and not VOCAB_OF_QUIZ_RE.match(block[i]["text"])
        and not DEF_RE.match(block[i]["text"])
    ):
        candidates.append(block[i])
        i += 1
    if len(candidates) > 4:
        overflow, candidates = candidates[: len(candidates) - 4], candidates[-4:]
        entry.quiz_question = f"{entry.quiz_question} {' '.join(c['text'] for c in overflow)}".strip()
        entry.warnings.append(
            f"Quiz question spanned extra line(s) ({', '.join(c['text'] for c in overflow)}); merged into the question text."
        )
    for label, c in zip(LABELS, candidates):
        entry.quiz_options.append(QuizOption(label=label, text=c["text"].strip(), is_bold=c["bold"]))
    for label in LABELS[len(candidates):]:
        entry.warnings.append(f"Missing quiz option {label}.")
    if len(entry.quiz_options) != 4:
        entry.warnings.append(f"Expected 4 quiz options, found {len(entry.quiz_options)}.")

    # --- skip the "Vocab of/for quiz" marker line ---
    if i < len(block) and VOCAB_OF_QUIZ_RE.match(block[i]["text"]):
        i += 1

    # --- trailing option definitions (== wrong answers) ---
    while i < len(block):
        text = block[i]["text"]
        m = DEF_RE.match(text)
        if not m:
            i += 1
            continue
        def_word = m.group("word").strip()
        def_rest = m.group("rest").strip()
        j = i + 1
        if j < len(block) and re.match(r"^\([^)]*\)$", block[j]["text"]) and DEVANAGARI_RE.search(block[j]["text"]):
            def_rest = f"{def_rest} {block[j]['text']}".strip()
            j += 1
        def_meaning, def_translation = _split_translation(def_rest)
        full_desc = def_meaning + (f" ({def_translation})" if def_translation else "")

        match = next((o for o in entry.quiz_options if o.text.strip().lower() == def_word.lower()), None)
        if match:
            match.description = full_desc
        else:
            entry.warnings.append(f"Trailing definition for '{def_word}' doesn't match any quiz option.")
        i = j

    # The correct answer is the one option written in bold. Cross-check
    # against the trailing-definition heuristic (the correct option is the
    # one that *doesn't* get its own definition afterward), and fall back to
    # that heuristic entirely when the author forgot to bold the answer.
    bold_opts = [o for o in entry.quiz_options if o.is_bold]
    no_def_opts = [o for o in entry.quiz_options if o.description is None]

    if len(bold_opts) == 1:
        bold_opts[0].is_correct = True
        if len(no_def_opts) == 1 and no_def_opts[0] is not bold_opts[0]:
            entry.warnings.append(
                f"Bold option '{bold_opts[0].text}' disagrees with the trailing-definition heuristic "
                f"(which pointed to '{no_def_opts[0].text}'); trusted the bold formatting."
            )
    else:
        if len(bold_opts) > 1:
            entry.warnings.append(
                f"{len(bold_opts)} quiz options are bold ({', '.join(o.text for o in bold_opts)}); "
                "falling back to the trailing-definition heuristic."
            )
        if len(no_def_opts) == 1:
            no_def_opts[0].is_correct = True
        elif len(no_def_opts) == 0:
            entry.warnings.append("Every quiz option has a trailing definition; couldn't identify the correct answer.")
        else:
            entry.warnings.append(
                f"{len(no_def_opts)} quiz options have no trailing definition and none are bold "
                f"({', '.join(o.text for o in no_def_opts)}); can't tell which is correct."
            )

    return entry
