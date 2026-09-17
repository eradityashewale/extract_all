"""
Parses an "Idioms - N (date).docx" file into structured idiom entries.

Expected shape per idiom (repeated ~10x per file), matching what these files
actually contain:

    1. Idiom: meaning (translation)
    Example: example sentence
    Quiz - Idiom
        option A
        option B      <- the correct option is written in bold
        option C
        option D

    ... then at the end of the file:

    Photos:
    1. Idiom
    <image>
    2. Idiom
    <image>
    ...

Source files are messy (in at least one file the "Example:" line is written
after the quiz options instead of before). This parser tolerates that by
classifying each line rather than assuming a fixed order, and records a
`warnings` list per idiom so problem entries can be spotted instead of
imported wrong.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import docx
from docx.oxml.ns import qn

DEVANAGARI_RE = re.compile(r"[ऀ-ॿ]")
HEADER_RE = re.compile(r"^(\d+)\.\s*(?P<idiom>.+?):\s*(?P<rest>.*)$")
PHOTO_HEADER_RE = re.compile(r"^(\d+)\.\s*(?P<idiom>.+?):?\s*$")
TRANSLATION_RE = re.compile(r"^(?P<meaning>.*?)\s*\((?P<translation>[^)]*[ऀ-ॿ][^)]*)\)\s*$")

LABELS = ["A", "B", "C", "D"]


@dataclass
class QuizOption:
    label: str
    text: str
    is_correct: bool = False
    is_bold: bool = False  # scratch flag used while parsing; not persisted


@dataclass
class IdiomEntry:
    idiom_number: int
    idiom: str
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


def parse_docx(path: str) -> tuple[list[IdiomEntry], list[str]]:
    doc = docx.Document(path)
    paras = _get_paragraphs(doc)

    file_warnings: list[str] = []

    # Split into the definitions section and the "Photos" section.
    split_idx = next(
        (i for i, p in enumerate(paras) if p["text"].lower().startswith("photo")),
        len(paras),
    )
    body, photos = paras[:split_idx], paras[split_idx + 1 :]

    # --- find headword line indexes ---
    header_idxs = [i for i, p in enumerate(body) if HEADER_RE.match(p["text"])]
    if not header_idxs:
        file_warnings.append("No numbered idiom headers found at all.")
        return [], file_warnings

    entries: list[IdiomEntry] = []
    for n, start in enumerate(header_idxs):
        end = header_idxs[n + 1] if n + 1 < len(header_idxs) else len(body)
        block = body[start:end]
        entries.append(_parse_block(block))

    # --- match images from the photos section ---
    current_idiom_idx = -1
    photo_idioms_seen: list[str] = []
    for p in photos:
        m = PHOTO_HEADER_RE.match(p["text"])
        if m:
            photo_idioms_seen.append(m.group("idiom").strip())
            current_idiom_idx += 1
            continue
        if p["rids"]:
            if current_idiom_idx < 0 or current_idiom_idx >= len(entries):
                file_warnings.append(f"Image found with no matching idiom slot (rid={p['rids'][0]}).")
                continue
            target = entries[current_idiom_idx]
            expected_idiom = photo_idioms_seen[current_idiom_idx] if current_idiom_idx < len(photo_idioms_seen) else None
            if expected_idiom and expected_idiom.lower() != target.idiom.lower():
                file_warnings.append(
                    f"Photo section idiom '{expected_idiom}' at position {current_idiom_idx + 1} "
                    f"doesn't match parsed idiom '{target.idiom}' at the same position; "
                    "assigned by position anyway."
                )
            target.image_rid = p["rids"][0]

    for e in entries:
        if e.image_rid is None:
            e.warnings.append("No image found for this idiom.")

    return entries, file_warnings


def _parse_block(block: list[dict]) -> IdiomEntry:
    header = HEADER_RE.match(block[0]["text"])
    number = int(header.group(1))
    idiom = header.group("idiom").strip()
    rest = header.group("rest").strip()

    meaning, translation = _split_translation(rest)
    entry = IdiomEntry(idiom_number=number, idiom=idiom, meaning=meaning, meaning_translation=translation)

    examples: list[str] = []
    quiz_question: str | None = None
    quiz_options: list[QuizOption] = []
    unexpected: list[str] = []

    for p in block[1:]:
        text, bold = p["text"], p["bold"]
        tl = text.lower()
        if tl.startswith("example"):
            ex = re.sub(r"^examples?\s*[:\-–]?\s*", "", text, flags=re.IGNORECASE).strip()
            if ex:
                examples.append(ex)
        elif tl.startswith("quiz") and quiz_question is None:
            quiz_question = re.sub(r"^quiz\s*[-–]?\s*", "", text, flags=re.IGNORECASE).strip()
        elif quiz_question is not None and len(quiz_options) < 4:
            quiz_options.append(QuizOption(label=LABELS[len(quiz_options)], text=text.strip(), is_bold=bold))
        else:
            unexpected.append(text)

    if not examples:
        entry.warnings.append("No 'Example' line found.")
    else:
        entry.example = examples[0]
        if len(examples) > 1:
            entry.warnings.append(f"Multiple 'Example' lines found ({len(examples)}); using the first.")

    if quiz_question is None:
        entry.warnings.append("No 'Quiz' question found.")
    entry.quiz_question = quiz_question
    entry.quiz_options = quiz_options

    if len(quiz_options) != 4:
        entry.warnings.append(f"Expected 4 quiz options, found {len(quiz_options)}.")

    if unexpected:
        entry.warnings.append(f"Unrecognized line(s) in block: {unexpected}")

    bold_opts = [o for o in quiz_options if o.is_bold]
    if len(bold_opts) == 1:
        bold_opts[0].is_correct = True
    elif len(bold_opts) == 0:
        entry.warnings.append("No quiz option is bold; can't tell which is correct.")
    else:
        entry.warnings.append(
            f"{len(bold_opts)} quiz options are bold ({', '.join(o.text for o in bold_opts)}); "
            "can't tell which is correct."
        )

    return entry
