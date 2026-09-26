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

# "1.Word (pos) - meaning" or, without a part of speech, "5. Gambol – meaning".
# Without a pos a plain hyphen only counts as the separator when it has a
# space on one side, so hyphenated words stay intact.
_SEP = r"(?:\s*\((?P<pos>[^)]{1,15})\)\s*[:\-–—]|\s*[:–—]|\s+-|-\s)"
HEADER_RE = re.compile(r"^(\d+)\.\s*(?P<word>[^():–—]+?)" + _SEP + r"\s*(?P<rest>.+)$")
# Photo headings: "1.Word", "2." (word left blank) -- or, in some files, just "Word".
PHOTO_HEADER_RE = re.compile(r"^(\d+)\.\s*(?P<word>.*?)[\s:\-–—]*$")
PHOTO_MARKER_RE = re.compile(r"^(?:[a-z]+\s+with\s+photos|photos:?)$", re.IGNORECASE)
TRANSLATION_RE = re.compile(r"^(?P<meaning>.*?)\s*\((?P<translation>[^)]*[ऀ-ॿ][^)]*)\)\s*$")
EXAMPLE_LABEL_RE = re.compile(r"^examples?\s*[:\-–—]?\s*", re.IGNORECASE)
DEVANAGARI_RE = re.compile(r"[ऀ-ॿ]")
# "Word (pos) - meaning", or "Word: meaning" / "Word – meaning" with no pos.
DEF_RE = re.compile(
    r"^(?P<word>[^(:–—]{1,40}?)(?:\s*\((?P<pos>[^)]+)\)\s*[-–:]?|\s*[:–—]|\s+-|-\s)\s*(?P<rest>.*)$"
)
HINT_RE = re.compile(r"^hint\b", re.IGNORECASE)
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

    # Some files have no marker at all (e.g. a second "Vocab for quiz" where the
    # marker should be); then the photos start at the heading of the first image.
    if not photos:
        first_img = next((i for i, p in enumerate(paras) if p["rids"]), None)
        if first_img is not None:
            if first_img > 0 and not paras[first_img]["text"] and len(paras[first_img - 1]["text"]) < 40:
                first_img -= 1
            body, photos = paras[:first_img], paras[first_img:]

    # Entry numbers only go up, so a stray numbered line inside a block isn't
    # mistaken for a new entry.
    header_idxs = []
    last_number = 0
    for i, p in enumerate(body):
        m = HEADER_RE.match(p["text"])
        if m and int(m.group(1)) > last_number:
            header_idxs.append(i)
            last_number = int(m.group(1))
    if not header_idxs:
        file_warnings.append("No numbered entries found at all.")
        return [], file_warnings

    entries: list[OwsEntry] = []
    for n, start in enumerate(header_idxs):
        end = header_idxs[n + 1] if n + 1 < len(header_idxs) else len(body)
        entries.append(_parse_block(body[start:end]))

    # --- match images from the photos section ---
    # Each heading is matched to an entry by word; when the heading is blank or
    # names something else (e.g. a quiz option), fall back to its position.
    # Older files leave the headings unnumbered, and some put the image in the
    # same paragraph as its heading.
    numbered = any(PHOTO_HEADER_RE.match(p["text"]) for p in photos)
    slot, target = -1, None
    for p in photos:
        m = PHOTO_HEADER_RE.match(p["text"])
        if m or (not numbered and p["text"]):
            slot += 1
            heading = (m.group("word") if m else p["text"]).strip()
            target = _match_photo_heading(heading, slot, entries)
            if target is None:
                file_warnings.append(f"Photo heading '{p['text']}' at position {slot + 1} has no matching entry.")
            elif heading and _norm(heading) != _norm(target.word):
                file_warnings.append(
                    f"Photo section word '{heading}' at position {slot + 1} doesn't match parsed word "
                    f"'{target.word}'; assigned by position."
                )
        if p["rids"]:
            if target is None:
                file_warnings.append(f"Image found with no matching entry slot (rid={p['rids'][0]}).")
            elif target.image_rid is None:  # extra images for the same entry: keep the first
                target.image_rid = p["rids"][0]

    for e in entries:
        if e.image_rid is None:
            e.warnings.append("No image found for this entry.")

    return entries, file_warnings


def _norm(word: str) -> str:
    return re.sub(r"[^a-z]", "", word.lower())


def _match_photo_heading(heading: str, slot: int, entries: list[OwsEntry]) -> OwsEntry | None:
    h = _norm(heading)
    if h:
        free = [e for e in entries if e.image_rid is None]
        for e in free:
            if _norm(e.word) == h:
                return e
        for e in free:  # "Forfeited" -> "Forfeit"
            w = _norm(e.word)
            if len(w) >= 4 and (h.startswith(w) or w.startswith(h)):
                return e
    return entries[slot] if slot < len(entries) else None


def _parse_block(block: list[dict]) -> OwsEntry:
    header = HEADER_RE.match(block[0]["text"])
    number = int(header.group(1))
    word = header.group("word").strip()
    pos = header.group("pos").strip() if header.group("pos") else None
    meaning, translation = _split_translation(header.group("rest").strip())

    entry = OwsEntry(number=number, word=word, part_of_speech=pos, meaning=meaning, meaning_translation=translation)

    # The example usually comes before the quiz, but some entries put it after.
    example = None
    for k in range(1, len(block)):
        text = block[k]["text"]
        if text.lower().startswith("example"):
            inline = EXAMPLE_LABEL_RE.sub("", text).strip()
            if inline:
                example = inline
            elif k + 1 < len(block):
                example = block[k + 1]["text"].strip()
            break

    i = next((k for k in range(1, len(block)) if QUIZ_LINE_RE.match(block[k]["text"])), len(block))
    if i < len(block):
        entry.quiz_question = re.sub(r"^quiz\s*[-–:]?\s*", "", block[i]["text"], flags=re.IGNORECASE).strip()
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
    # the leading overflow as part of the question. A "Hint", "Example" or
    # repeated "Quiz -" line after the options also ends them.
    candidates = []
    while (
        i < len(block)
        and not VOCAB_OF_QUIZ_RE.match(block[i]["text"])
        and not DEF_RE.match(block[i]["text"])
        and not HINT_RE.match(block[i]["text"])
        and not QUIZ_LINE_RE.match(block[i]["text"])
        and not block[i]["text"].lower().startswith("example")
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

    # --- skip ahead to the "Vocab of/for quiz" marker line, if there is one ---
    marker = next((k for k in range(i, len(block)) if VOCAB_OF_QUIZ_RE.match(block[k]["text"])), None)
    if marker is not None:
        i = marker + 1

    # --- trailing option definitions (== wrong answers) ---
    while i < len(block):
        text = block[i]["text"]
        m = DEF_RE.match(text)
        if not m or HINT_RE.match(text):
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

    # The option that is the headword itself is the answer; this beats the bold
    # formatting, which is occasionally on the wrong option.
    headword_opts = [o for o in entry.quiz_options if _norm(o.text) == _norm(entry.word)]
    if len(headword_opts) == 1:
        headword_opts[0].is_correct = True
        bold_opts = [o for o in entry.quiz_options if o.is_bold]
        if bold_opts and headword_opts[0] not in bold_opts:
            entry.warnings.append(
                f"Bold option '{bold_opts[0].text}' isn't the headword; used the headword option as the answer."
            )
        return entry

    # Otherwise the correct answer is the one option written in bold. Cross-check
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
