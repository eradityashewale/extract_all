"""
Parses a "Vocab - N with photos.docx" file into structured word entries.

Expected shape per word (repeated ~10x per file), matching what these files
actually contain:

    1. Word (pos) - meaning. (translation)
    Examples -
        example sentence 1
        example sentence 2
    Synonyms - syn1, syn2, syn3
    Hint - mnemonic text                 <- optional
    Quiz - Word means
        option A
        option B
        option C
        option D
    OptionWord (pos) - meaning.          <- definition for each WRONG option
    OptionWord (pos) - meaning.          <- (the correct option has no
                                             definition here, since it's
                                             already defined as the headword)

    ... then at the end of the file:

    Vocab with photos -
    1. Word
    <image>
    2. Word
    <image>
    ...

Source files are messy (typos like "Exampless", inconsistent dashes/spaces,
translations that wrap onto their own paragraph). This parser is written to
tolerate that and record a `warnings` list per word rather than silently
guessing, so problem files can be spotted instead of imported wrong.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import docx
from docx.oxml.ns import qn

DEVANAGARI_RE = re.compile(r"[ऀ-ॿ]")
HEADER_RE = re.compile(r"^(\d+)\.\s*(?P<word>[^(]+?)\s*\((?P<pos>[^)]+)\)\s*[-–]?\s*(?P<rest>.*)$")
DEF_RE = re.compile(r"^(?P<word>[^(]+?)\s*\((?P<pos>[^)]+)\)\s*[-–]?\s*(?P<rest>.*)$")
PHOTO_HEADER_RE = re.compile(r"^(\d+)\.\s*(?P<word>.+)$")
LEADING_NUM_RE = re.compile(r"^\d+[.\)]\s*")
TRANSLATION_RE = re.compile(r"^(?P<meaning>.*?)\s*\((?P<translation>[^)]*[ऀ-ॿ][^)]*)\)\s*$")


@dataclass
class QuizOption:
    label: str
    text: str
    is_correct: bool = False
    description: str | None = None
    is_bold: bool = False  # scratch flag used while parsing; not persisted


@dataclass
class WordEntry:
    word_number: int
    word: str
    part_of_speech: str | None
    meaning: str
    meaning_translation: str | None = None
    hint: str | None = None
    examples: list[str] = field(default_factory=list)
    synonyms: list[str] = field(default_factory=list)
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


def parse_docx(path: str) -> tuple[list[WordEntry], list[str]]:
    doc = docx.Document(path)
    paras = _get_paragraphs(doc)

    file_warnings: list[str] = []

    # Split into the definitions section and the "Vocab with photos" section.
    split_idx = next(
        (i for i, p in enumerate(paras) if p["text"].lower().startswith("vocab with photo")),
        len(paras),
    )
    body, photos = paras[:split_idx], paras[split_idx + 1 :]

    # --- find headword line indexes ---
    header_idxs = [i for i, p in enumerate(body) if HEADER_RE.match(p["text"])]
    if not header_idxs:
        file_warnings.append("No numbered word headers found at all.")
        return [], file_warnings

    entries: list[WordEntry] = []
    for n, start in enumerate(header_idxs):
        end = header_idxs[n + 1] if n + 1 < len(header_idxs) else len(body)
        block = body[start:end]
        entries.append(_parse_block(block))

    # --- match images from the photos section ---
    current_word_idx = -1
    photo_words_seen: list[str] = []
    for p in photos:
        m = PHOTO_HEADER_RE.match(p["text"])
        if m:
            photo_words_seen.append(m.group("word").strip())
            current_word_idx += 1
            continue
        if p["rids"]:
            if current_word_idx < 0 or current_word_idx >= len(entries):
                file_warnings.append(f"Image found with no matching word slot (rid={p['rids'][0]}).")
                continue
            target = entries[current_word_idx]
            expected_word = photo_words_seen[current_word_idx] if current_word_idx < len(photo_words_seen) else None
            if expected_word and expected_word.lower() != target.word.lower():
                file_warnings.append(
                    f"Photo section word '{expected_word}' at position {current_word_idx + 1} "
                    f"doesn't match parsed word '{target.word}' at the same position; "
                    "assigned by position anyway."
                )
            target.image_rid = p["rids"][0]

    for e in entries:
        if e.image_rid is None:
            e.warnings.append("No image found for this word.")

    return entries, file_warnings


def _parse_block(block: list[dict]) -> WordEntry:
    header = HEADER_RE.match(block[0]["text"])
    number = int(header.group(1))
    word = header.group("word").strip()
    pos = header.group("pos").strip()
    rest = header.group("rest").strip()

    i = 1
    # An orphan "(translation)" that wrapped onto its own paragraph.
    if i < len(block) and re.match(r"^\([^)]*\)$", block[i]["text"]) and DEVANAGARI_RE.search(block[i]["text"]):
        rest = f"{rest} {block[i]['text']}".strip()
        i += 1

    meaning, translation = _split_translation(rest)
    entry = WordEntry(word_number=number, word=word, part_of_speech=pos, meaning=meaning, meaning_translation=translation)

    # --- Examples ---
    while i < len(block) and not block[i]["text"].lower().startswith("example"):
        i += 1
    if i < len(block):
        i += 1  # skip the "Examples -" header itself
        while i < len(block) and not block[i]["text"].lower().startswith("synonym"):
            ex = LEADING_NUM_RE.sub("", block[i]["text"]).strip()
            if ex:
                entry.examples.append(ex)
            i += 1
    else:
        entry.warnings.append("No 'Examples' section found.")

    # --- Synonyms ---
    if i < len(block) and block[i]["text"].lower().startswith("synonym"):
        syn_text = re.sub(r"^synonyms?\s*[-–]\s*", "", block[i]["text"], flags=re.IGNORECASE)
        entry.synonyms = [s.strip(" .") for s in re.split(r"[,/]", syn_text) if s.strip(" .")]
        i += 1
    else:
        entry.warnings.append("No 'Synonyms' line found.")

    # --- Hint (optional) ---
    if i < len(block) and block[i]["text"].lower().startswith("hint"):
        entry.hint = re.sub(r"^hint\s*[-–]?\s*", "", block[i]["text"], flags=re.IGNORECASE).strip()
        i += 1

    # --- Quiz question ---
    if i < len(block) and block[i]["text"].lower().startswith("quiz"):
        entry.quiz_question = re.sub(r"^quiz\s*[-–]?\s*", "", block[i]["text"], flags=re.IGNORECASE).strip()
        i += 1
    else:
        entry.warnings.append("No 'Quiz' question found.")

    # --- 4 options ---
    labels = ["A", "B", "C", "D"]
    for label in labels:
        if i < len(block) and not block[i]["text"].lower().startswith("quiz"):
            entry.quiz_options.append(
                QuizOption(label=label, text=block[i]["text"].strip(), is_bold=block[i]["bold"])
            )
            i += 1
        else:
            entry.warnings.append(f"Missing quiz option {label}.")
    if len(entry.quiz_options) != 4:
        entry.warnings.append(f"Expected 4 quiz options, found {len(entry.quiz_options)}.")

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

    # The correct answer is the one option written in bold, in its own quiz-option
    # paragraph. Cross-check against the trailing-definition heuristic (the correct
    # option is the one that *doesn't* get its own definition afterward), and fall
    # back to that heuristic entirely when the author forgot to bold the answer.
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
