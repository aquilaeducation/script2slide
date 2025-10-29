import re

# Section headers
SLIDE_RE = re.compile(r"^##\s*Slide:\s*(.+)$", re.I)
# Keep the same authoring header, but we'll emit legacy "quiz" blocks:
QUIZ_SINGLE_RE = re.compile(r"^##\s*Quiz:\s*Single\s*Choice\s*$", re.I)

# Generic "Field: value" matcher
FIELD_RE = re.compile(r"^(\w+):\s*(.*)$")


def parse_script(text: str):
    """
    Parse authoring text into a list of normalized blocks.

    Slide
    -----
    ## Slide: <title>
    Text:
    <freeform body copy... may be multiple lines>
    Bullets:
    - bullet one
    - bullet two
    Narration: (alias: Notes:)
    <freeform notes... may be multiple lines>
    Image: https://example.com/image.png
    Alt: Short alt text

    You may also omit "Text:" and just write freeform lines under the slide;
    they'll be captured as body text.

    Quiz (Multiple answers allowed)
    --------------------------------
    ## Quiz: Single Choice
    Title: Optional quiz title
    Question: Which option is best?
    A: Answer A
    B: Answer B
    C: Answer C
    D: Answer D
    Answer: C        # single answer
    # or:
    Answer: A,C      # multiple answers allowed (select all that apply)
    FeedbackCorrect: Nice job!
    FeedbackIncorrect: Give it another look.

    Returns a list of blocks. Slide dict shape:
      {
        "type": "slide",
        "title": str,
        "text": [str, ...],
        "bullets": [str, ...],
        "narration": str,
        "image": str,
        "alt": str
      }

    Quiz dict shape (legacy form that downstream exporters already support):
      {
        "type": "quiz",
        "title": str,
        "question": str,
        "A": str, "B": str, "C": str, "D": str,
        "answer": "C" or "A,C",   # single or multiple
        "feedback_correct": str,
        "feedback_incorrect": str
      }
    """
    # Normalize line endings and strip trailing spaces; keep empty lines
    lines = [ln.rstrip() for ln in text.splitlines()]
    i = 0
    blocks = []

    def at_boundary(j: int) -> bool:
        """True if the current line starts a new section/field list or is blank break."""
        if j >= len(lines):
            return True
        ln = lines[j].strip()
        if ln == "":
            return True
        if ln.startswith("- "):
            return False
        if SLIDE_RE.match(ln) or QUIZ_SINGLE_RE.match(ln) or FIELD_RE.match(ln):
            return True
        return False

    def collect_bullets(start: int):
        """Collect bullet lines beginning with '- ' until blank or next section/field."""
        out = []
        j = start
        while j < len(lines):
            ln = lines[j].strip()
            if ln.startswith("- "):
                out.append(ln[2:].strip())
                j += 1
                continue
            # stop on blank or any new section/field
            if ln == "" or FIELD_RE.match(ln) or SLIDE_RE.match(ln) or QUIZ_SINGLE_RE.match(ln):
                break
            # If it's not a bullet and not a field/section, it's stray text: stop bullets
            break
        return out, j

    def collect_textblock(start: int):
        """
        Collect contiguous freeform text lines until blank or next field/section.
        Lines are kept as-is (trimmed), joined by newline by the caller if needed.
        """
        out = []
        j = start
        while j < len(lines):
            raw = lines[j]
            ln = raw.strip()
            if ln == "":
                break
            if ln.startswith("- "):
                break
            if FIELD_RE.match(ln) or SLIDE_RE.match(ln) or QUIZ_SINGLE_RE.match(ln):
                break
            out.append(raw.strip())
            j += 1
        return out, j

    current = None

    while i < len(lines):
        ln_raw = lines[i]
        ln = ln_raw.strip()

        # Section starts
        m_slide = SLIDE_RE.match(ln)
        m_quiz_s = QUIZ_SINGLE_RE.match(ln)

        if m_slide:
            # flush previous block
            if current:
                blocks.append(current)
            current = {
                "type": "slide",
                "title": m_slide.group(1).strip(),
                "text": [],
                "bullets": [],
                "narration": "",
                "image": "",
                "alt": "",
            }
            i += 1
            continue

        if m_quiz_s:
            if current:
                blocks.append(current)
            # Emit legacy 'quiz' block (A-D + 'answer' which may be multiple letters)
            current = {
                "type": "quiz",
                "title": "",
                "question": "",
                "A": "",
                "B": "",
                "C": "",
                "D": "",
                "answer": "",  # e.g., "C" or "A,C"
                "feedback_correct": "",
                "feedback_incorrect": "",
            }
            i += 1
            continue

        # Inside a block, parse fields or freeform
        if current:
            fm = FIELD_RE.match(ln)
            if fm:
                key = fm.group(1).lower()
                val = fm.group(2).strip()

                if current["type"] == "slide":
                    if key in ("narration", "notes"):
                        # Multi-line narration: same-line value + following text block
                        narr_lines = [val] if val else []
                        more, i2 = collect_textblock(i + 1)
                        narr_lines += more
                        current["narration"] = "\n".join([t for t in narr_lines if t]).strip()
                        i = i2
                        continue

                    elif key in ("text", "body", "content"):
                        # Multi-line body text: same-line value + following text block
                        text_lines = [val] if val else []
                        more, i2 = collect_textblock(i + 1)
                        text_lines += more
                        current["text"] = [t for t in text_lines if t]
                        i = i2
                        continue

                    elif key == "bullets":
                        bullets, i2 = collect_bullets(i + 1)
                        current["bullets"] = bullets
                        i = i2
                        continue

                    elif key == "image":
                        current["image"] = val
                    elif key == "alt":
                        current["alt"] = val

                else:
                    # quiz (legacy shape)
                    if key == "title":
                        current["title"] = val
                    elif key == "question":
                        current["question"] = val
                    elif key in ("a", "b", "c", "d"):
                        # store directly into A/B/C/D
                        current[key.upper()] = val
                    elif key in ("correct", "answer"):
                        # allow single or multiple letters like "C" or "A,C"
                        current["answer"] = val.replace(" ", "").upper()
                    elif key == "feedbackcorrect":
                        current["feedback_correct"] = val
                    elif key == "feedbackincorrect":
                        current["feedback_incorrect"] = val

            else:
                # Freeform lines in a slide become body text (if not labeled)
                if current["type"] == "slide" and ln != "":
                    text_chunk, i2 = collect_textblock(i)
                    if text_chunk:
                        current.setdefault("text", [])
                        current["text"].extend(text_chunk)
                        i = i2
                        continue

        i += 1

    # flush last block
    if current:
        blocks.append(current)

    # Normalize types and defaults
    for b in blocks:
        if b.get("type") == "slide":
            b["text"] = b.get("text") or []
            b["bullets"] = b.get("bullets") or []
            b["narration"] = (b.get("narration") or "").strip()
            b["image"] = b.get("image") or ""
            b["alt"] = b.get("alt") or ""

    return blocks
