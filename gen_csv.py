# gen_csv.py
import io
import csv
from typing import List, Dict, Any, Optional


def _get_letter(idx: Optional[int]) -> str:
    return "" if idx is None or idx < 0 else "ABCD"[idx] if idx < 4 else ""


def _row_from_quiz_single(b: Dict[str, Any]) -> List[str]:
    title = b.get("title", "")
    question = b.get("question", "")
    choices = b.get("choices") or []
    # Normalize to exactly 4 options
    choices = [str(c or "") for c in choices[:4]] + [""] * max(0, 4 - len(choices))
    correct = _get_letter(b.get("correctIndex"))
    fb_ok = b.get("feedbackCorrect", "")
    fb_bad = b.get("feedbackIncorrect", "")
    return [title, question, choices[0], choices[1], choices[2], choices[3], correct, fb_ok, fb_bad]


def _row_from_legacy_quiz(b: Dict[str, Any]) -> List[str]:
    """Support a legacy shape where options are keyed A/B/C/D and answer is a letter."""
    title = b.get("title", "")
    question = b.get("question", "")
    A = b.get("A", "")
    B = b.get("B", "")
    C = b.get("C", "")
    D = b.get("D", "")
    # Some older fields might be 'Answer' or 'Correct'
    correct = b.get("Answer") or b.get("answer") or b.get("Correct") or b.get("correct") or ""
    correct = str(correct).strip()[:1].upper()
    fb_ok = b.get("feedbackcorrect", "") or b.get("FeedbackCorrect", "")
    fb_bad = b.get("feedbackincorrect", "") or b.get("FeedbackIncorrect", "")
    return [title, question, A, B, C, D, correct, fb_ok, fb_bad]


def build_quiz_csv(blocks: List[Dict[str, Any]]) -> io.BytesIO:
    """
    Build a CSV containing all single-choice quizzes from parsed blocks.
    Columns: Title, Question, A, B, C, D, Correct, FeedbackCorrect, FeedbackIncorrect
    """
    mem = io.StringIO(newline="")
    writer = csv.writer(mem)
    writer.writerow(["Title", "Question", "A", "B", "C", "D", "Correct", "FeedbackCorrect", "FeedbackIncorrect"])

    for b in blocks:
        if not isinstance(b, dict):
            continue

        t = b.get("type", "")

        if t == "quiz_single":
            writer.writerow(_row_from_quiz_single(b))
        elif t == "quiz":  # legacy support
            writer.writerow(_row_from_legacy_quiz(b))
        # ignore non-quiz blocks

    data = mem.getvalue().encode("utf-8-sig")  # BOM so Excel opens cleanly
    return io.BytesIO(data)