import io
import csv
from typing import List, Dict, Any, Optional


def _get_letter(idx: Optional[int]) -> str:
    return "" if idx is None or idx < 0 else "ABCD"[idx] if idx < 4 else ""


def _row_from_quiz_single(b: Dict[str, Any]) -> List[str]:
    """
    Backward-compat: write quiz_single as a row (single answer by index).
    """
    title = b.get("title", "")
    question = b.get("question", "")

    choices = b.get("choices") or []
    # Normalize to exactly 4 options
    choices = [str(c or "") for c in choices[:4]] + [""] * max(0, 4 - len(choices))
    correct = _get_letter(b.get("correctIndex"))
    fb_ok = b.get("feedbackCorrect", "")
    fb_bad = b.get("feedbackIncorrect", "")
    return [title, question, choices[0], choices[1], choices[2], choices[3], correct, fb_ok, fb_bad]


def _row_from_quiz_legacy(b: Dict[str, Any]) -> List[str]:
    """
    Legacy 'quiz' shape: A..D + answer (single or multiple letters like "A,C").
    """
    title = b.get("title", "")
    question = b.get("question", "")
    A = b.get("A", "")
    B = b.get("B", "")
    C = b.get("C", "")
    D = b.get("D", "")
    # Keep multiple letters intact (e.g., "A,C")
    correct = b.get("Answer") or b.get("answer") or b.get("Correct") or b.get("correct") or ""
    correct = str(correct).strip().upper()
    fb_ok = b.get("feedback_correct", "") or b.get("FeedbackCorrect", "")
    fb_bad = b.get("feedback_incorrect", "") or b.get("FeedbackIncorrect", "")
    return [title, question, A, B, C, D, correct, fb_ok, fb_bad]


def build_quiz_csv(blocks: List[Dict[str, Any]]) -> io.BytesIO:
    """
    Build a CSV containing all quizzes.
    Columns: Title, Question, A, B, C, D, Correct (Single or Multiple), FeedbackCorrect, FeedbackIncorrect
    """
    mem = io.StringIO(newline="")
    writer = csv.writer(mem)
    writer.writerow(["Title", "Question", "A", "B", "C", "D", "Correct (Single or Multiple)", "FeedbackCorrect", "FeedbackIncorrect"])

    for b in blocks:
        if not isinstance(b, dict):
            continue
        t = b.get("type", "")
        if t == "quiz":
            writer.writerow(_row_from_quiz_legacy(b))
        elif t == "quiz_single":
            # still supported for compatibility if any upstream still emits it
            writer.writerow(_row_from_quiz_single(b))
        # ignore non-quiz blocks

    data = mem.getvalue().encode("utf-8-sig")  # BOM so Excel opens cleanly
    return io.BytesIO(data)
