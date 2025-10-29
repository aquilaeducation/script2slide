# app.py
from flask import Flask, request, send_file
from pathlib import Path
import io
import zipfile
import re

from gen_pptx import build_storyboard_pptx
from gen_csv import build_quiz_csv
from parser_md import parse_script  # keep this if your parser file is parser_md.py

# --- Flask setup ---
app = Flask(__name__)

# Cap request size to avoid accidental DoS (adjust if you need bigger uploads)
app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 * 1024  # 2 MB

# Resolve index.html relative to this file (works no matter cwd)
BASE_DIR = Path(__file__).resolve().parent
INDEX_PATH = BASE_DIR / "index.html"


# -------- utilities --------
_slug_re = re.compile(r"[^A-Za-z0-9._-]+")

def _sanitize_filename(name: str, default: str) -> str:
    if not name:
        return default
    name = name.strip()
    name = _slug_re.sub("-", name)
    return name or default


def _split_body_to_lines(body: str):
    """
    Normalize row 'body' content into logical lines:
    - Convert explicit '\\n' into newlines (already done upstream)
    - Split on actual newlines
    - Also split on <br> and <br/> HTML breaks if present
    - Trim empty lines
    """
    raw = (body or "")
    # Support common HTML line breaks if a sheet exported HTML
    raw = raw.replace("<br/>", "\n").replace("<br>", "\n").replace("<BR/>", "\n").replace("<BR>", "\n")
    lines = [ln.strip() for ln in raw.splitlines()]
    return [ln for ln in lines if ln]


# -------- routes --------
@app.get("/healthz")
def healthz():
    return "ok", 200


@app.get("/")
def index():
    if INDEX_PATH.is_file():
        return send_file(INDEX_PATH, mimetype="text/html; charset=utf-8")
    return "<h1>index.html not found</h1><p>Put index.html next to app.py.</p>", 500


@app.post("/export")
def export():
    # detect upload vs. pasted text
    uploaded = request.files.get("upload")
    raw = request.form.get("script", "")
    fmt = request.form.get("format", "pptx")

    base_name = _sanitize_filename(request.form.get("filename", ""), "export")
    font_name = (request.form.get("font_name") or "Calibri").strip()
    font_color = (request.form.get("font_color") or "#111111").strip()
    bg_color = (request.form.get("bg_color") or "#FFFFFF").strip()

    def _auto_blocks_from_table(file_storage):
        import pandas as pd

        name = (file_storage.filename or "").lower()
        # Read CSV/Excel
        if name.endswith(".csv"):
            df = pd.read_csv(file_storage)
        elif name.endswith(".tsv"):
            df = pd.read_csv(file_storage, sep="\t")
        elif name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(file_storage)  # requires openpyxl for .xlsx
        else:
            raise ValueError("Unsupported file type. Use .csv, .tsv, .xlsx or .xls")

        blocks = []
        BULLET_PREFIXES = ("- ", "• ", "* ", "– ", "— ")

        def row_to_block(title, body, narration):
            title_raw = (str(title).strip() if title is not None else "")
            title_norm = title_raw or "Slide"
            body = str(body or "").replace("\\n", "\n")  # honor explicit \n from spreadsheets
            narration = str(narration or "")

            # ---------- QUIZ DETECTION ----------
            # Accept "QUIZ:", "QUIZ SINGLE:", or "QUIZ MULTI:" prefixes in the Title column.
            title_upper = title_raw.upper()
            if title_upper.startswith("QUIZ"):
                # Allow "QUIZ:", "QUIZ SINGLE:", "QUIZ MULTI:" forms
                quiz_title = title_raw
                # Strip the leading prefix label to leave a cleaner quiz title if provided as "QUIZ: Something"
                # e.g., "QUIZ: Password Hygiene" -> "Password Hygiene"
                if ":" in quiz_title:
                    quiz_title = quiz_title.split(":", 1)[1].strip()

                # Parse fielded lines in the body:
                #   Question: ...
                #   A: ...
                #   B: ...
                #   C: ...
                #   D: ...
                #   Answer: A,C
                #   FeedbackCorrect: ...
                #   FeedbackIncorrect: ...
                fields = {"title": quiz_title}
                for ln in _split_body_to_lines(body):
                    if ":" in ln:
                        k, v = ln.split(":", 1)
                        fields[k.strip().lower()] = v.strip()

                return {
                    "type": "quiz",
                    "title": fields.get("title", ""),
                    "question": fields.get("question", ""),
                    "A": fields.get("a", ""),
                    "B": fields.get("b", ""),
                    "C": fields.get("c", ""),
                    "D": fields.get("d", ""),
                    # keep multiple answers like "A,C"
                    "answer": (fields.get("answer", "") or "").replace(" ", "").upper(),
                    "feedback_correct": fields.get("feedbackcorrect", ""),
                    "feedback_incorrect": fields.get("feedbackincorrect", ""),
                }

            # ---------- SLIDE (DEFAULT) ----------
            # Normalize lines for slide text/bullets
            lines = _split_body_to_lines(body)

            # If single-line with semicolons and no newlines, treat as a list
            if len(lines) <= 1 and ";" in body and "\n" not in body:
                parts = [p.strip() for p in body.split(";") if p.strip()]
                lines = parts

            # Check if it's bulleted
            is_bulleted = False
            bullets = []
            for ln in lines:
                matched_prefix = next((p for p in BULLET_PREFIXES if ln.startswith(p)), None)
                if matched_prefix:
                    is_bulleted = True
                    bullets.append(ln[len(matched_prefix):].strip())
                else:
                    # If any line starts with a dash-like token (e.g., "-text" without space), try to normalize
                    if ln.startswith("-") and not ln.startswith("- "):
                        is_bulleted = True
                        bullets.append(ln[1:].strip())
                    else:
                        bullets.append(ln)

            block = {"type": "slide", "title": title_norm, "narration": narration}
            if is_bulleted:
                block["bullets"] = bullets
            else:
                # treat as body text (each line → new paragraph)
                block["text"] = "\n".join(bullets)
            return block

        # Iterate rows by position to be robust even without headers
        for _, row in df.iterrows():
            t = row.iloc[0] if len(row) > 0 else ""
            b = row.iloc[1] if len(row) > 1 else ""
            n = row.iloc[2] if len(row) > 2 else ""
            blocks.append(row_to_block(t, b, n))
        return blocks

    # Build blocks either from upload or from pasted script
    if uploaded and uploaded.filename:
        try:
            blocks = _auto_blocks_from_table(uploaded)
        except Exception as e:
            return f"Failed to read file: {e}", 400
    else:
        # Use your existing plain-text parser
        blocks = parse_script(raw)

    # Export as before
    if fmt == "pptx":
        pptx_bytes = build_storyboard_pptx(
            blocks,
            course_title=base_name,
            font_name=font_name,
            font_color=font_color,
            bg_color=bg_color,
        )
        return send_file(
            pptx_bytes,
            as_attachment=True,
            download_name=f"{base_name}.pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    if fmt == "csv":
        csv_bytes = build_quiz_csv(blocks)
        return send_file(
            csv_bytes,
            as_attachment=True,
            download_name=f"{base_name}.csv",
            mimetype="text/csv",
        )

    if fmt == "zip":
        mem = io.BytesIO()
        with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as zf:
            pptx_bytes = build_storyboard_pptx(
                blocks,
                course_title=base_name,
                font_name=font_name,
                font_color=font_color,
                bg_color=bg_color,
            )
            csv_bytes = build_quiz_csv(blocks)
            zf.writestr(f"{base_name}.pptx", pptx_bytes.getvalue())
            zf.writestr(f"{base_name}.csv", csv_bytes.getvalue())
        mem.seek(0)
        return send_file(mem, as_attachment=True, download_name=f"{base_name}.zip", mimetype="application/zip")

    return "Unknown format", 400


if __name__ == "__main__":
    # Local dev server (not for production edge)
    app.run(host="0.0.0.0", port=8080, debug=False)
