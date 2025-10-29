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
    raw = request.form.get("script", "")
    fmt = request.form.get("format", "pptx")

    base_name = _sanitize_filename(request.form.get("filename", ""), "export")
    font_name = (request.form.get("font_name") or "Calibri").strip()
    font_color = (request.form.get("font_color") or "#111111").strip()
    bg_color = (request.form.get("bg_color") or "#FFFFFF").strip()

    blocks = parse_script(raw)

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
