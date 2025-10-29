# gen_pptx.py
from typing import List, Dict, Any, Optional
import io
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor


def _to_lines(val) -> List[str]:
    if not val:
        return []
    if isinstance(val, str):
        return [ln.strip() for ln in val.splitlines() if ln.strip()]
    if isinstance(val, list):
        out: List[str] = []
        for v in val:
            if isinstance(v, str):
                out.extend([ln.strip() for ln in v.splitlines() if ln.strip()])
        return out
    return []


def _clear_text_frame(tf):
    if tf is None:
        return
    try:
        tf.clear()
    except Exception:
        pass
    try:
        tf.word_wrap = True
        tf.auto_size = False
    except Exception:
        pass


def _hex_to_rgb(hexstr: str) -> RGBColor:
    s = hexstr.strip().lstrip("#")
    if len(s) == 3:
        s = "".join(ch * 2 for ch in s)
    try:
        r = int(s[0:2], 16); g = int(s[2:4], 16); b = int(s[4:6], 16)
        return RGBColor(r, g, b)
    except Exception:
        return RGBColor(255, 255, 255)


def _apply_font(p, *, size_pt: int, bold: bool, name: str, color: RGBColor):
    if p is None:
        return
    p.font.size = Pt(size_pt)
    p.font.bold = bold
    try: p.font.name = name
    except Exception: pass
    try: p.font.color.rgb = color
    except Exception: pass
    try:
        p.level = p.level or 0
        p._set_bullet(False)
    except Exception:
        pass


def _add_para(tf, text: str, *, size_pt: int, bold: bool, name: str, color: RGBColor):
    if tf is None: return
    p = tf.add_paragraph()
    p.text = text
    _apply_font(p, size_pt=size_pt, bold=bold, name=name, color=color)


def _add_bullet(tf, text: str, *, size_pt: int, level: int, name: str, color: RGBColor):
    if tf is None: return
    p = tf.add_paragraph()
    p.text = text
    try:
        p.level = level
        p._set_bullet(True)
    except Exception:
        if not p.text.startswith("• "):
            p.text = "• " + p.text
    _apply_font(p, size_pt=size_pt, bold=False, name=name, color=color)


def _ph_val(name: str):
    return getattr(PP_PLACEHOLDER, name, None)


def _find_title_placeholder(slide):
    for nm in ("TITLE", "CENTER_TITLE"):
        tv = _ph_val(nm)
        if tv is None: continue
        for shp in getattr(slide, "placeholders", []):
            try:
                if shp.placeholder_format.type == tv:
                    return shp
            except Exception:
                pass
    return None


def _content_bounds(prs: Presentation):
    slide_w = prs.slide_width
    slide_h = prs.slide_height
    left_margin = Inches(0.75)
    right_margin = Inches(0.75)
    top_margin = Inches(1.8)
    bottom_margin = Inches(1.0)
    left = left_margin
    top = top_margin
    width = slide_w - left_margin - right_margin
    height = slide_h - top_margin - bottom_margin
    return left, top, width, height


def _apply_slide_bg(slide, color_rgb: RGBColor):
    """Fill slide background with a solid color."""
    try:
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = color_rgb
    except Exception:
        pass


def build_storyboard_pptx(
    blocks: List[Dict[str, Any]],
    course_title: str = "Course",
    *,
    font_name: str = "Calibri",
    font_color: str = "#111111",
    bg_color: str = "#FFFFFF",
) -> io.BytesIO:
    prs = Presentation()
    slide_w = prs.slide_width
    text_rgb = _hex_to_rgb(font_color)
    bg_rgb = _hex_to_rgb(bg_color)

    # Title slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    _apply_slide_bg(title_slide, bg_rgb)

    title_ph = _find_title_placeholder(title_slide)
    if title_ph and title_ph.has_text_frame:
        tf = title_ph.text_frame
        _clear_text_frame(tf)
        p = tf.add_paragraph(); p.text = course_title
        _apply_font(p, size_pt=44, bold=True, name=font_name, color=text_rgb)
    else:
        tb = title_slide.shapes.add_textbox(Inches(0.75), Inches(0.7), slide_w - Inches(1.5), Inches(1.1))
        tf = tb.text_frame
        _clear_text_frame(tf)
        _add_para(tf, course_title, size_pt=44, bold=True, name=font_name, color=text_rgb)

    try:
        sub = title_slide.placeholders[1]
        if sub and sub.has_text_frame:
            _clear_text_frame(sub.text_frame)
            _add_para(sub.text_frame, "Storyboard generated automatically", size_pt=20, bold=False, name=font_name, color=text_rgb)
    except Exception:
        pass

    # Content slides
    for b in blocks:
        if not isinstance(b, dict) or b.get("type") != "slide": continue

        title = (b.get("title") or "Slide").strip()
        body_text = _to_lines(b.get("text"))
        bullets = _to_lines(b.get("bullets"))
        narration = (b.get("narration") or "").strip()

        s = prs.slides.add_slide(prs.slide_layouts[1])
        _apply_slide_bg(s, bg_rgb)

        tph = _find_title_placeholder(s)
        if tph and tph.has_text_frame:
            tf_title = tph.text_frame
            _clear_text_frame(tf_title)
            _add_para(tf_title, title, size_pt=40, bold=True, name=font_name, color=text_rgb)
        else:
            tb = s.shapes.add_textbox(Inches(0.75), Inches(0.6), slide_w - Inches(1.5), Inches(1.1))
            tf_title = tb.text_frame
            _clear_text_frame(tf_title)
            _add_para(tf_title, title, size_pt=40, bold=True, name=font_name, color=text_rgb)

        left, top, width, height = _content_bounds(prs)
        if body_text and bullets:
            text_height = int(height * 0.55)
            bullets_height = height - text_height - Inches(0.2)
        else:
            text_height = height
            bullets_height = 0

        if body_text:
            tb_text = s.shapes.add_textbox(left, top, width, text_height)
            tf_text = tb_text.text_frame
            _clear_text_frame(tf_text)
            for line in body_text:
                _add_para(tf_text, line, size_pt=28, bold=False, name=font_name, color=text_rgb)

        if bullets:
            top_bul = top + (text_height + Inches(0.2) if body_text else 0)
            tb_bul = s.shapes.add_textbox(left, top_bul, width, bullets_height or height)
            tf_bul = tb_bul.text_frame
            _clear_text_frame(tf_bul)
            for line in bullets:
                _add_bullet(tf_bul, line, size_pt=24, level=0, name=font_name, color=text_rgb)

        try:
            notes_tf = s.notes_slide.notes_text_frame
            _clear_text_frame(notes_tf)
            if narration:
                for line in narration.splitlines():
                    _add_para(notes_tf, line, size_pt=14, bold=False, name=font_name, color=text_rgb)
        except Exception:
            pass

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf