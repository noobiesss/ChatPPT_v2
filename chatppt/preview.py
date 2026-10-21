"""HTML/CSS wireframe preview of one slide (16:9) plus a text-overflow estimate, without building a .pptx.

Edit when: a renderer in pptx_render.py changes its geometry or font sizes (keep the numbers here in sync),
or when tuning the overflow estimate. Pure functions (no Streamlit); positions are in inches like pptx_render.
"""

import html
import math
import unicodedata
from typing import Dict, List, Optional, Sequence, Tuple

from chatppt.pptx_template import _ensure_white_text_contrast

SLIDE_W, SLIDE_H = 13.333, 7.5
_PT_W = SLIDE_W * 72  # slide width in points: 1pt = 100cqw / _PT_W
_LINE = 1.2  # PowerPoint single line spacing, in em
_INSET_X, _INSET_Y = 0.1, 0.05  # default text-box insets (inches)

# Scoped stylesheet; st.html is not iframed, so every class is prefixed
PREVIEW_CSS = """
<style>
.cpv-slide { position: relative; width: 100%; aspect-ratio: 16 / 9; container-type: inline-size; overflow: hidden;
  border-radius: 6px; box-shadow: 0 1px 4px rgba(0,0,0,.25); font-family: var(--cpv-font, system-ui, sans-serif); }
.cpv-slide * { box-sizing: border-box; }
.cpv-b { position: absolute; }
.cpv-t { position: absolute; display: flex; flex-direction: column; padding: 0.375cqw 0.75cqw;
  line-height: 1.2; overflow: visible; word-break: break-word; }
.cpv-t p { margin: 0; white-space: pre-wrap; }
.cpv-over { outline: 0.3cqw dashed #E5484D; outline-offset: -0.15cqw; background: rgba(229,72,77,.10); }
.cpv-img { display: flex; align-items: center; justify-content: center; text-align: center;
  background: repeating-linear-gradient(45deg, rgba(128,128,128,.18) 0 1cqw, rgba(128,128,128,.08) 1cqw 2cqw); }
.cpv-tbl { position: absolute; border-collapse: collapse; table-layout: fixed; }
.cpv-tbl td { padding: 0.375cqw 0.9cqw; vertical-align: middle; border: 0.12cqw solid rgba(255,255,255,.7);
  line-height: 1.2; word-break: break-word; }
</style>
"""


def _hex(color) -> str:
    return "#{:02X}{:02X}{:02X}".format(color[0], color[1], color[2])


def _pos(x: float, y: float, w: float, h: float) -> str:
    # Inches -> percent of the slide, so the wireframe scales with whatever width the page gives it
    return (f"left:{x / SLIDE_W * 100:.3f}%;top:{y / SLIDE_H * 100:.3f}%;"
            f"width:{w / SLIDE_W * 100:.3f}%;height:{h / SLIDE_H * 100:.3f}%;")


def _fs(pt: float) -> str:
    # cqw = 1% of the .cpv-slide container's width, so text keeps the same size relative to the slide
    return f"font-size:{pt * 100 / _PT_W:.4f}cqw;"


# --- Overflow estimate -------------------------------------------------------------------------------

def _char_em(ch: str) -> float:
    """Approximate advance width of one character in em."""
    # Wide / full-width characters (CJK, full-width punctuation) take a full em
    if unicodedata.east_asian_width(ch) in ("W", "F"):
        return 1.0
    if ch == " ":
        return 0.28
    if ch.isupper() or ch.isdigit():
        return 0.62
    if ch.isalpha():
        return 0.5
    return 0.35


def _lines(text: str, pt: float, bold: bool, width_in: float) -> int:
    # Wrapped lines per hard line = text width in points (sum of em widths x font size) / usable box width
    avail =max((width_in - 2 * _INSET_X) * 72, 1.0)
    total = 0
    for line in (text or "").split("\n") or [""]:
        em = sum(_char_em(c) for c in line) * (1.06 if bold else 1.0) * 1.05  # 5% slack for word wrapping
        total += max(1, math.ceil(em * pt / avail))
    return total


def text_height_pt(paras: Sequence[dict], width_in: float) -> float:
    """Estimated height of stacked paragraphs ({text, pt, bold, after}) in a box of width_in inches."""
    return sum(_lines(p["text"], p["pt"], p.get("bold", False), width_in) * p["pt"] * _LINE + p.get("after", 0)
               for p in paras)


def overflows(paras: Sequence[dict], w: float, h: float) -> bool:
    return text_height_pt(paras, w) > (h - 2 * _INSET_Y) * 72 + 2


# --- Drawing primitives ------------------------------------------------------------------------------

class _Canvas:
    def __init__(self) -> None:
        self.parts: List[str] = []
        self.warnings: List[str] = []

    def rect(self, x, y, w, h, fill, border=None, radius=0.0, oval=False, extra="") -> None:
        style = _pos(x, y, w, h) + f"background:{fill};"
        if border:
            style += f"border:0.1cqw solid {border};"
        if oval:
            style += "border-radius:50%;"
        elif radius:
            style += f"border-radius:{radius / SLIDE_W * 100:.3f}cqw;"
        self.parts.append(f'<div class="cpv-b" style="{style}{extra}"></div>')

    def text(self, x, y, w, h, paras: List[dict], warn: Optional[str] = None, align="left",
             anchor="top", check=True) -> None:
        over = check and overflows(paras, w, h)
        if over and warn:
            self.warnings.append(warn)
        justify = {"top": "flex-start", "middle": "center", "bottom": "flex-end"}[anchor]
        body = "".join(
            f'<p style="{_fs(p["pt"])}color:{p["color"]};font-weight:{700 if p.get("bold") else 400};'
            f'{"font-style:italic;" if p.get("italic") else ""}margin-bottom:{p.get("after", 0) * 100 / _PT_W:.4f}cqw;'
            f'text-align:{p.get("align", align)};">{html.escape(p["text"])}</p>'
            for p in paras
        )
        cls = "cpv-t cpv-over" if over else "cpv-t"
        self.parts.append(f'<div class="{cls}" style="{_pos(x, y, w, h)}justify-content:{justify};">{body}</div>')

    def card(self, x, y, w, h, theme) -> None:
        self.rect(x, y, w, h, _hex(theme["card_bg"]), _hex(theme["card_border"]), radius=0.04 * min(w, h))


def _p(text, pt, color, bold=False, after=0.0, **kw) -> dict:
    return {"text": str(text or ""), "pt": pt, "color": color, "bold": bold, "after": after, **kw}


# --- Layouts (geometry mirrors pptx_render.py) -------------------------------------------------------

def _header(c: _Canvas, s: dict, idx: int, total: int, t: Dict[str, str], builtin: bool) -> None:
    if builtin:
        c.rect(0, 0, SLIDE_W, 0.12, t["accent"])
    c.text(0.8, 0.42, 11.7, 1.05, [_p(s.get("action_title", ""), 26, t["title"], bold=True)],
           warn="Title is too long for its box (it runs into the content).")
    c.rect(0.8, 1.52, 11.7, 0.03, t["accent"])
    takeaway = (s.get("bottom_takeaway") or "").strip()
    if takeaway:
        c.rect(0.8, 6.35, 10.8, 0.68, t["takeaway_bg"], radius=0.1)
        c.text(1.0, 6.42, 10.4, 0.55, [_p(f"💡 Key Takeaway: {takeaway}", 13.5, t["takeaway_text"], bold=True)],
               warn="Takeaway banner text is too long (it spills out of the banner).")
    c.text(11.7, 6.45, 0.8, 0.45, [_p(f"{idx + 1} / {total}", 11, t["accent"])], align="right", check=False)


def _point_paras(pt: dict, t, head_pt, det_pt, after) -> List[dict]:
    return [_p(pt.get("heading", ""), head_pt, t["title"], bold=True, after=after),
            _p(pt.get("detail", ""), det_pt, t["body"])]


def _bullet_cards(c, s, t, theme):
    points = (s.get("points") or [])[:4]
    count = max(len(points), 1)
    card_h = (4.35 - 0.16 * (count - 1)) / count
    for i, pt in enumerate(points):
        top = 1.78 + i * (card_h + 0.16)
        c.card(0.8, top, 11.7, card_h, theme)
        c.rect(1.05, top + (card_h - 0.46) / 2, 0.46, 0.46, t["accent_strong"], oval=True)
        c.text(1.05, top + (card_h - 0.46) / 2, 0.46, 0.46, [_p(i + 1, 14, "#FFFFFF", bold=True)],
               align="center", anchor="middle", check=False)
        c.text(1.7, top + 0.08, 10.5, card_h - 0.16, _point_paras(pt, t, 17, 14, 3),
               warn=f"Card {i + 1} text overflows its card.")


def _three_cards(c, s, t, theme):
    points = list((s.get("points") or [])[:3])
    while len(points) < 3:
        points.append({"heading": f"Pillar {len(points) + 1}", "detail": "Key supporting detail."})
    for i, pt in enumerate(points):
        left = 0.8 + i * 4.0
        c.card(left, 1.8, 3.7, 4.3, theme)
        c.rect(left + 0.25, 2.08, 0.8, 0.08, t["accent"])
        c.text(left + 0.25, 2.3, 3.2, 3.6, _point_paras(pt, t, 19, 14.5, 12), warn=f"Card {i + 1} text overflows its card.")


def _two_column(c, s, t, theme):
    for i, side in enumerate(("left_column", "right_column")):
        col = s.get(side) or {}
        left = 0.8 + i * 6.05
        c.card(left, 1.8, 5.65, 4.3, theme)
        c.rect(left + 0.2, 2.0, 5.25, 0.55, t["accent_strong"] if i else t["takeaway_bg"], radius=0.08)
        c.text(left + 0.2, 2.0, 5.25, 0.55, [_p(col.get("column_title", ""), 16, "#FFFFFF", bold=True)],
               align="center", anchor="middle", warn=f"{'Right' if i else 'Left'} column header is too long.")
        paras = []
        for item in (col.get("items") or [])[:4]:
            paras += [_p(f"▪ {item.get('heading', '')}", 15.5, t["title"], bold=True, after=2),
                      _p(f"   {item.get('detail', '')}", 13.5, t["body"], after=10)]
        c.text(left + 0.25, 2.7, 5.15, 3.2, paras, warn=f"{'Right' if i else 'Left'} column text overflows its card.")


def _kpi_metrics(c, s, t, theme):
    kpis = (s.get("kpis") or [])[:4]
    count = max(len(kpis), 1)
    card_w = (11.7 - 0.28 * (count - 1)) / count
    for i, k in enumerate(kpis):
        left = 0.8 + i * (card_w + 0.28)
        c.card(left, 1.85, card_w, 4.2, theme)
        c.text(left + 0.2, 2.35, card_w - 0.4, 3.4, [
            _p(k.get("value", ""), 46 if count <= 3 else 38, t["accent"], bold=True, after=12),
            _p(k.get("label", ""), 18, t["title"], bold=True, after=10),
            _p(k.get("context", ""), 13.5, t["body"]),
        ], align="center", warn=f"KPI {i + 1} text overflows its card.")


def _side_points(c, points, left, card_w, pad_x, pad_y, head_pt, det_pt, t, theme):
    points = points[:3]
    count = max(len(points), 1)
    card_h = (4.3 - 0.15 * (count - 1)) / count
    for i, pt in enumerate(points):
        top = 1.8 + i * (card_h + 0.15)
        c.card(left, top, card_w, card_h, theme)
        c.text(left + pad_x, top + pad_y, card_w - 2 * pad_x, card_h - 2 * pad_y, _point_paras(pt, t, head_pt, det_pt, 4),
               warn=f"Insight card {i + 1} text overflows its card.")


def _chart(c, s, t, theme):
    ch = s.get("chart") or {}
    cats, vals = list(ch.get("categories") or []), list(ch.get("values") or [])
    n = min(len(cats), len(vals))
    cats, vals = cats[:n], vals[:n]
    c.card(0.8, 1.8, 6.4, 4.3, theme)
    c.text(1.0, 1.95, 6.0, 0.45, [_p(ch.get("chart_title", ""), 13, t["title"], bold=True)], align="center", check=False)
    kind = ch.get("chart_type", "bar")
    if n and kind == "pie":
        # One conic-gradient stop per slice, alternating accent and body colors
        total =sum(max(v, 0) for v in vals) or 1
        stops, acc = [], 0.0
        for i, v in enumerate(vals):
            start, acc = acc, acc + max(v, 0) / total * 100
            stops.append(f"{t['accent']} {start:.2f}% {acc:.2f}%" if i % 2 == 0 else f"{t['body']} {start:.2f}% {acc:.2f}%")
        c.rect(2.6, 2.55, 2.8, 2.8, "conic-gradient(" + ",".join(stops) + ")", oval=True)
    elif n:
        top_v = max(max(vals), 0) or 1
        slot = 5.4 / n
        for i, v in enumerate(vals):
            bar_h = max(v, 0) / top_v * 2.9
            x = 1.3 + i * slot
            if kind == "line":
                c.rect(x + slot / 2 - 0.07, 5.45 - bar_h - 0.07, 0.14, 0.14, t["accent"], oval=True)
            else:
                c.rect(x + slot * 0.2, 5.45 - bar_h, slot * 0.6, bar_h, t["accent"])
            c.text(x, 5.5, slot, 0.35, [_p(cats[i], 9, t["body"])], align="center", check=False)
    _side_points(c, s.get("points") or [], 7.45, 5.05, 0.2, 0.1, 16, 13, t, theme)


def _image_split(c, s, t, theme, image_label):
    img = s.get("image") or {}
    if img.get("placement") == "right":
        text_left, img_left = 0.8, 6.95
    else:
        img_left, text_left = 0.8, 6.55
    c.card(img_left, 1.8, 5.55, 4.3, theme)
    c.parts.append(
        f'<div class="cpv-b cpv-img" style="{_pos(img_left + 0.15, 1.95, 5.25, 3.5)}{_fs(13)}color:{t["body"]};">'
        f"{html.escape(image_label or 'Image')}</div>"
    )
    if img.get("caption"):
        c.text(img_left + 0.2, 5.5, 5.15, 0.5, [_p(f"▲ {img['caption']}", 11.5, t["accent"], italic=True)],
               align="center", warn="Image caption is too long (it wraps below the card).")
    _side_points(c, s.get("points") or [], text_left, 5.95, 0.22, 0.12, 16.5, 13.5, t, theme)


def _timeline(c, s, t, theme):
    steps = (s.get("timeline_steps") or [])[:5]
    if len(steps) < 2:
        _bullet_cards(c, s, t, theme)
        return
    n = len(steps)
    col_w = 11.7 / n
    title_pt = {2: 18, 3: 17, 4: 15.5}.get(n, 14.5)
    desc_pt = {2: 14.5, 3: 14, 4: 13}.get(n, 12)
    c.rect(0.8 + col_w / 2, 2.42 + 0.26 - 0.03, col_w * (n - 1), 0.06, t["card_border"])
    for i, step in enumerate(steps):
        col_left = 0.8 + i * col_w
        cx = col_left + col_w / 2
        c.text(col_left + 0.1, 1.82, col_w - 0.2, 0.5, [_p(step.get("stage_label", ""), 13, t["accent"], bold=True)],
               align="center", warn=f"Step {i + 1} stage label is too long.")
        c.rect(cx - 0.26, 2.42, 0.52, 0.52, t["accent_strong"], border=t["card_bg"], oval=True)
        c.text(cx - 0.26, 2.42, 0.52, 0.52, [_p(i + 1, 14, "#FFFFFF", bold=True)], align="center", anchor="middle",
               check=False)
        card_left, card_w = col_left + 0.1, col_w - 0.2
        c.card(card_left, 3.15, card_w, 2.95, theme)
        c.rect(card_left + 0.2, 3.37, 0.6, 0.07, t["accent"])
        c.text(card_left + 0.15, 3.53, card_w - 0.3, 2.45, [
            _p(step.get("title", ""), title_pt, t["title"], bold=True, after=6),
            _p(step.get("description", ""), desc_pt, t["body"]),
        ], warn=f"Step {i + 1} text overflows its card.")


def _table(c, s, t, theme):
    headers = (s.get("table_headers") or [])[:5]
    rows = [list(r)[:len(headers)] for r in (s.get("table_rows") or [])[:6]]
    if len(headers) < 2 or not rows:
        _bullet_cards(c, s, t, theme)
        return
    rows = [r + [""] * (len(headers) - len(r)) for r in rows]
    n_cols, n_rows = len(headers), len(rows)
    header_h = 0.6
    row_h = min(0.85, (4.3 - header_h) / n_rows)
    first_w = 11.7 * (0.4 if n_cols == 2 else 0.26)
    widths = [first_w] + [(11.7 - first_w) / (n_cols - 1)] * (n_cols - 1)
    body_pt = 13.5 if n_cols <= 3 and n_rows <= 4 else 12 if n_rows <= 5 else 11

    # Rows grow to fit their text in PowerPoint: estimate the real table height
    def row_need(cells, pt, bold):
        return max(text_height_pt([_p(x, pt, "", bold=bold)], w - 0.14) / 72 + 0.1 for x, w in zip(cells, widths))

    est = max(header_h, row_need(headers, 14.5, True)) + sum(max(row_h, row_need(r, body_pt, False)) for r in rows)
    over = 1.8 + est > 6.3
    if over:
        c.warnings.append("Table text is too long: rows grow and the table runs into the takeaway banner.")
    cols = "".join(f'<col style="width:{w / 11.7 * 100:.2f}%">' for w in widths)
    head = "".join(
        f'<td style="background:{t["accent_strong"]};color:#FFFFFF;font-weight:700;{_fs(14.5)}'
        f'text-align:{"left" if j == 0 else "center"};height:{header_h / SLIDE_W * 100:.3f}cqw;">{html.escape(h)}</td>'
        for j, h in enumerate(headers))
    body = ""
    for r, row in enumerate(rows, start=1):
        fill = t["card_bg"] if r % 2 else t["bg_content"]
        body += "<tr>" + "".join(
            f'<td style="background:{fill};color:{t["title"] if j == 0 else t["body"]};font-weight:{700 if j == 0 else 400};'
            f'{_fs(body_pt)}text-align:{"left" if j == 0 else "center"};height:{row_h / SLIDE_W * 100:.3f}cqw;">'
            f"{html.escape(x)}</td>" for j, x in enumerate(row)) + "</tr>"
    x, y = 0.8, 1.8
    cls = "cpv-tbl cpv-over" if over else "cpv-tbl"
    c.parts.append(f'<table class="{cls}" style="left:{x / SLIDE_W * 100:.3f}%;top:{y / SLIDE_H * 100:.3f}%;'
                   f'width:{11.7 / SLIDE_W * 100:.3f}%;"><colgroup>{cols}</colgroup><tr>{head}</tr>{body}</table>')


def _title_slide(c, s, t, kind):
    if kind == "cover":
        c.rect(1.1, 2.2, 1.4, 0.14, t["accent"])
        c.text(1.1, 2.5, 11.0, 2.2, [_p(s.get("action_title", ""), 42, t["title_cover"], bold=True)],
               warn="Cover title is too long for its box.")
        if (s.get("subtitle_or_context") or "").strip():
            c.text(1.1, 4.8, 10.5, 1.4, [_p(s["subtitle_or_context"], 20, t["subtitle_cover"])],
                   warn="Cover subtitle is too long for its box.")
    elif kind == "section_divider":
        c.rect(1.1, 2.6, 0.9, 0.12, t["accent"], radius=0.06)
        c.text(1.1, 2.9, 11.0, 1.6, [_p(s.get("action_title", ""), 38, t["title_cover"], bold=True)],
               warn="Section title is too long for its box.")
        if (s.get("subtitle_or_context") or "").strip():
            c.text(1.1, 4.5, 10.0, 1.2, [_p(s["subtitle_or_context"], 18, t["subtitle_cover"])],
                   warn="Section overview is too long for its box.")
    else:  # closing
        title = (s.get("action_title") or "").strip() or "Thank You"
        size = 60 if len(title) <= 24 else 44 if len(title) <= 48 else 36
        c.text(1.1, 2.3, SLIDE_W - 2.2, 2.0, [_p(title, size, t["title_cover"], bold=True)], align="center",
               anchor="middle", warn="Closing phrase is too long for its box.")
        c.rect((SLIDE_W - 1.4) / 2, 4.5, 1.4, 0.14, t["accent"])
        if (s.get("subtitle_or_context") or "").strip():
            c.text(1.1, 4.9, SLIDE_W - 2.2, 1.0, [_p(s["subtitle_or_context"], 22, t["subtitle_cover"])],
                   align="center", warn="Presenter line is too long for its box.")


def render_slide_preview(
    slide: dict,
    theme: Dict[str, object],
    idx: int,
    total: int,
    builtin_theme: bool = True,
    image_label: Optional[str] = None,
) -> Tuple[str, List[str]]:
    """
    Return (html, warnings) for one slide. `theme` is a pptx theme dict (RGBColor values, optional fonts).
    The html needs PREVIEW_CSS on the page once. Warnings list the text boxes estimated to overflow.
    """
    # RGBColor is a 3-tuple, so this converts every color in the theme to CSS hex and skips the font names
    t = {k: _hex(v) for k, v in theme.items() if isinstance(v, tuple) and len(v) == 3}
    t["title"], t["body"] = t["title_content"], t["body_content"]
    t["accent_strong"] = _hex(_ensure_white_text_contrast(theme["accent"]))
    layout = slide.get("layout_type", "bullet_cards")
    c = _Canvas()
    if layout in ("cover", "section_divider", "closing"):
        bg = t["bg_section"] if layout == "section_divider" else t["bg_cover"]
        _title_slide(c, slide, t, layout)
    else:
        bg = t["bg_content"]
        _header(c, slide, idx, total, t, builtin_theme)
        draw = {
            "three_cards": _three_cards, "two_column": _two_column, "kpi_metrics": _kpi_metrics, "chart_slide": _chart,
            "timeline_flow": _timeline, "comparison_table": _table,
        }.get(layout)
        if layout == "image_split":
            _image_split(c, slide, t, theme, image_label)
        elif draw:
            draw(c, slide, t, theme)
        else:
            _bullet_cards(c, slide, t, theme)
    font = theme.get("font_body") or theme.get("font_heading")
    font_css = f"--cpv-font:'{html.escape(font)}',system-ui,sans-serif;" if font else ""
    return (f'<div class="cpv-slide" style="background:{bg};{font_css}">{"".join(c.parts)}</div>', c.warnings)
