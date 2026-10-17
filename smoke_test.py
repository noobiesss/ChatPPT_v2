"""Offline smoke test for ChatPPT V2 (no API keys, no network).

Run from the project root after ANY code change:   python smoke_test.py
It checks that (1) every module imports, (2) a deck with all 8 layouts builds and re-opens,
(3) custom .pptx template inheritance still strips border frames and keeps template fonts,
(4) a few pure helpers still behave. Exit code 0 = all good.
"""

import importlib
import io
import pkgutil
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt


def check_imports() -> None:
    import chatppt

    names = [m.name for m in pkgutil.walk_packages(chatppt.__path__, "chatppt.")]
    for name in names:
        importlib.import_module(name)
    importlib.import_module("app")
    print(f"[ok] imported {len(names)} modules + app")


def sample_slides() -> list:
    pts = [{"heading": f"Point {i}", "detail": "Supporting detail text for this point."} for i in range(1, 4)]
    return [
        {"slide_number": 1, "layout_type": "cover", "action_title": "Deck Title", "subtitle_or_context": "Subtitle"},
        {"slide_number": 2, "layout_type": "section_divider", "action_title": "Section", "subtitle_or_context": "Overview"},
        {"slide_number": 3, "layout_type": "bullet_cards", "action_title": "Bullet cards title", "points": pts, "bottom_takeaway": "So what."},
        {"slide_number": 4, "layout_type": "three_cards", "action_title": "Three cards title", "points": pts, "bottom_takeaway": "So what."},
        {"slide_number": 5, "layout_type": "two_column", "action_title": "Two columns title",
         "left_column": {"column_title": "Before", "items": pts[:2]}, "right_column": {"column_title": "After", "items": pts[:2]},
         "bottom_takeaway": "So what."},
        {"slide_number": 6, "layout_type": "kpi_metrics", "action_title": "KPI title",
         "kpis": [{"value": "+62%", "label": "Growth", "context": "Year over year"}, {"value": "3.5x", "label": "Speed", "context": "Faster queries"}],
         "bottom_takeaway": "So what."},
        {"slide_number": 7, "layout_type": "chart_slide", "action_title": "Chart title", "points": pts[:2],
         "chart": {"chart_type": "bar", "chart_title": "Chart", "categories": ["A", "B", "C"], "series_name": "Rev ($M)", "values": [1.0, 2.0, 3.0]},
         "bottom_takeaway": "So what."},
        {"slide_number": 8, "layout_type": "image_split", "action_title": "Image split title", "points": pts[:2],
         "image": {"preferred_style": "stock_photo", "visual_prompt": "x", "search_keywords": "office", "caption": "Caption", "placement": "left"},
         "image_source_mode": "Reserve Photo Placeholder", "bottom_takeaway": "So what."},
    ]


def check_builtin_decks() -> None:
    from chatppt.pptx_build import build_presentation_pptx
    from chatppt.presets import BUILTIN_THEMES

    slides = sample_slides()
    for theme in list(BUILTIN_THEMES)[:2]:
        data = build_presentation_pptx(slides, theme, None, False, "", [])
        assert len(Presentation(io.BytesIO(data)).slides) == len(slides), theme
    print(f"[ok] built-in decks: {len(slides)} slides, all 8 layouts")


def _make_template() -> bytes:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    layout = prs.slide_layouts[6]
    tmp = prs.slides.add_slide(layout)  # the template's one sample slide (explicit Georgia / Verdana runs)
    frame = tmp.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.5), Inches(12.3), Inches(6.5))
    frame.fill.background()
    frame.name = "TemplateFrame"
    el = frame._element
    el.getparent().remove(el)
    prs.slide_masters[0].shapes._spTree.append(el)
    for name, size, text in (("Georgia", 36, "Big title"), ("Verdana", 14, "body text")):
        tb = tmp.shapes.add_textbox(Inches(1), Inches(1 if size > 20 else 3), Inches(8), Inches(1))
        run = tb.text_frame.paragraphs[0].add_run()
        run.text, run.font.size, run.font.name = text, Pt(size), name
    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()


def check_custom_template() -> None:
    from chatppt.pptx_build import build_presentation_pptx

    data = build_presentation_pptx(
        sample_slides()[:3], "Corporate Blue", None, False, "", [], custom_template_bytes=_make_template()
    )
    out = Presentation(io.BytesIO(data))
    leftovers = [s.name for s in out.slide_masters[0].shapes if not s.is_placeholder]
    assert not leftovers, f"master decorations still present: {leftovers}"
    assert len(out.slides) == 3
    fonts = {
        run.font.name
        for slide in out.slides
        for shape in slide.shapes
        if shape.has_text_frame
        for para in shape.text_frame.paragraphs
        for run in para.runs
    }
    assert {"Georgia", "Verdana"} <= fonts, f"template fonts not applied: {fonts}"
    print("[ok] custom template: border frame stripped, template fonts applied")


def check_helpers() -> None:
    from chatppt.llm_client import extract_clean_json_str
    from chatppt.outline import mentions_slide_count
    from chatppt.ui.editor_text import format_points_for_editor, parse_points_from_editor

    pts = [{"heading": "A", "detail": "B"}, {"heading": "C", "detail": "D"}]
    assert parse_points_from_editor(format_points_for_editor(pts)) == pts
    assert extract_clean_json_str('<think>hmm</think>```json\n{"a": 1}\n```') == '{"a": 1}'
    assert mentions_slide_count("make it 6 slides") and not mentions_slide_count("shorter titles")
    print("[ok] helpers")


if __name__ == "__main__":
    for fn in (check_imports, check_builtin_decks, check_custom_template, check_helpers):
        fn()
    print("ALL SMOKE TESTS PASSED")
    sys.exit(0)
