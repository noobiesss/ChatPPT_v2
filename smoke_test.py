"""Offline smoke test for ChatPPT V2 (no API keys, no network).

Run from the project root after ANY code change:   python smoke_test.py
It checks that (1) every module imports, (2) a deck with all 11 layouts builds and re-opens,
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
        {"slide_number": 9, "layout_type": "timeline_flow", "action_title": "Timeline title",
         "timeline_steps": [{"stage_label": f"Q{i}", "title": f"Step {i}", "description": "What happens."} for i in range(1, 5)],
         "bottom_takeaway": "So what."},
        {"slide_number": 10, "layout_type": "comparison_table", "action_title": "Table title",
         "table_headers": ["Criterion", "Option A", "Option B"],
         "table_rows": [["Cost", "Low", "High"], ["Speed", "Fast", "Slow"], ["Risk", "Medium", "Low"]],
         "bottom_takeaway": "So what."},
        {"slide_number": 11, "layout_type": "closing", "action_title": "Thank You", "subtitle_or_context": "Presented by Test"},
    ]


def check_builtin_decks() -> None:
    from chatppt.pptx_build import build_presentation_pptx
    from chatppt.presets import BUILTIN_THEMES

    slides = sample_slides()
    for theme in list(BUILTIN_THEMES)[:2]:
        data = build_presentation_pptx(slides, theme, None, False, "", [])
        out = Presentation(io.BytesIO(data))
        assert len(out.slides) == len(slides), theme
    # comparison_table renders a NATIVE table with the header row and every cell
    table_slide = out.slides[9]
    tables = [sh.table for sh in table_slide.shapes if sh.has_table]
    assert tables and len(tables[0].rows) == 4 and len(tables[0].columns) == 3
    assert tables[0].cell(0, 1).text == "Option A" and tables[0].cell(3, 2).text == "Low"
    # timeline_flow renders every stage label
    texts = {sh.text_frame.text for sh in out.slides[8].shapes if sh.has_text_frame}
    assert {"Q1", "Q4", "1", "4", "Step 1\nWhat happens."} <= texts, texts
    print(f"[ok] built-in decks: {len(slides)} slides, all 11 layouts (native table, timeline nodes)")


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
    from chatppt.outline import mentions_slide_count, parse_slide_count_change
    from chatppt.ui.editor_text import format_points_for_editor, parse_points_from_editor

    pts = [{"heading": "A", "detail": "B"}, {"heading": "C", "detail": "D"}]
    assert parse_points_from_editor(format_points_for_editor(pts)) == pts
    assert extract_clean_json_str('<think>hmm</think>```json\n{"a": 1}\n```') == '{"a": 1}'
    assert mentions_slide_count("make it 6 slides") and not mentions_slide_count("shorter titles")
    # Relative slide-count edits: new target = current deck size + delta (regression: "add one more slide" used to be trimmed back)
    assert parse_slide_count_change("add one more slide conclusion", 8) == 9
    assert parse_slide_count_change("add 2 more slides", 8) == 10
    assert parse_slide_count_change("add a conclusion slide", 8) == 9
    assert parse_slide_count_change("remove the last slide", 8) == 7
    assert parse_slide_count_change("\u52a0\u4e00\u9875\u7ed3\u8bba", 8) == 9
    assert parse_slide_count_change("make it 6 slides", 8) is None
    assert parse_slide_count_change("add a chart to slide 3", 8) is None
    assert parse_slide_count_change("shorter titles", 8) is None
    assert parse_slide_count_change("add a thank you slide", 8) == 9
    assert parse_slide_count_change("add thank you slide", 8) == 9
    assert parse_slide_count_change("remove the thank you slide", 8) == 7
    assert parse_slide_count_change("remove duplicates from slide 3", 8) is None
    print("[ok] helpers")


def check_closing_slide() -> None:
    from chatppt.outline import _trim_slides_to_count, sanitize_outline_dict
    from chatppt.schemas import LayoutType, PresentationOutline, SlideSpec

    def spec(i, layout):
        return SlideSpec(slide_number=i + 1, layout_type=layout, action_title=f"T{i}")

    # A misplaced closing slide is moved to the end, kept unique and given no takeaway banner
    layouts = ["cover", "closing", "bullet_cards", "bullet_cards", "bullet_cards"]
    out = sanitize_outline_dict(PresentationOutline(deck_title="d", slides=[spec(i, lt) for i, lt in enumerate(layouts)]))
    kinds = [s["layout_type"] for s in out]
    assert kinds[-1] == "closing" and kinds.count("closing") == 1, kinds
    assert out[-1]["bottom_takeaway"] is None
    # Trimming never drops the closing slide or the conclusion right before it
    deck = [spec(i, lt) for i, lt in enumerate(["cover", "bullet_cards", "bullet_cards", "bullet_cards", "closing"])]
    trimmed = _trim_slides_to_count(deck, 4)
    assert len(trimmed) == 4 and trimmed[-1].layout_type == LayoutType.CLOSING and trimmed[-2].action_title == "T3"
    print("[ok] closing slide: placement, no takeaway, trim-safe")


def check_new_layouts() -> None:
    from chatppt.outline import normalize_table, sanitize_outline_dict
    from chatppt.schemas import PresentationOutline, SlideSpec
    from chatppt.ui.editor_text import format_steps_for_editor, format_table_for_editor, parse_steps_from_editor, parse_table_from_editor

    # Ragged rows are padded / cut to the header count; a 1-column table is rejected
    headers, rows = normalize_table(["A", "B", "C"], [["1"], ["1", "2", "3", "4"], ["", ""]])
    assert headers == ["A", "B", "C"] and rows == [["1", "", ""], ["1", "2", "3"]]
    assert normalize_table(["Only"], [["x"]]) == ([], [])
    pts = [{"heading": "H1", "detail": "D1"}, {"heading": "H2", "detail": "D2"}, {"heading": "H3", "detail": "D3"}]
    slides = [
        SlideSpec(slide_number=1, layout_type="cover", action_title="c"),
        SlideSpec(slide_number=2, layout_type="comparison_table", action_title="bad table", points=pts),
        SlideSpec(slide_number=3, layout_type="timeline_flow", action_title="steps from points", points=pts),
        SlideSpec(slide_number=4, layout_type="timeline_flow", action_title="empty timeline"),
    ]
    out = sanitize_outline_dict(PresentationOutline(deck_title="d", slides=slides))
    assert out[1]["layout_type"] != "comparison_table"  # no table data -> falls back to cards
    assert out[2]["layout_type"] == "timeline_flow" and [s["title"] for s in out[2]["timeline_steps"]] == ["H1", "H2", "H3"]
    assert out[3]["layout_type"] != "timeline_flow"
    steps = [{"stage_label": "Q1", "title": "Pilot", "description": "Run it"}]
    assert parse_steps_from_editor(format_steps_for_editor(steps)) == steps
    assert parse_table_from_editor(format_table_for_editor(["A", "B"], [["x", "y"]])) == (["A", "B"], [["x", "y"]])
    print("[ok] timeline / table: sanitizer fallbacks, rectangular tables, editor round-trip")


def check_report_images() -> None:
    from PIL import Image

    from chatppt.extraction import prepare_report_image
    from chatppt.images import assign_report_images, resolve_slide_image_bytes
    from chatppt.schemas import IMAGE_MODE_PLACEHOLDER, IMAGE_MODE_REPORT_FIRST

    def png(w, h, mode="RGB"):
        buf = io.BytesIO()
        Image.new(mode, (w, h), "white").save(buf, format="PNG")
        return buf.getvalue()

    assert prepare_report_image(png(64, 64)) is None  # logo / icon size is skipped
    assert prepare_report_image(b"not an image") is None
    big = prepare_report_image(png(3000, 1500))
    assert big and max(big["width"], big["height"]) <= 2400  # huge scans are downscaled
    assert prepare_report_image(png(400, 300, "RGBA"))["bytes"][1:4] == b"PNG"  # alpha kept as PNG

    rep, img, txt = IMAGE_MODE_REPORT_FIRST, "image_split", "bullet_cards"
    deck = [{"layout_type": txt}, {"layout_type": img, "image_source_mode": rep},
            {"layout_type": img, "image_source_mode": rep, "report_image_index": 0},
            {"layout_type": img, "image_source_mode": IMAGE_MODE_PLACEHOLDER},
            {"layout_type": img, "image_source_mode": rep}]
    # Explicit pick wins, the others get the remaining figures in order; placeholder slides are skipped
    assert assign_report_images(deck, 3, rep) == {2: 0, 1: 1, 4: 2}
    assert assign_report_images(deck, 1, rep) == {2: 0}  # out of figures -> others fall back
    images = [{"name": "a", "bytes": b"A"}, {"name": "b", "bytes": b"B"}]
    slide = {"layout_type": img, "image_source_mode": rep, "report_image_index": 1}
    assert resolve_slide_image_bytes(slide, None, False, "", images) == b"B"
    print("[ok] report images: size filter, downscale, assignment, resolver")


def check_assistant() -> None:
    from chatppt.chat_assistant import build_assistant_prompt, parse_assistant_reply

    assert parse_assistant_reply("ACTION: REVISE").action == "revise"
    assert parse_assistant_reply("<think>hmm</think>ACTION: REVISE\n").action == "revise"
    dec = parse_assistant_reply("ACTION: ANSWER\nBecause the data is a comparison.")
    assert dec.action == "answer" and dec.reply == "Because the data is a comparison."
    assert parse_assistant_reply("**ACTION: ANSWER**\nHi there").reply == "Hi there"
    assert parse_assistant_reply("Just a plain answer").action == "answer"
    assert parse_assistant_reply("").action == "revise" and parse_assistant_reply("ACTION: ANSWER").action == "revise"
    prompt = build_assistant_prompt(
        "why three cards?", "SOURCE_MARKER text", sample_slides(), [{"role": "user", "content": "HISTORY_MARKER"}]
    )
    for marker in ("SOURCE_MARKER", "Three cards title", "KPI title", "HISTORY_MARKER", "why three cards?"):
        assert marker in prompt, marker
    # Targeted edits: PATCH with valid slide numbers; anything unusable degrades to a full revise
    dec = parse_assistant_reply("ACTION: PATCH: 3, 5", deck_size=9)
    assert dec.action == "patch" and dec.slides == [3, 5], dec
    assert parse_assistant_reply("**ACTION: PATCH 2-4**", deck_size=9).slides == [2, 3, 4]
    assert parse_assistant_reply("ACTION: PATCH: 12", deck_size=9).action == "revise"
    assert parse_assistant_reply("ACTION: PATCH:", deck_size=9).action == "revise"
    assert parse_assistant_reply("ACTION: PATCH: 1-9", deck_size=9).action == "revise"  # too many slides
    # A long source is cut to relevant excerpts: a fact far past the first 12k chars must still reach the model
    long_source = "Intro.\n\n" + "\n\n".join("Filler paragraph about logistics. " * 20 for _ in range(60))
    long_source += "\n\nChurn fell to 3.2% after the NEEDLE loyalty program.\n\n" + "Tail filler. " * 400
    prompt = build_assistant_prompt("what happened to churn after the loyalty program?", long_source, [], [])
    assert "NEEDLE" in prompt and len(prompt) < 14000
    print("[ok] assistant routing + prompt context + patch protocol")


def check_retrieval() -> None:
    from chatppt.retrieval import select_relevant_context, split_into_chunks, tokenize

    assert "增长" in tokenize("营收增长62%") and "62%" in tokenize("营收增长62%")
    assert all(len(c) <= 1000 for c in split_into_chunks("x" * 5000))
    assert select_relevant_context("short text", "anything", 100) == "short text"
    doc = "\n\n".join(f"Section {i}: generic operations text. " * 12 for i in range(40))
    doc += "\n\n最终客户流失率下降到 3.2%。"
    out = select_relevant_context(doc, "客户流失率是多少", 3000)
    assert "3.2%" in out and len(out) <= 3000 + 20 and out.startswith("Section 0")
    # Nothing matches -> head of the document, like the old truncation
    assert select_relevant_context(doc, "zzz", 3000).startswith(doc[:2000])
    print("[ok] retrieval: CJK bigrams, chunking, BM25 recall, head fallback")


def check_patch() -> None:
    import chatppt.outline as outline
    from chatppt.schemas import SlidePatch, SlideSpec

    deck = sample_slides()
    deck[3]["image_source_mode"] = "Reserve Photo Placeholder"
    new = SlideSpec(slide_number=4, layout_type="bullet_cards", action_title="Patched title",
                    points=[{"heading": "H", "detail": "D"}])
    stray_closing = SlideSpec(slide_number=5, layout_type="closing", action_title="Bye")
    real = outline._request_outline
    try:
        outline._request_outline = lambda *a, **k: SlidePatch(slides=[new, stray_closing])
        out = outline.patch_slides("Groq", "k", "m", None, deck, [4, 5], "rewrite", "source", "tone", "English")
    finally:
        outline._request_outline = real
    assert len(out) == len(deck)
    # Untouched slides are the very same objects: a patch can never drift them
    assert all(out[i] is deck[i] for i in range(len(deck)) if i not in (3, 4))
    assert out[3]["action_title"] == "Patched title" and out[3]["slide_number"] == 4
    assert out[3]["image_source_mode"] == "Reserve Photo Placeholder"  # user's per-slide choice kept
    assert out[4]["layout_type"] == "bullet_cards"  # closing is only allowed as the last slide
    print("[ok] targeted patch: only target slides replaced, closing guard, image mode kept")


if __name__ == "__main__":
    for fn in (check_imports, check_builtin_decks, check_custom_template, check_helpers, check_closing_slide,
               check_new_layouts, check_report_images, check_assistant, check_retrieval, check_patch):
        fn()
    print("ALL SMOKE TESTS PASSED")
    sys.exit(0)