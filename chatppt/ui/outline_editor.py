"""Per-slide outline editor (expanders, layout-specific fields) with move / insert / delete buttons and a live
wireframe preview per slide.

Edit when: changing what can be edited on each slide before rendering, the per-slide structure buttons,
or how the preview is placed.
"""

from typing import Optional

import streamlit as st

from chatppt.deck_ops import can_delete, can_insert, can_move, delete_slide, insert_slide, move_slide, push_history
from chatppt.images import assign_report_images
from chatppt.outline import normalize_table, normalize_timeline_steps
from chatppt.pptx_template import template_theme_for_preview
from chatppt.presets import BUILTIN_THEMES
from chatppt.preview import PREVIEW_CSS, render_slide_preview
from chatppt.schemas import (
    ChartType,
    IMAGE_MODE_OPTIONS,
    IMAGE_MODE_PLACEHOLDER,
    IMAGE_MODE_REPORT_FIRST,
    IMAGE_MODE_UNSPLASH_FIRST,
    LayoutType,
    VisualStyle,
)
from chatppt.ui.common import IMAGE_MODE_HELP, IMAGE_MODE_LABELS, TPL_COLORS_FROM_TEMPLATE, TPL_MODE_CUSTOM
from chatppt.ui.state import mark_slides_changed
from chatppt.ui.editor_text import (
    format_kpis_for_editor,
    format_points_for_editor,
    format_steps_for_editor,
    format_table_for_editor,
    parse_kpis_from_editor,
    parse_points_from_editor,
    parse_steps_from_editor,
    parse_table_from_editor,
)


def _render_report_image_picker(chat: dict, slide: dict, idx: int, kp: str, default_img_mode: str) -> None:
    """Dropdown to pin a figure from the uploaded report to this slide, with a thumbnail of what will be used."""
    images = chat.get("report_images") or []
    if not images:
        slide["report_image_index"] = None
        st.caption("No usable figures in the uploaded report (logos and icons are skipped), so this slide uses Unsplash / AI.")
        return
    options = [None] + list(range(len(images)))
    cur = slide.get("report_image_index")
    pick = st.selectbox(
        "Report figure",
        options=options,
        index=options.index(cur) if cur in options else 0,
        format_func=lambda k: "Auto (next unused figure)" if k is None
        else f"#{k + 1} · {images[k]['name']} ({images[k].get('width', '?')}×{images[k].get('height', '?')})",
        key=f"{kp}rptimg_{idx}",
    )
    slide["report_image_index"] = pick
    shown = assign_report_images(chat["slides"], len(images), default_img_mode).get(idx)
    if shown is None:
        st.caption("Every figure is already used by other slides, so this slide falls back to Unsplash / AI.")
    else:
        st.image(images[shown]["bytes"], width=240, caption=f"Figure #{shown + 1} · {images[shown]['name']}")


def _render_slide_ops(container, slides: list, idx: int, kp: str) -> Optional[str]:
    """Icon buttons next to a slide's expander. Returns the clicked operation, if any."""
    clicked = None
    with container.container(horizontal=True, gap="small"):
        for op, icon, tip, enabled in (
            ("up", ":material/arrow_upward:", "Move up (the cover always stays first)", can_move(slides, idx, -1)),
            ("down", ":material/arrow_downward:", "Move down (a closing slide always stays last)", can_move(slides, idx, 1)),
            ("insert", ":material/add:", "Insert a blank slide after this one", can_insert(slides)),
            ("delete", ":material/delete:", "Delete this slide (the cover cannot be deleted)", can_delete(slides, idx)),
        ):
            if st.button(icon, key=f"{kp}{op}_{idx}", help=tip, type="tertiary", disabled=not enabled):
                clicked = op
    return clicked


def _preview_theme() -> tuple:
    """(theme dict, is_builtin) matching what 'Render presentation' will use right now."""
    ss = st.session_state
    theme_name = ss.get("opt_theme") or next(iter(BUILTIN_THEMES))
    if ss.get("opt_tpl_mode") == TPL_MODE_CUSTOM and ss.get("custom_tpl_bytes"):
        use_colors = ss.get("opt_tpl_color_mode", TPL_COLORS_FROM_TEMPLATE) == TPL_COLORS_FROM_TEMPLATE
        return template_theme_for_preview(ss["custom_tpl_bytes"], theme_name, use_colors), False
    return BUILTIN_THEMES.get(theme_name, BUILTIN_THEMES["Corporate Blue"]), True


def _image_label(slide: dict, idx: int, assignment: dict, default_img_mode: str) -> str:
    """What the image box of an image_split slide will contain when rendered."""
    mode = slide.get("image_source_mode") or default_img_mode
    if mode == IMAGE_MODE_PLACEHOLDER:
        return "Photo placeholder"
    if mode == IMAGE_MODE_REPORT_FIRST and idx in assignment:
        return f"Report figure #{assignment[idx] + 1}"
    return "Unsplash / AI image"


def _apply_slide_op(chat: dict, op: str, idx: int, image_mode: str) -> None:
    """Apply a structure edit (one Undo away), keep the edited slide open, sync the slider, rerun."""
    push_history(chat)
    slides = chat["slides"]
    if op == "up":
        chat["slides"], focus = move_slide(slides, idx, -1), idx - 1
    elif op == "down":
        chat["slides"], focus = move_slide(slides, idx, 1), idx + 1
    elif op == "insert":
        chat["slides"], focus = insert_slide(slides, idx, image_mode)
    else:
        chat["slides"], focus = delete_slide(slides, idx), max(0, idx - 1)
    chat["focus_slide"] = focus
    mark_slides_changed(chat)
    st.rerun()


def render_outline_editor(chat: dict) -> None:
    """Interactive per-slide editor. Widget keys are scoped by chat id + generation counter."""
    kp = f"{chat['id']}_{chat['gen']}_"
    layout_options = [lt.value for lt in LayoutType]
    default_img_mode = st.session_state.settings.get("image_mode", IMAGE_MODE_UNSPLASH_FIRST)
    if default_img_mode not in IMAGE_MODE_OPTIONS:
        default_img_mode = IMAGE_MODE_UNSPLASH_FIRST

    updated_slides = []
    pending_op = None
    focus = chat.get("focus_slide", 0)
    total = len(chat["slides"])
    theme, builtin = _preview_theme()
    figures = assign_report_images(chat["slides"], len(chat.get("report_images") or []), default_img_mode)
    st.html(PREVIEW_CSS)  # style-only: Streamlit puts it in the event container, it takes no space

    for idx, slide in enumerate(chat["slides"]):
        current_layout = slide.get("layout_type", LayoutType.BULLET_CARDS.value)
        expander_label = f"Slide {idx + 1} [{current_layout.upper()}] — {slide.get('action_title', '')}"

        c_exp, c_ops = st.columns([13, 3], vertical_alignment="top")
        op = _render_slide_ops(c_ops, chat["slides"], idx, kp)
        if op:
            pending_op = (op, idx)

        # Overflow estimate from the last saved values marks the slide in the collapsed list too
        _, old_warnings = render_slide_preview(slide, theme, idx, total, builtin)
        exp = c_exp.expander(
            expander_label,
            expanded=(idx == focus),
            icon=":material/warning:" if old_warnings else None,
            key=f"{kp}exp_{idx}",
            on_change="rerun",  # tracks .open, so closed slides skip building their preview
        )
        with exp:
            preview_slot = st.container(width=560)
            top_c1, top_c2 = st.columns([1, 3])
            with top_c1:
                new_layout = st.selectbox(
                    "Slide layout type",
                    options=layout_options,
                    index=layout_options.index(current_layout) if current_layout in layout_options else 2,
                    key=f"{kp}layout_{idx}",
                )
            with top_c2:
                new_title = st.text_input(
                    "Action title (conclusion first)",
                    value=slide.get("action_title", ""),
                    key=f"{kp}title_{idx}",
                )

            slide["layout_type"] = new_layout
            slide["action_title"] = new_title

            # Dynamic Editor Fields Based on Selected Layout
            if new_layout in (LayoutType.COVER.value, LayoutType.SECTION_DIVIDER.value, LayoutType.CLOSING.value):
                sub_label = (
                    "Optional small line under the title (e.g. Presented by ...)"
                    if new_layout == LayoutType.CLOSING.value
                    else "Subtitle / section overview"
                )
                slide["subtitle_or_context"] = st.text_area(
                    sub_label,
                    value=slide.get("subtitle_or_context") or "",
                    height=80,
                    key=f"{kp}sub_{idx}",
                )

            elif new_layout in (LayoutType.BULLET_CARDS.value, LayoutType.THREE_CARDS.value):
                hint = "3 items" if new_layout == "three_cards" else "3-4 items"
                raw_pts = st.text_area(
                    f"Structured points ({hint}), one per line as `Heading :: Detail`",
                    value=format_points_for_editor(slide.get("points", [])),
                    height=120,
                    key=f"{kp}pts_{idx}",
                )
                slide["points"] = parse_points_from_editor(raw_pts)

            elif new_layout == LayoutType.TWO_COLUMN.value:
                lc1, lc2 = st.columns(2)
                left_col = slide.get("left_column") or {"column_title": "Left Column", "items": []}
                right_col = slide.get("right_column") or {"column_title": "Right Column", "items": []}

                with lc1:
                    l_title = st.text_input("Left column header", value=left_col.get("column_title", ""), key=f"{kp}l_title_{idx}")
                    l_pts = st.text_area(
                        "Left column points (`Heading :: Detail`)",
                        value=format_points_for_editor(left_col.get("items", [])),
                        height=110,
                        key=f"{kp}l_pts_{idx}",
                    )
                with lc2:
                    r_title = st.text_input("Right column header", value=right_col.get("column_title", ""), key=f"{kp}r_title_{idx}")
                    r_pts = st.text_area(
                        "Right column points (`Heading :: Detail`)",
                        value=format_points_for_editor(right_col.get("items", [])),
                        height=110,
                        key=f"{kp}r_pts_{idx}",
                    )
                slide["left_column"] = {"column_title": l_title, "items": parse_points_from_editor(l_pts)}
                slide["right_column"] = {"column_title": r_title, "items": parse_points_from_editor(r_pts)}

            elif new_layout == LayoutType.TIMELINE_FLOW.value:
                raw_steps = st.text_area(
                    "Timeline steps: 2 to 5 lines formatted as `Stage | Title | Description` "
                    "(e.g. `Q1 2026 | Pilot Rollout | Launch in two plants and measure downtime`)",
                    value=format_steps_for_editor(slide.get("timeline_steps", [])),
                    height=140,
                    key=f"{kp}steps_{idx}",
                )
                slide["timeline_steps"] = normalize_timeline_steps(parse_steps_from_editor(raw_steps))

            elif new_layout == LayoutType.COMPARISON_TABLE.value:
                raw_table = st.text_area(
                    "Comparison table: first line = column headers, then one row per line; separate cells with `|` "
                    "(2-5 columns, up to 6 rows)",
                    value=format_table_for_editor(slide.get("table_headers", []), slide.get("table_rows", [])),
                    height=160,
                    key=f"{kp}table_{idx}",
                )
                headers, rows = normalize_table(*parse_table_from_editor(raw_table))
                if not rows:
                    st.caption(":material/warning: Needs a header line with 2+ columns and at least one row; renders as cards until then.")
                slide["table_headers"], slide["table_rows"] = headers, rows

            elif new_layout == LayoutType.KPI_METRICS.value:
                raw_kpis = st.text_area(
                    "Hero KPI metrics, one per line as `Value | Label | Context` (e.g. `+62% | Latency drop | Across Q3 peak loads`)",
                    value=format_kpis_for_editor(slide.get("kpis", [])),
                    height=110,
                    key=f"{kp}kpis_{idx}",
                )
                slide["kpis"] = parse_kpis_from_editor(raw_kpis)

            elif new_layout == LayoutType.CHART_SLIDE.value:
                chart_obj = slide.get("chart") or {}
                cc1, cc2 = st.columns(2)
                with cc1:
                    c_types = [ct.value for ct in ChartType]
                    cur_ct = chart_obj.get("chart_type", "bar")
                    new_ct = st.segmented_control(
                        "Chart type",
                        options=c_types,
                        default=cur_ct if cur_ct in c_types else c_types[0],
                        required=True,
                        format_func={"bar": ":material/bar_chart: Bar", "line": ":material/show_chart: Line",
                                     "pie": ":material/pie_chart: Pie"}.get,
                        key=f"{kp}ctype_{idx}",
                    ) or c_types[0]
                    new_ctitle = st.text_input("Chart title", value=chart_obj.get("chart_title", ""), key=f"{kp}ctitle_{idx}")
                    new_sname = st.text_input("Series name & unit", value=chart_obj.get("series_name", "Value"), key=f"{kp}csname_{idx}")
                with cc2:
                    cats_str = st.text_input(
                        "Categories (comma-separated)",
                        value=", ".join(chart_obj.get("categories", ["Q1", "Q2", "Q3"])),
                        key=f"{kp}ccats_{idx}",
                    )
                    vals_str = st.text_input(
                        "Values (comma-separated numbers)",
                        value=", ".join(str(v) for v in chart_obj.get("values", [10.0, 20.0, 30.0])),
                        key=f"{kp}cvals_{idx}",
                    )

                parsed_cats = [c.strip() for c in cats_str.split(",") if c.strip()]
                parsed_vals = []
                for v in vals_str.split(","):
                    try:
                        parsed_vals.append(float(v.strip()))
                    except ValueError:
                        pass
                if not parsed_cats or not parsed_vals:
                    parsed_cats, parsed_vals = ["Item 1", "Item 2"], [50.0, 80.0]

                slide["chart"] = {
                    "chart_type": new_ct,
                    "chart_title": new_ctitle,
                    "series_name": new_sname,
                    "categories": parsed_cats,
                    "values": parsed_vals,
                }

                raw_c_pts = st.text_area(
                    "Right-side chart commentary (`Heading :: Detail`)",
                    value=format_points_for_editor(slide.get("points", [])),
                    height=100,
                    key=f"{kp}cpts_{idx}",
                )
                slide["points"] = parse_points_from_editor(raw_c_pts)

            elif new_layout == LayoutType.IMAGE_SPLIT.value:
                img_obj = slide.get("image") or {}
                ic1, ic2 = st.columns(2)

                with ic1:
                    cur_src = slide.get("image_source_mode") or default_img_mode
                    if cur_src not in IMAGE_MODE_OPTIONS:
                        cur_src = IMAGE_MODE_UNSPLASH_FIRST

                    chosen_src = st.segmented_control(
                        "Image mode",
                        options=IMAGE_MODE_OPTIONS,
                        default=cur_src,
                        required=True,
                        format_func=IMAGE_MODE_LABELS.get,
                        key=f"{kp}imgsrc_{idx}",
                        help=IMAGE_MODE_HELP,
                    ) or cur_src
                    slide["image_source_mode"] = chosen_src
                    if chosen_src == IMAGE_MODE_REPORT_FIRST:
                        _render_report_image_picker(chat, slide, idx, kp, default_img_mode)

                    place_opts = ["left", "right"]
                    cur_place = img_obj.get("placement", "left")
                    new_place = st.segmented_control(
                        "Image placement on slide",
                        options=place_opts,
                        default=cur_place if cur_place in place_opts else "left",
                        required=True,
                        format_func={"left": ":material/align_horizontal_left: Left",
                                     "right": ":material/align_horizontal_right: Right"}.get,
                        key=f"{kp}imgplace_{idx}",
                    ) or "left"

                    new_kw = st.text_input(
                        "Unsplash search keywords (English)",
                        value=img_obj.get("search_keywords", ""),
                        key=f"{kp}ukw_{idx}",
                    )
                    new_vprompt = st.text_area(
                        "AI fallback prompt (English)",
                        value=img_obj.get("visual_prompt", ""),
                        height=75,
                        key=f"{kp}vprompt_{idx}",
                    )

                with ic2:
                    new_caption = st.text_input(
                        "Image caption",
                        value=img_obj.get("caption", ""),
                        key=f"{kp}imgcap_{idx}",
                    )
                    raw_i_pts = st.text_area(
                        "Accompanying key points (`Heading :: Detail`)",
                        value=format_points_for_editor(slide.get("points", [])),
                        height=175,
                        key=f"{kp}ipts_{idx}",
                    )
                    slide["points"] = parse_points_from_editor(raw_i_pts)

                slide["image"] = {
                    "preferred_style": VisualStyle.STOCK_PHOTO.value,
                    "visual_prompt": new_vprompt,
                    "search_keywords": new_kw,
                    "caption": new_caption,
                    "placement": new_place,
                }

            # Bottom Takeaway Editor for all content layouts
            if new_layout not in (LayoutType.COVER.value, LayoutType.SECTION_DIVIDER.value, LayoutType.CLOSING.value):
                slide["bottom_takeaway"] = st.text_input(
                    "Bottom takeaway banner (the 'so what?')",
                    value=slide.get("bottom_takeaway") or "",
                    key=f"{kp}takeaway_{idx}",
                )

            if exp.open:
                # Drawn after every field was read, so the preview shows what was just typed
                html_preview, warnings = render_slide_preview(
                    slide, theme, idx, total, builtin, _image_label(slide, idx, figures, default_img_mode)
                )
                preview_slot.html(html_preview)
                for warning in warnings:
                    preview_slot.caption(f":material/warning: {warning}")

            updated_slides.append(slide)

    chat["slides"] = updated_slides
    # Applied after every field was read, so text typed right before the click is kept
    if pending_op:
        _apply_slide_op(chat, *pending_op, default_img_mode)