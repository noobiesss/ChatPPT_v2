"""Per-slide outline editor (expanders, layout-specific fields).

Edit when: changing what can be edited on each slide before rendering.
"""

import streamlit as st

from chatppt.schemas import ChartType, IMAGE_MODE_OPTIONS, IMAGE_MODE_UNSPLASH_FIRST, LayoutType, VisualStyle
from chatppt.ui.editor_text import format_kpis_for_editor, format_points_for_editor, parse_kpis_from_editor, parse_points_from_editor


def render_outline_editor(chat: dict) -> None:
    """Interactive per-slide editor. Widget keys are scoped by chat id + generation counter."""
    kp = f"{chat['id']}_{chat['gen']}_"
    layout_options = [lt.value for lt in LayoutType]
    default_img_mode = st.session_state.settings.get("image_mode", IMAGE_MODE_UNSPLASH_FIRST)
    if default_img_mode not in IMAGE_MODE_OPTIONS:
        default_img_mode = IMAGE_MODE_UNSPLASH_FIRST

    updated_slides = []

    for idx, slide in enumerate(chat["slides"]):
        current_layout = slide.get("layout_type", LayoutType.BULLET_CARDS.value)
        expander_label = f"Slide {idx + 1} [{current_layout.upper()}] — {slide.get('action_title', '')}"

        with st.expander(expander_label, expanded=(idx < 1)):
            top_c1, top_c2 = st.columns([1, 3])
            with top_c1:
                new_layout = st.selectbox(
                    "Slide Layout Type",
                    options=layout_options,
                    index=layout_options.index(current_layout) if current_layout in layout_options else 2,
                    key=f"{kp}layout_{idx}",
                )
            with top_c2:
                new_title = st.text_input(
                    "Action Title (Conclusion-First)",
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
                    else "Subtitle / Section Overview"
                )
                slide["subtitle_or_context"] = st.text_area(
                    sub_label,
                    value=slide.get("subtitle_or_context") or "",
                    height=80,
                    key=f"{kp}sub_{idx}",
                )

            elif new_layout in (LayoutType.BULLET_CARDS.value, LayoutType.THREE_CARDS.value):
                hint = "Enter 3 items (Heading :: Detail)" if new_layout == "three_cards" else "Enter 3-4 items (Heading :: Detail)"
                raw_pts = st.text_area(
                    f"Structured Points — One per line formatted as `{hint}`",
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
                    l_title = st.text_input("Left Column Header", value=left_col.get("column_title", ""), key=f"{kp}l_title_{idx}")
                    l_pts = st.text_area(
                        "Left Column Points (`Heading :: Detail`)",
                        value=format_points_for_editor(left_col.get("items", [])),
                        height=110,
                        key=f"{kp}l_pts_{idx}",
                    )
                with lc2:
                    r_title = st.text_input("Right Column Header", value=right_col.get("column_title", ""), key=f"{kp}r_title_{idx}")
                    r_pts = st.text_area(
                        "Right Column Points (`Heading :: Detail`)",
                        value=format_points_for_editor(right_col.get("items", [])),
                        height=110,
                        key=f"{kp}r_pts_{idx}",
                    )
                slide["left_column"] = {"column_title": l_title, "items": parse_points_from_editor(l_pts)}
                slide["right_column"] = {"column_title": r_title, "items": parse_points_from_editor(r_pts)}

            elif new_layout == LayoutType.KPI_METRICS.value:
                raw_kpis = st.text_area(
                    "Hero KPI Metrics — One per line formatted as `Value | Label | Context` (e.g. `+62% | Latency Drop | Across Q3 peak loads`)",
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
                    new_ct = st.selectbox(
                        "Chart Type",
                        options=c_types,
                        index=c_types.index(cur_ct) if cur_ct in c_types else 0,
                        key=f"{kp}ctype_{idx}",
                    )
                    new_ctitle = st.text_input("Chart Title", value=chart_obj.get("chart_title", ""), key=f"{kp}ctitle_{idx}")
                    new_sname = st.text_input("Series Name & Unit", value=chart_obj.get("series_name", "Value"), key=f"{kp}csname_{idx}")
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
                    "Right-Side Chart Commentary (`Heading :: Detail`)",
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

                    chosen_src = st.radio(
                        "Image Mode",
                        options=IMAGE_MODE_OPTIONS,
                        index=IMAGE_MODE_OPTIONS.index(cur_src),
                        key=f"{kp}imgsrc_{idx}",
                        help="Option 1 searches Unsplash first and falls back to AI if unavailable. Option 2 reserves a photo placeholder box.",
                    )
                    slide["image_source_mode"] = chosen_src

                    place_opts = ["left", "right"]
                    cur_place = img_obj.get("placement", "left")
                    new_place = st.selectbox(
                        "Image Placement on Slide",
                        options=place_opts,
                        index=place_opts.index(cur_place) if cur_place in place_opts else 0,
                        key=f"{kp}imgplace_{idx}",
                    )

                    new_kw = st.text_input(
                        "Unsplash Search Keywords (English)",
                        value=img_obj.get("search_keywords", ""),
                        key=f"{kp}ukw_{idx}",
                    )
                    new_vprompt = st.text_area(
                        "AI Fallback Prompt (English)",
                        value=img_obj.get("visual_prompt", ""),
                        height=75,
                        key=f"{kp}vprompt_{idx}",
                    )

                with ic2:
                    new_caption = st.text_input(
                        "Image Caption",
                        value=img_obj.get("caption", ""),
                        key=f"{kp}imgcap_{idx}",
                    )
                    raw_i_pts = st.text_area(
                        "Accompanying Key Points (`Heading :: Detail`)",
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
                    "💡 Bottom Takeaway Banner ('So What?')",
                    value=slide.get("bottom_takeaway") or "",
                    key=f"{kp}takeaway_{idx}",
                )

            updated_slides.append(slide)

    chat["slides"] = updated_slides