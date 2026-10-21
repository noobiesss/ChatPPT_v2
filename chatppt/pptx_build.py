"""build_presentation_pptx(): assembles the whole deck from outline dicts.

Edit when: changing deck-level flow (template vs built-in theme, per-layout dispatch, footers/page numbers).
"""

import io
from typing import List, Dict, Optional

from pptx import Presentation
from pptx.util import Inches

from chatppt.images import assign_report_images, resolve_slide_image_bytes
from chatppt.pptx_render import apply_slide_fonts, clear_existing_slides, draw_header_and_takeaway, render_bullet_cards_slide, render_chart_slide, render_closing_slide, render_comparison_table_slide, render_cover_slide, render_image_split_slide, render_kpi_metrics_slide, render_section_divider_slide, render_three_cards_slide, render_timeline_flow_slide, render_two_column_slide
from chatppt.pptx_template import _capture_template_prototypes, _create_slide_with_template, _scale_template_to_widescreen, _strip_template_decorations, extract_template_style
from chatppt.presets import BUILTIN_THEMES
from chatppt.schemas import IMAGE_MODE_UNSPLASH_FIRST, LayoutType


def build_presentation_pptx(
    slides_data: List[dict],
    theme_name: str,
    gemini_api_key: Optional[str],
    use_imagen: bool,
    unsplash_access_key: str,
    extracted_report_images: List[Dict[str, bytes]],
    custom_template_bytes: Optional[bytes] = None,
    use_template_colors: bool = True,
    default_image_mode: str = IMAGE_MODE_UNSPLASH_FIRST,
) -> bytes:
    """
    Assemble the complete .pptx presentation across all 11 visual layouts.
    Supports both Built-in Color Themes and Custom Uploaded .pptx Templates (inheriting
    Slide Masters, role-specific layouts, slide backgrounds, theme colors, and fonts).
    Sample-slide decorations (illustrations, frames, widgets, logos) are intentionally NOT inherited.
    """
    prototypes: Dict[str, dict] = {}
    is_custom_template = bool(custom_template_bytes)

    if is_custom_template:
        # Order matters: rescale first so captured geometry is 16:9, read the style and snapshot the
        # sample slides' backgrounds before decorations and sample slides are removed
        prs = Presentation(io.BytesIO(custom_template_bytes))
        _scale_template_to_widescreen(prs, 13.333, 7.5)
        theme = extract_template_style(
            prs,
            fallback_theme_name=theme_name,
            use_template_colors=use_template_colors,
        )
        prototypes = _capture_template_prototypes(prs)
        _strip_template_decorations(prs)
        clear_existing_slides(prs)
        apply_solid_bg = False  # Preserve custom template's master, layout, and prototype backgrounds
    else:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        theme = dict(BUILTIN_THEMES.get(theme_name, BUILTIN_THEMES["Corporate Blue"]))
        apply_solid_bg = True

    total_slides = len(slides_data)
    report_images = extracted_report_images or []
    # Report figures for image_split slides in "Report Images First" mode (explicit picks + auto-fill)
    report_assignment = assign_report_images(slides_data, len(report_images), default_image_mode)

    for idx, slide_data in enumerate(slides_data):
        layout = slide_data.get("layout_type", LayoutType.BULLET_CARDS.value)
        if layout in (LayoutType.COVER.value, LayoutType.CLOSING.value):
            role = "cover"  # the closing slide reuses the template's cover background
        elif layout == LayoutType.SECTION_DIVIDER.value:
            role = "section"
        else:
            role = "content"

        slide = _create_slide_with_template(prs, role, prototypes, is_custom_template)

        if layout == LayoutType.COVER.value:
            render_cover_slide(slide, slide_data, theme, apply_bg=apply_solid_bg)

        elif layout == LayoutType.SECTION_DIVIDER.value:
            render_section_divider_slide(slide, slide_data, theme, apply_bg=apply_solid_bg)

        elif layout == LayoutType.CLOSING.value:
            render_closing_slide(slide, slide_data, theme, apply_bg=apply_solid_bg)

        else:
            draw_header_and_takeaway(slide, slide_data, idx, total_slides, theme, apply_bg=apply_solid_bg)

            if layout == LayoutType.THREE_CARDS.value:
                render_three_cards_slide(slide, slide_data, theme)

            elif layout == LayoutType.TWO_COLUMN.value:
                render_two_column_slide(slide, slide_data, theme)

            elif layout == LayoutType.KPI_METRICS.value:
                render_kpi_metrics_slide(slide, slide_data, theme)

            elif layout == LayoutType.CHART_SLIDE.value:
                render_chart_slide(slide, slide_data, theme)

            elif layout == LayoutType.TIMELINE_FLOW.value:
                render_timeline_flow_slide(slide, slide_data, theme)

            elif layout == LayoutType.COMPARISON_TABLE.value:
                render_comparison_table_slide(slide, slide_data, theme)

            elif layout == LayoutType.IMAGE_SPLIT.value:
                img_bytes = resolve_slide_image_bytes(
                    slide_data={**slide_data, "report_image_index": report_assignment.get(idx)},
                    gemini_api_key=gemini_api_key,
                    use_imagen=use_imagen,
                    unsplash_access_key=unsplash_access_key,
                    extracted_report_images=report_images,
                    default_image_mode=default_image_mode,
                )
                render_image_split_slide(slide, slide_data, theme, img_bytes=img_bytes)

            else:
                render_bullet_cards_slide(slide, slide_data, theme)

        apply_slide_fonts(slide, theme)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output.getvalue()