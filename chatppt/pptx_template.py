"""Custom .pptx template inheritance: theme/colors/fonts, backgrounds, decoration stripping, layout prototypes.

Edit when: the uploaded template's look is not reproduced correctly (borders, fonts, colors, backgrounds).
"""

import copy
import io
import xml.etree.ElementTree as ET
from typing import List, Dict, Optional, Tuple

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches
import streamlit as st

from chatppt.presets import BUILTIN_THEMES


_NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


_NS_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


_NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _hex_to_rgb_tuple(hex_str: str) -> Optional[Tuple[int, int, int]]:
    """Parse a 6-character hex string like '1A2B3C' into an (R, G, B) tuple."""
    cleaned = (hex_str or "").strip().lstrip("#")
    if len(cleaned) != 6:
        return None
    try:
        return int(cleaned[0:2], 16), int(cleaned[2:4], 16), int(cleaned[4:6], 16)
    except ValueError:
        return None


def _rgb_to_hex(color: RGBColor) -> str:
    return f"#{color[0]:02X}{color[1]:02X}{color[2]:02X}"


def _luminance(rgb: Tuple[int, int, int]) -> float:
    """Return perceived luminance (0..255) for contrast decisions."""
    return 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]


def _blend_rgb(c1: Tuple[int, int, int], c2: Tuple[int, int, int], ratio: float) -> RGBColor:
    """Linearly interpolate from c1 toward c2 by ratio (0.0 to 1.0)."""
    r = int(round(c1[0] * (1.0 - ratio) + c2[0] * ratio))
    g = int(round(c1[1] * (1.0 - ratio) + c2[1] * ratio))
    b = int(round(c1[2] * (1.0 - ratio) + c2[2] * ratio))
    return RGBColor(max(0, min(255, r)), max(0, min(255, g)), max(0, min(255, b)))


def _relative_luminance(rgb: Tuple[int, int, int]) -> float:
    """WCAG relative luminance (0..1)."""
    def _lin(v: int) -> float:
        c = v / 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * _lin(rgb[0]) + 0.7152 * _lin(rgb[1]) + 0.0722 * _lin(rgb[2])


def _ensure_white_text_contrast(color: RGBColor, min_ratio: float = 4.5) -> RGBColor:
    """
    Darken `color` step by step (toward black, keeping its hue) until white text on it reaches
    the requested WCAG contrast ratio. Used for the Key Takeaway banner and column headers.
    """
    rgb = (color[0], color[1], color[2])
    for _ in range(30):
        ratio = 1.05 / (_relative_luminance(rgb) + 0.05)
        if ratio >= min_ratio:
            break
        rgb = (int(rgb[0] * 0.92), int(rgb[1] * 0.92), int(rgb[2] * 0.92))
    return RGBColor(*rgb)


def _parse_theme_xml_colors_and_fonts(prs: Presentation) -> Tuple[Dict[str, Tuple[int, int, int]], Optional[str], Optional[str]]:
    """
    Inspect the Slide Master's Office Theme XML (theme1.xml) to extract scheme colors
    (dk1, lt1, dk2, lt2, accent1..accent6) and major/minor Latin font typefaces.
    """
    colors: Dict[str, Tuple[int, int, int]] = {}
    font_heading: Optional[str] = None
    font_body: Optional[str] = None

    try:
        if not prs.slide_masters:
            return colors, font_heading, font_body
        master_part = prs.slide_masters[0].part
        theme_blob = None
        for rel in master_part.rels.values():
            if "theme" in (rel.reltype or "").lower() and hasattr(rel, "target_part"):
                theme_blob = rel.target_part.blob
                break
        if not theme_blob:
            return colors, font_heading, font_body

        root = ET.fromstring(theme_blob)
        clr_scheme = root.find(f".//{_NS_A}clrScheme")
        if clr_scheme is not None:
            for child in clr_scheme:
                tag_name = child.tag.split("}")[-1]
                srgb = child.find(f"{_NS_A}srgbClr")
                sys_clr = child.find(f"{_NS_A}sysClr")
                hex_val = None
                if srgb is not None:
                    hex_val = srgb.get("val")
                elif sys_clr is not None:
                    hex_val = sys_clr.get("lastClr") or sys_clr.get("val")
                rgb_tuple = _hex_to_rgb_tuple(hex_val or "")
                if rgb_tuple is not None:
                    colors[tag_name] = rgb_tuple

        font_scheme = root.find(f".//{_NS_A}fontScheme")
        if font_scheme is not None:
            major_latin = font_scheme.find(f"{_NS_A}majorFont/{_NS_A}latin")
            minor_latin = font_scheme.find(f"{_NS_A}minorFont/{_NS_A}latin")
            if major_latin is not None:
                tf = (major_latin.get("typeface") or "").strip()
                if tf and not tf.startswith("+"):
                    font_heading = tf
            if minor_latin is not None:
                tf = (minor_latin.get("typeface") or "").strip()
                if tf and not tf.startswith("+"):
                    font_body = tf
    except Exception:
        pass

    return colors, font_heading, font_body


def _extract_solid_fill_rgb_from_el(el) -> Optional[Tuple[int, int, int]]:
    """Extract the first explicit srgbClr inside a solidFill child of an XML element."""
    if el is None:
        return None
    try:
        for node in el.iter(f"{_NS_A}solidFill"):
            for sub in node:
                if sub.tag.endswith("srgbClr"):
                    rgb = _hex_to_rgb_tuple(sub.get("val", ""))
                    if rgb is not None:
                        return rgb
    except Exception:
        pass
    return None


def _sample_slide_colors_and_fonts(prs: Presentation) -> Dict[str, object]:
    """
    Inspect existing slides and layouts for explicit background colors, vibrant accent shape fills,
    and explicitly assigned font names (vital for Canva / Slidesgo templates that bypass theme1.xml).
    """
    info: Dict[str, object] = {
        "cover_bg": None,
        "content_bg": None,
        "accent": None,
        "font": None,
        "font_heading": None,
        "font_body": None,
    }
    try:
        slide_w = float(prs.slide_width or Inches(13.333))
        slide_h = float(prs.slide_height or Inches(7.5))
        slide_area = max(slide_w * slide_h, 1.0)

        n_slides = len(prs.slides)
        font_runs_slides: List[Tuple[str, float]] = []
        font_runs_other: List[Tuple[str, float]] = []
        containers = list(prs.slides) + list(prs.slide_layouts) + list(prs.slide_masters)
        for s_idx, container in enumerate(containers):
            c_sld = container._element.find(f"{_NS_P}cSld")
            bg_el = c_sld.find(f"{_NS_P}bg") if c_sld is not None else None
            bg_rgb = _extract_solid_fill_rgb_from_el(bg_el)
            if bg_rgb is not None:
                if s_idx == 0 and info["cover_bg"] is None:
                    info["cover_bg"] = bg_rgb
                elif info["content_bg"] is None:
                    info["content_bg"] = bg_rgb

            for shape in getattr(container, "shapes", []):
                try:
                    w = float(shape.width or 0)
                    h = float(shape.height or 0)
                    area_ratio = (w * h) / slide_area
                    sp_pr = shape._element.find(f"{_NS_P}spPr")
                    fill_rgb = _extract_solid_fill_rgb_from_el(sp_pr)
                    if fill_rgb is not None:
                        # Full-bleed rectangle acts as a slide background
                        if area_ratio >= 0.75 and not getattr(shape, "is_placeholder", False):
                            if s_idx == 0 and info["cover_bg"] is None:
                                info["cover_bg"] = fill_rgb
                            elif s_idx > 0 and info["content_bg"] is None:
                                info["content_bg"] = fill_rgb
                        else:
                            # Look for a saturated non-grey accent color
                            spread = max(fill_rgb) - min(fill_rgb)
                            lum = _luminance(fill_rgb)
                            if spread >= 35 and 25 < lum < 235 and info["accent"] is None:
                                info["accent"] = fill_rgb

                    if getattr(shape, "has_text_frame", False):
                        for para in shape.text_frame.paragraphs:
                            for run in para.runs:
                                fname = (run.font.name or "").strip()
                                if not fname or fname.startswith("+") or not (run.text or "").strip():
                                    continue
                                size_pt = float(run.font.size.pt) if run.font.size else 0.0
                                bucket = font_runs_slides if s_idx < n_slides else font_runs_other
                                bucket.append((fname, size_pt))
                except Exception:
                    continue

        # Fonts that are really used on the sample slides win over theme1.xml: Canva / Slidesgo
        # templates often leave theme fonts at defaults and set the visible fonts on each run.
        from collections import Counter
        runs = font_runs_slides or font_runs_other
        if runs:
            sized = [r for r in runs if r[1] > 0]
            heading = None
            if sized:
                biggest = max(sized, key=lambda r: r[1])
                if biggest[1] >= 18:
                    heading = biggest[0]
            small = [name for name, size in runs if size < 18]
            body = Counter(small).most_common(1)[0][0] if small else None
            overall = Counter(name for name, _ in runs).most_common(1)[0][0]
            info["font_heading"] = heading or body or overall
            info["font_body"] = body or heading or overall
            info["font"] = info["font_heading"]
    except Exception:
        pass
    return info


def extract_template_style(
    prs: Presentation,
    fallback_theme_name: str = "Corporate Blue",
    use_template_colors: bool = True,
) -> Dict[str, object]:
    """
    Build a complete theme dictionary (colors + font_heading + font_body) derived from an uploaded
    .pptx template, falling back gracefully to `fallback_theme_name` for any missing tokens.
    """
    base_theme: Dict[str, object] = dict(BUILTIN_THEMES.get(fallback_theme_name, BUILTIN_THEMES["Corporate Blue"]))
    scheme_colors, font_heading, font_body = _parse_theme_xml_colors_and_fonts(prs)
    sampled = _sample_slide_colors_and_fonts(prs)

    resolved_heading = sampled.get("font_heading") or font_heading or sampled.get("font")
    resolved_body = sampled.get("font_body") or font_body or resolved_heading
    base_theme["font_heading"] = resolved_heading
    base_theme["font_body"] = resolved_body

    if not use_template_colors:
        return base_theme

    # Determine primary accent color: prefer sampled decorative shape color or non-default scheme accent1
    OFFICE_DEFAULT_ACCENT1 = (68, 114, 196)  # #4472C4 default Office blue
    scheme_acc1 = scheme_colors.get("accent1")
    sampled_acc = sampled.get("accent")
    if sampled_acc and (scheme_acc1 is None or scheme_acc1 == OFFICE_DEFAULT_ACCENT1):
        accent_tuple = sampled_acc
    else:
        accent_tuple = scheme_acc1 or sampled_acc or scheme_colors.get("accent2")

    # Determine content background color (light vs dark deck)
    content_bg_tuple = sampled.get("content_bg")
    if content_bg_tuple is None:
        lt1 = scheme_colors.get("lt1", (255, 255, 255))
        content_bg_tuple = lt1

    cover_bg_tuple = sampled.get("cover_bg") or scheme_colors.get("dk2") or scheme_colors.get("dk1") or content_bg_tuple

    if accent_tuple is None and sampled.get("content_bg") is None and not scheme_colors:
        return base_theme

    if accent_tuple is None:
        acc_c = base_theme["accent"]
        accent_tuple = (acc_c[0], acc_c[1], acc_c[2])

    content_is_dark = _luminance(content_bg_tuple) < 128
    cover_is_dark = _luminance(cover_bg_tuple) < 128

    if content_is_dark:
        card_bg = _blend_rgb(content_bg_tuple, (255, 255, 255), 0.08)
        card_border = _blend_rgb(content_bg_tuple, accent_tuple, 0.35)
        title_content = RGBColor(245, 247, 252)
        body_content = RGBColor(200, 206, 220)
        takeaway_bg = _ensure_white_text_contrast(_blend_rgb(content_bg_tuple, accent_tuple, 0.28))
        takeaway_text = RGBColor(255, 255, 255)
    else:
        dk_text = scheme_colors.get("dk1") or (22, 33, 48)
        if _luminance(dk_text) > 140:
            dk_text = (22, 33, 48)
        card_bg = RGBColor(255, 255, 255)
        card_border = _blend_rgb((255, 255, 255), accent_tuple, 0.25)
        title_content = RGBColor(*dk_text)
        body_content = _blend_rgb(dk_text, (120, 128, 140), 0.35)
        # Derive the banner from the template accent (not from the near-black body text color) so it
        # follows the template style; darken only as far as needed for readable white text.
        takeaway_bg = _ensure_white_text_contrast(_blend_rgb(accent_tuple, (15, 23, 42), 0.30))
        takeaway_text = RGBColor(255, 255, 255)

    if cover_is_dark:
        title_cover = RGBColor(255, 255, 255)
        subtitle_cover = _blend_rgb((255, 255, 255), accent_tuple, 0.22)
    else:
        title_cover = title_content
        subtitle_cover = body_content

    base_theme.update({
        "bg_cover": RGBColor(*cover_bg_tuple),
        "bg_section": _blend_rgb(cover_bg_tuple, accent_tuple, 0.18),
        "bg_content": RGBColor(*content_bg_tuple),
        "card_bg": card_bg,
        "card_border": card_border,
        "accent": RGBColor(*accent_tuple),
        "title_cover": title_cover,
        "subtitle_cover": subtitle_cover,
        "title_content": title_content,
        "body_content": body_content,
        "takeaway_bg": takeaway_bg,
        "takeaway_text": takeaway_text,
    })
    return base_theme


@st.cache_data(show_spinner=False, max_entries=4)
def template_theme_for_preview(template_bytes: bytes, fallback_theme_name: str, use_template_colors: bool) -> Dict[str, object]:
    """Theme dict (colors + fonts) of an uploaded template, cached for the per-rerun wireframe previews."""
    try:
        prs = Presentation(io.BytesIO(template_bytes))
        return extract_template_style(prs, fallback_theme_name=fallback_theme_name, use_template_colors=use_template_colors)
    except Exception:  # noqa: BLE001 - a broken template previews with the fallback palette
        return dict(BUILTIN_THEMES.get(fallback_theme_name, BUILTIN_THEMES["Corporate Blue"]))


@st.cache_data(show_spinner=False)
def inspect_template_summary(template_bytes: bytes) -> Dict[str, str]:
    """Return a concise summary of an uploaded .pptx template (slide count, fonts, accent color)."""
    try:
        prs = Presentation(io.BytesIO(template_bytes))
        style = extract_template_style(prs, use_template_colors=True)
        font_h = style.get("font_heading") or "Default"
        accent_hex = _rgb_to_hex(style["accent"])
        bg_hex = _rgb_to_hex(style["bg_content"])
        return {
            "slides": str(len(prs.slides)),
            "layouts": str(len(prs.slide_layouts)),
            "font": str(font_h),
            "accent": accent_hex,
            "bg": bg_hex,
        }
    except Exception:
        return {"slides": "?", "layouts": "?", "font": "Default", "accent": "#0072CE", "bg": "#F4F7FB"}


def _scale_xml_xfrm(root_el, scale_x: float, scale_y: float) -> None:
    """Scale DrawingML <a:off> and <a:ext> coordinates inside an XML subtree."""
    for xfrm in root_el.iter(f"{_NS_A}xfrm"):
        # Only scale top-level shape transforms, not child coordinate spaces inside groups
        off = xfrm.find(f"{_NS_A}off")
        ext = xfrm.find(f"{_NS_A}ext")
        ch_off = xfrm.find(f"{_NS_A}chOff")
        if ch_off is not None:
            continue
        if off is not None:
            try:
                x = int(off.get("x", "0"))
                y = int(off.get("y", "0"))
                off.set("x", str(int(round(x * scale_x))))
                off.set("y", str(int(round(y * scale_y))))
            except ValueError:
                pass
        if ext is not None:
            try:
                cx = int(ext.get("cx", "0"))
                cy = int(ext.get("cy", "0"))
                ext.set("cx", str(int(round(cx * scale_x))))
                ext.set("cy", str(int(round(cy * scale_y))))
            except ValueError:
                pass


def _scale_template_to_widescreen(prs: Presentation, target_w_in: float = 13.333, target_h_in: float = 7.5) -> None:
    """
    Normalize a presentation's canvas to 16:9 (13.333 x 7.5 in) and proportionally scale shapes
    on Slide Masters, Slide Layouts, and existing Slides so 10x5.625 (Canva/Google Slides) or 4:3
    templates fill the entire widescreen slide without leaving blank margins.
    """
    target_w = int(Inches(target_w_in))
    target_h = int(Inches(target_h_in))
    orig_w = int(prs.slide_width or target_w)
    orig_h = int(prs.slide_height or target_h)

    if abs(orig_w - target_w) > 1000 or abs(orig_h - target_h) > 1000:
        scale_x = target_w / float(orig_w)
        scale_y = target_h / float(orig_h)
        for master in prs.slide_masters:
            _scale_xml_xfrm(master._element, scale_x, scale_y)
        for layout in prs.slide_layouts:
            _scale_xml_xfrm(layout._element, scale_x, scale_y)
        for slide in prs.slides:
            _scale_xml_xfrm(slide._element, scale_x, scale_y)

    prs.slide_width = target_w
    prs.slide_height = target_h


def _has_visible_fill(shape) -> bool:
    """True for pictures and for shapes with an explicit solid/gradient/picture/pattern fill (not outline-only)."""
    try:
        if hasattr(shape, "image"):
            return True
    except Exception:
        pass
    try:
        sp_pr = shape._element.find(f"{_NS_P}spPr")
    except Exception:
        return False
    if sp_pr is None or sp_pr.find(f"{_NS_A}noFill") is not None:
        return False
    return any(sp_pr.find(f"{_NS_A}{tag}") is not None for tag in ("solidFill", "gradFill", "blipFill", "pattFill"))


def _is_full_bleed_background_shape(shape, slide_w_emu: int, slide_h_emu: int) -> bool:
    """
    Return True only if a shape on a sample template slide acts as the slide BACKGROUND:
    a text-free picture or vector shape that covers (nearly) the entire slide canvas.

    Everything else on the sample slide (illustrations, border frames, header/footer widgets,
    logos, sample text) is deliberately ignored: those elements were positioned for the sample
    content and collide with generated content, so they are not inherited.
    """
    if getattr(shape, "is_placeholder", False):
        return False
    if getattr(shape, "has_table", False) or getattr(shape, "has_chart", False):
        return False
    if getattr(shape, "has_text_frame", False) and (shape.text_frame.text or "").strip():
        return False

    try:
        width = int(shape.width or 0)
        height = int(shape.height or 0)
    except Exception:
        return False
    if width <= 0 or height <= 0 or slide_w_emu <= 0 or slide_h_emu <= 0:
        return False

    if not _has_visible_fill(shape):
        # Outline-only rectangles / groups are border frames, not backgrounds.
        return False

    return width >= int(slide_w_emu * 0.92) and height >= int(slide_h_emu * 0.92)


def _strip_template_decorations(prs: Presentation) -> None:
    """
    Remove non-placeholder decorations (border frames, lines, logos, ornaments, text labels) from every
    slide layout and the slide master. Only placeholders and full-bleed filled/picture backgrounds stay,
    so master/layout artwork cannot collide with the generated content (same rule as sample slides).
    """
    slide_w_emu = int(prs.slide_width or Inches(13.333))
    slide_h_emu = int(prs.slide_height or Inches(7.5))
    for container in list(prs.slide_layouts) + list(prs.slide_masters):
        for shape in list(container.shapes):
            try:
                if getattr(shape, "is_placeholder", False):
                    continue
                if _is_full_bleed_background_shape(shape, slide_w_emu, slide_h_emu):
                    continue
                el = shape._element
                parent = el.getparent()
                if parent is not None:
                    parent.remove(el)
            except Exception:
                continue


def _snapshot_slide_prototype(slide, slide_w_emu: int, slide_h_emu: int) -> dict:
    """
    Capture a prototype slide's layout reference, its custom <p:bg> XML, and any full-bleed
    background picture/shape before the sample slides are cleared.
    """
    c_sld = slide._element.find(f"{_NS_P}cSld")
    bg_el = c_sld.find(f"{_NS_P}bg") if c_sld is not None else None
    rels_map = {r_id: rel for r_id, rel in slide.part.rels.items()}

    decor_items = []
    for shape in slide.shapes:
        if not _is_full_bleed_background_shape(shape, slide_w_emu, slide_h_emu):
            continue
        try:
            if hasattr(shape, "image"):
                decor_items.append({
                    "kind": "picture",
                    "blob": shape.image.blob,
                })
            else:
                decor_items.append({
                    "kind": "xml",
                    "el": copy.deepcopy(shape._element),
                })
        except Exception:
            continue

    return {
        "layout": slide.slide_layout,
        "bg_el": copy.deepcopy(bg_el) if bg_el is not None else None,
        "rels": rels_map,
        "decor_items": decor_items,
    }


def _capture_template_prototypes(prs: Presentation) -> Dict[str, dict]:
    """
    Inspect existing slides in an uploaded .pptx template and capture role prototypes
    ('cover', 'section', 'content') so their backgrounds survive slide clearing.
    """
    prototypes: Dict[str, dict] = {}
    slides = list(prs.slides)
    if not slides:
        return prototypes

    slide_w_emu = int(prs.slide_width or Inches(13.333))
    slide_h_emu = int(prs.slide_height or Inches(7.5))
    prototypes["cover"] = _snapshot_slide_prototype(slides[0], slide_w_emu, slide_h_emu)

    if len(slides) == 1:
        prototypes["content"] = prototypes["cover"]
        prototypes["section"] = prototypes["cover"]
        return prototypes

    # Pick a representative content slide (prefer Slide 2 or a slide with multiple shapes)
    content_slide = slides[1]
    for s in slides[1:]:
        lname = (s.slide_layout.name or "").lower()
        if "section" in lname or "divider" in lname:
            prototypes.setdefault("section", _snapshot_slide_prototype(s, slide_w_emu, slide_h_emu))
        elif "content" in lname or "title" in lname:
            content_slide = s
            break

    prototypes["content"] = _snapshot_slide_prototype(content_slide, slide_w_emu, slide_h_emu)
    prototypes.setdefault("section", prototypes["cover"])
    return prototypes


def _remap_embed_rids(el, old_rels: dict, target_slide_part) -> None:
    """Remap any r:embed or r:link attributes in a cloned XML element onto target_slide_part."""
    for attr_name in (f"{_NS_R}embed", f"{_NS_R}link", f"{_NS_R}id"):
        for node in el.iter():
            old_rid = node.get(attr_name)
            if old_rid and old_rid in old_rels:
                rel = old_rels[old_rid]
                try:
                    if getattr(rel, "is_external", False):
                        new_rid = target_slide_part.rels.get_or_add_ext_rel(rel.reltype, rel._target)
                    else:
                        new_rid = target_slide_part.relate_to(rel.target_part, rel.reltype)
                    node.set(attr_name, new_rid)
                except Exception:
                    pass


def _select_layout_for_role(prs: Presentation, role: str, prototypes: Dict[str, dict]):
    """
    Choose the best SlideLayout for 'cover', 'section', or 'content'.
    When a prototype slide exists for that role, use its layout; otherwise match layout names
    instead of blindly picking the Blank layout (which often hides master graphics).
    """
    proto = prototypes.get(role)
    if proto and proto.get("layout") is not None:
        return proto["layout"]

    layouts = list(prs.slide_layouts)
    if not layouts:
        raise RuntimeError("Presentation has no slide layouts.")

    if role == "cover":
        for lyt in layouts:
            n = (lyt.name or "").lower()
            if "title slide" in n or "cover" in n:
                return lyt
        return layouts[0]

    if role == "section":
        for lyt in layouts:
            n = (lyt.name or "").lower()
            if "section" in n or "divider" in n or "header" in n:
                return lyt
        return layouts[min(2, len(layouts) - 1)]

    # Content role: prefer a layout that does not hide master shapes (showMasterSp != "0")
    preferred_keywords = ("title and content", "title only", "content", "custom")
    for kw in preferred_keywords:
        for lyt in layouts:
            n = (lyt.name or "").lower()
            show_master = lyt._element.get("showMasterSp", "1")
            if kw in n and show_master != "0":
                return lyt

    for lyt in layouts:
        if lyt._element.get("showMasterSp", "1") != "0":
            return lyt

    return layouts[min(6, len(layouts) - 1)]


def _create_slide_with_template(
    prs: Presentation,
    role: str,
    prototypes: Dict[str, dict],
    is_custom_template: bool,
):
    """
    Add a new slide using the appropriate layout, strip empty layout placeholders,
    and replay the captured prototype background (<p:bg> and full-bleed background shapes).
    """
    if not is_custom_template:
        layout_idx = 6 if len(prs.slide_layouts) > 6 else len(prs.slide_layouts) - 1
        slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
        for shape in list(slide.placeholders):
            sp = shape._element
            sp.getparent().remove(sp)
        return slide

    layout = _select_layout_for_role(prs, role, prototypes)
    slide = prs.slides.add_slide(layout)

    # Remove empty placeholder boxes inherited from the layout so they don't overlap our cards
    for shape in list(slide.placeholders):
        sp = shape._element
        sp.getparent().remove(sp)

    proto = prototypes.get(role) or prototypes.get("content")
    if not proto:
        return slide

    old_rels = proto.get("rels") or {}

    # 1. Replay slide-level <p:bg> if the prototype slide had a custom background fill/image
    if proto.get("bg_el") is not None:
        try:
            c_sld = slide._element.find(f"{_NS_P}cSld")
            if c_sld is not None:
                existing_bg = c_sld.find(f"{_NS_P}bg")
                if existing_bg is not None:
                    c_sld.remove(existing_bg)
                new_bg = copy.deepcopy(proto["bg_el"])
                _remap_embed_rids(new_bg, old_rels, slide.part)
                c_sld.insert(0, new_bg)
        except Exception:
            pass

    # 2. Replay full-bleed background pictures/shapes only (stretched to the full canvas)
    slide_w = int(prs.slide_width)
    slide_h = int(prs.slide_height)
    for item in proto.get("decor_items") or []:
        try:
            if item["kind"] == "picture":
                pic = slide.shapes.add_picture(
                    io.BytesIO(item["blob"]), 0, 0, width=slide_w, height=slide_h
                )
                # Send the background picture behind everything else on the slide
                sp_tree = slide.shapes._spTree
                sp_tree.remove(pic._element)
                sp_tree.insert(2, pic._element)
            elif item["kind"] == "xml":
                new_el = copy.deepcopy(item["el"])
                _remap_embed_rids(new_el, old_rels, slide.part)
                sp_tree = slide.shapes._spTree
                sp_tree.insert(2, new_el)
        except Exception:
            continue

    return slide
