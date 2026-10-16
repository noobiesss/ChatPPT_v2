import io
import json
import urllib.parse
import urllib.request
from enum import Enum
from typing import List, Dict, Optional

import streamlit as st
from PIL import Image
from pydantic import BaseModel, Field
from google import genai
from google.genai import types
from pypdf import PdfReader
from docx import Document

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION


# =====================================================================
# 1. Enumerations & Pydantic Schemas for Gemini Structured Output
# =====================================================================

class LayoutType(str, Enum):
    COVER = "cover"
    SECTION_DIVIDER = "section_divider"
    BULLET_CARDS = "bullet_cards"
    THREE_CARDS = "three_cards"
    TWO_COLUMN = "two_column"
    KPI_METRICS = "kpi_metrics"
    CHART_SLIDE = "chart_slide"
    IMAGE_SPLIT = "image_split"


class ChartType(str, Enum):
    BAR = "bar"
    LINE = "line"
    PIE = "pie"


class ImagePlacement(str, Enum):
    LEFT = "left"
    RIGHT = "right"


class VisualStyle(str, Enum):
    STOCK_PHOTO = "stock_photo"
    AI_ILLUSTRATION = "ai_illustration"


class StructuredPoint(BaseModel):
    heading: str = Field(
        description="Short, bold anchor phrase (3 to 6 words) summarizing the point, e.g., '3x Faster Query Speed'."
    )
    detail: str = Field(
        description="Concise explanation (12 to 25 words max). Avoid long paragraphs."
    )


class ColumnContent(BaseModel):
    column_title: str = Field(
        description="Header for this comparison column, e.g., 'Legacy System' or 'Proposed Architecture'."
    )
    items: List[StructuredPoint] = Field(
        description="2 to 4 structured points belonging to this column."
    )


class KPIMetric(BaseModel):
    value: str = Field(
        description="High-impact short metric string (max 8 chars), e.g., '+62%', '$2.4M', '99.9%', '3.5x'."
    )
    label: str = Field(
        description="Short metric title (2 to 4 words), e.g., 'Latency Reduction'."
    )
    context: str = Field(
        description="Brief 5 to 12 word context explaining the metric."
    )


class ChartSpec(BaseModel):
    chart_type: ChartType = Field(
        description="Select 'bar' for category comparisons, 'line' for time trends, or 'pie' for percentage shares."
    )
    chart_title: str = Field(
        description="Descriptive title for the chart graphic itself."
    )
    categories: List[str] = Field(
        description="Labels for each data point on the X-axis or pie slices (2 to 6 items)."
    )
    series_name: str = Field(
        description="Name of the data series including unit, e.g., 'Revenue ($M)' or 'Response Time (ms)'."
    )
    values: List[float] = Field(
        description="Numerical values corresponding 1-to-1 with 'categories'. Length MUST match 'categories'."
    )


class ImageSpec(BaseModel):
    preferred_style: VisualStyle = Field(
        default=VisualStyle.AI_ILLUSTRATION,
        description="Smart Hybrid Router: Choose 'stock_photo' for real-world subjects, people, offices, hardware, or cities. "
                    "Choose 'ai_illustration' for abstract technical architectures, software workflows, or conceptual diagrams."
    )
    visual_prompt: str = Field(
        description="Detailed English prompt (20 to 35 words) for an AI image generator describing a clean, "
                    "modern illustration or diagram that visually explains this slide's concept. "
                    "Specify 'clean minimal corporate style, no text inside image'."
    )
    search_keywords: str = Field(
        description="2 to 4 concise English keywords for Unsplash stock photo search, e.g., 'data center server room' or 'team strategy meeting'."
    )
    caption: str = Field(
        description="Short caption (8 to 15 words, in the target presentation language) displayed below the image."
    )
    placement: ImagePlacement = Field(
        default=ImagePlacement.LEFT,
        description="Alternate between 'left' and 'right' across different image slides."
    )


class SlideSpec(BaseModel):
    slide_number: int = Field(description="Sequential slide number starting from 1.")
    layout_type: LayoutType = Field(description="Visual layout archetype selected for this slide.")
    action_title: str = Field(
        description="Conclusion-first Action Title (6 to 14 words) stating the core takeaway, NOT a vague label."
    )
    subtitle_or_context: Optional[str] = Field(
        default=None,
        description="Used for 'cover' (presentation subtitle) or 'section_divider' (section overview)."
    )
    points: Optional[List[StructuredPoint]] = Field(
        default=None,
        description="Populate for 'bullet_cards' (3-4 items), 'three_cards' (3 items), "
                    "'chart_slide' (2-3 insights), or 'image_split' (2-3 key takeaways)."
    )
    left_column: Optional[ColumnContent] = Field(
        default=None,
        description="Populate ONLY when layout_type is 'two_column'."
    )
    right_column: Optional[ColumnContent] = Field(
        default=None,
        description="Populate ONLY when layout_type is 'two_column'."
    )
    kpis: Optional[List[KPIMetric]] = Field(
        default=None,
        description="Populate ONLY when layout_type is 'kpi_metrics' (2 to 4 headline numbers)."
    )
    chart: Optional[ChartSpec] = Field(
        default=None,
        description="Populate ONLY when layout_type is 'chart_slide'."
    )
    image: Optional[ImageSpec] = Field(
        default=None,
        description="Populate ONLY when layout_type is 'image_split'."
    )
    bottom_takeaway: Optional[str] = Field(
        default=None,
        description="Single punchy sentence (10 to 20 words) summarizing the 'So What?' in a bottom banner. "
                    "Leave None for 'cover' and 'section_divider'."
    )


class PresentationOutline(BaseModel):
    deck_title: str = Field(description="Overall title of the slide deck.")
    slides: List[SlideSpec] = Field(description="Ordered list of slides matching the requested slide count.")


# =====================================================================
# 2. Built-in 16:9 Color Themes
# =====================================================================

BUILTIN_THEMES: Dict[str, Dict[str, RGBColor]] = {
    "Corporate Blue": {
        "bg_cover": RGBColor(15, 37, 55),
        "bg_section": RGBColor(22, 52, 77),
        "bg_content": RGBColor(244, 247, 251),
        "card_bg": RGBColor(255, 255, 255),
        "card_border": RGBColor(215, 225, 238),
        "accent": RGBColor(0, 114, 206),
        "title_cover": RGBColor(255, 255, 255),
        "subtitle_cover": RGBColor(185, 215, 240),
        "title_content": RGBColor(15, 37, 55),
        "body_content": RGBColor(55, 65, 80),
        "takeaway_bg": RGBColor(15, 37, 55),
        "takeaway_text": RGBColor(255, 255, 255),
    },
    "Modern Dark": {
        "bg_cover": RGBColor(16, 16, 22),
        "bg_section": RGBColor(24, 24, 34),
        "bg_content": RGBColor(22, 22, 30),
        "card_bg": RGBColor(34, 35, 48),
        "card_border": RGBColor(58, 60, 82),
        "accent": RGBColor(99, 102, 241),
        "title_cover": RGBColor(255, 255, 255),
        "subtitle_cover": RGBColor(165, 170, 195),
        "title_content": RGBColor(242, 244, 255),
        "body_content": RGBColor(195, 200, 220),
        "takeaway_bg": RGBColor(48, 50, 78),
        "takeaway_text": RGBColor(240, 242, 255),
    },
    "Minimalist Light": {
        "bg_cover": RGBColor(245, 245, 243),
        "bg_section": RGBColor(235, 235, 232),
        "bg_content": RGBColor(252, 252, 250),
        "card_bg": RGBColor(255, 255, 255),
        "card_border": RGBColor(220, 220, 218),
        "accent": RGBColor(35, 35, 35),
        "title_cover": RGBColor(20, 20, 20),
        "subtitle_cover": RGBColor(95, 95, 95),
        "title_content": RGBColor(20, 20, 20),
        "body_content": RGBColor(65, 65, 65),
        "takeaway_bg": RGBColor(35, 35, 35),
        "takeaway_text": RGBColor(255, 255, 255),
    },
    "Emerald Forest": {
        "bg_cover": RGBColor(18, 43, 32),
        "bg_section": RGBColor(26, 58, 44),
        "bg_content": RGBColor(243, 248, 245),
        "card_bg": RGBColor(255, 255, 255),
        "card_border": RGBColor(205, 225, 212),
        "accent": RGBColor(40, 135, 85),
        "title_cover": RGBColor(255, 255, 255),
        "subtitle_cover": RGBColor(185, 220, 200),
        "title_content": RGBColor(18, 43, 32),
        "body_content": RGBColor(45, 65, 55),
        "takeaway_bg": RGBColor(18, 43, 32),
        "takeaway_text": RGBColor(255, 255, 255),
    },
}


# =====================================================================
# 3. Document Text & Embedded Image Extraction Helpers
# =====================================================================

def extract_text_from_uploaded_file(uploaded_file) -> str:
    """Extract plain text from uploaded PDF, DOCX, TXT, or MD files."""
    if uploaded_file is None:
        return ""

    filename = uploaded_file.name.lower()
    file_bytes = uploaded_file.getvalue()

    if filename.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(file_bytes))
        return "\n".join([page.extract_text() or "" for page in reader.pages])

    elif filename.endswith(".docx"):
        doc = Document(io.BytesIO(file_bytes))
        return "\n".join([para.text for para in doc.paragraphs if para.text.strip()])

    elif filename.endswith((".txt", ".md")):
        return file_bytes.decode("utf-8", errors="ignore")

    return ""


def extract_images_from_report(uploaded_file) -> List[Dict[str, bytes]]:
    """
    Extract embedded images from an uploaded PDF or DOCX report so users can reuse
    original diagrams, charts, or photos directly inside their slides.
    """
    extracted: List[Dict[str, bytes]] = []
    if uploaded_file is None:
        return extracted

    filename = uploaded_file.name.lower()
    file_bytes = uploaded_file.getvalue()

    try:
        if filename.endswith(".pdf"):
            reader = PdfReader(io.BytesIO(file_bytes))
            img_counter = 1
            for page in reader.pages:
                for image_file_object in page.images:
                    extracted.append({
                        "name": f"Report_Image_{img_counter} ({image_file_object.name})",
                        "bytes": image_file_object.data
                    })
                    img_counter += 1

        elif filename.endswith(".docx"):
            doc = Document(io.BytesIO(file_bytes))
            img_counter = 1
            for rel in doc.part.rels.values():
                if "image" in rel.target_ref:
                    extracted.append({
                        "name": f"Report_Image_{img_counter}",
                        "bytes": rel.target_part.blob
                    })
                    img_counter += 1
    except Exception:
        pass

    return extracted


# =====================================================================
# 4. Gemini Outline Router & Sanitizer
# =====================================================================

SYSTEM_INSTRUCTION = """
You are a Principal Presentation Designer at a top-tier consulting firm (McKinsey / BCG style).
Transform the user's source report into a visually diverse, executive-ready slide outline.

### STRICT DESIGN & ROUTING RULES:

1. ACTION TITLES (Conclusion-First):
   - Every content slide MUST have an `action_title` (6 to 14 words) that states the core finding or conclusion
     (e.g., "Automated Caching Reduced Peak Database Latency by 62%"), NEVER a generic label like "Performance Analysis".

2. SMART LAYOUT ROUTING (Vary layouts so no more than 2 consecutive slides share the same type):
   - `cover`: MUST be used for Slide 1. Put the main title in `action_title` and subtitle/context in `subtitle_or_context`.
   - `section_divider`: Use sparingly in decks with 8 or more slides to transition between major themes.
   - `image_split`: Use for 20% to 35% of slides—specifically when explaining physical products, system architectures,
     workflows, user scenarios, or concepts that audiences understand faster with a visual graphic.
     Populate BOTH `image` AND `points` (2 to 3 structured takeaways).
     Inside `image`, set `preferred_style`:
       * Choose `'stock_photo'` for real-world scenes, human teams, offices, hardware, retail, or cities (provide tangible `search_keywords` for Unsplash).
       * Choose `'ai_illustration'` for abstract software architectures, algorithms, data flows, or conceptual metaphors (provide a detailed `visual_prompt`).
   - `three_cards`: Use when presenting 3 parallel pillars, modules, features, or strategic recommendations (`points` = 3 items).
   - `two_column`: Use for comparisons (Pros vs. Cons, Before vs. After, Problem vs. Solution). Populate BOTH `left_column` and `right_column`.
   - `kpi_metrics`: Use when highlighting 2 to 4 headline numbers/statistics from the text (`kpis`).
   - `chart_slide`: Use when the source text contains a series of comparable numbers (trends, benchmarks, or shares).
     Populate `chart` AND `points` (2 to 3 key insights explaining the chart).
   - `bullet_cards`: Use for sequential explanations or general structured takeaways (3 to 4 `points`).

3. DATA INTEGRITY (NO HALLUCINATED NUMBERS):
   - ONLY choose `chart_slide` or `kpi_metrics` if real numerical figures exist in the source text.
   - In `ChartSpec`, the length of `categories` and `values` MUST be identical.

4. SCANNABLE COPYWRITING & TAKEAWAY BANNER:
   - Break every point into a short bold `heading` (3-6 words) and a concise `detail` (12-25 words).
   - Include a `bottom_takeaway` on all content slides summarizing the strategic implication ("So What?").
"""


def sanitize_outline_dict(outline: PresentationOutline) -> List[dict]:
    """
    Validate and repair LLM output edge cases and return a clean list of slide dictionaries
    with all sub-structures initialized for safe UI editing and PPTX rendering.
    """
    sanitized_slides = []

    for idx, slide in enumerate(outline.slides):
        s_dict = slide.model_dump()
        s_dict["slide_number"] = idx + 1
        layout = s_dict["layout_type"]

        if idx == 0:
            layout = LayoutType.COVER.value

        # Ensure default points exist
        if not s_dict.get("points"):
            s_dict["points"] = [
                {"heading": "Key Finding 1", "detail": s_dict.get("subtitle_or_context") or s_dict["action_title"]},
                {"heading": "Key Finding 2", "detail": "Supporting evidence and operational impact from the report."},
                {"heading": "Key Finding 3", "detail": "Strategic recommendation and next implementation phase."},
            ]

        # Ensure default two_column structures exist
        if not s_dict.get("left_column"):
            s_dict["left_column"] = {
                "column_title": "Current State / Challenges",
                "items": s_dict["points"][:2]
            }
        if not s_dict.get("right_column"):
            s_dict["right_column"] = {
                "column_title": "Target State / Solutions",
                "items": s_dict["points"][-2:]
            }

        # Ensure default KPIs exist
        if not s_dict.get("kpis"):
            s_dict["kpis"] = [
                {"value": "100%", "label": "Core Target", "context": "Primary benchmark highlighted in the report"},
                {"value": "3.0x", "label": "Efficiency Gain", "context": "Expected improvement across key workflows"},
                {"value": "-40%", "label": "Overhead Reduction", "context": "Estimated operational savings"},
            ]

        # Validate or initialize chart structure
        chart = s_dict.get("chart")
        if not chart or not chart.get("categories") or not chart.get("values"):
            if layout == LayoutType.CHART_SLIDE.value:
                layout = LayoutType.BULLET_CARDS.value
            s_dict["chart"] = {
                "chart_type": ChartType.BAR.value,
                "chart_title": "Performance Comparison",
                "categories": ["Category A", "Category B", "Category C"],
                "series_name": "Metric Value",
                "values": [45.0, 72.0, 95.0],
            }
        else:
            min_len = min(len(chart["categories"]), len(chart["values"]))
            if min_len < 2:
                if layout == LayoutType.CHART_SLIDE.value:
                    layout = LayoutType.BULLET_CARDS.value
            else:
                chart["categories"] = chart["categories"][:min_len]
                chart["values"] = chart["values"][:min_len]

        # Validate or initialize image spec
        if not s_dict.get("image"):
            s_dict["image"] = {
                "preferred_style": VisualStyle.AI_ILLUSTRATION.value,
                "visual_prompt": f"Clean modern corporate illustration representing {s_dict['action_title']}, minimal vector style, no text",
                "search_keywords": "modern technology business",
                "caption": s_dict["action_title"],
                "placement": ImagePlacement.LEFT.value if idx % 2 == 0 else ImagePlacement.RIGHT.value,
            }
        elif not s_dict["image"].get("preferred_style"):
            s_dict["image"]["preferred_style"] = VisualStyle.AI_ILLUSTRATION.value

        s_dict["layout_type"] = layout
        # Default UI image source selection to Smart Hybrid (Auto)
        s_dict["image_source_mode"] = "Smart Hybrid (Auto)"
        s_dict["selected_report_image_idx"] = 0
        s_dict["custom_uploaded_image_bytes"] = None

        sanitized_slides.append(s_dict)

    return sanitized_slides


def generate_smart_outline(
    api_key: str,
    source_text: str,
    num_slides: int,
    tone: str,
    language: str,
    model_name: str = "gemini-3.8-flash"
) -> List[dict]:
    """Call Gemini with Structured Outputs to generate a multi-layout presentation outline."""
    client = genai.Client(api_key=api_key)

    user_prompt = f"""
    Create a {num_slides}-slide presentation outline based on the source document below.

    Target Specifications:
    - Total Slides: Exactly {num_slides} slides (Slide 1 MUST be 'cover').
    - Presentation Tone: {tone}.
    - Output Language: {language} (Write all titles, headings, details, chart labels, captions, and takeaways in {language};
      keep ONLY `visual_prompt` and `search_keywords` in English).

    Source Document:
    \"\"\"
    {source_text}
    \"\"\"
    """

    response = client.models.generate_content(
        model=model_name,
        contents=user_prompt,
        config=types.GenerateContentConfig(
            system_instruction=SYSTEM_INSTRUCTION,
            response_mime_type="application/json",
            response_schema=PresentationOutline,
            temperature=0.3,
        ),
    )

    parsed_outline = PresentationOutline.model_validate_json(response.text)
    return sanitize_outline_dict(parsed_outline)


# =====================================================================
# 5. Smart Hybrid Image Engine (Unsplash API + Gemini AI + Fallback)
# =====================================================================

def fetch_unsplash_image_bytes(search_keywords: str, unsplash_access_key: Optional[str]) -> Optional[bytes]:
    """
    Query Unsplash Search API for a high-resolution landscape stock photo matching search_keywords.
    Returns raw image bytes or None if no key/results are available.
    """
    if not unsplash_access_key or not search_keywords:
        return None

    try:
        query = urllib.parse.quote(search_keywords.strip())
        api_url = (
            f"https://api.unsplash.com/search/photos"
            f"?query={query}&per_page=1&orientation=landscape&content_filter=high"
        )
        req = urllib.request.Request(
            api_url,
            headers={
                "Authorization": f"Client-ID {unsplash_access_key.strip()}",
                "Accept-Version": "v1",
                "User-Agent": "ChatPPT-V2",
            },
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            results = data.get("results", [])
            if not results:
                return None
            img_url = results[0]["urls"]["regular"]

        img_req = urllib.request.Request(img_url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(img_req, timeout=10) as img_resp:
            return img_resp.read()
    except Exception:
        return None


def generate_ai_image_bytes(
    visual_prompt: str,
    api_key: Optional[str] = None,
    use_imagen: bool = False
) -> Optional[bytes]:
    """
    Generate a 4:3 conceptual illustration for a slide.
    Tries Google Imagen 3 first (if enabled), and falls back to Pollinations AI stream.
    """
    enhanced_prompt = (
        f"{visual_prompt}, clean modern corporate presentation illustration, "
        "minimalist high-contrast aesthetic, no text, no watermarks"
    )

    if use_imagen and api_key:
        try:
            client = genai.Client(api_key=api_key)
            result = client.models.generate_images(
                model="imagen-3.0-generate-002",
                prompt=enhanced_prompt,
                config=types.GenerateImagesConfig(
                    number_of_images=1,
                    output_mime_type="image/jpeg",
                    aspect_ratio="4:3",
                ),
            )
            for generated_image in result.generated_images:
                return generated_image.image.image_bytes
        except Exception:
            pass

    try:
        encoded_prompt = urllib.parse.quote(enhanced_prompt)
        url = f"https://image.pollinations.ai/prompt/{encoded_prompt}?width=800&height=600&nologo=true&seed=42"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=12) as response:
            return response.read()
    except Exception:
        return None


def resolve_slide_image_bytes(
    slide_data: dict,
    api_key: str,
    use_imagen: bool,
    unsplash_access_key: str,
    extracted_report_images: List[Dict[str, bytes]],
) -> Optional[bytes]:
    """
    Smart Hybrid Image Resolver:
    Routes image retrieval across Custom Uploads, Extracted Report Figures,
    Unsplash Stock Photography, and AI Concept Illustrations with automatic fallback.
    """
    src_mode = slide_data.get("image_source_mode", "Smart Hybrid (Auto)")
    image_spec = slide_data.get("image") or {}
    preferred_style = image_spec.get("preferred_style", VisualStyle.AI_ILLUSTRATION.value)
    visual_prompt = image_spec.get("visual_prompt") or slide_data.get("action_title", "Corporate illustration")
    search_keywords = image_spec.get("search_keywords") or slide_data.get("action_title", "business technology")

    # 1. Explicit Custom Upload
    if src_mode == "Upload Custom Image" and slide_data.get("custom_uploaded_image_bytes"):
        return slide_data["custom_uploaded_image_bytes"]

    # 2. Explicit Report Embedded Image
    if src_mode == "Use Image from Report" and extracted_report_images:
        sel_idx = min(slide_data.get("selected_report_image_idx", 0), len(extracted_report_images) - 1)
        return extracted_report_images[sel_idx]["bytes"]

    # 3. Explicit Placeholder Only
    if src_mode == "Visual Placeholder Only":
        return None

    # 4. Explicit Unsplash Stock Photo (falls back to AI if Unsplash has no match or no key)
    if src_mode == "Unsplash Stock Photo":
        img_bytes = fetch_unsplash_image_bytes(search_keywords, unsplash_access_key)
        if img_bytes:
            return img_bytes
        return generate_ai_image_bytes(visual_prompt, api_key=api_key, use_imagen=use_imagen)

    # 5. Explicit AI Generated Illustration
    if src_mode == "AI Generated Illustration":
        return generate_ai_image_bytes(visual_prompt, api_key=api_key, use_imagen=use_imagen)

    # 6. Smart Hybrid (Auto) — Follows Gemini's per-slide `preferred_style` with cross-fallback
    if preferred_style == VisualStyle.STOCK_PHOTO.value and unsplash_access_key:
        img_bytes = fetch_unsplash_image_bytes(search_keywords, unsplash_access_key)
        if img_bytes:
            return img_bytes
        return generate_ai_image_bytes(visual_prompt, api_key=api_key, use_imagen=use_imagen)
    else:
        img_bytes = generate_ai_image_bytes(visual_prompt, api_key=api_key, use_imagen=use_imagen)
        if img_bytes:
            return img_bytes
        return fetch_unsplash_image_bytes(search_keywords, unsplash_access_key)


def add_picture_fitted(
    slide,
    img_bytes: bytes,
    box_left: float,
    box_top: float,
    box_width: float,
    box_height: float
) -> bool:
    """
    Insert an image into a bounding box (in Inches) while preserving its native aspect ratio
    and centering it inside the box. Converts unsupported modes (WEBP/RGBA) to RGB automatically.
    """
    try:
        with Image.open(io.BytesIO(img_bytes)) as pil_img:
            if pil_img.mode in ("RGBA", "P") or pil_img.format not in ("JPEG", "PNG"):
                rgb_img = pil_img.convert("RGB")
                buf = io.BytesIO()
                rgb_img.save(buf, format="JPEG", quality=92)
                img_bytes = buf.getvalue()
                img_w, img_h = rgb_img.size
            else:
                img_w, img_h = pil_img.size

        img_ratio = img_w / float(img_h)
        box_ratio = box_width / float(box_height)

        if img_ratio > box_ratio:
            final_w = box_width
            final_h = box_width / img_ratio
        else:
            final_h = box_height
            final_w = box_height * img_ratio

        offset_left = box_left + (box_width - final_w) / 2.0
        offset_top = box_top + (box_height - final_h) / 2.0

        slide.shapes.add_picture(
            io.BytesIO(img_bytes),
            Inches(offset_left),
            Inches(offset_top),
            width=Inches(final_w),
            height=Inches(final_h),
        )
        return True
    except Exception:
        return False


# =====================================================================
# 6. Multi-Layout PowerPoint Rendering Engine (8 Archetypes)
# =====================================================================

def set_solid_background(slide, color: RGBColor) -> None:
    """Apply a solid background color to a slide."""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color


def clear_existing_slides(prs: Presentation) -> None:
    """Remove sample slides from an uploaded .pptx template while keeping Slide Masters intact."""
    for i in range(len(prs.slides) - 1, -1, -1):
        r_id = prs.slides._sldIdLst[i].rId
        prs.part.drop_rel(r_id)
        del prs.slides._sldIdLst[i]


def get_blank_or_clean_slide(prs: Presentation):
    """Select the cleanest available layout from a Presentation and remove leftover empty placeholders."""
    layout_idx = 6 if len(prs.slide_layouts) > 6 else len(prs.slide_layouts) - 1
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    for shape in list(slide.placeholders):
        sp = shape._element
        sp.getparent().remove(sp)
    return slide


def draw_card_shape(slide, left: float, top: float, width: float, height: float, theme: Dict[str, RGBColor]):
    """Draw a styled rounded rectangle card container."""
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    card.adjustments[0] = 0.04
    card.fill.solid()
    card.fill.fore_color.rgb = theme["card_bg"]
    card.line.color.rgb = theme["card_border"]
    card.line.width = Pt(1)
    return card


def draw_header_and_takeaway(
    slide,
    slide_data: dict,
    slide_idx: int,
    total_slides: int,
    theme: Dict[str, RGBColor],
    apply_bg: bool = True
) -> None:
    """Render the top accent bar, Action Title, bottom takeaway banner, and slide number."""
    if apply_bg:
        set_solid_background(slide, theme["bg_content"])

    # Top accent strip
    top_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12)
    )
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = theme["accent"]
    top_bar.line.fill.background()

    # Action Title
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.42), Inches(11.7), Inches(1.05))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = slide_data.get("action_title", "")
    p_title.font.size = Pt(26)
    p_title.font.bold = True
    p_title.font.color.rgb = theme["title_content"]

    # Accent underline below title
    divider = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.52), Inches(11.7), Inches(0.03)
    )
    divider.fill.solid()
    divider.fill.fore_color.rgb = theme["accent"]
    divider.line.fill.background()

    # Bottom Takeaway Banner ("So What?")
    takeaway = (slide_data.get("bottom_takeaway") or "").strip()
    if takeaway:
        banner = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.35), Inches(10.8), Inches(0.68)
        )
        banner.adjustments[0] = 0.15
        banner.fill.solid()
        banner.fill.fore_color.rgb = theme["takeaway_bg"]
        banner.line.fill.background()

        tb_take = slide.shapes.add_textbox(Inches(1.0), Inches(6.42), Inches(10.4), Inches(0.55))
        tf_take = tb_take.text_frame
        tf_take.word_wrap = True
        p_take = tf_take.paragraphs[0]
        p_take.text = f"💡 Key Takeaway: {takeaway}"
        p_take.font.size = Pt(13.5)
        p_take.font.bold = True
        p_take.font.color.rgb = theme["takeaway_text"]

    # Slide Number Footer
    footer_box = slide.shapes.add_textbox(Inches(11.7), Inches(6.45), Inches(0.8), Inches(0.45))
    tf_footer = footer_box.text_frame
    p_footer = tf_footer.paragraphs[0]
    p_footer.text = f"{slide_idx + 1} / {total_slides}"
    p_footer.alignment = PP_ALIGN.RIGHT
    p_footer.font.size = Pt(11)
    p_footer.font.color.rgb = theme["accent"]


def render_cover_slide(slide, slide_data: dict, theme: Dict[str, RGBColor], apply_bg: bool = True) -> None:
    """Render Layout 1: Cover Slide."""
    if apply_bg:
        set_solid_background(slide, theme["bg_cover"])

    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(1.1), Inches(2.2), Inches(1.4), Inches(0.14)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = theme["accent"]
    bar.line.fill.background()

    title_box = slide.shapes.add_textbox(Inches(1.1), Inches(2.5), Inches(11.0), Inches(2.2))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = slide_data.get("action_title", "Presentation Title")
    p_title.font.size = Pt(42)
    p_title.font.bold = True
    p_title.font.color.rgb = theme["title_cover"] if apply_bg else theme["title_content"]

    subtitle = (slide_data.get("subtitle_or_context") or "").strip()
    if subtitle:
        sub_box = slide.shapes.add_textbox(Inches(1.1), Inches(4.8), Inches(10.5), Inches(1.4))
        tf_sub = sub_box.text_frame
        tf_sub.word_wrap = True
        p_sub = tf_sub.paragraphs[0]
        p_sub.text = subtitle
        p_sub.font.size = Pt(20)
        p_sub.font.color.rgb = theme["subtitle_cover"] if apply_bg else theme["body_content"]


def render_section_divider_slide(slide, slide_data: dict, theme: Dict[str, RGBColor], apply_bg: bool = True) -> None:
    """Render Layout 2: Section Divider Slide."""
    if apply_bg:
        set_solid_background(slide, theme["bg_section"])

    accent_pill = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.1), Inches(2.6), Inches(0.9), Inches(0.12)
    )
    accent_pill.fill.solid()
    accent_pill.fill.fore_color.rgb = theme["accent"]
    accent_pill.line.fill.background()

    title_box = slide.shapes.add_textbox(Inches(1.1), Inches(2.9), Inches(11.0), Inches(1.6))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = slide_data.get("action_title", "Section Title")
    p_title.font.size = Pt(38)
    p_title.font.bold = True
    p_title.font.color.rgb = theme["title_cover"] if apply_bg else theme["title_content"]

    context = (slide_data.get("subtitle_or_context") or "").strip()
    if context:
        sub_box = slide.shapes.add_textbox(Inches(1.1), Inches(4.5), Inches(10.0), Inches(1.2))
        tf_sub = sub_box.text_frame
        tf_sub.word_wrap = True
        p_sub = tf_sub.paragraphs[0]
        p_sub.text = context
        p_sub.font.size = Pt(18)
        p_sub.font.color.rgb = theme["subtitle_cover"] if apply_bg else theme["body_content"]


def render_bullet_cards_slide(slide, slide_data: dict, theme: Dict[str, RGBColor]) -> None:
    """Render Layout 3: Stacked Horizontal Structured Cards."""
    points = (slide_data.get("points") or [])[:4]
    count = max(len(points), 1)
    avail_h = 4.35
    gap = 0.16
    card_h = (avail_h - gap * (count - 1)) / count

    for idx, pt in enumerate(points):
        top = 1.78 + idx * (card_h + gap)
        draw_card_shape(slide, 0.8, top, 11.7, card_h, theme)

        # Left number badge
        badge = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(1.05), Inches(top + (card_h - 0.46) / 2), Inches(0.46), Inches(0.46)
        )
        badge.fill.solid()
        badge.fill.fore_color.rgb = theme["accent"]
        badge.line.fill.background()
        tf_b = badge.text_frame
        p_b = tf_b.paragraphs[0]
        p_b.text = str(idx + 1)
        p_b.alignment = PP_ALIGN.CENTER
        p_b.font.size = Pt(14)
        p_b.font.bold = True
        p_b.font.color.rgb = RGBColor(255, 255, 255)

        # Heading & Detail text
        tb = slide.shapes.add_textbox(Inches(1.7), Inches(top + 0.08), Inches(10.5), Inches(card_h - 0.16))
        tf = tb.text_frame
        tf.word_wrap = True
        p_head = tf.paragraphs[0]
        p_head.text = pt.get("heading", "")
        p_head.font.size = Pt(17)
        p_head.font.bold = True
        p_head.font.color.rgb = theme["title_content"]
        p_head.space_after = Pt(3)

        p_det = tf.add_paragraph()
        p_det.text = pt.get("detail", "")
        p_det.font.size = Pt(14)
        p_det.font.color.rgb = theme["body_content"]


def render_three_cards_slide(slide, slide_data: dict, theme: Dict[str, RGBColor]) -> None:
    """Render Layout 4: 3 Parallel Vertical Pillar Cards."""
    points = (slide_data.get("points") or [])[:3]
    while len(points) < 3:
        points.append({"heading": f"Pillar {len(points) + 1}", "detail": "Key supporting detail."})

    card_w = 3.7
    gap = 0.3
    card_h = 4.3

    for idx, pt in enumerate(points):
        left = 0.8 + idx * (card_w + gap)
        top = 1.8
        draw_card_shape(slide, left, top, card_w, card_h, theme)

        # Accent header bar inside each pillar card
        strip = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(left + 0.25), Inches(top + 0.28), Inches(0.8), Inches(0.08)
        )
        strip.fill.solid()
        strip.fill.fore_color.rgb = theme["accent"]
        strip.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(left + 0.25), Inches(top + 0.5), Inches(card_w - 0.5), Inches(card_h - 0.7))
        tf = tb.text_frame
        tf.word_wrap = True

        p_head = tf.paragraphs[0]
        p_head.text = pt.get("heading", "")
        p_head.font.size = Pt(19)
        p_head.font.bold = True
        p_head.font.color.rgb = theme["title_content"]
        p_head.space_after = Pt(12)

        p_det = tf.add_paragraph()
        p_det.text = pt.get("detail", "")
        p_det.font.size = Pt(14.5)
        p_det.font.color.rgb = theme["body_content"]


def render_two_column_slide(slide, slide_data: dict, theme: Dict[str, RGBColor]) -> None:
    """Render Layout 5: Side-by-Side Two-Column Comparison."""
    left_col = slide_data.get("left_column") or {"column_title": "Column A", "items": []}
    right_col = slide_data.get("right_column") or {"column_title": "Column B", "items": []}

    col_w = 5.65
    gap = 0.4
    card_h = 4.3

    for idx, col_data in enumerate([left_col, right_col]):
        left = 0.8 + idx * (col_w + gap)
        top = 1.8
        draw_card_shape(slide, left, top, col_w, card_h, theme)

        # Column Header Banner
        header_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left + 0.2), Inches(top + 0.2), Inches(col_w - 0.4), Inches(0.55)
        )
        header_shape.fill.solid()
        header_shape.fill.fore_color.rgb = theme["accent"] if idx == 1 else theme["takeaway_bg"]
        header_shape.line.fill.background()

        tf_h = header_shape.text_frame
        p_h = tf_h.paragraphs[0]
        p_h.text = col_data.get("column_title", "")
        p_h.alignment = PP_ALIGN.CENTER
        p_h.font.size = Pt(16)
        p_h.font.bold = True
        p_h.font.color.rgb = RGBColor(255, 255, 255)

        # Column Items
        tb = slide.shapes.add_textbox(Inches(left + 0.25), Inches(top + 0.9), Inches(col_w - 0.5), Inches(card_h - 1.1))
        tf = tb.text_frame
        tf.word_wrap = True

        for item_idx, item in enumerate((col_data.get("items") or [])[:4]):
            p_head = tf.paragraphs[0] if item_idx == 0 else tf.add_paragraph()
            p_head.text = f"▪ {item.get('heading', '')}"
            p_head.font.size = Pt(15.5)
            p_head.font.bold = True
            p_head.font.color.rgb = theme["title_content"]
            p_head.space_after = Pt(2)

            p_det = tf.add_paragraph()
            p_det.text = f"   {item.get('detail', '')}"
            p_det.font.size = Pt(13.5)
            p_det.font.color.rgb = theme["body_content"]
            p_det.space_after = Pt(10)


def render_kpi_metrics_slide(slide, slide_data: dict, theme: Dict[str, RGBColor]) -> None:
    """Render Layout 6: Hero KPI Metric Cards."""
    kpis = (slide_data.get("kpis") or [])[:4]
    count = max(len(kpis), 1)
    total_w = 11.7
    gap = 0.28
    card_w = (total_w - gap * (count - 1)) / count
    card_h = 4.2

    for idx, kpi in enumerate(kpis):
        left = 0.8 + idx * (card_w + gap)
        top = 1.85
        draw_card_shape(slide, left, top, card_w, card_h, theme)

        tb = slide.shapes.add_textbox(Inches(left + 0.2), Inches(top + 0.5), Inches(card_w - 0.4), Inches(card_h - 0.8))
        tf = tb.text_frame
        tf.word_wrap = True

        # Big Hero Metric Number
        p_val = tf.paragraphs[0]
        p_val.text = kpi.get("value", "")
        p_val.alignment = PP_ALIGN.CENTER
        p_val.font.size = Pt(46 if count <= 3 else 38)
        p_val.font.bold = True
        p_val.font.color.rgb = theme["accent"]
        p_val.space_after = Pt(12)

        # Metric Label
        p_lbl = tf.add_paragraph()
        p_lbl.text = kpi.get("label", "")
        p_lbl.alignment = PP_ALIGN.CENTER
        p_lbl.font.size = Pt(18)
        p_lbl.font.bold = True
        p_lbl.font.color.rgb = theme["title_content"]
        p_lbl.space_after = Pt(10)

        # Metric Context
        p_ctx = tf.add_paragraph()
        p_ctx.text = kpi.get("context", "")
        p_ctx.alignment = PP_ALIGN.CENTER
        p_ctx.font.size = Pt(13.5)
        p_ctx.font.color.rgb = theme["body_content"]


def render_chart_slide(slide, slide_data: dict, theme: Dict[str, RGBColor]) -> None:
    """Render Layout 7: Native Editable PowerPoint Chart (Left) + Key Insights (Right)."""
    chart_spec = slide_data.get("chart") or {}
    categories = chart_spec.get("categories") or ["A", "B", "C"]
    values = chart_spec.get("values") or [10.0, 20.0, 30.0]
    min_len = min(len(categories), len(values))
    categories = categories[:min_len]
    values = values[:min_len]

    # Left Card Frame for Chart
    draw_card_shape(slide, 0.8, 1.8, 6.4, 4.3, theme)

    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series(chart_spec.get("series_name", "Series 1"), values)

    c_type_str = chart_spec.get("chart_type", "bar")
    if c_type_str == ChartType.LINE.value:
        xl_type = XL_CHART_TYPE.LINE_MARKERS
    elif c_type_str == ChartType.PIE.value:
        xl_type = XL_CHART_TYPE.PIE
    else:
        xl_type = XL_CHART_TYPE.COLUMN_CLUSTERED

    chart_frame = slide.shapes.add_chart(
        xl_type, Inches(1.0), Inches(1.95), Inches(6.0), Inches(4.0), chart_data
    )
    chart = chart_frame.chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.TOP
    chart.legend.include_in_layout = False

    if chart_spec.get("chart_title"):
        chart.has_title = True
        chart.chart_title.text_frame.text = chart_spec["chart_title"]

    # Right Side Commentary Cards
    points = (slide_data.get("points") or [])[:3]
    count = max(len(points), 1)
    card_h = (4.3 - 0.15 * (count - 1)) / count

    for idx, pt in enumerate(points):
        top = 1.8 + idx * (card_h + 0.15)
        draw_card_shape(slide, 7.45, top, 5.05, card_h, theme)

        tb = slide.shapes.add_textbox(Inches(7.65), Inches(top + 0.1), Inches(4.65), Inches(card_h - 0.2))
        tf = tb.text_frame
        tf.word_wrap = True

        p_head = tf.paragraphs[0]
        p_head.text = pt.get("heading", "")
        p_head.font.size = Pt(16)
        p_head.font.bold = True
        p_head.font.color.rgb = theme["title_content"]
        p_head.space_after = Pt(4)

        p_det = tf.add_paragraph()
        p_det.text = pt.get("detail", "")
        p_det.font.size = Pt(13)
        p_det.font.color.rgb = theme["body_content"]


def render_image_split_slide(
    slide,
    slide_data: dict,
    theme: Dict[str, RGBColor],
    img_bytes: Optional[bytes] = None
) -> None:
    """Render Layout 8: Visual Image + Caption on one side, Structured Points on the other."""
    image_info = slide_data.get("image") or {}
    placement = image_info.get("placement", "left")
    caption_text = image_info.get("caption", "")
    visual_prompt = image_info.get("visual_prompt", "Visual Illustration")
    search_kw = image_info.get("search_keywords", "")

    if placement == "right":
        text_left, img_left = 0.8, 6.95
    else:
        img_left, text_left = 0.8, 6.55

    img_w, img_h = 5.55, 3.65
    img_top = 1.8

    # Framed Container for Image + Caption
    draw_card_shape(slide, img_left, img_top, img_w, 4.3, theme)

    inserted = False
    if img_bytes:
        inserted = add_picture_fitted(
            slide,
            img_bytes,
            box_left=img_left + 0.15,
            box_top=img_top + 0.15,
            box_width=img_w - 0.3,
            box_height=img_h - 0.15,
        )

    if not inserted:
        # Styled Visual Blueprint Placeholder
        ph_box = slide.shapes.add_textbox(
            Inches(img_left + 0.3), Inches(img_top + 0.6), Inches(img_w - 0.6), Inches(img_h - 0.7)
        )
        tf_ph = ph_box.text_frame
        tf_ph.word_wrap = True
        p_ph = tf_ph.paragraphs[0]
        p_ph.text = f"🖼️ [ Visual Concept Placeholder ]\n\nKeywords: {search_kw}\nPrompt: {visual_prompt}"
        p_ph.alignment = PP_ALIGN.CENTER
        p_ph.font.size = Pt(12.5)
        p_ph.font.color.rgb = theme["body_content"]

    if caption_text:
        cap_box = slide.shapes.add_textbox(
            Inches(img_left + 0.2), Inches(img_top + img_h + 0.05), Inches(img_w - 0.4), Inches(0.5)
        )
        tf_cap = cap_box.text_frame
        tf_cap.word_wrap = True
        p_cap = tf_cap.paragraphs[0]
        p_cap.text = f"▲ {caption_text}"
        p_cap.alignment = PP_ALIGN.CENTER
        p_cap.font.size = Pt(11.5)
        p_cap.font.italic = True
        p_cap.font.color.rgb = theme["accent"]

    # Opposite Side: Structured Key Takeaways
    points = (slide_data.get("points") or [])[:3]
    count = max(len(points), 1)
    card_h = (4.3 - 0.15 * (count - 1)) / count

    for idx, pt in enumerate(points):
        top = 1.8 + idx * (card_h + 0.15)
        draw_card_shape(slide, text_left, top, 5.95, card_h, theme)

        tb = slide.shapes.add_textbox(Inches(text_left + 0.22), Inches(top + 0.12), Inches(5.5), Inches(card_h - 0.24))
        tf = tb.text_frame
        tf.word_wrap = True

        p_head = tf.paragraphs[0]
        p_head.text = pt.get("heading", "")
        p_head.font.size = Pt(16.5)
        p_head.font.bold = True
        p_head.font.color.rgb = theme["title_content"]
        p_head.space_after = Pt(4)

        p_det = tf.add_paragraph()
        p_det.text = pt.get("detail", "")
        p_det.font.size = Pt(13.5)
        p_det.font.color.rgb = theme["body_content"]


def build_presentation_pptx(
    slides_data: List[dict],
    theme_name: str,
    api_key: str,
    use_imagen: bool,
    unsplash_access_key: str,
    extracted_report_images: List[Dict[str, bytes]],
    custom_template_bytes: Optional[bytes] = None,
) -> bytes:
    """
    Assemble the complete .pptx presentation across all 8 visual layouts.
    Supports both Built-in Color Themes and Custom Uploaded .pptx Templates.
    """
    theme = BUILTIN_THEMES.get(theme_name, BUILTIN_THEMES["Corporate Blue"])

    if custom_template_bytes:
        prs = Presentation(io.BytesIO(custom_template_bytes))
        clear_existing_slides(prs)
        apply_solid_bg = False  # Preserve custom template's master background graphics
    else:
        prs = Presentation()
        apply_solid_bg = True

    # Standardize canvas to 16:9 widescreen (13.333 x 7.5 inches)
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    total_slides = len(slides_data)

    for idx, slide_data in enumerate(slides_data):
        layout = slide_data.get("layout_type", LayoutType.BULLET_CARDS.value)
        slide = get_blank_or_clean_slide(prs)

        if layout == LayoutType.COVER.value:
            render_cover_slide(slide, slide_data, theme, apply_bg=apply_solid_bg)

        elif layout == LayoutType.SECTION_DIVIDER.value:
            render_section_divider_slide(slide, slide_data, theme, apply_bg=apply_solid_bg)

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

            elif layout == LayoutType.IMAGE_SPLIT.value:
                img_bytes = resolve_slide_image_bytes(
                    slide_data=slide_data,
                    api_key=api_key,
                    use_imagen=use_imagen,
                    unsplash_access_key=unsplash_access_key,
                    extracted_report_images=extracted_report_images,
                )
                render_image_split_slide(slide, slide_data, theme, img_bytes=img_bytes)

            else:
                render_bullet_cards_slide(slide, slide_data, theme)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output.getvalue()


# =====================================================================
# 7. Helper UI Functions for Structured Text Area Editing
# =====================================================================

def format_points_for_editor(points: List[dict]) -> str:
    """Serialize structured points as 'Heading :: Detail' lines for easy text-area editing."""
    lines = []
    for pt in points or []:
        h = (pt.get("heading") or "").strip()
        d = (pt.get("detail") or "").strip()
        lines.append(f"{h} :: {d}" if d else h)
    return "\n".join(lines)


def parse_points_from_editor(raw_text: str) -> List[dict]:
    """Parse 'Heading :: Detail' lines back into structured point dictionaries."""
    parsed = []
    for line in raw_text.split("\n"):
        line = line.strip()
        if not line:
            continue
        if "::" in line:
            parts = line.split("::", 1)
            parsed.append({"heading": parts[0].strip(), "detail": parts[1].strip()})
        else:
            parsed.append({"heading": line[:40], "detail": line})
    return parsed


def format_kpis_for_editor(kpis: List[dict]) -> str:
    """Serialize KPIs as 'Value | Label | Context' lines."""
    return "\n".join(
        f"{k.get('value', '')} | {k.get('label', '')} | {k.get('context', '')}"
        for k in (kpis or [])
    )


def parse_kpis_from_editor(raw_text: str) -> List[dict]:
    """Parse 'Value | Label | Context' lines back into KPI dictionaries."""
    parsed = []
    for line in raw_text.split("\n"):
        line = line.strip()
        if not line:
            continue
        parts = [p.strip() for p in line.split("|")]
        if len(parts) >= 3:
            parsed.append({"value": parts[0], "label": parts[1], "context": parts[2]})
        elif len(parts) == 2:
            parsed.append({"value": parts[0], "label": parts[1], "context": ""})
        else:
            parsed.append({"value": parts[0][:8], "label": "Metric", "context": parts[0]})
    return parsed


# =====================================================================
# 8. Main Streamlit Application
# =====================================================================

def main():
    st.set_page_config(page_title="ChatPPT V2 - Smart Visual Presentation Generator", page_icon="📊", layout="wide")

    if "slides_editor_data" not in st.session_state:
        st.session_state.slides_editor_data = None
    if "extracted_report_images" not in st.session_state:
        st.session_state.extracted_report_images = []
    if "pptx_bytes" not in st.session_state:
        st.session_state.pptx_bytes = None

    # --- Sidebar Configuration ---
    with st.sidebar:
        st.title("⚙️ ChatPPT V2 Settings")
        api_key = st.text_input("Gemini API Key", type="password", help="Enter your Google Gemini API key.")

        model_choice = st.selectbox(
            "Gemini Model",
            [
                "gemini-3.8-flash",
                "gemini-3.5-flash",
                "gemini-3.1-flash-lite",
                "gemini-2.5-flash",
                "Custom Model ID...",
            ],
            index=0,
            help="Select the active Gemini model for your API key.",
        )
        if model_choice == "Custom Model ID...":
            active_model = st.text_input("Enter Custom Model ID", value="gemini-3.8-flash").strip()
        else:
            active_model = model_choice

        st.subheader("1. Deck Parameters")
        num_slides = st.slider("Number of Slides", min_value=4, max_value=20, value=8, step=1)
        tone = st.selectbox(
            "Presentation Tone",
            [
                "Executive & Consulting (McKinsey Style)",
                "Professional & Corporate",
                "Academic & Analytical",
                "Persuasive & Venture Pitch",
                "Technical Architecture & Engineering",
            ],
        )
        output_lang = st.selectbox(
            "Output Language",
            ["English", "Simplified Chinese", "Traditional Chinese", "Malay", "Japanese"],
        )

        st.subheader("2. Visual Theme & Template")
        template_mode = st.radio(
            "Styling Mode",
            ["Built-in 16:9 Color Theme", "Upload Custom .pptx Template"],
        )
        selected_theme = st.selectbox("Color Palette", list(BUILTIN_THEMES.keys()))
        uploaded_template = None
        if template_mode == "Upload Custom .pptx Template":
            uploaded_template = st.file_uploader(
                "Upload .pptx Template",
                type=["pptx"],
                help="Preserves your template's Slide Master background while using the selected Color Palette for cards and charts.",
            )

        st.subheader("3. Smart Hybrid Image Engine")
        unsplash_key = st.text_input(
            "Unsplash Access Key (Optional)",
            type="password",
            help="Enter your free Unsplash Access Key for real stock photography. If left blank, Smart Hybrid Mode automatically falls back to AI generation.",
        )
        use_imagen = st.checkbox(
            "Use Google Imagen 3 (via Gemini Key)",
            value=False,
            help="When enabled, tries Google Imagen 3 first for AI illustrations and falls back to Pollinations AI if unavailable.",
        )

    # --- Main Content Area ---
    st.title("📊 ChatPPT V2 — Smart Visual Slide Generator")
    st.caption(
        "Automatically routes your report into 8 consulting-grade layouts with Smart Hybrid Visuals "
        "(Unsplash Real Stock Photography + Gemini AI Concept Illustrations + Extracted Report Figures)."
    )

    # Step 1: Upload or Paste Source Material
    st.header("Step 1: Provide Your Report or Notes")
    col1, col2 = st.columns(2)

    with col1:
        uploaded_report = st.file_uploader(
            "Upload Report (.pdf, .docx, .txt, .md) — Embedded images are extracted automatically!",
            type=["pdf", "docx", "txt", "md"],
        )

    with col2:
        pasted_text = st.text_area(
            "Or Paste Report Text / Data Directly",
            height=160,
            placeholder="Paste your report, experimental metrics, or project summary here...",
        )

    if st.button("✨ Step 1: Analyze Report & Generate Smart Visual Outline", type="primary", use_container_width=True):
        if not api_key:
            st.error("Please enter your Gemini API Key in the sidebar first.")
        else:
            file_text = extract_text_from_uploaded_file(uploaded_report)
            combined_text = "\n\n".join(filter(None, [file_text.strip(), pasted_text.strip()]))

            if not combined_text:
                st.warning("Please upload a document or paste text first.")
            else:
                with st.spinner(f"Extracting embedded images & designing multi-layout slide deck with {active_model}..."):
                    try:
                        st.session_state.extracted_report_images = extract_images_from_report(uploaded_report)
                        st.session_state.slides_editor_data = generate_smart_outline(
                            api_key=api_key,
                            source_text=combined_text,
                            num_slides=num_slides,
                            tone=tone,
                            language=output_lang,
                            model_name=active_model,
                        )
                        st.session_state.pptx_bytes = None
                        img_count = len(st.session_state.extracted_report_images)
                        st.success(
                            f"Smart outline generated using {active_model}! Extracted {img_count} embedded image(s) from your uploaded file."
                        )
                    except Exception as e:
                        st.error(f"Failed to generate outline: {e}")

    # Step 2: Interactive Multi-Layout Outline Editor
    if st.session_state.slides_editor_data:
        st.divider()
        st.header("Step 2: Review & Customize Layouts, Charts, and Smart Hybrid Visuals")
        st.info(
            "💡 You can change any slide's **Layout Type**, edit structured points (`Heading :: Detail`), "
            "or customize **Smart Hybrid Visuals** (which automatically choose between Unsplash Stock Photos and AI Illustrations)."
        )

        layout_options = [lt.value for lt in LayoutType]
        report_imgs = st.session_state.extracted_report_images

        if report_imgs:
            with st.expander(f"🖼️ Preview {len(report_imgs)} Embedded Image(s) Extracted from Your Report"):
                img_cols = st.columns(min(len(report_imgs), 4))
                for r_idx, r_img in enumerate(report_imgs):
                    with img_cols[r_idx % 4]:
                        st.image(r_img["bytes"], caption=f"[{r_idx}] {r_img['name']}", use_container_width=True)

        updated_slides = []

        for idx, slide in enumerate(st.session_state.slides_editor_data):
            current_layout = slide.get("layout_type", LayoutType.BULLET_CARDS.value)
            expander_label = f"Slide {idx + 1} [{current_layout.upper()}] — {slide.get('action_title', '')}"

            with st.expander(expander_label, expanded=(idx < 2)):
                top_c1, top_c2 = st.columns([1, 3])
                with top_c1:
                    new_layout = st.selectbox(
                        "Slide Layout Type",
                        options=layout_options,
                        index=layout_options.index(current_layout) if current_layout in layout_options else 2,
                        key=f"layout_{idx}",
                    )
                with top_c2:
                    new_title = st.text_input(
                        "Action Title (Conclusion-First)",
                        value=slide.get("action_title", ""),
                        key=f"title_{idx}",
                    )

                slide["layout_type"] = new_layout
                slide["action_title"] = new_title

                # Dynamic Editor Fields Based on Selected Layout
                if new_layout in (LayoutType.COVER.value, LayoutType.SECTION_DIVIDER.value):
                    slide["subtitle_or_context"] = st.text_area(
                        "Subtitle / Section Overview",
                        value=slide.get("subtitle_or_context") or "",
                        height=80,
                        key=f"sub_{idx}",
                    )

                elif new_layout in (LayoutType.BULLET_CARDS.value, LayoutType.THREE_CARDS.value):
                    hint = "Enter 3 items (Heading :: Detail)" if new_layout == "three_cards" else "Enter 3-4 items (Heading :: Detail)"
                    raw_pts = st.text_area(
                        f"Structured Points — One per line formatted as `{hint}`",
                        value=format_points_for_editor(slide.get("points", [])),
                        height=120,
                        key=f"pts_{idx}",
                    )
                    slide["points"] = parse_points_from_editor(raw_pts)

                elif new_layout == LayoutType.TWO_COLUMN.value:
                    lc1, lc2 = st.columns(2)
                    left_col = slide.get("left_column") or {"column_title": "Left Column", "items": []}
                    right_col = slide.get("right_column") or {"column_title": "Right Column", "items": []}

                    with lc1:
                        l_title = st.text_input("Left Column Header", value=left_col.get("column_title", ""), key=f"l_title_{idx}")
                        l_pts = st.text_area(
                            "Left Column Points (`Heading :: Detail`)",
                            value=format_points_for_editor(left_col.get("items", [])),
                            height=110,
                            key=f"l_pts_{idx}",
                        )
                    with lc2:
                        r_title = st.text_input("Right Column Header", value=right_col.get("column_title", ""), key=f"r_title_{idx}")
                        r_pts = st.text_area(
                            "Right Column Points (`Heading :: Detail`)",
                            value=format_points_for_editor(right_col.get("items", [])),
                            height=110,
                            key=f"r_pts_{idx}",
                        )
                    slide["left_column"] = {"column_title": l_title, "items": parse_points_from_editor(l_pts)}
                    slide["right_column"] = {"column_title": r_title, "items": parse_points_from_editor(r_pts)}

                elif new_layout == LayoutType.KPI_METRICS.value:
                    raw_kpis = st.text_area(
                        "Hero KPI Metrics — One per line formatted as `Value | Label | Context` (e.g. `+62% | Latency Drop | Across Q3 peak loads`)",
                        value=format_kpis_for_editor(slide.get("kpis", [])),
                        height=110,
                        key=f"kpis_{idx}",
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
                            key=f"ctype_{idx}",
                        )
                        new_ctitle = st.text_input("Chart Title", value=chart_obj.get("chart_title", ""), key=f"ctitle_{idx}")
                        new_sname = st.text_input("Series Name & Unit", value=chart_obj.get("series_name", "Value"), key=f"csname_{idx}")
                    with cc2:
                        cats_str = st.text_input(
                            "Categories (comma-separated)",
                            value=", ".join(chart_obj.get("categories", ["Q1", "Q2", "Q3"])),
                            key=f"ccats_{idx}",
                        )
                        vals_str = st.text_input(
                            "Values (comma-separated numbers)",
                            value=", ".join(str(v) for v in chart_obj.get("values", [10.0, 20.0, 30.0])),
                            key=f"cvals_{idx}",
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
                        key=f"cpts_{idx}",
                    )
                    slide["points"] = parse_points_from_editor(raw_c_pts)

                elif new_layout == LayoutType.IMAGE_SPLIT.value:
                    img_obj = slide.get("image") or {}
                    ai_rec_style = img_obj.get("preferred_style", VisualStyle.AI_ILLUSTRATION.value)
                    ic1, ic2 = st.columns(2)

                    with ic1:
                        src_choices = [
                            "Smart Hybrid (Auto)",
                            "Unsplash Stock Photo",
                            "AI Generated Illustration",
                            "Upload Custom Image",
                            "Visual Placeholder Only",
                        ]
                        if report_imgs:
                            src_choices.insert(1, "Use Image from Report")

                        cur_src = slide.get("image_source_mode", "Smart Hybrid (Auto)")
                        if cur_src not in src_choices:
                            cur_src = src_choices[0]

                        chosen_src = st.selectbox(
                            f"Image Source (Gemini Auto-Recommendation: `{ai_rec_style}`)",
                            options=src_choices,
                            index=src_choices.index(cur_src),
                            key=f"imgsrc_{idx}",
                        )
                        slide["image_source_mode"] = chosen_src

                        place_opts = ["left", "right"]
                        cur_place = img_obj.get("placement", "left")
                        new_place = st.selectbox(
                            "Image Placement on Slide",
                            options=place_opts,
                            index=place_opts.index(cur_place) if cur_place in place_opts else 0,
                            key=f"imgplace_{idx}",
                        )

                        new_kw = st.text_input(
                            "Unsplash Search Keywords (English)",
                            value=img_obj.get("search_keywords", ""),
                            key=f"ukw_{idx}",
                        )
                        new_vprompt = st.text_area(
                            "AI Illustration Prompt (English)",
                            value=img_obj.get("visual_prompt", ""),
                            height=75,
                            key=f"vprompt_{idx}",
                        )

                        if chosen_src == "Use Image from Report" and report_imgs:
                            rep_labels = [f"[{i}] {r['name']}" for i, r in enumerate(report_imgs)]
                            sel_rep = st.selectbox(
                                "Select Extracted Report Image",
                                options=range(len(report_imgs)),
                                format_func=lambda x: rep_labels[x],
                                key=f"repimg_{idx}",
                            )
                            slide["selected_report_image_idx"] = sel_rep

                        if chosen_src == "Upload Custom Image":
                            up_custom_img = st.file_uploader(
                                "Upload Slide Image (.png, .jpg, .webp)",
                                type=["png", "jpg", "jpeg", "webp"],
                                key=f"upimg_{idx}",
                            )
                            if up_custom_img is not None:
                                slide["custom_uploaded_image_bytes"] = up_custom_img.getvalue()

                    with ic2:
                        style_opts = [VisualStyle.STOCK_PHOTO.value, VisualStyle.AI_ILLUSTRATION.value]
                        new_pref_style = st.selectbox(
                            "Smart Hybrid Preferred Style",
                            options=style_opts,
                            index=style_opts.index(ai_rec_style) if ai_rec_style in style_opts else 1,
                            key=f"prefstyle_{idx}",
                            help="In 'Smart Hybrid (Auto)' mode, 'stock_photo' queries Unsplash first (then falls back to AI), while 'ai_illustration' generates a custom graphic directly.",
                        )
                        new_caption = st.text_input(
                            "Image Caption",
                            value=img_obj.get("caption", ""),
                            key=f"imgcap_{idx}",
                        )
                        raw_i_pts = st.text_area(
                            "Accompanying Key Points (`Heading :: Detail`)",
                            value=format_points_for_editor(slide.get("points", [])),
                            height=135,
                            key=f"ipts_{idx}",
                        )
                        slide["points"] = parse_points_from_editor(raw_i_pts)

                    slide["image"] = {
                        "preferred_style": new_pref_style,
                        "visual_prompt": new_vprompt,
                        "search_keywords": new_kw,
                        "caption": new_caption,
                        "placement": new_place,
                    }

                # Bottom Takeaway Editor for all content layouts
                if new_layout not in (LayoutType.COVER.value, LayoutType.SECTION_DIVIDER.value):
                    slide["bottom_takeaway"] = st.text_input(
                        "💡 Bottom Takeaway Banner ('So What?')",
                        value=slide.get("bottom_takeaway") or "",
                        key=f"takeaway_{idx}",
                    )

                updated_slides.append(slide)

        st.session_state.slides_editor_data = updated_slides

        # Step 3: Render & Download
        st.divider()
        st.header("Step 3: Render Visual PowerPoint (.pptx)")

        if st.button("🚀 Render Complete Presentation (.pptx)", type="primary", use_container_width=True):
            if template_mode == "Upload Custom .pptx Template" and uploaded_template is None:
                st.error("Please upload a .pptx template file in the sidebar, or switch to 'Built-in 16:9 Color Theme'.")
            else:
                with st.spinner("Rendering layouts, native charts, and Smart Hybrid visuals into PowerPoint..."):
                    try:
                        tpl_bytes = uploaded_template.getvalue() if uploaded_template else None
                        st.session_state.pptx_bytes = build_presentation_pptx(
                            slides_data=st.session_state.slides_editor_data,
                            theme_name=selected_theme,
                            api_key=api_key,
                            use_imagen=use_imagen,
                            unsplash_access_key=unsplash_key,
                            extracted_report_images=st.session_state.extracted_report_images,
                            custom_template_bytes=tpl_bytes,
                        )
                        st.success("Presentation rendered successfully! Click below to download.")
                    except Exception as e:
                        st.error(f"Error rendering PPTX: {e}")

        if st.session_state.pptx_bytes:
            st.download_button(
                label="📥 Download ChatPPT_V2_Presentation.pptx",
                data=st.session_state.pptx_bytes,
                file_name="ChatPPT_V2_Presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )


if __name__ == "__main__":
    main()