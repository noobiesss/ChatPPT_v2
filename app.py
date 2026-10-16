import copy
import io
import json
import re
import time
import uuid
import urllib.parse
import urllib.request
import urllib.error
import xml.etree.ElementTree as ET
from enum import Enum
from typing import List, Dict, Optional, Tuple

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
# 1. Enumerations & Pydantic Schemas for Structured Output
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
# 2. Multi-Provider LLM Presets & Built-in 16:9 Color Themes
# =====================================================================

LLM_PROVIDERS: Dict[str, Dict] = {
    "Google Gemini": {
        "base_url": None,  # Handled natively via google-genai SDK
        "models": [
            "gemini-3.8-flash",
            "gemini-3.5-flash",
            "gemini-3.1-flash-lite",
            "gemini-2.5-flash",
            "Custom Model ID...",
        ],
        "key_label": "Google Gemini API Key",
        "key_help": "Enter your Google AI Studio / Vertex Gemini API Key.",
    },
    "Groq (Fast & Free Tier)": {
        "base_url": "https://api.groq.com/openai/v1/chat/completions",
        "models": [
            "llama-3.3-70b-versatile",
            "openai/gpt-oss-120b",
            "meta-llama/llama-4-scout-17b-16e-instruct",
            "qwen/qwen3-32b",
            "llama-3.1-8b-instant",
            "Custom Model ID...",
        ],
        "key_label": "Groq API Key (gsk_...)",
        "key_help": "Get a high-speed key from console.groq.com.",
    },
    "OpenAI (Paid)": {
        "base_url": "https://api.openai.com/v1/chat/completions",
        "models": [
            "gpt-4o",
            "gpt-4o-mini",
            "gpt-4.1",
            "gpt-4.1-mini",
            "Custom Model ID...",
        ],
        "key_label": "OpenAI API Key (sk-...)",
        "key_help": "Enter your OpenAI Platform API Key.",
    },
    "DeepSeek (Paid)": {
        "base_url": "https://api.deepseek.com/chat/completions",
        "models": [
            "deepseek-chat",
            "deepseek-reasoner",
            "Custom Model ID...",
        ],
        "key_label": "DeepSeek API Key",
        "key_help": "Enter your DeepSeek Platform API Key.",
    },
    "OpenRouter (All-in-One)": {
        "base_url": "https://openrouter.ai/api/v1/chat/completions",
        "models": [
            "anthropic/claude-3.7-sonnet",
            "google/gemini-2.5-pro",
            "meta-llama/llama-3.3-70b-instruct",
            "deepseek/deepseek-chat",
            "Custom Model ID...",
        ],
        "key_label": "OpenRouter API Key (sk-or-...)",
        "key_help": "Access Claude, Gemini, Llama, and DeepSeek with one OpenRouter key.",
    },
    "Custom OpenAI-Compatible API": {
        "base_url": "https://api.openai.com/v1/chat/completions",
        "models": ["Custom Model ID..."],
        "key_label": "Custom Provider API Key",
        "key_help": "Works with SiliconFlow, Together AI, Fireworks, vLLM, or Ollama.",
    },
}

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
# 4. Multi-Provider LLM Outline Router & Sanitizer
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


# ---------------------------------------------------------------------
# Live Model Discovery (avoids stale hard-coded model lists)
# ---------------------------------------------------------------------

# Substrings that identify non-chat models (speech, embeddings, images, moderation...)
NON_CHAT_KEYWORDS = (
    "whisper", "tts", "audio", "speech", "transcribe", "embed", "embedding",
    "guard", "moderation", "dall-e", "imagen", "veo", "image", "realtime",
    "orpheus", "aqa", "rerank", "babbage", "davinci", "live", "robotics",
    "computer-use", "search-preview", "omni-moderation",
)


def _is_chat_model(model_id: str) -> bool:
    lowered = model_id.lower()
    return not any(kw in lowered for kw in NON_CHAT_KEYWORDS)


def _derive_models_url(chat_url: str) -> str:
    """Turn a '.../chat/completions' endpoint into its sibling '.../models' endpoint."""
    url = chat_url.strip().rstrip("/")
    if url.endswith("/chat/completions"):
        return url[: -len("/chat/completions")] + "/models"
    return url + "/models"


@st.cache_data(ttl=600, show_spinner=False)
def fetch_live_models(provider: str, api_key: str, base_url: str) -> List[str]:
    """
    Query the provider's model-listing endpoint and return chat-capable model IDs.
    Raises on any failure (st.cache_data never caches exceptions), so callers can fall back.
    """
    api_key = (api_key or "").strip()

    if provider == "Google Gemini":
        client = genai.Client(api_key=api_key)
        ids = []
        for m in client.models.list():
            actions = getattr(m, "supported_actions", None) or []
            name = (getattr(m, "name", "") or "").replace("models/", "")
            if "generateContent" in actions and name.startswith("gemini") and _is_chat_model(name):
                ids.append(name)
        return sorted(set(ids), reverse=True)

    req = urllib.request.Request(
        _derive_models_url(base_url),
        headers={
            "Authorization": f"Bearer {api_key}",
            "User-Agent": "ChatPPT-V2/2.0",
        },
        method="GET",
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        data = json.loads(resp.read().decode("utf-8"))

    items = data.get("data", data) if isinstance(data, dict) else data
    ids = []
    for item in items:
        if not isinstance(item, dict) or not item.get("id"):
            continue
        # Groq flags retired/inactive models via an "active" field
        if item.get("active") is False:
            continue
        model_id = item["id"]
        if not _is_chat_model(model_id):
            continue
        # OpenAI lists many legacy/non-chat families; keep only chat-capable prefixes
        if provider == "OpenAI (Paid)" and not model_id.startswith(("gpt-", "chatgpt-", "o1", "o3", "o4")):
            continue
        ids.append(model_id)
    return sorted(set(ids))


def resolve_model_options(provider: str, api_key: str, base_url: str, static_models: List[str]):
    """
    Return (model_list, status_message). Prefers the live list; falls back to the static preset.
    The static list's order is kept for entries that still exist so sensible defaults stay on top.
    """
    static_clean = [m for m in static_models if m != "Custom Model ID..."]
    needs_key = provider != "OpenRouter (All-in-One)"
    if needs_key and not (api_key or "").strip():
        return static_models, "Enter your API key to load the live model list."
    if provider == "Custom OpenAI-Compatible API" and not (base_url or "").strip():
        return static_models, "Enter an endpoint URL to load models."

    try:
        live = fetch_live_models(provider, api_key, base_url or "")
    except Exception as exc:  # noqa: BLE001 - any network/auth/parse failure falls back
        return static_models, f"Could not fetch live models ({type(exc).__name__}); showing presets."
    if not live:
        return static_models, "Provider returned no chat models; showing presets."

    preferred = [m for m in static_clean if m in live]
    ordered = preferred + [m for m in live if m not in preferred]
    return ordered + ["Custom Model ID..."], f"Loaded {len(live)} live models from {provider}."


def extract_clean_json_str(raw_text: str) -> str:
    """
    Strip <think>...</think> reasoning blocks and Markdown code fences from LLM responses
    so Pydantic can reliably parse JSON across DeepSeek, Qwen, Llama, and OpenAI models.
    """
    cleaned = re.sub(r"<think>.*?</think>", "", raw_text, flags=re.DOTALL).strip()
    fence_match = re.search(r"```(?:json)?\s*(\{.*\})\s*```", cleaned, flags=re.DOTALL)
    if fence_match:
        return fence_match.group(1).strip()

    first_brace = cleaned.find("{")
    last_brace = cleaned.rfind("}")
    if first_brace != -1 and last_brace != -1 and last_brace > first_brace:
        return cleaned[first_brace:last_brace + 1]
    return cleaned


def _parses_as_outline(text: str) -> bool:
    """Return True if the text is valid JSON that satisfies the PresentationOutline schema."""
    try:
        PresentationOutline.model_validate_json(text)
        return True
    except Exception:  # noqa: BLE001
        return False


def repair_json_str(raw_text: str) -> str:
    """
    Best-effort repair for slightly malformed LLM JSON (the most common failures are a stray
    closing brace/bracket or missing closers at the end). Returns the original text unchanged
    if no repair produces a schema-valid outline.
    """
    if _parses_as_outline(raw_text):
        return raw_text

    # Scan once, ignoring characters inside string literals
    stack: List[str] = []
    closer_positions: List[int] = []
    in_string = False
    escaped = False
    for idx, ch in enumerate(raw_text):
        if in_string:
            if escaped:
                escaped = False
            elif ch == "\\":
                escaped = True
            elif ch == '"':
                in_string = False
            continue
        if ch == '"':
            in_string = True
        elif ch in "{[":
            stack.append("}" if ch == "{" else "]")
        elif ch in "}]":
            closer_positions.append(idx)
            if stack and stack[-1] == ch:
                stack.pop()

    # Strategy 1: truncated output -> close any unterminated string/containers
    candidate = raw_text + ('"' if in_string else "") + "".join(reversed(stack))
    if _parses_as_outline(candidate):
        return candidate

    # Strategy 2: a single stray closing brace/bracket -> try removing each one
    for pos in closer_positions:
        candidate = raw_text[:pos] + raw_text[pos + 1:]
        if _parses_as_outline(candidate):
            return candidate

    return raw_text


def call_openai_compatible_json(
    base_url: str,
    api_key: str,
    model_name: str,
    system_prompt: str,
    user_prompt: str,
) -> str:
    """
    Zero-dependency REST client for OpenAI-compatible chat completion endpoints
    (Groq, OpenAI, DeepSeek, OpenRouter, SiliconFlow, etc.).
    Automatically retries without response_format if a reasoning model rejects json_object mode.
    """
    schema_json = json.dumps(PresentationOutline.model_json_schema(), ensure_ascii=False)
    full_system_prompt = (
        f"{system_prompt}\n\n"
        "### OUTPUT FORMAT REQUIREMENT:\n"
        "You MUST respond ONLY with a valid JSON object that strictly conforms to the following JSON Schema. "
        "Do NOT wrap the JSON in markdown explanations or add extra text outside the JSON object.\n"
        f"JSON Schema:\n{schema_json}"
    )

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key.strip()}",
        "HTTP-Referer": "https://chatppt-v2.streamlit.app",
        "X-Title": "ChatPPT V2",
        "User-Agent": "ChatPPT-V2/2.0",
    }

    for use_json_mode in (True, False):
        payload = {
            "model": model_name,
            "messages": [
                {"role": "system", "content": full_system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "temperature": 0.3,
        }
        if use_json_mode:
            payload["response_format"] = {"type": "json_object"}

        req = urllib.request.Request(
            base_url.strip(),
            data=json.dumps(payload).encode("utf-8"),
            headers=headers,
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=90) as resp:
                resp_data = json.loads(resp.read().decode("utf-8"))
                content = resp_data["choices"][0]["message"]["content"]
                return extract_clean_json_str(content)
        except urllib.error.HTTPError as http_err:
            err_body = http_err.read().decode("utf-8", errors="ignore")
            lowered_err = err_body.lower()
            if http_err.code == 400:
                # Groq-style "json_validate_failed": the server rejected the model's malformed JSON
                # but returns the raw text in `failed_generation`; try to repair it locally.
                try:
                    failed_text = json.loads(err_body).get("error", {}).get("failed_generation")
                except (ValueError, AttributeError):
                    failed_text = None
                if isinstance(failed_text, str) and failed_text.strip():
                    repaired = repair_json_str(extract_clean_json_str(failed_text))
                    if _parses_as_outline(repaired):
                        return repaired
                # Model rejected json_object mode or produced invalid JSON in strict mode:
                # retry once without response_format and let the local parser/repair handle it.
                if use_json_mode and (
                    "response_format" in lowered_err
                    or "json_validate_failed" in lowered_err
                    or "failed to generate json" in lowered_err
                ):
                    continue
            raise RuntimeError(f"Provider API HTTP {http_err.code}: {err_body}") from http_err


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


# Extra attempts when the model returns a different number of slides than requested
OUTLINE_COUNT_MAX_RETRIES = 2

# Matches slide-count mentions in free-text revision notes, e.g. "6 slides", "add 2 more pages", "8页"
_SLIDE_COUNT_MENTION_RE = re.compile(
    r"\d+\s*(?:more\s+|extra\s+|additional\s+|new\s+)?(?:slides?|pages?)|\d+\s*[\u9875\u9801\u5f20\u5f35]",
    re.IGNORECASE,
)


def mentions_slide_count(text: str) -> bool:
    """Return True if a revision instruction itself talks about a slide/page count."""
    return bool(_SLIDE_COUNT_MENTION_RE.search(text or ""))


def _request_outline(
    provider: str,
    api_key: str,
    model_name: str,
    user_prompt: str,
    custom_base_url: Optional[str],
) -> PresentationOutline:
    """Send one outline request to the selected provider and return the validated outline."""
    if provider == "Google Gemini":
        client = genai.Client(api_key=api_key)
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
        raw_json = extract_clean_json_str(response.text)
    else:
        endpoint_url = custom_base_url or LLM_PROVIDERS[provider]["base_url"]
        raw_json = call_openai_compatible_json(
            base_url=endpoint_url,
            api_key=api_key,
            model_name=model_name,
            system_prompt=SYSTEM_INSTRUCTION,
            user_prompt=user_prompt,
        )
    return PresentationOutline.model_validate_json(repair_json_str(raw_json))


def _build_count_correction(outline: PresentationOutline, num_slides: int) -> str:
    """Build a corrective note that tells the model exactly how its previous outline missed the target."""
    got = len(outline.slides)
    titles = "\n".join(f"  {i + 1}. {s.action_title}" for i, s in enumerate(outline.slides))
    if got < num_slides:
        fix = f"Add {num_slides - got} more slide(s) covering other important parts of the source document."
    else:
        fix = f"Merge or remove {got - num_slides} slide(s), keeping the most important content."
    return (
        "\n\nCORRECTION REQUIRED: your previous answer contained "
        f"{got} slides, but EXACTLY {num_slides} slides are required.\n"
        f"Previous slide titles:\n{titles}\n"
        f"{fix} Return the COMPLETE outline again (not only the changes) with exactly {num_slides} "
        "objects in the `slides` array. Keep every text field short so the answer is not cut off."
    )


def _trim_slides_to_count(slides: List[SlideSpec], target: int) -> List[SlideSpec]:
    """
    Deterministically shorten an over-long deck: drop section dividers first (never the cover),
    then drop slides right before the closing slide so the cover and the conclusion survive.
    """
    trimmed = list(slides)
    for idx in range(len(trimmed) - 1, 0, -1):
        if len(trimmed) <= target:
            break
        if trimmed[idx].layout_type == LayoutType.SECTION_DIVIDER:
            del trimmed[idx]
    while len(trimmed) > target and len(trimmed) > 2:
        del trimmed[-2]
    return trimmed


def generate_smart_outline(
    provider: str,
    api_key: str,
    source_text: str,
    num_slides: int,
    tone: str,
    language: str,
    model_name: str,
    custom_base_url: Optional[str] = None,
    enforce_count: bool = True,
) -> Tuple[List[dict], Optional[str]]:
    """
    Multi-Provider Outline Generator:
    Routes request to either Google Gemini SDK or any OpenAI-Compatible endpoint (Groq, OpenAI, DeepSeek, etc.)
    and validates the returned JSON against PresentationOutline.

    When enforce_count is True the slide count is verified: the model is asked again (up to
    OUTLINE_COUNT_MAX_RETRIES times) if it misses the target, and an over-long result is trimmed.
    Returns (slides, warning). `warning` is a user-facing note if the final count still differs
    from the request, otherwise None.
    """
    if enforce_count:
        count_rule = (
            f"- Total Slides: Exactly {num_slides} slides (Slide 1 MUST be 'cover'). "
            f"The `slides` array MUST contain exactly {num_slides} objects numbered 1 to {num_slides}; "
            "count them before you answer."
        )
    else:
        count_rule = (
            "- Total Slides: if the revision notes at the end of the source document ask for a specific "
            f"slide count, follow them; otherwise use {num_slides} slides. Slide 1 MUST be 'cover'."
        )

    base_prompt = f"""
    Create a {num_slides}-slide presentation outline based on the source document below.

    Target Specifications:
    {count_rule}
    - Presentation Tone: {tone}.
    - Output Language: {language} (Write all titles, headings, details, chart labels, captions, and takeaways in {language};
      keep ONLY `visual_prompt` and `search_keywords` in English).

    Source Document:
    \"\"\"
    {source_text}
    \"\"\"
    """

    outline = _request_outline(provider, api_key, model_name, base_prompt, custom_base_url)
    warning: Optional[str] = None

    if enforce_count and len(outline.slides) != num_slides:
        best = outline
        current = outline
        for _ in range(OUTLINE_COUNT_MAX_RETRIES):
            try:
                current = _request_outline(
                    provider,
                    api_key,
                    model_name,
                    base_prompt + _build_count_correction(current, num_slides),
                    custom_base_url,
                )
            except Exception:  # noqa: BLE001 - keep the best earlier result instead of failing the whole job
                break
            if abs(len(current.slides) - num_slides) < abs(len(best.slides) - num_slides):
                best = current
            if len(current.slides) == num_slides:
                break
        outline = best

        got = len(outline.slides)
        if got > num_slides:
            outline = outline.model_copy(update={"slides": _trim_slides_to_count(outline.slides, num_slides)})
            warning = (
                f"The model kept returning {got} slides, so I trimmed the deck to the requested "
                f"{num_slides} (section dividers and middle slides were dropped first)."
            )
        elif got < num_slides:
            warning = (
                f"The model returned only {got} of the {num_slides} requested slides, even after retrying. "
                "Try a larger model, or ask for more with a follow-up such as “add 2 more slides”."
            )

    return sanitize_outline_dict(outline), warning


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
    gemini_api_key: Optional[str] = None,
    use_imagen: bool = False
) -> Optional[bytes]:
    """
    Generate a 4:3 conceptual illustration for a slide.
    Tries Google Imagen 3 first (if enabled and a Gemini API key is available),
    and falls back to Pollinations AI stream.
    """
    enhanced_prompt = (
        f"{visual_prompt}, clean modern corporate presentation illustration, "
        "minimalist high-contrast aesthetic, no text, no watermarks"
    )

    if use_imagen and gemini_api_key:
        try:
            client = genai.Client(api_key=gemini_api_key)
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
    gemini_api_key: Optional[str],
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
        return generate_ai_image_bytes(visual_prompt, gemini_api_key=gemini_api_key, use_imagen=use_imagen)

    # 5. Explicit AI Generated Illustration
    if src_mode == "AI Generated Illustration":
        return generate_ai_image_bytes(visual_prompt, gemini_api_key=gemini_api_key, use_imagen=use_imagen)

    # 6. Smart Hybrid (Auto) — Follows LLM's per-slide `preferred_style` with cross-fallback
    if preferred_style == VisualStyle.STOCK_PHOTO.value and unsplash_access_key:
        img_bytes = fetch_unsplash_image_bytes(search_keywords, unsplash_access_key)
        if img_bytes:
            return img_bytes
        return generate_ai_image_bytes(visual_prompt, gemini_api_key=gemini_api_key, use_imagen=use_imagen)
    else:
        img_bytes = generate_ai_image_bytes(visual_prompt, gemini_api_key=gemini_api_key, use_imagen=use_imagen)
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
# 6. Multi-Layout PowerPoint Rendering Engine & Template Inheritance
# =====================================================================

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
    }
    try:
        slide_w = float(prs.slide_width or Inches(13.333))
        slide_h = float(prs.slide_height or Inches(7.5))
        slide_area = max(slide_w * slide_h, 1.0)

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

                    if info["font"] is None and getattr(shape, "has_text_frame", False):
                        for para in shape.text_frame.paragraphs:
                            for run in para.runs:
                                fname = (run.font.name or "").strip()
                                if fname and not fname.startswith("+"):
                                    info["font"] = fname
                                    break
                except Exception:
                    continue
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

    resolved_heading = font_heading or sampled.get("font")
    resolved_body = font_body or resolved_heading
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
        dk1 = scheme_colors.get("dk1", (15, 37, 55))
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
        takeaway_bg = _blend_rgb(content_bg_tuple, accent_tuple, 0.28)
        takeaway_text = RGBColor(255, 255, 255)
    else:
        dk_text = scheme_colors.get("dk1") or (22, 33, 48)
        if _luminance(dk_text) > 140:
            dk_text = (22, 33, 48)
        card_bg = RGBColor(255, 255, 255)
        card_border = _blend_rgb((255, 255, 255), accent_tuple, 0.25)
        title_content = RGBColor(*dk_text)
        body_content = _blend_rgb(dk_text, (120, 128, 140), 0.35)
        takeaway_bg = RGBColor(*dk_text) if _luminance(dk_text) < 95 else _blend_rgb(accent_tuple, (15, 23, 42), 0.45)
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


def _is_decorative_or_branding_shape(shape, slide_h_emu: int) -> bool:
    """
    Return True if a shape on a sample template slide is part of the visual template
    (background image, decorative vector graphic, border frame, logo, or small header/footer tag)
    rather than sample body text.
    """
    if getattr(shape, "is_placeholder", False):
        return False
    if getattr(shape, "has_table", False) or getattr(shape, "has_chart", False):
        return False

    # Pictures without text (background art, corner logos, decorative illustrations)
    if hasattr(shape, "image"):
        return True

    if not getattr(shape, "has_text_frame", False):
        return True

    txt = (shape.text_frame.text or "").strip()
    if not txt:
        return True

    # Keep short header/footer branding labels near the top or bottom edge of the slide
    top = int(shape.top or 0)
    height = int(shape.height or 0)
    near_top = top < int(slide_h_emu * 0.08)
    near_bottom = (top + height) > int(slide_h_emu * 0.90)
    if (near_top or near_bottom) and len(txt) <= 32 and "\n" not in txt:
        return True

    return False


def _snapshot_slide_prototype(slide, slide_h_emu: int) -> dict:
    """
    Capture a prototype slide's layout reference, custom <p:bg> XML, relationships, and
    decorative/background shapes (pictures & non-text vector shapes) before slides are cleared.
    """
    c_sld = slide._element.find(f"{_NS_P}cSld")
    bg_el = c_sld.find(f"{_NS_P}bg") if c_sld is not None else None
    rels_map = {r_id: rel for r_id, rel in slide.part.rels.items()}

    decor_items = []
    for shape in slide.shapes:
        if not _is_decorative_or_branding_shape(shape, slide_h_emu):
            continue
        try:
            if hasattr(shape, "image"):
                decor_items.append({
                    "kind": "picture",
                    "blob": shape.image.blob,
                    "left": int(shape.left or 0),
                    "top": int(shape.top or 0),
                    "width": int(shape.width or 0),
                    "height": int(shape.height or 0),
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
    ('cover', 'section', 'content') so their backgrounds, frames, and logos survive slide clearing.
    """
    prototypes: Dict[str, dict] = {}
    slides = list(prs.slides)
    if not slides:
        return prototypes

    slide_h_emu = int(prs.slide_height or Inches(7.5))
    prototypes["cover"] = _snapshot_slide_prototype(slides[0], slide_h_emu)

    if len(slides) == 1:
        prototypes["content"] = prototypes["cover"]
        prototypes["section"] = prototypes["cover"]
        return prototypes

    # Pick a representative content slide (prefer Slide 2 or a slide with multiple shapes)
    content_slide = slides[1]
    for s in slides[1:]:
        lname = (s.slide_layout.name or "").lower()
        if "section" in lname or "divider" in lname:
            prototypes.setdefault("section", _snapshot_slide_prototype(s, slide_h_emu))
        elif "content" in lname or "title" in lname:
            content_slide = s
            break

    prototypes["content"] = _snapshot_slide_prototype(content_slide, slide_h_emu)
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
    and replay any captured prototype background (<p:bg>) and decorative/branding shapes.
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

    # 2. Replay decorative background images, vector frames, and corner logos
    for item in proto.get("decor_items") or []:
        try:
            if item["kind"] == "picture":
                slide.shapes.add_picture(
                    io.BytesIO(item["blob"]),
                    item["left"],
                    item["top"],
                    width=item["width"],
                    height=item["height"],
                )
            elif item["kind"] == "xml":
                new_el = copy.deepcopy(item["el"])
                _remap_embed_rids(new_el, old_rels, slide.part)
                slide.shapes._spTree.append(new_el)
        except Exception:
            continue

    return slide


def apply_slide_fonts(slide, theme: Dict[str, object]) -> None:
    """Apply the template's heading and body fonts across all text frames on a rendered slide."""
    font_heading = theme.get("font_heading")
    font_body = theme.get("font_body") or font_heading
    if not font_heading and not font_body:
        return

    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        for para in shape.text_frame.paragraphs:
            is_heading = bool(para.font.bold) or (para.font.size is not None and para.font.size >= Pt(16))
            target_font = font_heading if is_heading else font_body
            if target_font:
                para.font.name = target_font
                for run in para.runs:
                    run.font.name = target_font


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


def draw_card_shape(slide, left: float, top: float, width: float, height: float, theme: Dict[str, object]):
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
    theme: Dict[str, object],
    apply_bg: bool = True
) -> None:
    """Render the top accent bar, Action Title, bottom takeaway banner, and slide number."""
    if apply_bg:
        set_solid_background(slide, theme["bg_content"])
        # Top accent strip (only drawn on built-in themes so custom template top banners aren't covered)
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


def render_cover_slide(slide, slide_data: dict, theme: Dict[str, object], apply_bg: bool = True) -> None:
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
    p_title.font.color.rgb = theme["title_cover"]

    subtitle = (slide_data.get("subtitle_or_context") or "").strip()
    if subtitle:
        sub_box = slide.shapes.add_textbox(Inches(1.1), Inches(4.8), Inches(10.5), Inches(1.4))
        tf_sub = sub_box.text_frame
        tf_sub.word_wrap = True
        p_sub = tf_sub.paragraphs[0]
        p_sub.text = subtitle
        p_sub.font.size = Pt(20)
        p_sub.font.color.rgb = theme["subtitle_cover"]


def render_section_divider_slide(slide, slide_data: dict, theme: Dict[str, object], apply_bg: bool = True) -> None:
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
    p_title.font.color.rgb = theme["title_cover"]

    context = (slide_data.get("subtitle_or_context") or "").strip()
    if context:
        sub_box = slide.shapes.add_textbox(Inches(1.1), Inches(4.5), Inches(10.0), Inches(1.2))
        tf_sub = sub_box.text_frame
        tf_sub.word_wrap = True
        p_sub = tf_sub.paragraphs[0]
        p_sub.text = context
        p_sub.font.size = Pt(18)
        p_sub.font.color.rgb = theme["subtitle_cover"]


def render_bullet_cards_slide(slide, slide_data: dict, theme: Dict[str, object]) -> None:
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


def render_three_cards_slide(slide, slide_data: dict, theme: Dict[str, object]) -> None:
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


def render_two_column_slide(slide, slide_data: dict, theme: Dict[str, object]) -> None:
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


def render_kpi_metrics_slide(slide, slide_data: dict, theme: Dict[str, object]) -> None:
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


def render_chart_slide(slide, slide_data: dict, theme: Dict[str, object]) -> None:
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
    theme: Dict[str, object],
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
    gemini_api_key: Optional[str],
    use_imagen: bool,
    unsplash_access_key: str,
    extracted_report_images: List[Dict[str, bytes]],
    custom_template_bytes: Optional[bytes] = None,
    use_template_colors: bool = True,
) -> bytes:
    """
    Assemble the complete .pptx presentation across all 8 visual layouts.
    Supports both Built-in Color Themes and Custom Uploaded .pptx Templates (inheriting
    Slide Masters, role-specific layouts, prototype slide artwork/logos, theme colors, and fonts).
    """
    prototypes: Dict[str, dict] = {}
    is_custom_template = bool(custom_template_bytes)

    if is_custom_template:
        prs = Presentation(io.BytesIO(custom_template_bytes))
        _scale_template_to_widescreen(prs, 13.333, 7.5)
        theme = extract_template_style(
            prs,
            fallback_theme_name=theme_name,
            use_template_colors=use_template_colors,
        )
        prototypes = _capture_template_prototypes(prs)
        clear_existing_slides(prs)
        apply_solid_bg = False  # Preserve custom template's master, layout, and prototype backgrounds
    else:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        theme = dict(BUILTIN_THEMES.get(theme_name, BUILTIN_THEMES["Corporate Blue"]))
        apply_solid_bg = True

    total_slides = len(slides_data)

    for idx, slide_data in enumerate(slides_data):
        layout = slide_data.get("layout_type", LayoutType.BULLET_CARDS.value)
        if layout == LayoutType.COVER.value:
            role = "cover"
        elif layout == LayoutType.SECTION_DIVIDER.value:
            role = "section"
        else:
            role = "content"

        slide = _create_slide_with_template(prs, role, prototypes, is_custom_template)

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
                    gemini_api_key=gemini_api_key,
                    use_imagen=use_imagen,
                    unsplash_access_key=unsplash_access_key,
                    extracted_report_images=extracted_report_images,
                )
                render_image_split_slide(slide, slide_data, theme, img_bytes=img_bytes)

            else:
                render_bullet_cards_slide(slide, slide_data, theme)

        apply_slide_fonts(slide, theme)

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
# 8. Chat-Style UI (Sidebar chats, Composer, Settings dialog)
# =====================================================================

CUSTOM_PROVIDER = "Custom OpenAI-Compatible API"
OPENROUTER_PROVIDER = "OpenRouter (All-in-One)"

TONE_OPTIONS = {
    "Executive (McKinsey)": "Executive & Consulting (McKinsey Style)",
    "Professional": "Professional & Corporate",
    "Academic": "Academic & Analytical",
    "Venture Pitch": "Persuasive & Venture Pitch",
    "Technical": "Technical Architecture & Engineering",
}
LANGUAGE_OPTIONS = ["English", "Simplified Chinese", "Traditional Chinese", "Malay", "Japanese"]
REVISION_MAX_CHARS = 600  # Short follow-up messages are treated as revision instructions

APP_CSS = """
<style>
/* Theme-neutral styling: text and page colors come from the active Streamlit theme (light or dark).
   Surfaces use translucent grey so contrast is preserved in both themes. */
header[data-testid="stHeader"] { background: transparent; }
.block-container { max-width: 880px; padding-top: 2rem; padding-bottom: 1rem; }

/* Sidebar */
[data-testid="stSidebar"] .stButton > button {
    justify-content: flex-start; text-align: left; border: none; background: transparent;
    box-shadow: none; font-weight: 400; padding: 0.45rem 0.7rem; border-radius: 10px;
}
[data-testid="stSidebar"] .stButton > button:hover { background: rgba(128, 128, 128, 0.16); }
.st-key-new_chat_btn button {
    background: #C96442 !important; color: #FFFFFF !important; font-weight: 600 !important;
    justify-content: center !important;
}
.st-key-new_chat_btn button:hover { background: #B5573A !important; }
.st-key-settings_btn button {
    border: 1px solid rgba(128, 128, 128, 0.40) !important; justify-content: center !important;
}
.sidebar-label {
    font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.08em;
    opacity: 0.6; margin: 1.1rem 0 0.3rem 0.4rem;
}

/* Hero (empty chat state) */
.hero { text-align: center; padding: 3rem 0 1.5rem; }
.hero h1 { font-size: 2.1rem; font-weight: 600; letter-spacing: -0.02em; margin-bottom: 0.4rem; }
.hero p { opacity: 0.7; margin: 0; }
.steps { display: grid; grid-template-columns: repeat(3, 1fr); gap: 0.8rem; margin: 1.6rem 0 0.5rem; }
.step {
    background: rgba(128, 128, 128, 0.09); border: 1px solid rgba(128, 128, 128, 0.28);
    border-radius: 14px; padding: 0.9rem 1rem; text-align: left;
}
.step b { display: block; margin-bottom: 0.2rem; }
.step span { opacity: 0.72; font-size: 0.88rem; }
@media (max-width: 700px) { .steps { grid-template-columns: 1fr; } }

/* Composer dock (sticky at the bottom of the main area) */
.st-key-composer_dock {
    position: sticky; bottom: 0; z-index: 100; padding: 1.5rem 0 0.75rem;
    background: linear-gradient(to top, __DOCK_BG__ 78%, transparent);
}
.st-key-composer {
    background: rgba(128, 128, 128, 0.10); border: 1px solid rgba(128, 128, 128, 0.35);
    border-radius: 20px; padding: 0.7rem 1rem 0.6rem;
}
.st-key-composer textarea { font-size: 1rem; }
.st-key-composer [data-baseweb="textarea"], .st-key-composer [data-baseweb="base-input"] {
    border: none !important; background: transparent !important;
}
.st-key-composer .stButton > button, .st-key-composer [data-testid="stPopover"] > div > button {
    border-radius: 999px; font-size: 0.82rem; padding: 0.25rem 0.6rem; white-space: nowrap;
    overflow: hidden; text-overflow: ellipsis;
}
.st-key-composer [data-baseweb="select"] > div { border-radius: 999px; min-height: 2rem; font-size: 0.82rem; }

/* Active chat highlight (injected per run) */
.st-key-chat_btn___ACTIVE__ button { background: rgba(128, 128, 128, 0.26) !important; font-weight: 600 !important; }
</style>
"""


def _dock_background() -> str:
    """Opaque color for the sticky composer's fade-out, matched to the active Streamlit theme."""
    try:
        theme_type = st.context.theme.type  # "light", "dark" or None (requires Streamlit >= 1.46)
    except Exception:  # noqa: BLE001
        theme_type = None
    if theme_type == "dark":
        return "#0E1117"
    return st.get_option("theme.backgroundColor") or "#FFFFFF"


def inject_css(active_chat_id: str) -> None:
    """Inject the global stylesheet and highlight the active chat button in the sidebar."""
    css = APP_CSS.replace("__ACTIVE__", active_chat_id).replace("__DOCK_BG__", _dock_background())
    st.markdown(css, unsafe_allow_html=True)


# ---------------------------------------------------------------------
# Session state helpers
# ---------------------------------------------------------------------

def create_chat() -> None:
    """Create a new empty chat and make it the active one."""
    ss = st.session_state
    cid = uuid.uuid4().hex[:8]
    ss.chats[cid] = {
        "id": cid,
        "title": "New chat",
        "messages": [],
        "slides": None,
        "report_images": [],
        "pptx_bytes": None,
        "source_text": "",
        "gen": 0,
        "composer_ver": 0,
        "pending": None,
    }
    ss.chat_order.insert(0, cid)
    ss.active_chat = cid


def delete_chat(cid: str) -> None:
    ss = st.session_state
    ss.chats.pop(cid, None)
    if cid in ss.chat_order:
        ss.chat_order.remove(cid)
    if not ss.chats:
        create_chat()
    elif ss.active_chat not in ss.chats:
        ss.active_chat = ss.chat_order[0]


def init_state() -> None:
    ss = st.session_state
    if "settings" not in ss:
        ss.settings = {
            "api_keys": {},
            "custom_base_url": LLM_PROVIDERS[CUSTOM_PROVIDER]["base_url"] or "",
            "unsplash_key": "",
            "use_imagen": False,
        }
    ss.setdefault("custom_tpl_bytes", None)
    ss.setdefault("custom_tpl_name", None)
    if "chats" not in ss:
        ss.chats = {}
        ss.chat_order = []
        ss.active_chat = None
    if not ss.chats:
        create_chat()
    if ss.active_chat not in ss.chats:
        ss.active_chat = ss.chat_order[0]


# ---------------------------------------------------------------------
# Settings dialog (all API keys live here)
# ---------------------------------------------------------------------

@st.dialog("⚙️ Settings", width="large")
def settings_dialog() -> None:
    s = st.session_state.settings
    tab_keys, tab_images, tab_about = st.tabs(["🔑 API Keys", "🖼️ Images", "ℹ️ About"])

    new_keys: Dict[str, str] = {}
    with tab_keys:
        st.caption(
            "Keys are kept in this browser session's server memory only. They are never written to disk "
            "and are cleared when the session ends."
        )
        for i, (name, cfg) in enumerate(LLM_PROVIDERS.items()):
            saved = bool(s["api_keys"].get(name))
            new_keys[name] = st.text_input(
                f"{name}  {'✅' if saved else ''}".strip(),
                value=s["api_keys"].get(name, ""),
                type="password",
                placeholder=cfg["key_label"],
                help=cfg["key_help"],
                key=f"set_key_{i}",
            ).strip()
        custom_url = st.text_input(
            "Custom endpoint URL (used by 'Custom OpenAI-Compatible API')",
            value=s["custom_base_url"],
            help="Full chat completions URL, e.g. https://api.siliconflow.cn/v1/chat/completions",
            key="set_custom_url",
        ).strip()

    with tab_images:
        unsplash_key = st.text_input(
            "Unsplash Access Key (optional)",
            value=s["unsplash_key"],
            type="password",
            help="Enables real stock photography. If blank, Smart Hybrid mode falls back to AI generation.",
            key="set_unsplash",
        ).strip()
        use_imagen = st.checkbox(
            "Use Google Imagen 3 for AI illustrations (optional)",
            value=s["use_imagen"],
            help="Falls back to Pollinations AI if unavailable. Uses the Google Gemini key from the API Keys tab.",
            key="set_imagen",
        )

    with tab_about:
        st.markdown(
            "**ChatPPT V2** turns reports into consulting-style PowerPoint decks.\n\n"
            "- Chats live in this browser session; refreshing the page clears them.\n"
            "- Model lists are fetched live from each provider once its key is saved."
        )

    c1, c2 = st.columns(2)
    if c1.button("Save", type="primary", use_container_width=True, key="settings_save"):
        s["api_keys"] = {k: v for k, v in new_keys.items() if v}
        s["custom_base_url"] = custom_url
        s["unsplash_key"] = unsplash_key
        s["use_imagen"] = use_imagen
        fetch_live_models.clear()
        st.session_state.pop("model_fetch_fail", None)
        st.rerun()
    if c2.button("Cancel", use_container_width=True, key="settings_cancel"):
        st.rerun()


# ---------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------

def render_sidebar() -> None:
    ss = st.session_state
    with st.sidebar:
        st.markdown("### 📊 ChatPPT")
        if st.button("＋  New chat", key="new_chat_btn", use_container_width=True):
            create_chat()
            st.rerun()
        if st.button("⚙️  Settings", key="settings_btn", use_container_width=True):
            settings_dialog()

        st.markdown('<div class="sidebar-label">Chats</div>', unsafe_allow_html=True)
        for cid in list(ss.chat_order):
            title = ss.chats[cid]["title"]
            title = title if len(title) <= 24 else title[:23] + "…"
            c1, c2 = st.columns([6, 1], vertical_alignment="center")
            if c1.button(f"💬  {title}", key=f"chat_btn_{cid}", use_container_width=True):
                ss.active_chat = cid
                st.rerun()
            if c2.button("✕", key=f"del_{cid}", help="Delete chat"):
                delete_chat(cid)
                st.rerun()


# ---------------------------------------------------------------------
# Chat rendering
# ---------------------------------------------------------------------

def render_hero() -> None:
    st.markdown(
        """
        <div class="hero">
          <h1>What are we presenting today?</h1>
          <p>Drop in a report or paste your notes, and I'll design a consulting-style deck.</p>
          <div class="steps">
            <div class="step"><b>1 · Add your source</b><span>Attach a PDF, DOCX, TXT or MD file, or paste text.</span></div>
            <div class="step"><b>2 · Review the outline</b><span>Tweak layouts, charts, KPIs and images per slide.</span></div>
            <div class="step"><b>3 · Download the .pptx</b><span>Fully editable shapes and native charts.</span></div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_messages(chat: dict) -> None:
    for m in chat["messages"]:
        is_user = m["role"] == "user"
        with st.chat_message(m["role"], avatar=None if is_user else "📊"):
            if m.get("attachment"):
                st.markdown(f"📎 `{m['attachment']}`")
            if m.get("content"):
                st.markdown(m["content"])
            if m.get("meta"):
                st.caption(m["meta"])


def render_outline_editor(chat: dict) -> None:
    """Interactive per-slide editor. Widget keys are scoped by chat id + generation counter."""
    kp = f"{chat['id']}_{chat['gen']}_"
    layout_options = [lt.value for lt in LayoutType]
    report_imgs = chat["report_images"]

    if report_imgs:
        with st.expander(f"🖼️ Preview {len(report_imgs)} Embedded Image(s) Extracted from Your Report"):
            img_cols = st.columns(min(len(report_imgs), 4))
            for r_idx, r_img in enumerate(report_imgs):
                with img_cols[r_idx % 4]:
                    st.image(r_img["bytes"], caption=f"[{r_idx}] {r_img['name']}", use_container_width=True)

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
            if new_layout in (LayoutType.COVER.value, LayoutType.SECTION_DIVIDER.value):
                slide["subtitle_or_context"] = st.text_area(
                    "Subtitle / Section Overview",
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
                        f"Image Source (AI Auto-Recommendation: `{ai_rec_style}`)",
                        options=src_choices,
                        index=src_choices.index(cur_src),
                        key=f"{kp}imgsrc_{idx}",
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
                        "AI Illustration Prompt (English)",
                        value=img_obj.get("visual_prompt", ""),
                        height=75,
                        key=f"{kp}vprompt_{idx}",
                    )

                    if chosen_src == "Use Image from Report" and report_imgs:
                        rep_labels = [f"[{i}] {r['name']}" for i, r in enumerate(report_imgs)]
                        sel_rep = st.selectbox(
                            "Select Extracted Report Image",
                            options=range(len(report_imgs)),
                            format_func=lambda x: rep_labels[x],
                            key=f"{kp}repimg_{idx}",
                        )
                        slide["selected_report_image_idx"] = sel_rep

                    if chosen_src == "Upload Custom Image":
                        up_custom_img = st.file_uploader(
                            "Upload Slide Image (.png, .jpg, .webp)",
                            type=["png", "jpg", "jpeg", "webp"],
                            key=f"{kp}upimg_{idx}",
                        )
                        if up_custom_img is not None:
                            slide["custom_uploaded_image_bytes"] = up_custom_img.getvalue()

                with ic2:
                    style_opts = [VisualStyle.STOCK_PHOTO.value, VisualStyle.AI_ILLUSTRATION.value]
                    new_pref_style = st.selectbox(
                        "Smart Hybrid Preferred Style",
                        options=style_opts,
                        index=style_opts.index(ai_rec_style) if ai_rec_style in style_opts else 1,
                        key=f"{kp}prefstyle_{idx}",
                        help="In 'Smart Hybrid (Auto)' mode, 'stock_photo' queries Unsplash first (then falls back to AI), while 'ai_illustration' generates a custom graphic directly.",
                    )
                    new_caption = st.text_input(
                        "Image Caption",
                        value=img_obj.get("caption", ""),
                        key=f"{kp}imgcap_{idx}",
                    )
                    raw_i_pts = st.text_area(
                        "Accompanying Key Points (`Heading :: Detail`)",
                        value=format_points_for_editor(slide.get("points", [])),
                        height=135,
                        key=f"{kp}ipts_{idx}",
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
                    key=f"{kp}takeaway_{idx}",
                )

            updated_slides.append(slide)

    chat["slides"] = updated_slides


def render_export_block(chat: dict) -> None:
    ss = st.session_state
    s = ss.settings
    cid = chat["id"]

    use_custom_tpl = ss.get("opt_tpl_mode") == "Custom .pptx template"
    tpl_bytes = ss.get("custom_tpl_bytes")
    tpl_name = ss.get("custom_tpl_name")
    use_tpl_colors = ss.get("opt_tpl_color_mode", "Auto-extract colors from .pptx") == "Auto-extract colors from .pptx"

    st.markdown("**Export**")
    if use_custom_tpl and tpl_bytes:
        info = inspect_template_summary(tpl_bytes)
        st.caption(
            f"🎨 Using custom template **`{tpl_name}`** (font: `{info['font']}`, accent: `{info['accent']}`, "
            f"color mode: {'template colors' if use_tpl_colors else ss.get('opt_theme', 'Corporate Blue')}). "
            "Send a short message below to revise the outline."
        )
    else:
        st.caption(
            f"🎨 Styling: **{ss.get('opt_theme', list(BUILTIN_THEMES.keys())[0])}** (change in the 🎨 button below). "
            "Send a short message below (e.g. “make it 6 slides with more charts”) to regenerate from the same source."
        )

    if st.button("🚀 Render presentation (.pptx)", type="primary", use_container_width=True, key=f"render_{cid}"):
        if use_custom_tpl and not tpl_bytes:
            st.error("Upload a .pptx template in the 🎨 menu, or switch to a built-in theme.")
        else:
            with st.spinner("Rendering layouts, native charts and visuals into PowerPoint..."):
                try:
                    chat["pptx_bytes"] = build_presentation_pptx(
                        slides_data=chat["slides"],
                        theme_name=ss.get("opt_theme", list(BUILTIN_THEMES.keys())[0]),
                        gemini_api_key=s["api_keys"].get("Google Gemini"),
                        use_imagen=s["use_imagen"],
                        unsplash_access_key=s["unsplash_key"],
                        extracted_report_images=chat["report_images"],
                        custom_template_bytes=tpl_bytes if use_custom_tpl else None,
                        use_template_colors=use_tpl_colors,
                    )
                    st.success("Presentation rendered! Download it below.")
                except Exception as e:  # noqa: BLE001
                    st.error(f"Error rendering PPTX: {e}")

    if chat["pptx_bytes"]:
        st.download_button(
            "📥 Download ChatPPT_V2_Presentation.pptx",
            data=chat["pptx_bytes"],
            file_name="ChatPPT_V2_Presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            key=f"dl_{cid}_{chat['gen']}",
        )


def render_deck_block(chat: dict) -> None:
    with st.chat_message("assistant", avatar="📊"):
        st.markdown("**Slide outline** — expand any slide to edit it, then render the deck.")
        render_outline_editor(chat)
        st.divider()
        render_export_block(chat)


# ---------------------------------------------------------------------
# Composer (message box with upload / template / tone / model controls)
# ---------------------------------------------------------------------

def _short(text: str, n: int = 16) -> str:
    text = (text or "").split("/")[-1]
    return text if len(text) <= n else text[: n - 1] + "…"


def get_model_options(provider: str, api_key: str, base_url: str, static_models: List[str]):
    """Live model list with a 60s negative cache so a failing network doesn't slow every rerun."""
    fails = st.session_state.setdefault("model_fetch_fail", {})
    sig = (provider, hash(api_key), base_url)
    if fails.get(sig, 0) > time.time():
        return static_models, "Live list unavailable right now; showing presets."
    options, status = resolve_model_options(provider, api_key, base_url, static_models)
    if status.startswith("Could not fetch"):
        fails[sig] = time.time() + 60
    return options, status


def render_model_picker() -> dict:
    """Provider + model selectors (rendered inside a popover). Returns the active model config."""
    s = st.session_state.settings
    provider = st.selectbox("Provider", list(LLM_PROVIDERS.keys()), key="opt_provider")
    cfg = LLM_PROVIDERS[provider]
    api_key = (s["api_keys"].get(provider) or "").strip()
    base_url = (s["custom_base_url"] if provider == CUSTOM_PROVIDER else cfg["base_url"]) or ""

    if not api_key:
        st.warning("No API key for this provider yet. Open ⚙️ Settings in the sidebar.")

    options, status = get_model_options(provider, api_key, base_url, cfg["models"])
    choice = st.selectbox("Model", options, key=f"opt_model_{provider}")
    st.caption(status)
    if st.button("🔄 Refresh model list", key="refresh_models"):
        fetch_live_models.clear()
        st.session_state.pop("model_fetch_fail", None)
        st.rerun()

    if choice == "Custom Model ID...":
        default_custom = "gemini-3.8-flash" if provider == "Google Gemini" else "gpt-4o-mini"
        model = st.text_input("Custom model ID", value=default_custom, key=f"opt_custom_model_{provider}").strip()
    else:
        model = choice
    return {"provider": provider, "api_key": api_key, "base_url": base_url, "model": model}


def handle_send(chat: dict, text: str, report, model_cfg: dict, tone: str, language: str, num_slides: int) -> None:
    """Validate the message, extract source text/images, and queue a generation job."""
    text = (text or "").strip()
    if not text and report is None:
        st.toast("Attach a report or paste some text first.", icon="⚠️")
        return
    if not model_cfg["api_key"]:
        st.toast(f"Add your {model_cfg['provider']} API key first.", icon="🔑")
        settings_dialog()
        return

    is_revision = (
        report is None and bool(chat["slides"]) and bool(chat["source_text"]) and len(text) <= REVISION_MAX_CHARS
    )

    if is_revision:
        source_for_llm = (
            chat["source_text"]
            + "\n\nUser revision instructions (apply these when designing the deck): "
            + text
        )
    else:
        with st.spinner("Reading your file..."):
            file_text = extract_text_from_uploaded_file(report) if report is not None else ""
            chat["report_images"] = extract_images_from_report(report) if report is not None else []
        source = "\n\n".join(filter(None, [file_text.strip(), text]))
        if not source:
            st.toast("Couldn't read any text from that file.", icon="⚠️")
            return
        chat["source_text"] = source
        source_for_llm = source

    if not chat["messages"]:
        chat["title"] = (report.name if report is not None else text.split("\n")[0])[:40] or "New chat"

    # A revision that itself mentions a slide count (e.g. "make it 6 slides") overrides the slider
    enforce_count = not (is_revision and mentions_slide_count(text))
    slide_label = f"{num_slides} slides" if enforce_count else "slide count per your request"
    meta = f"{slide_label} · {language} · {_short(model_cfg['provider'], 12)} / {_short(model_cfg['model'], 24)}"
    shown = text if len(text) <= 500 else text[:500] + f"… ({len(text):,} characters)"
    if is_revision:
        shown = f"✏️ {shown}"
    chat["messages"].append(
        {
            "role": "user",
            "content": shown,
            "attachment": report.name if report is not None else None,
            "meta": meta,
        }
    )
    chat["pending"] = {
        "provider": model_cfg["provider"],
        "api_key": model_cfg["api_key"],
        "model": model_cfg["model"],
        "base_url": model_cfg["base_url"] if model_cfg["provider"] == CUSTOM_PROVIDER else None,
        "source_text": source_for_llm,
        "num_slides": num_slides,
        "enforce_count": enforce_count,
        "tone": tone,
        "language": language,
    }
    chat["composer_ver"] += 1  # new widget keys => text box and file uploader reset
    st.rerun()


def render_composer(chat: dict) -> None:
    ss = st.session_state
    cid, ver = chat["id"], chat["composer_ver"]
    text_key, file_key = f"composer_text_{cid}_{ver}", f"composer_file_{cid}_{ver}"
    busy = chat["pending"] is not None

    # Labels are derived from last-run widget state so popover buttons reflect current choices
    attached = ss.get(file_key)
    prov_name = ss.get("opt_provider", list(LLM_PROVIDERS.keys())[0])
    model_state = ss.get(f"opt_model_{prov_name}")
    if model_state == "Custom Model ID...":
        model_state = ss.get(f"opt_custom_model_{prov_name}")
    is_custom_tpl = ss.get("opt_tpl_mode") == "Custom .pptx template"
    if is_custom_tpl and ss.get("custom_tpl_name"):
        tpl_label = f"🎨 {_short(ss['custom_tpl_name'], 13)}"
    elif is_custom_tpl:
        tpl_label = "🎨 Custom .pptx"
    else:
        tpl_label = f"🎨 {_short(ss.get('opt_theme', 'Theme'), 14)}"

    with st.container(key="composer_dock"):
        with st.container(key="composer"):
            text = st.text_area(
                "Message",
                key=text_key,
                height=90,
                placeholder=(
                    "Paste your report or notes here… "
                    if not chat["slides"]
                    else "Ask for changes, e.g. “make it 6 slides and add more charts”… "
                ),
                label_visibility="collapsed",
                disabled=busy,
            )
            c_file, c_tpl, c_tone, c_model, c_deck, c_send = st.columns(
                [1.15, 1.2, 1.55, 1.45, 1.35, 1.0], vertical_alignment="center"
            )

            with c_file.popover(f"📎 {_short(attached.name, 12)}" if attached else "📎 Attach", use_container_width=True):
                report = st.file_uploader(
                    "Report file",
                    type=["pdf", "docx", "txt", "md"],
                    key=file_key,
                    help="Embedded images in PDF/DOCX files are extracted automatically.",
                )
                st.caption("PDF · DOCX · TXT · MD — embedded figures can be reused on slides.")

            with c_tpl.popover(tpl_label, use_container_width=True):
                st.radio("Styling", ["Built-in 16:9 theme", "Custom .pptx template"], key="opt_tpl_mode")
                if ss.get("opt_tpl_mode") == "Custom .pptx template":
                    tpl_upload = st.file_uploader(
                        "Upload .pptx template",
                        type=["pptx"],
                        key="opt_tpl_file",
                        help="Inherits Slide Master, role layouts, background art, logos, theme colors, and fonts.",
                    )
                    if tpl_upload is not None:
                        new_bytes = tpl_upload.getvalue()
                        if ss.get("custom_tpl_bytes") != new_bytes or ss.get("custom_tpl_name") != tpl_upload.name:
                            ss["custom_tpl_bytes"] = new_bytes
                            ss["custom_tpl_name"] = tpl_upload.name
                            st.rerun()

                    if ss.get("custom_tpl_bytes"):
                        info = inspect_template_summary(ss["custom_tpl_bytes"])
                        st.caption(
                            f"✅ Loaded `{ss['custom_tpl_name']}` · Font: `{info['font']}` · "
                            f"Accent: `{info['accent']}` · BG: `{info['bg']}`"
                        )
                        if st.button("🗑️ Remove template", key="clear_tpl_btn", use_container_width=True):
                            ss["custom_tpl_bytes"] = None
                            ss["custom_tpl_name"] = None
                            st.rerun()

                    st.radio(
                        "Template card & accent colors",
                        ["Auto-extract colors from .pptx", "Use fallback palette below"],
                        key="opt_tpl_color_mode",
                    )
                    st.selectbox("Fallback color palette", list(BUILTIN_THEMES.keys()), key="opt_theme")
                else:
                    st.selectbox("Color palette", list(BUILTIN_THEMES.keys()), key="opt_theme")

            tone_label = c_tone.selectbox(
                "Tone", list(TONE_OPTIONS.keys()), key="opt_tone", label_visibility="collapsed"
            )

            with c_model.popover(f"🤖 {_short(model_state or 'Model', 14)}", use_container_width=True):
                model_cfg = render_model_picker()

            with c_deck.popover(
                f"⚙ {ss.get('opt_slides', 8)} slides", use_container_width=True
            ):
                num_slides = st.slider("Number of slides", 4, 20, 8, key="opt_slides")
                language = st.selectbox("Output language", LANGUAGE_OPTIONS, key="opt_language")

            send = c_send.button(
                "Generate ➤", type="primary", use_container_width=True, key=f"send_{cid}", disabled=busy
            )

    if send:
        handle_send(
            chat,
            text,
            report,
            model_cfg,
            TONE_OPTIONS[tone_label],
            ss.get("opt_language", "English"),
            int(ss.get("opt_slides", 8)),
        )


def run_pending(chat: dict, container) -> None:
    """Run the queued outline generation inside the chat area, then rerun to show the result."""
    p = chat["pending"]
    with container:
        with st.chat_message("assistant", avatar="📊"):
            with st.spinner(f"Designing your deck with {p['provider']} · {p['model']}…"):
                try:
                    slides, count_warning = generate_smart_outline(
                        provider=p["provider"],
                        api_key=p["api_key"],
                        source_text=p["source_text"],
                        num_slides=p["num_slides"],
                        tone=p["tone"],
                        language=p["language"],
                        model_name=p["model"],
                        custom_base_url=p["base_url"],
                        enforce_count=p.get("enforce_count", True),
                    )
                    chat["slides"] = slides
                    chat["gen"] += 1
                    chat["pptx_bytes"] = None
                    if slides and slides[0].get("action_title"):
                        chat["title"] = slides[0]["action_title"][:40]
                    n_img = len(chat["report_images"])
                    note = f" Found {n_img} embedded image(s) in your file." if n_img else ""
                    if count_warning:
                        note += f"\n\n⚠️ {count_warning}"
                    chat["messages"].append(
                        {
                            "role": "assistant",
                            "content": f"Your outline is ready — **{len(slides)} slides**.{note} "
                            "Review and tweak it below, then render the .pptx.",
                        }
                    )
                except Exception as e:  # noqa: BLE001
                    err = str(e)
                    err = err if len(err) <= 1200 else err[:1200] + "…"
                    chat["messages"].append(
                        {"role": "assistant", "content": f"❌ I couldn't generate the outline.\n\n```\n{err}\n```"}
                    )
                chat["pending"] = None
    st.rerun()


# ---------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------

def main():
    st.set_page_config(
        page_title="ChatPPT V2",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    init_state()
    chat = st.session_state.chats[st.session_state.active_chat]
    inject_css(chat["id"])
    render_sidebar()

    chat_area = st.container()
    with chat_area:
        if not chat["messages"] and chat["pending"] is None:
            render_hero()
        render_messages(chat)
        if chat["slides"] and chat["pending"] is None:
            render_deck_block(chat)

    render_composer(chat)

    if chat["pending"] is not None:
        run_pending(chat, chat_area)


if __name__ == "__main__":
    main()