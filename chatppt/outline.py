"""Outline generation: LLM call + retries, slide-count enforcement and outline sanitizing.

Edit when: changing how the outline is requested, validated, trimmed, or normalized.
"""

import re
from typing import List, Optional, Tuple

from google import genai
from google.genai import types

from chatppt.images import _derive_fallback_keywords
from chatppt.llm_client import call_openai_compatible_json, extract_clean_json_str, repair_json_str
from chatppt.presets import LLM_PROVIDERS
from chatppt.prompts import SYSTEM_INSTRUCTION
from chatppt.schemas import ChartType, IMAGE_MODE_OPTIONS, IMAGE_MODE_UNSPLASH_FIRST, ImagePlacement, LayoutType, PresentationOutline, SlideSpec, VisualStyle


def sanitize_outline_dict(
    outline: PresentationOutline,
    default_image_mode: str = IMAGE_MODE_UNSPLASH_FIRST,
) -> List[dict]:
    """
    Validate and repair LLM output edge cases and return a clean list of slide dictionaries
    with all sub-structures initialized for safe UI editing and PPTX rendering.
    """
    if default_image_mode not in IMAGE_MODE_OPTIONS:
        default_image_mode = IMAGE_MODE_UNSPLASH_FIRST

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
        fallback_kw = _derive_fallback_keywords(s_dict.get("action_title", ""))
        if not s_dict.get("image"):
            s_dict["image"] = {
                "preferred_style": VisualStyle.STOCK_PHOTO.value,
                "visual_prompt": (
                    f"Clean modern corporate illustration representing {s_dict['action_title']}, "
                    "minimal vector style, no text"
                ),
                "search_keywords": fallback_kw,
                "caption": s_dict["action_title"],
                "placement": ImagePlacement.LEFT.value if idx % 2 == 0 else ImagePlacement.RIGHT.value,
            }
        else:
            s_dict["image"]["preferred_style"] = VisualStyle.STOCK_PHOTO.value
            if not (s_dict["image"].get("search_keywords") or "").strip():
                s_dict["image"]["search_keywords"] = fallback_kw
            if not (s_dict["image"].get("visual_prompt") or "").strip():
                s_dict["image"]["visual_prompt"] = (
                    f"Clean modern corporate illustration representing {s_dict['action_title']}, "
                    "minimal vector style, no text"
                )

        s_dict["layout_type"] = layout
        # Default to Option 1 (Unsplash Photo First, Fallback to AI) unless the user selected Option 2
        s_dict["image_source_mode"] = default_image_mode

        sanitized_slides.append(s_dict)

    # Guarantee at least one image_split slide in decks of 4+ slides so photos/placeholders appear
    if len(sanitized_slides) >= 4 and not any(
        s["layout_type"] == LayoutType.IMAGE_SPLIT.value for s in sanitized_slides
    ):
        candidate_indices = [
            i for i in range(1, len(sanitized_slides))
            if sanitized_slides[i]["layout_type"] in (LayoutType.BULLET_CARDS.value, LayoutType.THREE_CARDS.value)
        ]
        if not candidate_indices:
            candidate_indices = [
                i for i in range(1, len(sanitized_slides))
                if sanitized_slides[i]["layout_type"] != LayoutType.SECTION_DIVIDER.value
            ]
        if candidate_indices:
            mid_idx = candidate_indices[len(candidate_indices) // 2]
            sanitized_slides[mid_idx]["layout_type"] = LayoutType.IMAGE_SPLIT.value

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


# Relative slide-count edits in revision notes, e.g. "add one more slide", "remove the last slide", "\u52a0\u4e00\u9875"
_EN_QTY = r"\d+|an?|one|two|three|four|five|six"
_EN_SKIP = r"(?!(?:to|on|in|of|for|at|the|this|that|each|every)\b)"
_EN_ADD_RE = re.compile(
    rf"\b(?:add|insert|append|include|create)\s+(?P<n1>{_EN_QTY})\s+(?:(?:more|extra|additional|new|another)\s+)?"
    rf"(?:{_EN_SKIP}[\w-]+\s+){{0,3}}?(?:slides?|pages?)\b(?!\s*\d)"
    rf"|\b(?P<n2>\d+|one|two|three|four|five|six)\s+(?:more|extra|additional|new)\s+(?:slides?|pages?)\b"
    rf"|\banother\s+(?:{_EN_SKIP}[\w-]+\s+){{0,2}}?(?:slide|page)\b(?!\s*\d)",
    re.IGNORECASE,
)
_EN_REMOVE_RE = re.compile(
    rf"\b(?:remove|delete|drop|cut)\s+(?:(?P<n>{_EN_QTY})\s+)?(?:(?:the|this|that|last|first|final|extra)\s+)*(?:slides?|pages?)\b",
    re.IGNORECASE,
)
# Chinese: verbs add = \u52a0/\u589e\u52a0/\u6dfb\u52a0/\u65b0\u589e/\u8865\u5145 ; remove = \u5220\u9664/\u5220\u6389/\u5220/\u53bb\u6389/\u51cf\u5c11/\u51cf ; unit = \u9875/\u9801/\u5f20/\u5f35
_ZH_QTY = r"\d+|[\u4e00\u4e8c\u4e24\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]"
_ZH_GAP = r"(?:(?![\u5230\u5728\u7b2c\u6bcf])\w){0,4}?"
_ZH_ADD_RE = re.compile(
    rf"(?:\u52a0|\u589e\u52a0|\u6dfb\u52a0|\u65b0\u589e|\u8865\u5145)\s*(?P<n>{_ZH_QTY})?\s*[\u4e2a\u500b]?\s*{_ZH_GAP}[\u9875\u9801\u5f20\u5f35]"
)
_ZH_REMOVE_RE = re.compile(
    rf"(?:\u5220\u9664|\u5220\u6389|\u5220|\u53bb\u6389|\u51cf\u5c11|\u51cf)\s*(?P<n>{_ZH_QTY})?\s*[\u4e2a\u500b]?\s*{_ZH_GAP}[\u9875\u9801\u5f20\u5f35]"
)
_QTY_WORDS = {
    "a": 1, "an": 1, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
    "\u4e00": 1, "\u4e8c": 2, "\u4e24": 2, "\u4e09": 3, "\u56db": 4, "\u4e94": 5,
    "\u516d": 6, "\u4e03": 7, "\u516b": 8, "\u4e5d": 9, "\u5341": 10,
}
# Same range as the "Number of slides" slider in the composer
MIN_SLIDES, MAX_SLIDES = 4, 20


def _qty_value(token: Optional[str]) -> int:
    """Convert a quantity token ("2", "one", "\u4e24") to an int; a missing quantity means 1."""
    if not token:
        return 1
    token = token.lower()
    return int(token) if token.isdigit() else _QTY_WORDS.get(token, 1)


def parse_slide_count_change(text: str, current_count: int) -> Optional[int]:
    """
    Detect a RELATIVE slide-count edit in a revision note ("add one more slide", "remove the last slide").
    Returns the new absolute target (current_count + delta, clamped to the slider range),
    or None when the note contains no relative change.
    """
    text = text or ""
    delta = 0
    for m in _EN_ADD_RE.finditer(text):
        delta += _qty_value(m.groupdict().get("n1") or m.groupdict().get("n2"))
    for m in _ZH_ADD_RE.finditer(text):
        delta += _qty_value(m.group("n"))
    for m in _EN_REMOVE_RE.finditer(text):
        delta -= _qty_value(m.group("n"))
    for m in _ZH_REMOVE_RE.finditer(text):
        delta -= _qty_value(m.group("n"))
    if delta == 0:
        return None
    return max(MIN_SLIDES, min(MAX_SLIDES, current_count + delta))


def _request_outline(
    provider: str,
    api_key: str,
    model_name: str,
    user_prompt: str,
    custom_base_url: Optional[str],
) -> PresentationOutline:
    """
    Send an outline request to the selected provider and return the validated PresentationOutline.
    Retries once automatically if the model returns transient malformed or empty JSON.
    """
    last_err: Optional[Exception] = None
    for attempt in range(2):
        try:
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
                raw_json = extract_clean_json_str(getattr(response, "text", None) or "")
                if not raw_json:
                    raise RuntimeError(f"Google Gemini (`{model_name}`) returned an empty response.")
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
        except Exception as exc:  # noqa: BLE001
            last_err = exc
            if attempt == 0:
                continue
            raise last_err


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
    default_image_mode: str = IMAGE_MODE_UNSPLASH_FIRST,
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

    return sanitize_outline_dict(outline, default_image_mode=default_image_mode), warning