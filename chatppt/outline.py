"""Outline generation: LLM call + retries, slide-count enforcement and outline sanitizing.

Edit when: changing how the outline is requested, validated, trimmed, or normalized.
"""

import json
import re
from typing import List, Optional, Tuple, Type

from google import genai
from google.genai import types
from pydantic import BaseModel

from chatppt.chat_assistant import build_deck_digest
from chatppt.images import _derive_fallback_keywords
from chatppt.llm_client import call_openai_compatible_json, extract_clean_json_str, repair_json_str
from chatppt.presets import LLM_PROVIDERS
from chatppt.prompts import SYSTEM_INSTRUCTION
from chatppt.retrieval import select_relevant_context
from chatppt.schemas import ChartType, IMAGE_MODE_OPTIONS, IMAGE_MODE_UNSPLASH_FIRST, ImagePlacement, LayoutType, PresentationOutline, SlideSpec, SlidePatch, VisualStyle


# Readability limits shared by the sanitizer, the renderer and the editor
TIMELINE_MAX_STEPS = 5
TABLE_MAX_COLS, TABLE_MAX_ROWS = 5, 6


def normalize_timeline_steps(steps: Optional[List[dict]]) -> List[dict]:
    """Drop empty steps, fill missing keys, keep at most TIMELINE_MAX_STEPS."""
    out = []
    for step in steps or []:
        step = step or {}
        title = str(step.get("title") or "").strip()
        desc = str(step.get("description") or "").strip()
        if not title and not desc:
            continue
        label = str(step.get("stage_label") or "").strip() or f"Step {len(out) + 1}"
        out.append({"stage_label": label, "title": title or desc[:40], "description": desc})
    return out[:TIMELINE_MAX_STEPS]


def _steps_from_points(points: List[dict]) -> List[dict]:
    return normalize_timeline_steps(
        [{"stage_label": f"Step {i + 1}", "title": p.get("heading"), "description": p.get("detail")}
         for i, p in enumerate(points or [])]
    )


def normalize_table(headers: Optional[List[str]], rows: Optional[List[List[str]]]) -> Tuple[List[str], List[List[str]]]:
    """
    Make the table rectangular: strip cells, drop empty rows, pad/cut every row to the header count,
    cap columns and rows. Returns ([], []) when fewer than 2 columns or no data row remain.
    """
    headers = [str(h or "").strip() for h in (headers or [])][:TABLE_MAX_COLS]
    while headers and not headers[-1]:
        headers.pop()
    if len(headers) < 2:
        return [], []
    width = len(headers)
    clean_rows = []
    for row in rows or []:
        cells = [str(c if c is not None else "").strip() for c in (row or [])][:width]
        if not any(cells):
            continue
        clean_rows.append(cells + [""] * (width - len(cells)))
    clean_rows = clean_rows[:TABLE_MAX_ROWS]
    return (headers, clean_rows) if clean_rows else ([], [])


def _sanitize_slide(slide: SlideSpec, idx: int, default_image_mode: str) -> dict:
    """Repair one slide (0-based deck position `idx`) and fill every sub-structure the editor and renderer expect."""
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

    # Validate timeline steps; a model that wrote `points` instead of steps still gets a timeline
    steps = normalize_timeline_steps(s_dict.get("timeline_steps"))
    if len(steps) < 2 and layout == LayoutType.TIMELINE_FLOW.value:
        if len(slide.points or []) >= 2:
            steps = _steps_from_points(s_dict["points"])
        else:
            layout = LayoutType.BULLET_CARDS.value
    s_dict["timeline_steps"] = steps or _steps_from_points(s_dict["points"])

    # Validate table: rectangular grid with 2-5 columns and at least 1 row, else fall back to cards
    headers, rows = normalize_table(s_dict.get("table_headers"), s_dict.get("table_rows"))
    if not rows and layout == LayoutType.COMPARISON_TABLE.value:
        layout = LayoutType.BULLET_CARDS.value
    if not rows:
        headers = ["Topic", "Detail"]
        rows = [[p["heading"], p["detail"]] for p in s_dict["points"][:TABLE_MAX_ROWS]]
    s_dict["table_headers"], s_dict["table_rows"] = headers, rows

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

    if layout == LayoutType.CLOSING.value:
        s_dict["bottom_takeaway"] = None  # a closing slide never carries a takeaway banner
    s_dict["layout_type"] = layout
    # Default to Option 1 (Unsplash Photo First, Fallback to AI) unless the user selected Option 2
    s_dict["image_source_mode"] = default_image_mode

    return s_dict


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

    # A closing (thank-you) slide is only valid as the single, last slide
    slides = list(outline.slides)
    closing_slides = [s for s in slides if s.layout_type == LayoutType.CLOSING]
    other_slides = [s for s in slides if s.layout_type != LayoutType.CLOSING]
    if closing_slides and other_slides:
        slides = other_slides + [closing_slides[-1]]

    for idx, slide in enumerate(slides):
        sanitized_slides.append(_sanitize_slide(slide, idx, default_image_mode))

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
                if sanitized_slides[i]["layout_type"] not in (LayoutType.SECTION_DIVIDER.value, LayoutType.CLOSING.value)
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
_EN_SKIP = r"(?!(?:to|on|in|of|for|at|the|this|that|each|every|from|with|about)\b)"
_EN_ADD_RE = re.compile(
    rf"\b(?:add|insert|append|include|create)\s+(?P<n1>{_EN_QTY})\s+(?:(?:more|extra|additional|new|another)\s+)?"
    rf"(?:{_EN_SKIP}[\w-]+\s+){{0,3}}?(?:slides?|pages?)\b(?!\s*\d)"
    rf"|\b(?P<n2>\d+|one|two|three|four|five|six)\s+(?:more|extra|additional|new)\s+(?:slides?|pages?)\b"
    rf"|\banother\s+(?:{_EN_SKIP}[\w-]+\s+){{0,2}}?(?:slide|page)\b(?!\s*\d)"
    r"|\b(?:add|insert|append|include|create)\s+(?:the\s+)?(?:thank[\s-]?you|thanks|closing|ending|q\s*&\s*a)\s+(?:slides?|pages?)\b",
    re.IGNORECASE,
)
_EN_REMOVE_RE = re.compile(
    rf"\b(?:remove|delete|drop|cut)\s+(?:(?P<n>{_EN_QTY})\s+)?(?:(?:the|this|that|last|first|final|extra)\s+)*"
    rf"(?:{_EN_SKIP}[\w-]+\s+){{0,2}}?(?:slides?|pages?)\b",
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
    response_model: Type[BaseModel] = PresentationOutline,
) -> BaseModel:
    """
    Send an outline request to the selected provider and return the validated `response_model`
    (the full PresentationOutline by default, SlidePatch for targeted edits).
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
                        response_schema=response_model,
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
                    response_model=response_model,
                )
            return response_model.model_validate_json(repair_json_str(raw_json, response_model))
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
    # Keep the closing slide AND the conclusion before it when the deck ends with a thank-you slide
    protected_tail = 2 if trimmed and trimmed[-1].layout_type == LayoutType.CLOSING else 1
    while len(trimmed) > target and len(trimmed) > protected_tail + 1:
        del trimmed[-(protected_tail + 1)]
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


# Source excerpt budget for a targeted edit (only chunks relevant to the edited slides are sent)
PATCH_SOURCE_MAX_CHARS = 12000

# Fields each layout actually uses; the sanitizer fills the others with placeholders the model must not see
_LAYOUT_FIELDS = {
    LayoutType.BULLET_CARDS.value: ("points",),
    LayoutType.THREE_CARDS.value: ("points",),
    LayoutType.TWO_COLUMN.value: ("left_column", "right_column"),
    LayoutType.KPI_METRICS.value: ("kpis",),
    LayoutType.CHART_SLIDE.value: ("points", "chart"),
    LayoutType.IMAGE_SPLIT.value: ("points", "image"),
    LayoutType.TIMELINE_FLOW.value: ("timeline_steps",),
    LayoutType.COMPARISON_TABLE.value: ("table_headers", "table_rows"),
}


def _slide_for_prompt(slide: dict) -> dict:
    """The parts of a slide dict that are visible on the rendered slide (what the model should edit)."""
    keys = ("slide_number", "layout_type", "action_title", "subtitle_or_context", "bottom_takeaway")
    keys += _LAYOUT_FIELDS.get(slide.get("layout_type"), ())
    return {k: slide.get(k) for k in keys if slide.get(k) not in (None, "", [])}


def patch_slides(
    provider: str,
    api_key: str,
    model_name: str,
    custom_base_url: Optional[str],
    slides: List[dict],
    targets: List[int],
    instruction: str,
    source_text: str,
    tone: str,
    language: str,
) -> List[dict]:
    """
    Targeted edit: ask the model for ONLY the slides numbered in `targets` (1-based) and replace them in place.
    Every other slide dict is returned as the very same object, so untouched slides cannot drift.
    Raises if the model does not return one slide per target (the caller then regenerates the whole deck).
    """
    deck = list(slides or [])
    targets = sorted({n for n in targets if 1 <= n <= len(deck)})
    if not targets:
        raise ValueError("No valid slide numbers to edit.")
    current = [deck[n - 1] for n in targets]
    nums = ", ".join(str(n) for n in targets)

    query = f"{instruction}\n{build_deck_digest(current)}"
    source = select_relevant_context(source_text, query, PATCH_SOURCE_MAX_CHARS)
    current_json = json.dumps([_slide_for_prompt(s) for s in current], ensure_ascii=False, indent=1)
    prompt = f"""
    You are editing ONE PART of an existing {len(deck)}-slide deck. Rewrite ONLY slide(s) {nums}.

    User instruction: {instruction}

    Rules:
    - Return a JSON object whose `slides` array contains exactly {len(targets)} slide object(s), one per slide {nums},
      in that order, each keeping its original `slide_number`. NEVER return other slides.
    - Apply the instruction. Keep everything it does not ask to change (facts, numbers, wording, layout).
    - If the instruction asks for another layout, switch `layout_type` and fill every field that layout needs,
      reusing the slide's existing content and the source excerpts. Never invent numbers.
    - Slide 1 must stay `cover`. Use `closing` only for the last slide ({len(deck)}) and only if asked.
    - Tone: {tone}. Output language: {language} (keep ONLY `visual_prompt` and `search_keywords` in English).

    Whole deck, for context only:
    {build_deck_digest(deck)}

    Current content of the slide(s) to rewrite:
    {current_json}

    Source document excerpts ([...] marks skipped text):
    \"\"\"
    {source}
    \"\"\"
    """
    result = _request_outline(provider, api_key, model_name, prompt, custom_base_url, response_model=SlidePatch)

    got = list(result.slides)
    by_number = {s.slide_number: s for s in got}
    if len(got) == len(targets) and all(n in by_number for n in targets):
        new_specs = [by_number[n] for n in targets]
    elif len(got) == len(targets):
        new_specs = got  # right count but renumbered: trust the order we asked for
    else:
        raise RuntimeError(f"The model returned {len(got)} slide(s) for the {len(targets)} requested.")

    patched = list(deck)
    for n, spec in zip(targets, new_specs):
        if spec.layout_type == LayoutType.CLOSING and n != len(deck):
            spec = spec.model_copy(update={"layout_type": LayoutType.BULLET_CARDS})
        old = deck[n - 1]
        # Keep the image option the user picked for this slide in the editor
        mode = old.get("image_source_mode") if old.get("image_source_mode") in IMAGE_MODE_OPTIONS else IMAGE_MODE_UNSPLASH_FIRST
        patched[n - 1] = _sanitize_slide(spec, n - 1, mode)
        if old.get("report_image_index") is not None:
            patched[n - 1]["report_image_index"] = old["report_image_index"]
    return patched
