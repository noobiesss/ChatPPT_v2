"""Conversational assistant: decides whether a short message is a question, a targeted slide edit, or a deck
instruction, and answers questions.

Edit when: changing how questions are answered, what context the assistant sees, or how messages are routed.
"""

import re
from dataclasses import dataclass, field
from typing import List, Optional

from chatppt.llm_client import call_gemini_text, call_openai_compatible_text, strip_think_tags
from chatppt.presets import LLM_PROVIDERS, PATCH_MAX_SLIDES
from chatppt.prompts import ASSISTANT_SYSTEM_INSTRUCTION
from chatppt.retrieval import select_relevant_context

# Context limits keep the routing call cheap and fast
SOURCE_CONTEXT_MAX_CHARS = 12000
HISTORY_MAX_MESSAGES = 6
HISTORY_MESSAGE_MAX_CHARS = 800

_ACTION_RE = re.compile(
    r"^[\s*`#>]*ACTION\s*:\s*\**\s*(ANSWER|REVISE|PATCH)\b([^\n]*)\n?(.*)", re.IGNORECASE | re.DOTALL
)
# Slide numbers after PATCH: "3", "2, 5", "3-4" (ranges are expanded)
_SLIDE_NUM_RE = re.compile(r"(\d+)\s*(?:-|–|to)\s*(\d+)|(\d+)", re.IGNORECASE)


@dataclass
class AssistantDecision:
    """
    action is "answer" (show `reply`, leave the deck alone), "patch" (rewrite only `slides`, 1-based)
    or "revise" (regenerate the whole outline).
    """

    action: str
    reply: str = ""
    slides: List[int] = field(default_factory=list)


def parse_patch_targets(spec: str) -> List[int]:
    """Parse the slide list after `ACTION: PATCH`, e.g. ": 2, 5" or "3-4" -> sorted unique numbers."""
    nums = set()
    for m in _SLIDE_NUM_RE.finditer(spec or ""):
        if m.group(3):
            nums.add(int(m.group(3)))
        else:
            lo, hi = sorted((int(m.group(1)), int(m.group(2))))
            nums.update(range(lo, min(hi, lo + PATCH_MAX_SLIDES) + 1))
    return sorted(nums)


def build_deck_digest(slides: List[dict]) -> str:
    """Compact text view of the current outline so the assistant can answer questions about it."""
    lines: List[str] = []
    for s in slides or []:
        lines.append(f"{s.get('slide_number')}. [{s.get('layout_type')}] {s.get('action_title', '')}")
        if s.get("subtitle_or_context"):
            lines.append(f"   context: {s['subtitle_or_context']}")
        layout = s.get("layout_type")
        if layout == "two_column":
            for side in ("left_column", "right_column"):
                col = s.get(side) or {}
                items = "; ".join(f"{i.get('heading')}: {i.get('detail')}" for i in col.get("items", []))
                lines.append(f"   {col.get('column_title', side)}: {items}")
        elif layout == "kpi_metrics":
            lines.append("   KPIs: " + "; ".join(f"{k.get('value')} {k.get('label')}" for k in s.get("kpis", [])))
        elif layout not in ("cover", "section_divider"):
            for pt in s.get("points", []):
                lines.append(f"   - {pt.get('heading')}: {pt.get('detail')}")
        if layout == "chart_slide" and s.get("chart"):
            ch = s["chart"]
            data = ", ".join(f"{c}={v}" for c, v in zip(ch.get("categories", []), ch.get("values", [])))
            lines.append(f"   chart ({ch.get('chart_type')}) {ch.get('chart_title')}: {data}")
        if layout == "image_split" and s.get("image"):
            lines.append(f"   image: {s['image'].get('caption')}")
        if s.get("bottom_takeaway"):
            lines.append(f"   takeaway: {s['bottom_takeaway']}")
    return "\n".join(lines)


def build_assistant_prompt(user_text: str, source_text: str, slides: List[dict], history: List[dict]) -> str:
    """
    Assemble the single user prompt: source, current outline, recent chat, new message.
    A long source is cut down to the excerpts most relevant to this message (and the previous user turn,
    so follow-ups like "and the second one?" still find their topic).
    """
    parts: List[str] = []
    recent = [m for m in (history or []) if m.get("content")][-HISTORY_MAX_MESSAGES:]
    source = (source_text or "").strip()
    if source:
        if len(source) > SOURCE_CONTEXT_MAX_CHARS:
            prev_user = next((str(m["content"]) for m in reversed(recent) if m.get("role") == "user"), "")
            source = select_relevant_context(source, f"{user_text}\n{prev_user}", SOURCE_CONTEXT_MAX_CHARS)
            parts.append(f"## Source document (excerpts relevant to the message; [...] marks skipped text)\n{source}")
        else:
            parts.append(f"## Source document\n{source}")
    else:
        parts.append("## Source document\n(none yet)")
    digest = build_deck_digest(slides)
    parts.append(f"## Current deck outline\n{digest}" if digest else "## Current deck outline\n(no deck yet)")

    if recent:
        rows = []
        for m in recent:
            who = "User" if m.get("role") == "user" else "Assistant"
            rows.append(f"{who}: {str(m['content'])[:HISTORY_MESSAGE_MAX_CHARS]}")
        parts.append("## Recent conversation\n" + "\n".join(rows))

    parts.append(f"## New user message\n{user_text}")
    return "\n\n".join(parts)


def parse_assistant_reply(raw: Optional[str], deck_size: Optional[int] = None) -> AssistantDecision:
    """
    Parse the `ACTION: ANSWER|REVISE|PATCH: <n,...>` protocol. Unknown or empty replies fall back to safe defaults.
    A PATCH with no numbers, too many slides, or (when deck_size is given) a number outside the deck becomes REVISE.
    """
    text = strip_think_tags(raw)
    if not text:
        return AssistantDecision("revise")
    m = _ACTION_RE.match(text)
    if not m:
        # The model ignored the protocol but wrote something: treat it as a plain answer
        return AssistantDecision("answer", text)
    action = m.group(1).lower()
    if action == "patch":
        targets = parse_patch_targets(m.group(2))
        in_deck = deck_size is None or all(1 <= n <= deck_size for n in targets)
        if targets and in_deck and len(targets) <= PATCH_MAX_SLIDES:
            return AssistantDecision("patch", slides=targets)
        return AssistantDecision("revise")
    rest = m.group(3).strip()
    if action == "revise" or not rest:
        return AssistantDecision("revise")
    return AssistantDecision("answer", rest)


def route_message(
    provider: str,
    api_key: str,
    model_name: str,
    base_url: Optional[str],
    user_text: str,
    source_text: str,
    slides: List[dict],
    history: List[dict],
) -> AssistantDecision:
    """
    One LLM call that answers the question, names the slides to patch, or says "revise" (= regenerate the outline).
    Any failure falls back to "revise", which is the app's behavior before Q&A existed.
    """
    try:
        prompt = build_assistant_prompt(user_text, source_text, slides, history)
        if provider == "Google Gemini":
            raw = call_gemini_text(api_key, model_name, ASSISTANT_SYSTEM_INSTRUCTION, prompt)
        else:
            endpoint = base_url or LLM_PROVIDERS[provider]["base_url"]
            raw = call_openai_compatible_text(endpoint, api_key, model_name, ASSISTANT_SYSTEM_INSTRUCTION, prompt)
        return parse_assistant_reply(raw, deck_size=len(slides or []))
    except Exception:  # noqa: BLE001
        return AssistantDecision("revise")