"""Conversational assistant: decides whether a short message is a question or a deck instruction, and answers questions.

Edit when: changing how questions are answered, what context the assistant sees, or how messages are routed.
"""

import re
from dataclasses import dataclass
from typing import List, Optional

from chatppt.llm_client import call_gemini_text, call_openai_compatible_text, strip_think_tags
from chatppt.presets import LLM_PROVIDERS
from chatppt.prompts import ASSISTANT_SYSTEM_INSTRUCTION

# Context limits keep the routing call cheap and fast
SOURCE_CONTEXT_MAX_CHARS = 12000
HISTORY_MAX_MESSAGES = 6
HISTORY_MESSAGE_MAX_CHARS = 800

_ACTION_RE = re.compile(r"^[\s*`#>]*ACTION\s*:\s*\**\s*(ANSWER|REVISE)\b[^\n]*\n?(.*)", re.IGNORECASE | re.DOTALL)


@dataclass
class AssistantDecision:
    """action is "answer" (show `reply`, leave the deck alone) or "revise" (run the outline generation)."""

    action: str
    reply: str = ""


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
    """Assemble the single user prompt: source (truncated), current outline, recent chat, new message."""
    parts: List[str] = []
    source = (source_text or "").strip()
    if source:
        if len(source) > SOURCE_CONTEXT_MAX_CHARS:
            source = source[:SOURCE_CONTEXT_MAX_CHARS] + "\n[... source truncated ...]"
        parts.append(f"## Source document\n{source}")
    else:
        parts.append("## Source document\n(none yet)")
    digest = build_deck_digest(slides)
    parts.append(f"## Current deck outline\n{digest}" if digest else "## Current deck outline\n(no deck yet)")

    recent = [m for m in (history or []) if m.get("content")][-HISTORY_MAX_MESSAGES:]
    if recent:
        rows = []
        for m in recent:
            who = "User" if m.get("role") == "user" else "Assistant"
            rows.append(f"{who}: {str(m['content'])[:HISTORY_MESSAGE_MAX_CHARS]}")
        parts.append("## Recent conversation\n" + "\n".join(rows))

    parts.append(f"## New user message\n{user_text}")
    return "\n\n".join(parts)


def parse_assistant_reply(raw: Optional[str]) -> AssistantDecision:
    """Parse the `ACTION: ANSWER|REVISE` protocol. Unknown or empty replies fall back to safe defaults."""
    text = strip_think_tags(raw)
    if not text:
        return AssistantDecision("revise")
    m = _ACTION_RE.match(text)
    if not m:
        # The model ignored the protocol but wrote something: treat it as a plain answer
        return AssistantDecision("answer", text)
    rest = m.group(2).strip()
    if m.group(1).lower() == "revise" or not rest:
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
    One LLM call that either answers the question or says "revise" (= regenerate the outline).
    Any failure falls back to "revise", which is the app's behavior before Q&A existed.
    """
    try:
        prompt = build_assistant_prompt(user_text, source_text, slides, history)
        if provider == "Google Gemini":
            raw = call_gemini_text(api_key, model_name, ASSISTANT_SYSTEM_INSTRUCTION, prompt)
        else:
            endpoint = base_url or LLM_PROVIDERS[provider]["base_url"]
            raw = call_openai_compatible_text(endpoint, api_key, model_name, ASSISTANT_SYSTEM_INSTRUCTION, prompt)
        return parse_assistant_reply(raw)
    except Exception:  # noqa: BLE001
        return AssistantDecision("revise")