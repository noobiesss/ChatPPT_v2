"""Conversational assistant: decides whether a short message is a question, a targeted slide edit, or a deck
instruction, and answers questions.

Edit when: changing how questions are answered, what context the assistant sees, or how messages are routed.
"""

import re  # regular expressions, used to read the model's "ACTION: ..." first line
from dataclasses import dataclass, field  # small typed container for the routing decision
from typing import List, Optional

# Plain-text (non-JSON) model calls for Gemini and OpenAI-compatible providers, and the <think> cleaner
from chatppt.llm_client import call_gemini_text, call_openai_compatible_text, strip_think_tags
# Provider table (endpoint URLs) and the largest number of slides one targeted edit may rewrite
from chatppt.presets import LLM_PROVIDERS, PATCH_MAX_SLIDES
# System prompt that explains the ANSWER / REVISE / PATCH protocol to the model
from chatppt.prompts import ASSISTANT_SYSTEM_INSTRUCTION
# BM25 excerpt picker, used when the source document is too long to send whole
from chatppt.retrieval import select_relevant_context

# Context limits keep the routing call cheap and fast
SOURCE_CONTEXT_MAX_CHARS = 12000  # longest source text sent as-is; longer sources are cut to relevant excerpts
HISTORY_MAX_MESSAGES = 6  # only the last few chat messages are sent as conversation context
HISTORY_MESSAGE_MAX_CHARS = 800  # each of those messages is truncated to this length

# "ACTION: <verb>" on the first line, tolerating Markdown decoration (**, `, #, >) that models like to add.
# Group 1 = verb, group 2 = rest of that line (PATCH slide list), group 3 = everything after (the answer text)
_ACTION_RE = re.compile(
    # ^[\s*`#>]*      leading spaces / Markdown symbols the model may put before the keyword
    # ACTION\s*:\s*   the keyword and its colon, with optional spaces
    # \**\s*          optional closing bold markers ("**ACTION:** ANSWER")
    # (ANSWER|...)\b  group 1: the verb, as a whole word
    # ([^\n]*)        group 2: whatever follows on the same line (the PATCH slide list)
    # \n?(.*)         group 3: the rest of the reply after the first line (DOTALL lets . cross newlines)
    r"^[\s*`#>]*ACTION\s*:\s*\**\s*(ANSWER|REVISE|PATCH)\b([^\n]*)\n?(.*)", re.IGNORECASE | re.DOTALL
)
# Slide numbers after PATCH: "3", "2, 5", "3-4" (ranges are expanded)
# Alternative 1 = a range "lo - hi" / "lo – hi" / "lo to hi" (groups 1 and 2); alternative 2 = one number (group 3)
_SLIDE_NUM_RE = re.compile(r"(\d+)\s*(?:-|–|to)\s*(\d+)|(\d+)", re.IGNORECASE)


@dataclass
class AssistantDecision:
    """
    action is "answer" (show `reply`, leave the deck alone), "patch" (rewrite only `slides`, 1-based)
    or "revise" (regenerate the whole outline).
    """

    action: str  # "answer", "patch" or "revise"
    reply: str = ""  # the answer text shown in the chat (only used for "answer")
    slides: List[int] = field(default_factory=list)  # 1-based slide numbers to rewrite (only used for "patch")


def parse_patch_targets(spec: str) -> List[int]:
    """Parse the slide list after `ACTION: PATCH`, e.g. ": 2, 5" or "3-4" -> sorted unique numbers."""
    nums = set()  # a set drops duplicates such as "3, 3"
    for m in _SLIDE_NUM_RE.finditer(spec or ""):  # every number or range in the text, left to right
        if m.group(3):  # a single number matched
            nums.add(int(m.group(3)))
        else:  # a range matched: group 1 and group 2 are its two ends
            lo, hi = sorted((int(m.group(1)), int(m.group(2))))  # sorted, so "5-3" works like "3-5"
            # Expand the range, but never past PATCH_MAX_SLIDES numbers (guards against e.g. "1-1000")
            nums.update(range(lo, min(hi, lo + PATCH_MAX_SLIDES) + 1))
    return sorted(nums)  # ascending list, the order the slides are rewritten in


def build_deck_digest(slides: List[dict]) -> str:
    """Compact text view of the current outline so the assistant can answer questions about it."""
    lines: List[str] = []  # one string per output line, joined at the end
    for s in slides or []:
        # Header line per slide: number, layout and title, e.g. "3. [kpi_metrics] Revenue grew 62%"
        lines.append(f"{s.get('slide_number')}. [{s.get('layout_type')}] {s.get('action_title', '')}")
        if s.get("subtitle_or_context"):  # cover / divider / closing subtitle, when present
            lines.append(f"   context: {s['subtitle_or_context']}")
        layout = s.get("layout_type")
        # Only the fields the slide's layout actually shows are listed: the sanitizer fills every layout's
        # fields with placeholders, and those placeholders must not look like real content to the model
        if layout == "two_column":
            for side in ("left_column", "right_column"):
                col = s.get(side) or {}
                # "Heading: detail; Heading: detail" for every item in this column
                items = "; ".join(f"{i.get('heading')}: {i.get('detail')}" for i in col.get("items", []))
                lines.append(f"   {col.get('column_title', side)}: {items}")
        elif layout == "kpi_metrics":
            # "+62% Growth; 3.5x Speed"
            lines.append("   KPIs: " + "; ".join(f"{k.get('value')} {k.get('label')}" for k in s.get("kpis", [])))
        elif layout == "timeline_flow":
            for st in s.get("timeline_steps") or []:  # one line per step, in order
                lines.append(f"   > {st.get('stage_label')}: {st.get('title')}: {st.get('description')}")
        elif layout == "comparison_table":
            lines.append("   table: " + " | ".join(s.get("table_headers") or []))  # header row
            for row in s.get("table_rows") or []:  # then each data row
                lines.append("   row: " + " | ".join(row))
        elif layout not in ("cover", "section_divider", "closing"):
            # Every other content layout (cards, chart, image split) shows its `points`
            for pt in s.get("points", []):
                lines.append(f"   - {pt.get('heading')}: {pt.get('detail')}")
        if layout == "chart_slide" and s.get("chart"):
            ch = s["chart"]
            # Pair each category with its value: "A=1.0, B=2.0, C=3.0"
            data = ", ".join(f"{c}={v}" for c, v in zip(ch.get("categories", []), ch.get("values", [])))
            lines.append(f"   chart ({ch.get('chart_type')}) {ch.get('chart_title')}: {data}")
        if layout == "image_split" and s.get("image"):
            lines.append(f"   image: {s['image'].get('caption')}")  # only the caption is visible on the slide
        if s.get("bottom_takeaway"):
            lines.append(f"   takeaway: {s['bottom_takeaway']}")
    return "\n".join(lines)


def build_assistant_prompt(user_text: str, source_text: str, slides: List[dict], history: List[dict]) -> str:
    """
    Assemble the single user prompt: source, current outline, recent chat, new message.
    A long source is cut down to the excerpts most relevant to this message (and the previous user turn,
    so follow-ups like "and the second one?" still find their topic).
    """
    parts: List[str] = []  # Markdown sections, joined with blank lines at the end
    # Last HISTORY_MAX_MESSAGES non-empty messages (empty ones carry nothing useful)
    recent = [m for m in (history or []) if m.get("content")][-HISTORY_MAX_MESSAGES:]
    source = (source_text or "").strip()
    if source:
        if len(source) > SOURCE_CONTEXT_MAX_CHARS:
            # The most recent user message before this one, or "" if there is none
            prev_user = next((str(m["content"]) for m in reversed(recent) if m.get("role") == "user"), "")
            # Search with both the new message and the previous user turn as the query
            source = select_relevant_context(source, f"{user_text}\n{prev_user}", SOURCE_CONTEXT_MAX_CHARS)
            parts.append(f"## Source document (excerpts relevant to the message; [...] marks skipped text)\n{source}")
        else:
            parts.append(f"## Source document\n{source}")  # short enough: send it whole
    else:
        parts.append("## Source document\n(none yet)")  # tells the model explicitly that there is no source
    digest = build_deck_digest(slides)
    parts.append(f"## Current deck outline\n{digest}" if digest else "## Current deck outline\n(no deck yet)")

    if recent:
        rows = []
        for m in recent:
            who = "User" if m.get("role") == "user" else "Assistant"
            rows.append(f"{who}: {str(m['content'])[:HISTORY_MESSAGE_MAX_CHARS]}")  # truncated per message
        parts.append("## Recent conversation\n" + "\n".join(rows))

    parts.append(f"## New user message\n{user_text}")  # last, so it is the freshest thing the model reads
    return "\n\n".join(parts)


def parse_assistant_reply(raw: Optional[str], deck_size: Optional[int] = None) -> AssistantDecision:
    """
    Parse the `ACTION: ANSWER|REVISE|PATCH: <n,...>` protocol. Unknown or empty replies fall back to safe defaults.
    A PATCH with no numbers, too many slides, or (when deck_size is given) a number outside the deck becomes REVISE.
    """
    text = strip_think_tags(raw)  # reasoning models may prefix their reply with <think>...</think>
    if not text:
        return AssistantDecision("revise")  # nothing usable: behave like the app before Q&A existed
    m = _ACTION_RE.match(text)  # .match only looks at the start of the text
    if not m:
        # The model ignored the protocol but wrote something: treat it as a plain answer
        return AssistantDecision("answer", text)
    action = m.group(1).lower()  # "answer", "revise" or "patch"
    if action == "patch":
        targets = parse_patch_targets(m.group(2))  # slide numbers from the rest of the first line
        # Without a known deck size every number is accepted; otherwise all must lie inside the deck
        in_deck = deck_size is None or all(1 <= n <= deck_size for n in targets)
        if targets and in_deck and len(targets) <= PATCH_MAX_SLIDES:
            return AssistantDecision("patch", slides=targets)
        return AssistantDecision("revise")  # unusable PATCH: regenerate the whole outline instead
    rest = m.group(3).strip()  # the text after the first line
    if action == "revise" or not rest:
        # REVISE, or an ANSWER with no answer text (nothing to show), regenerates the deck
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
            # Gemini goes through its own SDK
            raw = call_gemini_text(api_key, model_name, ASSISTANT_SYSTEM_INSTRUCTION, prompt)
        else:
            # Every other provider speaks the OpenAI chat-completions format; a custom URL overrides the preset
            endpoint = base_url or LLM_PROVIDERS[provider]["base_url"]
            raw = call_openai_compatible_text(endpoint, api_key, model_name, ASSISTANT_SYSTEM_INSTRUCTION, prompt)
        # Pass the deck size so PATCH numbers outside the deck are rejected
        return parse_assistant_reply(raw, deck_size=len(slides or []))
    except Exception:  # noqa: BLE001
        return AssistantDecision("revise")  # network / auth / quota errors never block the user's request
