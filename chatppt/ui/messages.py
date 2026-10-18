"""Hero (empty state) and chat message rendering, with the outline block anchored in the conversation.

Edit when: changing the welcome screen, how chat messages look, or where the outline block appears.
"""

import base64
from functools import lru_cache
from typing import Callable, List, Optional

import streamlit as st

from chatppt.ui.common import ASSISTANT_AVATAR, LOGO_PATH


@lru_cache(maxsize=1)
def _logo_data_uri() -> str:
    with open(LOGO_PATH, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode("ascii")


def render_hero() -> None:
    st.markdown(
        f"""
        <div class="hero">
          <img class="hero-logo" src="{_logo_data_uri()}" alt="ChatPPT logo">
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


# Content prefixes of assistant messages written before messages carried a "kind"
_DECK_PREFIXES = ("Your outline is ready", "Updated **slide")


def message_kind(message: dict) -> str:
    """'deck' (the outline was created or changed), 'answer', 'error', or 'user'."""
    if message.get("role") != "assistant":
        return "user"
    if message.get("kind") in ("deck", "answer", "error"):
        return message["kind"]
    content = str(message.get("content") or "")
    if content.startswith(_DECK_PREFIXES):
        return "deck"
    if "Answer" in str(message.get("meta") or ""):
        return "answer"
    return "error"


def deck_anchor(messages: List[dict]) -> Optional[int]:
    """Index of the message the outline block follows: the last one that changed the deck (None = at the end)."""
    for i in range(len(messages) - 1, -1, -1):
        if message_kind(messages[i]) == "deck":
            return i
    return None


def render_messages(chat: dict, render_deck: Optional[Callable[[], None]] = None) -> None:
    """
    Render the conversation. `render_deck` draws the outline block right after the message that last changed
    the deck, so answering a question or a failed request never moves (or seems to re-send) the outline.
    """
    anchor = deck_anchor(chat["messages"]) if render_deck else None
    for i, m in enumerate(chat["messages"]):
        is_user = m["role"] == "user"
        with st.chat_message(m["role"], avatar=None if is_user else ASSISTANT_AVATAR):
            if m.get("attachment"):
                st.markdown(f":material/attach_file: `{m['attachment']}`")
            if m.get("content"):
                st.markdown(m["content"])
            if m.get("meta"):
                st.caption(m["meta"])
        if render_deck and i == anchor:
            render_deck()
    if render_deck and anchor is None:
        render_deck()
