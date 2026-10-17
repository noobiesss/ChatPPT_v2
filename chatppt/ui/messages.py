"""Hero (empty state) and chat message rendering.

Edit when: changing the welcome screen or how chat messages look.
"""

import base64
from functools import lru_cache

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


def render_messages(chat: dict) -> None:
    for m in chat["messages"]:
        is_user = m["role"] == "user"
        with st.chat_message(m["role"], avatar=None if is_user else ASSISTANT_AVATAR):
            if m.get("attachment"):
                st.markdown(f":material/attach_file: `{m['attachment']}`")
            if m.get("content"):
                st.markdown(m["content"])
            if m.get("meta"):
                st.caption(m["meta"])
