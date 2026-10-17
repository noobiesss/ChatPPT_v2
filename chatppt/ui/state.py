"""Streamlit session-state setup: chats and default settings.

Edit when: adding a session key, a default setting, or changing chat creation/deletion.
"""

import uuid

import streamlit as st

from chatppt.presets import CUSTOM_PROVIDER, LLM_PROVIDERS
from chatppt.schemas import IMAGE_MODE_OPTIONS, IMAGE_MODE_UNSPLASH_FIRST


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
            "provider": next(iter(LLM_PROVIDERS)),
            "models": {},
            "custom_base_url": LLM_PROVIDERS[CUSTOM_PROVIDER]["base_url"] or "",
            "image_mode": IMAGE_MODE_UNSPLASH_FIRST,
            "unsplash_key": "",
            "use_imagen": False,
        }
    ss.settings.setdefault("provider", next(iter(LLM_PROVIDERS)))
    ss.settings.setdefault("models", {})
    if ss.settings.get("image_mode") not in IMAGE_MODE_OPTIONS:
        ss.settings["image_mode"] = IMAGE_MODE_UNSPLASH_FIRST
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
