"""Streamlit session-state setup: chats and default settings.

Edit when: adding a session key, a default setting, changing chat creation/deletion, or restoring projects.
"""

import uuid
from typing import Optional

import streamlit as st

from chatppt.outline import MAX_SLIDES, MIN_SLIDES
from chatppt.presets import CUSTOM_PROVIDER, LLM_PROVIDERS
from chatppt.schemas import IMAGE_MODE_OPTIONS, IMAGE_MODE_UNSPLASH_FIRST


def new_chat(cid: Optional[str] = None) -> dict:
    """A fresh chat dict (also the base for chats restored from the browser)."""
    cid = cid or uuid.uuid4().hex[:8]
    return {
        "id": cid,
        "title": "New chat",
        "messages": [],
        "slides": None,
        "report_images": [],
        "pptx_bytes": None,
        "source_text": "",
        "gen": 0,  # outline version: part of the editor's widget keys, bumped to rebuild them from chat["slides"]
        "composer_ver": 0,  # part of the composer's widget keys, bumped to clear the text box and uploader
        "pending": None,  # the queued job (see ui/generation.py); not None while a request is running
        "slides_history": [],  # undo stack of outline snapshots (see chatppt/deck_ops.py)
        "slides_future": [],  # redo stack
        "updated_at": 0.0,  # last change, for the browser chat store (chatppt/chat_store.py)
    }


def create_chat() -> None:
    """Create a new empty chat and make it the active one."""
    ss = st.session_state
    chat = new_chat()
    ss.chats[chat["id"]] = chat
    ss.chat_order.insert(0, chat["id"])
    ss.active_chat = chat["id"]


def mark_slides_changed(chat: dict) -> None:
    """After the outline was replaced (undo/redo, restore, structure edit): fresh editor widgets, stale .pptx
    dropped, and the slide-count slider follows the deck (applied by the composer before the slider exists)."""
    chat["gen"] += 1
    chat["pptx_bytes"] = None
    if chat.get("slides"):
        st.session_state["opt_slides_sync"] = max(MIN_SLIDES, min(MAX_SLIDES, len(chat["slides"])))


def restore_project(loaded: dict) -> None:
    """Open a validated project (chatppt.project_io.load_project) as a new active chat.
    Must run in a widget callback: composer options can only be set before their widgets exist."""
    ss = st.session_state
    create_chat()
    chat = ss.chats[ss.active_chat]
    chat.update(loaded["chat"])
    settings = dict(loaded["settings"])
    ss.settings["models"].update(settings.pop("models", {}))
    ss.settings.update(settings)
    for key, value in loaded["options"].items():
        if key == "opt_slides":
            ss["opt_slides_sync"] = value
        else:
            ss[key] = value


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
            "remember_keys": True,  # keep API keys in this browser (chatppt/prefs.py)
            "remember_chats": True,  # keep chats in this browser (chatppt/chat_store.py)
        }
    ss.settings.setdefault("provider", next(iter(LLM_PROVIDERS)))
    ss.settings.setdefault("models", {})
    ss.settings.setdefault("remember_keys", True)
    ss.settings.setdefault("remember_chats", True)
    ss.setdefault("_session_nonce", uuid.uuid4().hex)  # identifies this browser session to the prefs bridge
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
