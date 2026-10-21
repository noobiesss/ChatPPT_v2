"""Chats kept in the user's own browser (IndexedDB): record layout, change detection and restore.

Edit when: changing what is saved per chat, the size limits, or how stored chats are validated on restore.
Pure functions (no Streamlit); the browser bridge lives in chatppt/ui/browser_chats.py.

Each chat is stored as up to three records so a small edit never re-sends large data:
  "<id>/chat"    title, messages, slides, updated_at   (rewritten on every change)
  "<id>/source"  source text                            (rewritten only when the source changes)
  "<id>/images"  report figures as base64               (rewritten only when the figures change)
"""

import hashlib
import json
from typing import Dict, List, Optional

from chatppt.project_io import ProjectFileError, decode_images, encode_images, encode_messages, load_messages, load_slides

STORE_VERSION = 1
PARTS = ("chat", "source", "images")
MAX_STORED_CHATS = 30
# Figures above this total size are not kept in the browser (the chat itself still is)
MAX_IMAGES_BYTES = 15 * 1024 * 1024


def record_key(chat_id: str, part: str) -> str:
    return f"{chat_id}/{part}"


def is_storable(chat: dict) -> bool:
    """Only chats with content, and never in the middle of a generation."""
    return bool(chat.get("messages") or chat.get("slides")) and chat.get("pending") is None


def _dumps(obj) -> str:
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"))


def _chat_body(chat: dict) -> dict:
    return {"title": chat.get("title", ""), "messages": encode_messages(chat.get("messages")),
            "slides": chat.get("slides") or []}


def _images_fit(chat: dict) -> bool:
    return sum(len(im.get("bytes") or b"") for im in chat.get("report_images") or []) <= MAX_IMAGES_BYTES


def part_signature(chat: dict, part: str) -> str:
    """Cheap fingerprint of one part, to detect changes without re-encoding large data every run."""
    if part == "chat":
        raw = _dumps(_chat_body(chat))
    elif part == "source":
        raw = chat.get("source_text") or ""
    else:
        figures = chat.get("report_images") or []
        # Names + byte sizes stand in for the image data (figures are replaced, never edited in place);
        # the size-cap flag is included so crossing MAX_IMAGES_BYTES also counts as a change
        raw = repr([(im.get("name"), len(im.get("bytes") or b"")) for im in figures]) + repr(_images_fit(chat))
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


def part_payload(chat: dict, part: str) -> str:
    """The JSON string stored for one part."""
    if part == "chat":
        return _dumps({"v": STORE_VERSION, **_chat_body(chat), "updated_at": chat.get("updated_at", 0)})
    if part == "source":
        return _dumps({"v": STORE_VERSION, "source_text": chat.get("source_text") or ""})
    figures = encode_images(chat.get("report_images")) if _images_fit(chat) else []
    return _dumps({"v": STORE_VERSION, "report_images": figures})


def _load(raw: Optional[str]) -> Optional[dict]:
    try:
        data = json.loads(raw) if isinstance(raw, str) else None
    except ValueError:
        return None
    return data if isinstance(data, dict) and data.get("v") == STORE_VERSION else None


def restore_chats(records: Dict[str, str]) -> List[dict]:
    """
    Rebuild chats from stored records, newest first (at most MAX_STORED_CHATS). Every chat is validated like an
    uploaded project; a broken or unknown record is skipped instead of failing the whole restore.
    """
    by_chat: Dict[str, Dict[str, str]] = {}
    for key, value in (records or {}).items():
        chat_id, _, part = str(key).partition("/")
        if chat_id and part in PARTS:
            by_chat.setdefault(chat_id, {})[part] = value
    chats = []
    for chat_id, parts in by_chat.items():
        body = _load(parts.get("chat"))
        if not body:
            continue
        try:
            slides = load_slides(body.get("slides") or [])
        except ProjectFileError:
            continue
        messages = load_messages(body.get("messages"))
        if not messages and not slides:
            continue
        source = _load(parts.get("source")) or {}
        images = _load(parts.get("images")) or {}
        updated = body.get("updated_at")
        chats.append({
            "id": chat_id[:32],
            "title": str(body.get("title") or "Restored chat")[:40],
            "messages": messages,
            "slides": slides or None,
            "source_text": str(source.get("source_text") or ""),
            "report_images": decode_images(images.get("report_images")),
            "updated_at": float(updated) if isinstance(updated, (int, float)) else 0.0,
        })
    chats.sort(key=lambda c: c["updated_at"], reverse=True)
    return chats[:MAX_STORED_CHATS]
