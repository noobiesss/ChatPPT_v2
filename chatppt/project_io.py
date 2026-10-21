"""Project files (.json): save a chat (outline, conversation, source text, report figures) plus non-secret
settings, and load them back after a page refresh or on another machine.

Edit when: changing what a saved project contains or how an uploaded project is validated.
API keys are never written to a project file.
"""

import base64
import json
from datetime import datetime, timezone
from typing import Dict, List

from chatppt.outline import MAX_SLIDES, MIN_SLIDES, _sanitize_slide
from chatppt.presets import BUILTIN_THEMES, LANGUAGE_OPTIONS, LLM_PROVIDERS, TONE_OPTIONS
from chatppt.schemas import IMAGE_MODE_OPTIONS, IMAGE_MODE_UNSPLASH_FIRST, SlideSpec

PROJECT_FORMAT = "chatppt-project"
PROJECT_VERSION = 1
# Settings that are safe to share; API keys / the Unsplash key are deliberately absent
SAFE_SETTING_KEYS = ("provider", "models", "custom_base_url", "image_mode", "use_imagen", "remember_chats")
# Composer options restored on load (template mode is not: the .pptx template itself is not saved)
OPTION_KEYS = ("opt_slides", "opt_language", "opt_tone", "opt_theme")
_MESSAGE_KEYS = ("role", "content", "meta", "attachment", "kind")


class ProjectFileError(ValueError):
    """The uploaded file is not a usable ChatPPT project. The message is user-facing."""


def encode_images(images: List[Dict[str, object]]) -> List[dict]:
    """Report figures as JSON-safe dicts (base64 data)."""
    return [
        {"name": im.get("name", ""), "width": im.get("width"), "height": im.get("height"),
         "data": base64.b64encode(im["bytes"]).decode("ascii")}
        for im in images or [] if im.get("bytes")
    ]


def encode_messages(messages: List[dict]) -> List[dict]:
    return [{k: m[k] for k in _MESSAGE_KEYS if m.get(k) is not None} for m in messages or []]


def export_project(chat: dict, settings: dict, options: dict) -> bytes:
    """Serialize one chat and the non-secret settings as UTF-8 JSON bytes."""
    data = {
        "format": PROJECT_FORMAT,
        "version": PROJECT_VERSION,
        "saved_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "chat": {
            "title": chat.get("title", ""),
            "messages": encode_messages(chat.get("messages")),
            "slides": chat.get("slides") or [],
            "source_text": chat.get("source_text", ""),
            "report_images": encode_images(chat.get("report_images")),
        },
        "settings": {k: settings[k] for k in SAFE_SETTING_KEYS if k in settings},
        "options": {k: options[k] for k in OPTION_KEYS if options.get(k) is not None},
    }
    return json.dumps(data, ensure_ascii=False, indent=1).encode("utf-8")


def load_slides(raw_slides) -> List[dict]:
    if not isinstance(raw_slides, list):
        raise ProjectFileError("The project has no slide list.")
    slides = []
    for i, raw in enumerate(raw_slides):
        try:
            spec = SlideSpec.model_validate(raw)
        except Exception as exc:  # noqa: BLE001
            raise ProjectFileError(f"Slide {i + 1} in the project is not valid: {str(exc)[:200]}") from exc
        # A project file is user-supplied: re-run the generation sanitizer so every slide is safe to edit and render
        mode = raw.get("image_source_mode")
        mode = mode if mode in IMAGE_MODE_OPTIONS else IMAGE_MODE_UNSPLASH_FIRST
        slide = _sanitize_slide(spec, i, mode)
        if isinstance(raw.get("report_image_index"), int):
            slide["report_image_index"] = raw["report_image_index"]
        slides.append(slide)
    return slides


def load_messages(raw_messages) -> List[dict]:
    messages = []
    for m in raw_messages if isinstance(raw_messages, list) else []:
        if isinstance(m, dict) and m.get("role") in ("user", "assistant") and isinstance(m.get("content", ""), str):
            messages.append({k: str(m[k]) for k in _MESSAGE_KEYS if m.get(k) is not None})
    return messages


def decode_images(raw_images) -> List[Dict[str, object]]:
    images = []
    for im in raw_images if isinstance(raw_images, list) else []:
        try:
            images.append({"name": str(im.get("name", "")), "width": im.get("width"), "height": im.get("height"),
                           "bytes": base64.b64decode(im["data"], validate=True)})
        except Exception:  # noqa: BLE001 - a broken figure is skipped, not fatal
            continue
    return images


def load_safe_settings(raw: dict) -> dict:
    """Keep only known, non-secret settings with valid values (shared with browser prefs)."""
    out = {}
    if raw.get("provider") in LLM_PROVIDERS:
        out["provider"] = raw["provider"]
    if isinstance(raw.get("models"), dict):
        out["models"] = {k: v for k, v in raw["models"].items() if k in LLM_PROVIDERS and isinstance(v, str)}
    if isinstance(raw.get("custom_base_url"), str):
        out["custom_base_url"] = raw["custom_base_url"]
    if raw.get("image_mode") in IMAGE_MODE_OPTIONS:
        out["image_mode"] = raw["image_mode"]
    for flag in ("use_imagen", "remember_chats"):
        if isinstance(raw.get(flag), bool):
            out[flag] = raw[flag]
    return out


def load_options(raw: dict) -> dict:
    """Keep only known composer options with valid values (shared with browser prefs)."""
    out = {}
    if isinstance(raw.get("opt_slides"), int):
        out["opt_slides"] = max(MIN_SLIDES, min(MAX_SLIDES, raw["opt_slides"]))
    for key, allowed in (("opt_language", LANGUAGE_OPTIONS), ("opt_tone", list(TONE_OPTIONS)),
                         ("opt_theme", list(BUILTIN_THEMES))):
        if raw.get(key) in allowed:
            out[key] = raw[key]
    return out


def load_project(raw_bytes: bytes) -> dict:
    """
    Parse and validate a project file. Returns {"chat": {...}, "settings": {...}, "options": {...}} with
    sanitized slides; raises ProjectFileError with a readable reason otherwise.
    """
    try:
        data = json.loads(raw_bytes.decode("utf-8-sig"))
    except Exception as exc:  # noqa: BLE001
        raise ProjectFileError("This file is not a ChatPPT project (it is not valid JSON).") from exc
    if not isinstance(data, dict) or data.get("format") != PROJECT_FORMAT:
        raise ProjectFileError("This JSON file is not a ChatPPT project.")
    if not isinstance(data.get("version"), int) or data["version"] > PROJECT_VERSION:
        raise ProjectFileError("This project was saved by a newer version of ChatPPT.")
    raw_chat = data.get("chat") if isinstance(data.get("chat"), dict) else {}
    slides = load_slides(raw_chat.get("slides"))
    chat = {
        "title": str(raw_chat.get("title") or "Restored project")[:40],
        "messages": load_messages(raw_chat.get("messages")),
        "slides": slides or None,
        "source_text": str(raw_chat.get("source_text") or ""),
        "report_images": decode_images(raw_chat.get("report_images")),
    }
    settings = load_safe_settings(data.get("settings") if isinstance(data.get("settings"), dict) else {})
    options = load_options(data.get("options") if isinstance(data.get("options"), dict) else {})
    return {"chat": chat, "settings": settings, "options": options}
