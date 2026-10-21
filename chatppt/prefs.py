"""Preferences remembered in the user's own browser (localStorage): settings, composer options and,
if the user allows it, API keys. Nothing is stored on the server.

Edit when: changing what is remembered between visits or how a stored payload is validated.
Pure functions (no Streamlit); the browser bridge lives in chatppt/ui/browser_prefs.py.
"""

import json
from typing import Optional

from chatppt.presets import LLM_PROVIDERS
from chatppt.project_io import OPTION_KEYS, SAFE_SETTING_KEYS, load_options, load_safe_settings

PREFS_VERSION = 1
_MAX_SECRET_LEN = 500


def collect_prefs(settings: dict, options: dict) -> dict:
    """What to store for this user. API keys / the Unsplash key only when settings["remember_keys"] is on."""
    prefs = {
        "v": PREFS_VERSION,
        "settings": {k: settings[k] for k in SAFE_SETTING_KEYS if k in settings},
        "options": {k: options[k] for k in OPTION_KEYS if options.get(k) is not None},
        "remember_keys": bool(settings.get("remember_keys", True)),
    }
    if prefs["remember_keys"]:
        prefs["api_keys"] = {k: v for k, v in (settings.get("api_keys") or {}).items() if v}
        prefs["unsplash_key"] = settings.get("unsplash_key") or ""
    return prefs


def prefs_json(prefs: dict) -> str:
    """Stable serialization, so an unchanged payload is never written twice."""
    return json.dumps(prefs, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def _secret(value) -> Optional[str]:
    if isinstance(value, str) and 0 < len(value.strip()) <= _MAX_SECRET_LEN:
        return value.strip()
    return None


def parse_prefs(raw: Optional[str]) -> Optional[dict]:
    """
    Validate a stored payload. Returns {"settings", "options", "remember_keys", "api_keys", "unsplash_key"}
    with only known keys and valid values, or None when nothing usable is stored.
    """
    if not raw:
        return None
    try:
        data = json.loads(raw)
    except (TypeError, ValueError):
        return None
    if not isinstance(data, dict) or data.get("v") != PREFS_VERSION:
        return None
    raw_keys = data.get("api_keys") if isinstance(data.get("api_keys"), dict) else {}
    return {
        "settings": load_safe_settings(data.get("settings") if isinstance(data.get("settings"), dict) else {}),
        "options": load_options(data.get("options") if isinstance(data.get("options"), dict) else {}),
        "remember_keys": data.get("remember_keys") is not False,  # missing means on (the default)
        # localStorage can be edited by hand: keep keys only for known providers and of a sane length
        "api_keys":{k: _secret(v) for k, v in raw_keys.items() if k in LLM_PROVIDERS and _secret(v)},
        "unsplash_key": _secret(data.get("unsplash_key")) or "",
    }
