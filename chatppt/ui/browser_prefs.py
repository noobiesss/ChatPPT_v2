"""Bridge to the browser's localStorage (Streamlit Custom Component v2, inline, invisible).

Edit when: changing when preferences are loaded from / written to the browser, or the Forget action.
What is stored is decided in chatppt/prefs.py.

Flow per browser session:
1. First run: the component reads localStorage once and sends it up as state "loaded" (one rerun).
2. `_on_loaded` (a callback, so it runs before any widget exists) applies it to Session State.
3. Every later run: if the preference payload changed, Python sends a numbered write down via `data`.
Nothing is written before step 2, so a fresh page can never overwrite what the browser remembered.
"""

import streamlit as st

from chatppt.prefs import collect_prefs, parse_prefs, prefs_json
from chatppt.project_io import OPTION_KEYS

COMPONENT_KEY = "browser_prefs"
STORAGE_KEY = "chatppt.prefs.v1"

_JS = """
const STORAGE_KEY = "chatppt.prefs.v1";

export default function (component) {
  const { data, setStateValue } = component;
  if (!data) return;
  // Per browser-session bookkeeping on window (the component is not iframed, so window is the app page)
  const loaded = (window.__chatpptPrefsLoaded = window.__chatpptPrefsLoaded || {});
  const versions = (window.__chatpptPrefsVersion = window.__chatpptPrefsVersion || {});
  const session = data.session || "default";

  if (data.need_load && !loaded[session]) {
    loaded[session] = true;
    let stored = null;
    let error = null;
    try {
      stored = window.localStorage.getItem(STORAGE_KEY);
    } catch (e) {
      error = String(e);
    }
    setStateValue("loaded", { stored: stored, error: error });
  }

  const version = data.version || 0;
  if (data.op && version > (versions[session] || 0)) {
    versions[session] = version;
    try {
      if (data.op === "write") {
        window.localStorage.setItem(STORAGE_KEY, data.payload);
      } else if (data.op === "clear") {
        window.localStorage.removeItem(STORAGE_KEY);
      }
    } catch (e) {
      // Storage blocked (private mode, browser policy): the app keeps working, it just forgets on refresh
    }
  }
}
"""

_BROWSER_PREFS = st.components.v2.component("chatppt_browser_prefs", html="<div></div>\n", js=_JS)


def _current_options() -> dict:
    ss = st.session_state
    return {k: ss.get(k) for k in OPTION_KEYS}


def _on_loaded() -> None:
    """Apply what the browser remembered. Runs as a callback, before any widget of this run exists."""
    ss = st.session_state
    loaded = (ss.get(COMPONENT_KEY) or {}).get("loaded") or {}
    ss["_prefs_loaded"] = True
    ss["_prefs_error"] = loaded.get("error")
    prefs = parse_prefs(loaded.get("stored"))
    if prefs:
        settings = dict(prefs["settings"])
        ss.settings["models"].update(settings.pop("models", {}))
        ss.settings.update(settings)
        ss.settings["remember_keys"] = prefs["remember_keys"]
        for provider, key in prefs["api_keys"].items():
            ss.settings["api_keys"].setdefault(provider, key)  # a key typed in this session wins
        if prefs["unsplash_key"] and not ss.settings.get("unsplash_key"):
            ss.settings["unsplash_key"] = prefs["unsplash_key"]
        for key, value in prefs["options"].items():
            ss[key] = value
    # What the browser holds now: nothing to write until something actually changes
    ss["_prefs_written"] = prefs_json(collect_prefs(ss.settings, _current_options()))


def forget_browser_prefs() -> None:
    """Callback: erase this browser's stored settings and keys and stop remembering until the next Save."""
    ss = st.session_state
    ss["_prefs_clear"] = True
    ss["_prefs_paused"] = True
    ss.settings["remember_keys"] = False


def resume_browser_prefs() -> None:
    """Called when the user saves settings: remembering starts again (with their remember-keys choice)."""
    st.session_state["_prefs_paused"] = False


def sync_browser_prefs() -> None:
    """Mount the invisible bridge (every run): load once, then write only when the preferences changed."""
    ss = st.session_state
    last = ss.get("_prefs_last_op") or {"op": None, "version": 0}
    if ss.pop("_prefs_clear", False):
        last = {"op": "clear", "version": last["version"] + 1}
        ss["_prefs_written"] = None
    elif ss.get("_prefs_loaded") and not ss.get("_prefs_paused"):
        payload = prefs_json(collect_prefs(ss.settings, _current_options()))
        if payload != ss.get("_prefs_written"):
            last = {"op": "write", "payload": payload, "version": last["version"] + 1}
            ss["_prefs_written"] = payload
    ss["_prefs_last_op"] = last
    # The latest operation is re-sent on every run (the browser skips versions it already applied),
    # so a write survives a rerun that interrupted the render it was first sent with
    data = {"session": ss["_session_nonce"], "need_load": not ss.get("_prefs_loaded"), **last}
    _BROWSER_PREFS(key=COMPONENT_KEY, data=data, on_loaded_change=_on_loaded, height="content")
