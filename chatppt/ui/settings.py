"""Settings dialog (provider -> API key -> model, image keys) and active-model resolution.

Edit when: changing the settings window, provider/model selection, or saved settings.
"""

import time
from typing import List

import streamlit as st

from chatppt.images import _clean_unsplash_key
from chatppt.llm_client import fetch_live_models, resolve_model_options
from chatppt.presets import CUSTOM_PROVIDER, LLM_PROVIDERS
from chatppt.schemas import IMAGE_MODE_OPTIONS, IMAGE_MODE_UNSPLASH_FIRST
from chatppt.ui.common import IMAGE_MODE_HELP, IMAGE_MODE_LABELS


def _refresh_model_list() -> None:
    fetch_live_models.clear()
    st.session_state.pop("model_fetch_fail", None)


def open_settings() -> None:
    """Open the settings dialog with fresh drafts (no stale widget state from a previous visit)."""
    ss = st.session_state
    ss.pop("_settings_draft", None)
    for key in [k for k in ss.keys() if isinstance(k, str) and k.startswith("set_")]:
        del ss[key]
    settings_dialog()


@st.dialog("Settings", icon=":material/settings:", width="large")
def settings_dialog() -> None:
    """One page, one Save: provider -> API key -> model, then image settings."""
    ss = st.session_state
    s = ss.settings
    draft = ss.setdefault("_settings_draft", {"keys": {}, "models": {}, "custom_url": None})
    providers = list(LLM_PROVIDERS.keys())

    st.caption(
        "Keys are kept in this browser session's server memory only. They are never written to disk "
        "and are cleared when the session ends."
    )

    # --- 1. Provider -> API key -> model ---
    st.markdown("##### 1 · AI model")
    current = s.get("provider") if s.get("provider") in providers else providers[0]

    def _provider_label(name: str) -> str:
        has_key = bool(draft["keys"].get(name) or s["api_keys"].get(name))
        return f"{name} · key saved" if has_key else name

    provider = st.selectbox(
        "Provider",
        providers,
        index=providers.index(current),
        format_func=_provider_label,
        key="set_provider",
    )
    cfg = LLM_PROVIDERS[provider]

    api_key = st.text_input(
        f"{provider} API key",
        value=draft["keys"].get(provider, s["api_keys"].get(provider, "")),
        type="password",
        placeholder=cfg["key_label"],
        help=cfg["key_help"],
        key=f"set_key_{provider}",
    ).strip()
    draft["keys"][provider] = api_key

    if provider == CUSTOM_PROVIDER:
        custom_url = st.text_input(
            "Endpoint URL",
            value=draft["custom_url"] if draft["custom_url"] is not None else s["custom_base_url"],
            help="Full chat completions URL, e.g. https://api.siliconflow.cn/v1/chat/completions",
            key="set_custom_url",
        ).strip()
        draft["custom_url"] = custom_url
        base_url = custom_url
    else:
        base_url = cfg["base_url"] or ""

    options, status = get_model_options(provider, api_key, base_url, cfg["models"])
    saved_model = draft["models"].get(provider) or s["models"].get(provider) or ""
    custom_default = ""
    if saved_model in options:
        model_index = options.index(saved_model)
    elif saved_model and "Custom Model ID..." in options:
        model_index = options.index("Custom Model ID...")
        custom_default = saved_model
    else:
        model_index = 0
    choice = st.selectbox("Model", options, index=model_index, key=f"set_model_{provider}")
    if choice == "Custom Model ID...":
        fallback_id = "gemini-3.8-flash" if provider == "Google Gemini" else "gpt-4o-mini"
        model = st.text_input(
            "Custom model ID", value=custom_default or fallback_id, key=f"set_custom_model_{provider}"
        ).strip()
    else:
        model = choice
    draft["models"][provider] = model

    cap_col, refresh_col = st.columns([5, 2], vertical_alignment="center")
    cap_col.caption(status)
    refresh_col.button("Refresh list", icon=":material/refresh:", key="set_refresh", on_click=_refresh_model_list, width="stretch")

    # --- 2. Images (three modes, defaulting to Unsplash / AI) ---
    st.divider()
    st.markdown("##### 2 · Images")
    cur_img_mode = s.get("image_mode", IMAGE_MODE_UNSPLASH_FIRST)
    if cur_img_mode not in IMAGE_MODE_OPTIONS:
        cur_img_mode = IMAGE_MODE_UNSPLASH_FIRST
    image_mode = st.segmented_control(
        "Slide image mode",
        IMAGE_MODE_OPTIONS,
        default=cur_img_mode,
        required=True,
        format_func=IMAGE_MODE_LABELS.get,
        help=IMAGE_MODE_HELP + " You can also pick a mode (and a report figure) per slide in the outline editor.",
        key="set_image_mode",
    ) or cur_img_mode
    unsplash_key = st.text_input(
        "Unsplash access key",
        value=s["unsplash_key"],
        type="password",
        help="Used by the Unsplash / AI and Report figures modes. If blank or no photo matches, an AI image is generated.",
        key="set_unsplash",
    ).strip()
    use_imagen = st.checkbox(
        "Use Google Imagen 3 for AI fallback",
        value=s["use_imagen"],
        help="When Unsplash has no photo (or no key), tries Google Imagen 3 first and falls back to Pollinations AI.",
        key="set_imagen",
    )

    c1, c2 = st.columns(2)
    if c1.button("Save", type="primary", width="stretch", key="settings_save"):
        keys = dict(s["api_keys"])
        for name, value in draft["keys"].items():
            if value:
                keys[name] = value
            else:
                keys.pop(name, None)
        s["api_keys"] = keys
        s["provider"] = provider
        for name, value in draft["models"].items():
            if value:
                s["models"][name] = value
        if draft["custom_url"] is not None:
            s["custom_base_url"] = draft["custom_url"]
        prev_image_mode = s.get("image_mode", IMAGE_MODE_UNSPLASH_FIRST)
        s["image_mode"] = image_mode
        s["unsplash_key"] = _clean_unsplash_key(unsplash_key)
        s["use_imagen"] = use_imagen
        # Apply updated global image mode to existing slides in open chats
        if image_mode != prev_image_mode:
            for chat_obj in ss.chats.values():
                if chat_obj.get("slides"):
                    for sl in chat_obj["slides"]:
                        sl["image_source_mode"] = image_mode
                    chat_obj["gen"] += 1
        _refresh_model_list()
        ss.pop("_settings_draft", None)
        st.rerun()
    if c2.button("Cancel", width="stretch", key="settings_cancel"):
        ss.pop("_settings_draft", None)
        st.rerun()


def get_model_options(provider: str, api_key: str, base_url: str, static_models: List[str]):
    """Live model list with a 60s negative cache so a failing network doesn't slow every rerun."""
    fails = st.session_state.setdefault("model_fetch_fail", {})
    sig = (provider, hash(api_key), base_url)
    if fails.get(sig, 0) > time.time():
        return static_models, "Live list unavailable right now; showing presets."
    options, status = resolve_model_options(provider, api_key, base_url, static_models)
    if status.startswith("Could not fetch"):
        fails[sig] = time.time() + 60
    return options, status


def get_active_model_cfg() -> dict:
    """Provider / key / model chosen in Settings."""
    s = st.session_state.settings
    providers = list(LLM_PROVIDERS.keys())
    provider = s.get("provider") if s.get("provider") in providers else providers[0]
    cfg = LLM_PROVIDERS[provider]
    api_key = (s["api_keys"].get(provider) or "").strip()
    base_url = (s["custom_base_url"] if provider == CUSTOM_PROVIDER else cfg["base_url"]) or ""
    model = s["models"].get(provider) or next((m for m in cfg["models"] if m != "Custom Model ID..."), "")
    return {"provider": provider, "api_key": api_key, "base_url": base_url, "model": model}
