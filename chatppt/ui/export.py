"""Deck block: undo / redo / save-project toolbar, outline editor, 'Render deck' / download controls.

Edit when: changing export options, the render button, or the download flow.
"""

import re
from typing import Optional

import streamlit as st

from chatppt.deck_ops import can_redo, can_undo, redo, undo
from chatppt.pptx_build import build_presentation_pptx
from chatppt.project_io import OPTION_KEYS, export_project
from chatppt.pptx_template import inspect_template_summary
from chatppt.presets import BUILTIN_THEMES
from chatppt.schemas import IMAGE_MODE_UNSPLASH_FIRST
from chatppt.ui.common import ASSISTANT_AVATAR, TPL_COLORS_FROM_TEMPLATE, TPL_MODE_CUSTOM
from chatppt.ui.outline_editor import render_outline_editor
from chatppt.ui.state import mark_slides_changed


def render_export_block(chat: dict) -> None:
    ss = st.session_state
    s = ss.settings
    cid = chat["id"]

    use_custom_tpl = ss.get("opt_tpl_mode") == TPL_MODE_CUSTOM
    tpl_bytes = ss.get("custom_tpl_bytes")
    tpl_name = ss.get("custom_tpl_name")
    use_tpl_colors = ss.get("opt_tpl_color_mode", TPL_COLORS_FROM_TEMPLATE) == TPL_COLORS_FROM_TEMPLATE

    st.markdown("**Export**")
    if use_custom_tpl and tpl_bytes:
        info = inspect_template_summary(tpl_bytes)
        st.caption(
            f":material/palette: Using custom template **`{tpl_name}`** (font: `{info['font']}`, accent: `{info['accent']}`, "
            f"color mode: {'template colors' if use_tpl_colors else ss.get('opt_theme', 'Corporate Blue')}). "
            "Send a short message below to revise the outline."
        )
    else:
        st.caption(
            f":material/palette: Styling: **{ss.get('opt_theme', list(BUILTIN_THEMES.keys())[0])}** "
            "(change it with the palette button below). "
            "Send a short message below (e.g. “make it 6 slides with more charts”) to regenerate from the same source."
        )

    if st.button(
        "Render presentation (.pptx)", type="primary", icon=":material/rocket_launch:", width="stretch",
        key=f"render_{cid}",
    ):
        if use_custom_tpl and not tpl_bytes:
            st.error("Upload a .pptx template with the palette button below, or switch to a built-in theme.")
        else:
            with st.spinner("Rendering layouts, native charts and visuals into PowerPoint..."):
                try:
                    chat["pptx_bytes"] = build_presentation_pptx(
                        slides_data=chat["slides"],
                        theme_name=ss.get("opt_theme", list(BUILTIN_THEMES.keys())[0]),
                        gemini_api_key=s["api_keys"].get("Google Gemini"),
                        use_imagen=s["use_imagen"],
                        unsplash_access_key=s["unsplash_key"],
                        extracted_report_images=chat["report_images"],
                        custom_template_bytes=tpl_bytes if use_custom_tpl else None,
                        use_template_colors=use_tpl_colors,
                        default_image_mode=s.get("image_mode", IMAGE_MODE_UNSPLASH_FIRST),
                    )
                    st.success("Presentation rendered! Download it below.")
                except Exception as e:  # noqa: BLE001
                    st.error(f"Error rendering PPTX: {e}")

    if chat["pptx_bytes"]:
        st.download_button(
            "Download ChatPPT_V2_Presentation.pptx",
            data=chat["pptx_bytes"],
            file_name="ChatPPT_V2_Presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            icon=":material/download:",
            width="stretch",
            key=f"dl_{cid}_{chat['gen']}",
        )


def _render_deck_toolbar(chat: dict) -> Optional[str]:
    """Undo / Redo / Save project row above the outline. Returns "undo" / "redo" when clicked."""
    ss = st.session_state
    cid = chat["id"]
    # Captured now: the download callable runs on another thread, where Session State is not available
    settings = dict(ss.settings)
    options = {k: ss.get(k) for k in OPTION_KEYS}
    action = None
    with st.container(horizontal=True, vertical_alignment="center"):
        if st.button("Undo", icon=":material/undo:", key=f"undo_{cid}", disabled=not can_undo(chat),
                     help="Go back to the previous outline version (before the last AI change or slide move)."):
            action = "undo"
        if st.button("Redo", icon=":material/redo:", key=f"redo_{cid}", disabled=not can_redo(chat),
                     help="Re-apply the version you just undid."):
            action = "redo"
        st.download_button(
            "Save project",
            data=lambda: export_project(chat, settings, options),
            file_name=f"{_safe_filename(chat.get('title'))}.chatppt.json",
            mime="application/json",
            icon=":material/save:",
            on_click="ignore",
            key=f"save_project_{cid}",
            help="Download the outline, conversation, source text and report figures as a .json project "
            "(API keys are not included). Open it again from the sidebar after a page refresh.",
        )
    return action


def _safe_filename(title: Optional[str]) -> str:
    cleaned = re.sub(r"[^\w\-]+", "_", title or "", flags=re.UNICODE).strip("_")
    return cleaned[:40] or "ChatPPT_project"


def render_deck_placeholder(chat: dict) -> None:
    """Shown in place of the outline while a message is being processed: same spot, no widgets.
    (Clicking an editor widget now would rerun the app and interrupt, then repeat, the running AI call.)"""
    with st.chat_message("assistant", avatar=ASSISTANT_AVATAR):
        with st.container(border=True):
            st.markdown(f"**Slide outline** · {len(chat['slides'])} slides")
            st.caption(":material/hourglass_top: Editing is paused while I work on your message.")


def render_deck_block(chat: dict) -> None:
    with st.chat_message("assistant", avatar=ASSISTANT_AVATAR):
        st.markdown("**Slide outline** — expand any slide to edit it, then render the deck.")
        action = _render_deck_toolbar(chat)
        # The editor runs first so text typed right before the click is part of the snapshot
        render_outline_editor(chat)
        if action:
            (undo if action == "undo" else redo)(chat)
            mark_slides_changed(chat)
            st.rerun()
        st.divider()
        render_export_block(chat)
