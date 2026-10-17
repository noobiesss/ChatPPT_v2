"""Deck block: outline editor + 'Render deck' / download controls.

Edit when: changing export options, the render button, or the download flow.
"""

import streamlit as st

from chatppt.pptx_build import build_presentation_pptx
from chatppt.pptx_template import inspect_template_summary
from chatppt.presets import BUILTIN_THEMES
from chatppt.schemas import IMAGE_MODE_UNSPLASH_FIRST
from chatppt.ui.outline_editor import render_outline_editor


def render_export_block(chat: dict) -> None:
    ss = st.session_state
    s = ss.settings
    cid = chat["id"]

    use_custom_tpl = ss.get("opt_tpl_mode") == "Custom .pptx template"
    tpl_bytes = ss.get("custom_tpl_bytes")
    tpl_name = ss.get("custom_tpl_name")
    use_tpl_colors = ss.get("opt_tpl_color_mode", "Auto-extract colors from .pptx") == "Auto-extract colors from .pptx"

    st.markdown("**Export**")
    if use_custom_tpl and tpl_bytes:
        info = inspect_template_summary(tpl_bytes)
        st.caption(
            f"🎨 Using custom template **`{tpl_name}`** (font: `{info['font']}`, accent: `{info['accent']}`, "
            f"color mode: {'template colors' if use_tpl_colors else ss.get('opt_theme', 'Corporate Blue')}). "
            "Send a short message below to revise the outline."
        )
    else:
        st.caption(
            f"🎨 Styling: **{ss.get('opt_theme', list(BUILTIN_THEMES.keys())[0])}** (change in the 🎨 button below). "
            "Send a short message below (e.g. “make it 6 slides with more charts”) to regenerate from the same source."
        )

    if st.button("🚀 Render presentation (.pptx)", type="primary", width="stretch", key=f"render_{cid}"):
        if use_custom_tpl and not tpl_bytes:
            st.error("Upload a .pptx template in the 🎨 menu, or switch to a built-in theme.")
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
            "📥 Download ChatPPT_V2_Presentation.pptx",
            data=chat["pptx_bytes"],
            file_name="ChatPPT_V2_Presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            width="stretch",
            key=f"dl_{cid}_{chat['gen']}",
        )


def render_deck_block(chat: dict) -> None:
    with st.chat_message("assistant", avatar="📊"):
        st.markdown("**Slide outline** — expand any slide to edit it, then render the deck.")
        render_outline_editor(chat)
        st.divider()
        render_export_block(chat)
