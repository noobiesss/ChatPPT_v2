"""Sticky bottom composer: text box, attach, template, model button, options, Generate.

Edit when: changing the input bar or its popovers.
"""

import streamlit as st

from chatppt.pptx_template import inspect_template_summary
from chatppt.presets import BUILTIN_THEMES, LANGUAGE_OPTIONS, TONE_OPTIONS
from chatppt.ui.common import _short
from chatppt.ui.generation import handle_send
from chatppt.ui.settings import get_active_model_cfg, open_settings


def render_composer(chat: dict) -> None:
    ss = st.session_state
    # A revision such as "add one more slide" queued a new target: apply it BEFORE the slider is created
    synced = ss.pop("opt_slides_sync", None)
    if synced is not None:
        ss["opt_slides"] = int(synced)
    cid, ver = chat["id"], chat["composer_ver"]
    text_key, file_key = f"composer_text_{cid}_{ver}", f"composer_file_{cid}_{ver}"
    busy = chat["pending"] is not None

    # Labels are derived from last-run widget state so popover buttons reflect current choices
    attached = ss.get(file_key)
    model_cfg = get_active_model_cfg()
    model_label = f"🤖 {_short(model_cfg['model'], 18)}" if model_cfg["api_key"] else "🔑 Set up model"
    is_custom_tpl = ss.get("opt_tpl_mode") == "Custom .pptx template"
    if is_custom_tpl and ss.get("custom_tpl_name"):
        tpl_label = f"🎨 {_short(ss['custom_tpl_name'], 13)}"
    elif is_custom_tpl:
        tpl_label = "🎨 Custom .pptx"
    else:
        tpl_label = f"🎨 {_short(ss.get('opt_theme', 'Theme'), 14)}"

    with st.container(key="composer_dock"):
        with st.container(key="composer"):
            text = st.text_area(
                "Message",
                key=text_key,
                height=90,
                placeholder=(
                    "Paste your report or notes, or ask me anything… "
                    if not chat["slides"]
                    else "Ask a question, or request changes, e.g. “why this layout?” or “add one more slide”… "
                ),
                label_visibility="collapsed",
                disabled=busy,
            )
            attach_label = f"📎 {_short(attached.name, 12)}" if attached else "📎 Attach"
            deck_label = f"⚙ {ss.get('opt_slides', 8)} slides"
            send_label = "Send ➤"
            # Column widths follow label length so button text is never cut off
            widths = [max(len(lbl), 8) + 5 for lbl in (attach_label, tpl_label, model_label, deck_label, send_label)]
            c_file, c_tpl, c_model, c_deck, c_send = st.columns(widths, vertical_alignment="center")

            with c_file.popover(attach_label, use_container_width=True):
                report = st.file_uploader(
                    "Report file",
                    type=["pdf", "docx", "txt", "md"],
                    key=file_key,
                    help="Attach a PDF, DOCX, TXT, or MD report.",
                )
                st.caption("PDF · DOCX · TXT · MD")

            with c_tpl.popover(tpl_label, use_container_width=True):
                st.radio("Styling", ["Built-in 16:9 theme", "Custom .pptx template"], key="opt_tpl_mode")
                if ss.get("opt_tpl_mode") == "Custom .pptx template":
                    tpl_upload = st.file_uploader(
                        "Upload .pptx template",
                        type=["pptx"],
                        key="opt_tpl_file",
                        help="Inherits the template's backgrounds, theme colors, and fonts (sample-slide decorations are not copied).",
                    )
                    if tpl_upload is not None:
                        new_bytes = tpl_upload.getvalue()
                        if ss.get("custom_tpl_bytes") != new_bytes or ss.get("custom_tpl_name") != tpl_upload.name:
                            ss["custom_tpl_bytes"] = new_bytes
                            ss["custom_tpl_name"] = tpl_upload.name
                            st.rerun()

                    if ss.get("custom_tpl_bytes"):
                        info = inspect_template_summary(ss["custom_tpl_bytes"])
                        st.caption(
                            f"✅ Loaded `{ss['custom_tpl_name']}` · Font: `{info['font']}` · "
                            f"Accent: `{info['accent']}` · BG: `{info['bg']}`"
                        )
                        if st.button("🗑️ Remove template", key="clear_tpl_btn", use_container_width=True):
                            ss["custom_tpl_bytes"] = None
                            ss["custom_tpl_name"] = None
                            st.rerun()

                    st.radio(
                        "Template card & accent colors",
                        ["Auto-extract colors from .pptx", "Use fallback palette below"],
                        key="opt_tpl_color_mode",
                    )
                    st.selectbox("Fallback color palette", list(BUILTIN_THEMES.keys()), key="opt_theme")
                else:
                    st.selectbox("Color palette", list(BUILTIN_THEMES.keys()), key="opt_theme")

            if c_model.button(model_label, key="model_btn", use_container_width=True, help="Change provider / model"):
                open_settings()

            with c_deck.popover(deck_label, use_container_width=True):
                num_slides = st.slider("Number of slides", 4, 20, 8, key="opt_slides")
                language = st.selectbox("Output language", LANGUAGE_OPTIONS, key="opt_language")
                tone_label = st.selectbox("Tone", list(TONE_OPTIONS.keys()), key="opt_tone")

            send = c_send.button(
                send_label, type="primary", use_container_width=True, key=f"send_{cid}", disabled=busy
            )
        st.markdown(
            '<div class="composer-note">ChatPPT can make mistakes. Please check responses.</div>',
            unsafe_allow_html=True,
        )

    if send:
        handle_send(
            chat,
            text,
            report,
            model_cfg,
            TONE_OPTIONS[tone_label],
            ss.get("opt_language", "English"),
            int(ss.get("opt_slides", 8)),
        )