"""Sticky bottom composer: text box, attach, template, model button, options, Generate.

Edit when: changing the input bar or its popovers.
"""

import streamlit as st

from chatppt.pptx_template import inspect_template_summary
from chatppt.presets import BUILTIN_THEMES, LANGUAGE_OPTIONS, TONE_OPTIONS
from chatppt.ui.common import TPL_COLORS_FALLBACK, TPL_COLORS_FROM_TEMPLATE, TPL_MODE_BUILTIN, TPL_MODE_CUSTOM, _short
from chatppt.ui.generation import handle_send
from chatppt.ui.settings import get_active_model_cfg, open_settings


def render_composer(chat: dict) -> None:
    ss = st.session_state
    # A revision such as "add one more slide" queued a new target: apply it BEFORE the slider is created
    synced = ss.pop("opt_slides_sync", None)
    if synced is not None:
        ss["opt_slides"] = int(synced)
    ss.setdefault("opt_slides", 8)  # single source of the default, so the slider itself carries none
    cid, ver = chat["id"], chat["composer_ver"]
    text_key, file_key = f"composer_text_{cid}_{ver}", f"composer_file_{cid}_{ver}"
    busy = chat["pending"] is not None

    # Labels are derived from last-run widget state so popover buttons reflect current choices
    attached = ss.get(file_key)
    model_cfg = get_active_model_cfg()
    has_key = bool(model_cfg["api_key"])
    model_label = _short(model_cfg["model"], 18) if has_key else "Set up model"
    is_custom_tpl = ss.get("opt_tpl_mode") == TPL_MODE_CUSTOM
    if is_custom_tpl and ss.get("custom_tpl_name"):
        tpl_label = _short(ss["custom_tpl_name"], 13)
    elif is_custom_tpl:
        tpl_label = "Custom .pptx"
    else:
        tpl_label = _short(ss.get("opt_theme", "Theme"), 14)

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
            attach_label = _short(attached.name, 12) if attached else "Attach"
            deck_label = f"{ss.get('opt_slides', 8)} slides"
            send_label = "Send"
            # Column widths follow label length (+ room for the icon) so button text is never cut off
            widths = [max(len(lbl), 8) + 7 for lbl in (attach_label, tpl_label, model_label, deck_label, send_label)]
            c_file, c_tpl, c_model, c_deck, c_send = st.columns(widths, vertical_alignment="center")

            with c_file.popover(attach_label, icon=":material/attach_file:", width="stretch"):
                report = st.file_uploader(
                    "Report file",
                    type=["pdf", "docx", "txt", "md"],
                    key=file_key,
                    help="Attach a PDF, DOCX, TXT, or MD report.",
                )
                st.caption("PDF · DOCX · TXT · MD")

            with c_tpl.popover(tpl_label, icon=":material/palette:", width="stretch"):
                st.segmented_control(
                    "Styling",
                    [TPL_MODE_BUILTIN, TPL_MODE_CUSTOM],
                    default=TPL_MODE_BUILTIN,
                    required=True,
                    format_func={TPL_MODE_BUILTIN: "Built-in theme", TPL_MODE_CUSTOM: "Custom .pptx"}.get,
                    key="opt_tpl_mode",
                )
                if ss.get("opt_tpl_mode") == TPL_MODE_CUSTOM:
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
                            f":material/check_circle: Loaded `{ss['custom_tpl_name']}` · Font: `{info['font']}` · "
                            f"Accent: `{info['accent']}` · BG: `{info['bg']}`"
                        )
                        if st.button("Remove template", icon=":material/delete:", key="clear_tpl_btn", width="stretch"):
                            ss["custom_tpl_bytes"] = None
                            ss["custom_tpl_name"] = None
                            st.rerun()

                    st.segmented_control(
                        "Template card & accent colors",
                        [TPL_COLORS_FROM_TEMPLATE, TPL_COLORS_FALLBACK],
                        default=TPL_COLORS_FROM_TEMPLATE,
                        required=True,
                        format_func={TPL_COLORS_FROM_TEMPLATE: "From template", TPL_COLORS_FALLBACK: "Fallback palette"}.get,
                        help="From template: colors are extracted from the .pptx. Fallback palette: use the palette chosen below.",
                        key="opt_tpl_color_mode",
                    )
                    st.selectbox("Fallback color palette", list(BUILTIN_THEMES.keys()), key="opt_theme")
                else:
                    st.selectbox("Color palette", list(BUILTIN_THEMES.keys()), key="opt_theme")

            model_icon = ":material/smart_toy:" if has_key else ":material/key:"
            if c_model.button(model_label, icon=model_icon, key="model_btn", width="stretch", help="Change provider / model"):
                open_settings()

            with c_deck.popover(deck_label, icon=":material/tune:", width="stretch"):
                num_slides = st.slider("Number of slides", 4, 20, key="opt_slides")
                language = st.selectbox("Output language", LANGUAGE_OPTIONS, key="opt_language")
                tone_label = st.selectbox("Tone", list(TONE_OPTIONS.keys()), key="opt_tone")

            send = c_send.button(
                send_label, type="primary", icon=":material/send:", icon_position="right", width="stretch",
                key=f"send_{cid}", disabled=busy,
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