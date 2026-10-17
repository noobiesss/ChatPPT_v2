"""Send handling and the queued generation run (validation, source reading, outline job).

Edit when: changing what happens after the user presses Generate.
"""

import streamlit as st

from chatppt.extraction import extract_images_from_report, extract_text_from_uploaded_file
from chatppt.outline import generate_smart_outline, mentions_slide_count
from chatppt.presets import CUSTOM_PROVIDER, REVISION_MAX_CHARS
from chatppt.schemas import IMAGE_MODE_UNSPLASH_FIRST
from chatppt.ui.common import _short
from chatppt.ui.settings import open_settings


def handle_send(chat: dict, text: str, report, model_cfg: dict, tone: str, language: str, num_slides: int) -> None:
    """Validate the message, extract source text/images, and queue a generation job."""
    text = (text or "").strip()
    if not text and report is None:
        st.toast("Attach a report or paste some text first.", icon="⚠️")
        return
    if not model_cfg["api_key"]:
        st.toast(f"Add your {model_cfg['provider']} API key first.", icon="🔑")
        open_settings()
        return

    is_revision = (
        report is None and bool(chat["slides"]) and bool(chat["source_text"]) and len(text) <= REVISION_MAX_CHARS
    )

    if is_revision:
        source_for_llm = (
            chat["source_text"]
            + "\n\nUser revision instructions (apply these when designing the deck): "
            + text
        )
    else:
        with st.spinner("Reading your file..."):
            file_text = extract_text_from_uploaded_file(report) if report is not None else ""
            chat["report_images"] = extract_images_from_report(report) if report is not None else []
        source = "\n\n".join(filter(None, [file_text.strip(), text]))
        if not source:
            st.toast("Couldn't read any text from that file.", icon="⚠️")
            return
        chat["source_text"] = source
        source_for_llm = source

    if not chat["messages"]:
        chat["title"] = (report.name if report is not None else text.split("\n")[0])[:40] or "New chat"

    # A revision that itself mentions a slide count (e.g. "make it 6 slides") overrides the slider
    enforce_count = not (is_revision and mentions_slide_count(text))
    slide_label = f"{num_slides} slides" if enforce_count else "slide count per your request"
    meta = f"{slide_label} · {language} · {_short(model_cfg['provider'], 12)} / {_short(model_cfg['model'], 24)}"
    shown = text if len(text) <= 500 else text[:500] + f"… ({len(text):,} characters)"
    if is_revision:
        shown = f"✏️ {shown}"
    chat["messages"].append(
        {
            "role": "user",
            "content": shown,
            "attachment": report.name if report is not None else None,
            "meta": meta,
        }
    )
    chat["pending"] = {
        "provider": model_cfg["provider"],
        "api_key": model_cfg["api_key"],
        "model": model_cfg["model"],
        "base_url": model_cfg["base_url"] if model_cfg["provider"] == CUSTOM_PROVIDER else None,
        "source_text": source_for_llm,
        "num_slides": num_slides,
        "enforce_count": enforce_count,
        "tone": tone,
        "language": language,
    }
    chat["composer_ver"] += 1  # new widget keys => text box and file uploader reset
    st.rerun()


def run_pending(chat: dict, container) -> None:
    """Run the queued outline generation inside the chat area, then rerun to show the result."""
    p = chat["pending"]
    default_img_mode = st.session_state.settings.get("image_mode", IMAGE_MODE_UNSPLASH_FIRST)
    with container:
        with st.chat_message("assistant", avatar="📊"):
            with st.spinner(f"Designing your deck with {p['provider']} · {p['model']}…"):
                try:
                    slides, count_warning = generate_smart_outline(
                        provider=p["provider"],
                        api_key=p["api_key"],
                        source_text=p["source_text"],
                        num_slides=p["num_slides"],
                        tone=p["tone"],
                        language=p["language"],
                        model_name=p["model"],
                        custom_base_url=p["base_url"],
                        enforce_count=p.get("enforce_count", True),
                        default_image_mode=default_img_mode,
                    )
                    chat["slides"] = slides
                    chat["gen"] += 1
                    chat["pptx_bytes"] = None
                    if slides and slides[0].get("action_title"):
                        chat["title"] = slides[0]["action_title"][:40]
                    note = ""
                    if count_warning:
                        note += f"\n\n⚠️ {count_warning}"
                    chat["messages"].append(
                        {
                            "role": "assistant",
                            "content": f"Your outline is ready — **{len(slides)} slides**.{note} "
                            "Review and tweak it below, then render the .pptx.",
                        }
                    )
                except Exception as e:  # noqa: BLE001
                    err = str(e)
                    err = err if len(err) <= 1200 else err[:1200] + "…"
                    chat["messages"].append(
                        {"role": "assistant", "content": f"❌ I couldn't generate the outline.\n\n```\n{err}\n```"}
                    )
                chat["pending"] = None
    st.rerun()
