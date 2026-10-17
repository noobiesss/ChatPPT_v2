"""Send handling and the queued run (question-vs-instruction routing, validation, source reading, outline job).

Edit when: changing what happens after the user presses Send.
"""

import streamlit as st

from chatppt.chat_assistant import route_message
from chatppt.extraction import extract_images_from_report, extract_text_from_uploaded_file
from chatppt.llm_client import TokenBudgetError
from chatppt.outline import generate_smart_outline, mentions_slide_count, parse_slide_count_change, patch_slides
from chatppt.presets import CUSTOM_PROVIDER, REVISION_MAX_CHARS
from chatppt.schemas import IMAGE_MODE_UNSPLASH_FIRST
from chatppt.ui.common import _short
from chatppt.ui.settings import open_settings


def _has_question_mark(text: str) -> bool:
    return "?" in text or "\uff1f" in text


def handle_send(chat: dict, text: str, report, model_cfg: dict, tone: str, language: str, num_slides: int) -> None:
    """Validate the message, prepare the source text, and queue a job (which may turn out to be a plain answer)."""
    text = (text or "").strip()
    if not text and report is None:
        st.toast("Attach a report or paste some text first.", icon="⚠️")
        return
    if not model_cfg["api_key"]:
        st.toast(f"Add your {model_cfg['provider']} API key first.", icon="🔑")
        open_settings()
        return

    has_deck = bool(chat["slides"]) and bool(chat["source_text"])
    short_text_only = report is None and bool(text) and len(text) <= REVISION_MAX_CHARS
    is_revision = has_deck and short_text_only
    prev_count = len(chat["slides"] or [])  # chat["slides"] is None until the first deck exists

    # A relative edit such as "add one more slide" means: new target = current deck size + delta
    relative_target = parse_slide_count_change(text, prev_count) if is_revision else None
    if relative_target is not None:
        num_slides = relative_target
    # A plain instruction ("add one more slide") skips the routing call; a question mark means "ask the model"
    force_revise = relative_target is not None and not _has_question_mark(text)
    # Short text-only messages are routed: the model decides between answering and (re)generating the deck
    route = short_text_only and not force_revise
    commit_source = None

    if is_revision:
        source_for_llm = (
            chat["source_text"]
            + "\n\nUser revision instructions (apply these when designing the deck): "
            + text
        )
        if relative_target is not None:
            deck_titles = "\n".join(
                f"  {i + 1}. {s.get('action_title', '')}" for i, s in enumerate(chat["slides"])
            )
            source_for_llm += (
                f"\n\nThe current deck has {prev_count} slides:\n{deck_titles}\n"
                f"The revised deck MUST have exactly {relative_target} slides. Keep ALL existing slides in their "
                "current order and apply ONLY the requested change. A requested thank-you / closing / Q&A slide must use "
                "the `closing` layout and be the LAST slide; a new conclusion/summary slide goes right before it "
                "(or last if there is no closing slide); other new slides go where they fit best. "
                "Do not replace or merge existing slides."
            )
    elif short_text_only:
        # No deck yet: if the model decides this is not a question, the short text itself becomes the source
        source_for_llm = text
        commit_source = text
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
    enforce_count = relative_target is not None or not (is_revision and mentions_slide_count(text))
    slide_label = f"{num_slides} slides" if enforce_count else "slide count per your request"
    meta_model = f"{_short(model_cfg['provider'], 12)} / {_short(model_cfg['model'], 24)}"
    meta_full = f"{slide_label} · {language} · {meta_model}"
    shown = text if len(text) <= 500 else text[:500] + f"… ({len(text):,} characters)"
    if is_revision and not route:
        shown = f"✏️ {shown}"
    chat["messages"].append(
        {
            "role": "user",
            "content": shown,
            "attachment": report.name if report is not None else None,
            # While routing, the slide count is not shown: the message may turn out to be a question
            "meta": meta_model if route else meta_full,
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
        "route": route,
        "user_text": text,
        "is_revision": is_revision,
        "meta_revise": meta_full,
        "meta_model": meta_model,
        "commit_source": commit_source,
        "sync_slides": relative_target,
    }
    chat["composer_ver"] += 1  # new widget keys => text box and file uploader reset
    st.rerun()


def _prepare_generation(chat: dict, p: dict) -> None:
    """Apply the side effects that only make sense once we know the message is a deck instruction."""
    if p.get("commit_source"):
        chat["source_text"] = p["commit_source"]
        chat["report_images"] = []
    if p.get("sync_slides") is not None:
        # The composer applies this before its slider is created, so the slider follows the new target
        st.session_state["opt_slides_sync"] = int(p["sync_slides"])
    if p.get("route") and chat["messages"] and chat["messages"][-1]["role"] == "user":
        last = chat["messages"][-1]
        last["meta"] = p["meta_revise"]
        if p.get("is_revision") and not last["content"].startswith("✏️"):
            last["content"] = f"✏️ {last['content']}"


def _run_patch(chat: dict, p: dict, targets: list) -> bool:
    """Rewrite only the target slides in place. Returns False on any failure so the caller regenerates instead."""
    nums = ", ".join(str(n) for n in targets)
    label = f"slide {nums}" if len(targets) == 1 else f"slides {nums}"
    with st.spinner(f"Updating {label} with {p['provider']} · {p['model']}…"):
        try:
            slides = patch_slides(
                provider=p["provider"],
                api_key=p["api_key"],
                model_name=p["model"],
                custom_base_url=p["base_url"],
                slides=chat["slides"],
                targets=targets,
                instruction=p["user_text"],
                source_text=chat["source_text"],
                tone=p["tone"],
                language=p["language"],
            )
        except Exception:  # noqa: BLE001 - fall back to the full regeneration
            return False
    chat["slides"] = slides
    chat["gen"] += 1  # new editor widget keys => the editor shows the patched slides
    chat["pptx_bytes"] = None
    if 1 in targets and slides[0].get("action_title"):
        chat["title"] = slides[0]["action_title"][:40]
    if chat["messages"] and chat["messages"][-1]["role"] == "user":
        last = chat["messages"][-1]
        last["meta"] = f"{label} · {p.get('meta_model', '')}"
        if not last["content"].startswith("✏️"):
            last["content"] = f"✏️ {last['content']}"
    chat["messages"].append(
        {
            "role": "assistant",
            "content": f"Updated **{label}**; all other slides are unchanged. Review it below, then render the .pptx.",
            "meta": "✏️ Targeted edit",
        }
    )
    return True


def run_pending(chat: dict, container) -> None:
    """Run the queued job in the chat area: answer a question, or generate the outline; then rerun."""
    p = chat["pending"]
    default_img_mode = st.session_state.settings.get("image_mode", IMAGE_MODE_UNSPLASH_FIRST)
    with container:
        with st.chat_message("assistant", avatar="📊"):
            if p.get("route"):
                with st.spinner("Thinking…"):
                    decision = route_message(
                        provider=p["provider"],
                        api_key=p["api_key"],
                        model_name=p["model"],
                        base_url=p["base_url"],
                        user_text=p["user_text"],
                        source_text=chat["source_text"],
                        slides=chat["slides"],
                        history=chat["messages"][:-1],  # everything before the message being answered
                    )
                if decision.action == "answer":
                    chat["messages"].append(
                        {
                            "role": "assistant",
                            "content": decision.reply,
                            "meta": "💬 Answer · outline unchanged" if chat["slides"] else "💬 Answer",
                        }
                    )
                    chat["pending"] = None
                    st.rerun()
                # A slide-count change ("add one more slide?") always needs the full regeneration
                if decision.action == "patch" and chat["slides"] and p.get("sync_slides") is None:
                    if _run_patch(chat, p, decision.slides):
                        chat["pending"] = None
                        st.rerun()
            _prepare_generation(chat, p)
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
                except TokenBudgetError as e:
                    # Readable explanation (plan limit vs. request size) instead of the raw HTTP body
                    chat["messages"].append({"role": "assistant", "content": f"⚠️ {e}"})
                except Exception as e:  # noqa: BLE001
                    err = str(e)
                    err = err if len(err) <= 1200 else err[:1200] + "…"
                    chat["messages"].append(
                        {"role": "assistant", "content": f"❌ I couldn't generate the outline.\n\n```\n{err}\n```"}
                    )
                chat["pending"] = None
    st.rerun()