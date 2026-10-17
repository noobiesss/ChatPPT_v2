# ChatPPT V2 — Project State (handoff for the next AI)

> **How to use:** paste/upload this file in a new chat together with ONLY the module file(s) the task needs (see 2 "Task -> files").
> Read it fully before editing. At the end of every session update **Last updated, Commit Log, Pending, Changelog** and return the file.

Last updated: 2026-10-17 (minimal opt-in `closing` thank-you layout)
Commit Log: `feat(layout): add minimal closing (thank-you) layout, opt-in only; last-slide enforcement, trim-safe, editor + renderer support`

---

## 0. User preferences (MUST FOLLOW)
- Reply to the user in **Chinese (简体中文)**. All **code and code comments strictly in English**.
- Deliver only the changed files as complete, ready-to-run modules (keep CRLF line endings when patching programmatically). Don't re-output untouched modules.
- Never rename/move a function or constant unless required; if you do, update every importer and section 2.
- After every change run `python smoke_test.py` (offline, no API keys) and report the result; give an updated `project_state.md` when state changes.

## 1. Goal & stack
"ChatPPT V2": Streamlit web app that turns reports (`.pdf/.docx/.txt/.md`) or pasted text into consulting-grade `.pptx` decks, and also answers questions about the source/deck in the same chat. Python 3, Streamlit >=1.46, `google-genai`, Pydantic v2, `python-pptx`, `pypdf`, `python-docx`, `Pillow`, `urllib`. Files: `app.py`, `chatppt/` package, `smoke_test.py`, `requirements.txt`, `README.md`, `.gitignore`, `.streamlit/config.toml`. Run: `streamlit run app.py`.

## 2. File map
Dependencies only point DOWN: `schemas/presets` -> `extraction/prompts/llm_client` -> `outline/images/chat_assistant` -> `pptx_template` -> `pptx_render` -> `pptx_build` -> `ui/*` -> `app.py`. Core modules never import UI code (may use `st.cache_data`).

| File | Holds |
| :--- | :--- |
| `app.py` | `main`: page config + flow (state, css, sidebar, messages, deck block, composer, `run_pending`) |
| `chatppt/schemas.py` | `LayoutType` (cover, section_divider, bullet_cards, three_cards, two_column, kpi_metrics, chart_slide, image_split, **closing**), `ChartType`, `SlideSpec`, `PresentationOutline`, sub-models, `IMAGE_MODE_*` |
| `chatppt/presets.py` | `LLM_PROVIDERS`, `BUILTIN_THEMES`, `TONE_OPTIONS`, `LANGUAGE_OPTIONS`, `CUSTOM_PROVIDER`, `REVISION_MAX_CHARS` |
| `chatppt/extraction.py` | `extract_text_from_uploaded_file`, `extract_images_from_report` |
| `chatppt/prompts.py` | `SYSTEM_INSTRUCTION` (outline design + routing rules), `ASSISTANT_SYSTEM_INSTRUCTION` (Q&A/router) |
| `chatppt/llm_client.py` | model discovery (`fetch_live_models`, `resolve_model_options`), JSON helpers (`extract_clean_json_str`, `repair_json_str`, `call_openai_compatible_json`), plain-text calls (`call_openai_compatible_text`, `call_gemini_text`, `strip_think_tags`) |
| `chatppt/chat_assistant.py` | `route_message`, `build_assistant_prompt`, `build_deck_digest`, `parse_assistant_reply`, `AssistantDecision` |
| `chatppt/outline.py` | `generate_smart_outline`, `_request_outline`, `_build_count_correction`, `_trim_slides_to_count`, `mentions_slide_count`, `parse_slide_count_change`, `sanitize_outline_dict` |
| `chatppt/images.py` | Unsplash fetch, AI fallback, `resolve_slide_image_bytes`, `add_picture_fitted`, `_derive_fallback_keywords` |
| `chatppt/pptx_template.py` | custom `.pptx` inheritance (colors, fonts, backgrounds, decoration stripping, prototypes) |
| `chatppt/pptx_render.py` | drawing helpers + 9 layout renderers (`render_closing_slide` etc.) |
| `chatppt/pptx_build.py` | `build_presentation_pptx` (deck assembly, layout dispatch) |
| `chatppt/ui/state.py` | `init_state`, chat create/delete |
| `chatppt/ui/css.py` | `APP_CSS` (theme-neutral, no hard-coded page colors) |
| `chatppt/ui/settings.py` | settings dialog, `get_active_model_cfg`, `open_settings` |
| `chatppt/ui/sidebar.py`, `messages.py` | sidebar; hero + chat messages (renders Markdown + `meta` caption) |
| `chatppt/ui/outline_editor.py` + `editor_text.py` | per-slide editor + text<->data helpers |
| `chatppt/ui/export.py` | deck block: editor + render/download (user presses a separate render button) |
| `chatppt/ui/generation.py` | `handle_send`, `_prepare_generation`, `run_pending` |
| `chatppt/ui/composer.py` + `common.py` | bottom input bar (Attach, Template, Model, `⚙ N slides` popover, **Send ➤**) |

**Task -> files to upload:**
- Outline quality / prompt wording -> `prompts.py` (+ `outline.py`)
- Q&A / question-vs-instruction routing -> `chat_assistant.py`, `prompts.py`, `ui/generation.py`
- Slide count / validation -> `outline.py` (+ `schemas.py`, `ui/generation.py`)
- Provider/model/JSON problems -> `llm_client.py` (+ `presets.py`)
- Template looks wrong -> `pptx_template.py`; a layout's look -> `pptx_render.py`
- NEW layout -> `schemas.py`, `prompts.py`, `outline.py`, `pptx_render.py`, `pptx_build.py`, `ui/outline_editor.py`
- Images -> `images.py` (+ `ui/settings.py`); settings -> `ui/settings.py`, `ui/state.py`, `presets.py`
- Input bar / send behavior -> `ui/composer.py`, `ui/generation.py`; styling -> `ui/css.py`

**Rules:** (1) keep layering above; (2) session keys are created in `ui/state.py::init_state`, settings read via `st.session_state.settings[...]`; (3) list new cross-module functions in this table; (4) finish with `python smoke_test.py` and add a check for bugs that could regress.

## 3. Key behaviors (what the code does today)
- **Send flow** (`handle_send` -> `run_pending`): text-only messages <= `REVISION_MAX_CHARS` are *routed*: `route_message` makes one plain-text LLM call (protocol first line `ACTION: ANSWER` + reply, or `ACTION: REVISE`; context = source (12k chars) + deck digest + last 6 messages). ANSWER -> assistant bubble (meta "💬 Answer · outline unchanged"), deck and slider untouched. REVISE -> outline generation. Attached file, long text, or a plain relative-count instruction without `?` skip routing. Router failure -> REVISE. Side effects (commit source for no-deck case, `opt_slides_sync`, user-message meta/✏️) happen in `_prepare_generation`, only after REVISE is known.
- **Slide count**: `generate_smart_outline(enforce_count=True)` = prompt rule + up to 2 corrective retries + deterministic trim. `mentions_slide_count` (absolute, "6 slides") disables enforcement. `parse_slide_count_change` (relative: "add one more slide", "add a thank you slide", "remove the last slide", "加一页", "删掉最后一页"; clamp 4-20) sets target = current deck size + delta and takes precedence; the revision prompt then lists the current slide titles and says keep all existing slides. `render_composer` applies `opt_slides_sync` to the slider before it is created.
- **closing layout** (opt-in only): big centered phrase + accent bar + optional one-line presenter note; no cards/points/takeaway/contact details. The prompt forbids adding it unprompted (default decks end with a content slide). `sanitize_outline_dict` keeps one closing slide, moves it last, clears its takeaway; `_trim_slides_to_count` protects closing + the conclusion before it; closing is never auto-promoted to `image_split`; `pptx_build` uses role `cover` for it so custom templates give it the cover background.
- **Sanitizer**: equalizes chart categories/values, fills missing structures, forces slide 1 = cover, guarantees >=1 `image_split` on 4+ slide decks, sets `image_source_mode`.
- **Images**: Option 1 `Unsplash Photo First (Fallback to AI)` (default; 2-word retry; then Imagen 3 -> Pollinations); Option 2 `Reserve Photo Placeholder`. Always insert via `add_picture_fitted`.
- **Custom template**: inherits widescreen-scaled backgrounds, sample-slide fonts, theme/sampled colors; strips master/layout/slide decorations; WCAG white-text guard.
- **LLM quirks**: reasoning models (`gpt-oss`, Groq `qwen3`, `deepseek-reasoner`) use 8192 tokens, `reasoning_effort` low/none, `<think>` stripping, `message.reasoning` fallback, retry without optional params.
- **UI**: theme-neutral CSS; `.streamlit/config.toml` defines `[theme.light]`/`[theme.dark]` (accent #C96442; never put colors in the bare `[theme]` table, it locks one mode); 50 MB upload limit there.

## 4. Known constraints / gotchas
- Gemini structured output forbids arbitrary `Dict`: chart data stays parallel lists.
- `chat["slides"]` is `None` until the first deck exists: use `bool(...)` or `or []`, never bare `len()`.
- The `opt_slides` slider has NO default value in `st.slider`; the default 8 is set once via `ss.setdefault("opt_slides", 8)` (avoids Streamlit's "default value + Session State API" warning). Session state can only be set before the widget is created in a run.
- Session state is per browser session; refresh clears chats and keys.
- Patching files: files use CRLF; preserve it.

## 5. Pending
- [ ] User to verify: default deck has NO thank-you slide; "add a thank you slide" -> 9 slides, last is a big centered "Thank You" (also with a custom .pptx template).
- [ ] User to verify Q&A: after a deck exists ask "why three cards on slide 4?" -> answer bubble, outline unchanged; "make the titles shorter" -> regenerates; first message "what can you do?" -> answer, no deck.
- [ ] User to verify: 8 slides, then "add one more slide conclusion" -> 9 slides (old 8 kept) and slider shows 9; no `opt_slides` warning in the terminal.
- [ ] User to check light/dark/system look after CSS and theme changes.
- [ ] Known edge: if a closing slide already exists and the user asks for another, the model may add a different slide (not handled).
- [ ] Optional: deploy (GitHub + Streamlit Community Cloud / Railway); further UI/render enhancements on request.

## 6. Changelog (condensed)
| Date | Change |
| :--- | :--- |
| 2026-10-16/17 | V2 core: 9 layouts (8 + closing), native charts, multi-provider BYOK + live model discovery, JSON repair, Claude-style chat UI; exact slide-count enforcement; custom `.pptx` template inheritance; 2-option image mode; reasoning-model fixes; monolith split into `chatppt/` package + `smoke_test.py` |
| 2026-10-17 | UI/theme: frosted composer dock works in light+dark, accent #C96442, footer note; `[theme.light]/[theme.dark]` in `.streamlit/config.toml` (System/Light/Dark selector) |
| 2026-10-17 | Fix: "add one more slide" replaced slide 8 -> relative count parsing (`parse_slide_count_change`), keep-existing-slides prompt, slider sync (`opt_slides_sync`) |
| 2026-10-17 | Feature: single **Send ➤** button with LLM routing between Q&A and deck revision (`chat_assistant.py`, plain-text calls in `llm_client.py`); fixed `None` slides crash and the `opt_slides` Streamlit warning |
| 2026-10-17 | Feature: opt-in minimal `closing` layout (schemas, prompts, outline, pptx_render, pptx_build, outline_editor); parser also understands "add/remove thank you slide"; smoke test has 9 layouts + `check_closing_slide` + `check_assistant` |