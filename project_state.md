# ChatPPT V2 — Project State (handoff for the next AI)

> **How to use:** paste/upload this file in a new chat together with ONLY the module file(s) the task needs (see 2 "Task -> files").
> Read it fully before editing. At the end of every session update **Last updated, Commit Log, Pending, Changelog** and return the file.

Last updated: 2026-10-17 (timeline_flow + comparison_table layouts, Report Images First image mode)
Commit Log: `feat(layout,images): timeline_flow and comparison_table layouts; report-figure image mode with auto assignment and per-slide picker`

---

## 0. User preferences (MUST FOLLOW)
- Reply to the user in **Chinese (简体中文)**. All **code and code comments strictly in English**.
- Deliver only the changed files as complete, ready-to-run modules (keep CRLF line endings when patching programmatically). Don't re-output untouched modules.
- Never rename/move a function or constant unless required; if you do, update every importer and section 2.
- After every change run `python smoke_test.py` (offline, no API keys) and report the result; give an updated `project_state.md` when state changes.

## 1. Goal & stack
"ChatPPT V2": Streamlit web app that turns reports (`.pdf/.docx/.txt/.md`) or pasted text into consulting-grade `.pptx` decks, and also answers questions about the source/deck in the same chat. Python 3, Streamlit >=1.57 (needed for `width="stretch"`; `use_container_width` is deprecated, never use it), `google-genai`, Pydantic v2, `python-pptx`, `pypdf`, `python-docx`, `Pillow`, `urllib`. Files: `app.py`, `chatppt/` package, `assets/` (logo), `smoke_test.py`, `requirements.txt`, `README.md`, `.gitignore`, `.streamlit/config.toml`. Run: `streamlit run app.py`.

## 2. File map
Dependencies only point DOWN: `schemas/presets` -> `extraction/retrieval/prompts/llm_client` -> `chat_assistant` -> `outline/images` -> `pptx_template` -> `pptx_render` -> `pptx_build` -> `ui/*` -> `app.py`. Core modules never import UI code (may use `st.cache_data`).

| File | Holds |
| :--- | :--- |
| `app.py` | `main`: page config + flow (state, css, sidebar, messages, deck block, composer, `run_pending`) |
| `chatppt/schemas.py` | `LayoutType` (cover, section_divider, bullet_cards, three_cards, two_column, kpi_metrics, chart_slide, image_split, closing, **timeline_flow**, **comparison_table**), `ChartType`, `SlideSpec` (+ `timeline_steps`, `table_headers`, `table_rows`), `TimelineStep`, `PresentationOutline`, `SlidePatch` (targeted-edit response), sub-models, `IMAGE_MODE_*` (3 modes incl. `IMAGE_MODE_REPORT_FIRST`) |
| `chatppt/presets.py` | `LLM_PROVIDERS`, `BUILTIN_THEMES`, `TONE_OPTIONS`, `LANGUAGE_OPTIONS`, `CUSTOM_PROVIDER`, `REVISION_MAX_CHARS`, `PATCH_MAX_SLIDES` |
| `chatppt/extraction.py` | `extract_text_from_uploaded_file`, `extract_images_from_report` (filtered, deduped, items `{name, bytes, width, height}`), `prepare_report_image` |
| `chatppt/retrieval.py` | `select_relevant_context` (BM25 excerpts of a long source within a char budget), `bm25_scores`, `split_into_chunks`, `tokenize` (words + CJK bigrams) |
| `chatppt/prompts.py` | `SYSTEM_INSTRUCTION` (outline design + routing rules), `ASSISTANT_SYSTEM_INSTRUCTION` (Q&A/router) |
| `chatppt/llm_client.py` | model discovery (`fetch_live_models`, `resolve_model_options`), JSON helpers (`extract_clean_json_str`, `repair_json_str`, `call_openai_compatible_json`; the last two take an optional schema, default `PresentationOutline`), plain-text calls (`call_openai_compatible_text`, `call_gemini_text`, `strip_think_tags`) |
| `chatppt/chat_assistant.py` | `route_message`, `build_assistant_prompt`, `build_deck_digest`, `parse_assistant_reply`, `parse_patch_targets`, `AssistantDecision` (action answer/patch/revise + `slides`) |
| `chatppt/outline.py` | `generate_smart_outline`, `patch_slides`, `_request_outline` (optional `response_model`), `_sanitize_slide`, `normalize_timeline_steps`, `normalize_table`, `TIMELINE_MAX_STEPS`, `TABLE_MAX_COLS/ROWS`, `_build_count_correction`, `_trim_slides_to_count`, `mentions_slide_count`, `parse_slide_count_change`, `sanitize_outline_dict` |
| `chatppt/deck_ops.py` | no-LLM outline ops (pure): `push_history`, `undo`, `redo`, `can_undo/redo` (`HISTORY_MAX` 20, deep snapshots in `chat["slides_history"]` / `["slides_future"]`); `move_slide`, `insert_slide`, `delete_slide`, `can_move/insert/delete`, `new_blank_slide` |
| `chatppt/project_io.py` | `export_project` / `load_project` (`.chatppt.json`, format `chatppt-project` v1; no API keys; slides re-validated via `SlideSpec` + `_sanitize_slide`), `ProjectFileError`, `OPTION_KEYS`, `load_safe_settings`, `load_options`, `encode_images` / `decode_images`, `encode_messages` / `load_messages`, `load_slides` (shared with the chat store) |
| `chatppt/prefs.py` | browser-remembered preferences (pure): `collect_prefs` (keys only if `settings["remember_keys"]`), `prefs_json` (stable), `parse_prefs` (validates via `project_io.load_safe_settings` / `load_options`, drops unknown providers and oversized secrets) |
| `chatppt/ui/browser_prefs.py` | invisible CCv2 inline component bridging to `localStorage["chatppt.prefs.v1"]`: `sync_browser_prefs` (mounted at the end of the sidebar every run), `forget_browser_prefs`, `resume_browser_prefs` |
| `chatppt/chat_store.py` | chats kept in the browser (pure): 3 records per chat (`<id>/chat`, `/source`, `/images`), `part_signature` (cheap per-part change detection), `part_payload`, `restore_chats` (validated, newest first, `MAX_STORED_CHATS` 30), `is_storable`, `MAX_IMAGES_BYTES` 15 MB |
| `chatppt/ui/browser_chats.py` | invisible CCv2 inline component bridging to IndexedDB `chatppt/records`: `sync_browser_chats` (end of `app.main`, after editor + composer, every run), `erase_browser_chats` |
| `chatppt/preview.py` | `render_slide_preview` (HTML/CSS 16:9 wireframe, same inch geometry as `pptx_render`, `cqw` font units) + overflow estimate (`text_height_pt`, `overflows`), `PREVIEW_CSS` |
| `chatppt/images.py` | Unsplash fetch, AI fallback, `assign_report_images`, `resolve_slide_image_bytes`, `add_picture_fitted`, `_derive_fallback_keywords` |
| `chatppt/pptx_template.py` | custom `.pptx` inheritance (colors, fonts, backgrounds, decoration stripping, prototypes) |
| `chatppt/pptx_render.py` | drawing helpers + 11 layout renderers (`render_closing_slide`, `render_timeline_flow_slide`, `render_comparison_table_slide` etc.) |
| `chatppt/pptx_build.py` | `build_presentation_pptx` (deck assembly, layout dispatch) |
| `chatppt/ui/state.py` | `init_state`, `new_chat` (chat dict factory), chat create/delete, `mark_slides_changed` (gen+1, drop pptx, slider sync), `restore_project` (callback only) |
| `chatppt/ui/css.py` | `APP_CSS` (theme-neutral, no hard-coded page colors) |
| `chatppt/ui/settings.py` | settings dialog, `get_active_model_cfg`, `open_settings` |
| `chatppt/ui/sidebar.py` (+ Open project popover), `messages.py` | sidebar; hero + chat messages (renders Markdown + `meta` caption) |
| `chatppt/ui/outline_editor.py` + `editor_text.py` | per-slide editor (+ `_render_report_image_picker`) + text<->data helpers (points, KPIs, `*_steps_*`, `*_table_*`) |
| `chatppt/ui/export.py` | deck block: Undo / Redo / Save project toolbar (`_render_deck_toolbar`), editor, render/download |
| `chatppt/ui/generation.py` | `handle_send`, `_prepare_generation`, `_run_patch`, `run_pending` |
| `assets/logo.png`, `assets/logo_square.png` | pixel-art brand logo (520x440 on a 20 px grid) + square version; paths `LOGO_PATH`, `LOGO_ICON_PATH` in `ui/common.py` |
| `chatppt/ui/composer.py` + `common.py` | bottom input bar (Attach, Template, Model, `N slides` popover, **Send**); `common.py` holds `_short`, `ASSISTANT_AVATAR`, `TPL_*` option values, `IMAGE_MODE_LABELS/HELP` |

**Task -> files to upload:**
- Outline quality / prompt wording -> `prompts.py` (+ `outline.py`)
- Q&A / question-vs-instruction routing / targeted patch -> `chat_assistant.py`, `prompts.py`, `ui/generation.py` (+ `outline.py`)
- Long-source context selection -> `retrieval.py` (+ `chat_assistant.py`)
- Slide count / validation -> `outline.py` (+ `schemas.py`, `ui/generation.py`)
- Provider/model/JSON problems -> `llm_client.py` (+ `presets.py`)
- Template looks wrong -> `pptx_template.py`; a layout's look -> `pptx_render.py`
- NEW layout -> `schemas.py`, `prompts.py`, `outline.py`, `pptx_render.py`, `pptx_build.py`, `ui/outline_editor.py`
- Images -> `images.py` (+ `ui/settings.py`); settings -> `ui/settings.py`, `ui/state.py`, `presets.py`
- Input bar / send behavior -> `ui/composer.py`, `ui/generation.py`; styling -> `ui/css.py`

**Rules:** (1) keep layering above; (2) session keys are created in `ui/state.py::init_state`, settings read via `st.session_state.settings[...]`; (3) list new cross-module functions in this table; (4) finish with `python smoke_test.py` and add a check for bugs that could regress.

## 3. Key behaviors (what the code does today)
- **Send flow** (`handle_send` -> `run_pending`): text-only messages <= `REVISION_MAX_CHARS` are *routed*: `route_message` makes one plain-text LLM call (protocol first line `ACTION: ANSWER` + reply, `ACTION: PATCH: <n,...>`, or `ACTION: REVISE`; context = source (<=12k chars, see Retrieval) + deck digest + last 6 messages). ANSWER -> assistant bubble (meta "💬 Answer · outline unchanged"), deck and slider untouched. PATCH -> `_run_patch`. REVISE -> outline generation. Attached file, long text, or a plain relative-count instruction without `?` skip routing. Router failure -> REVISE. Side effects (commit source for no-deck case, `opt_slides_sync`, user-message meta/✏️) happen in `_prepare_generation`, only after REVISE is known.
- **Targeted patch** (`ACTION: PATCH`): the router names <= `PATCH_MAX_SLIDES` (5) existing slides; numbers outside the deck, none, or too many -> REVISE. Never used for add/remove/reorder/count/deck-wide changes, and ignored when a relative count edit is pending (`sync_slides`). `patch_slides` sends the whole-deck digest, only the target slides' visible fields (`_slide_for_prompt`), and BM25 source excerpts; expects `SlidePatch` with one slide per target (matched by `slide_number`, else by order), sanitizes each with `_sanitize_slide`, keeps the old `image_source_mode`, turns a non-last `closing` into `bullet_cards`. All other slide dicts are returned as the same objects (100% unchanged). `gen` += 1, `pptx_bytes` reset, message "Updated slide N" (meta "✏️ Targeted edit"). Any patch failure falls back to the full regeneration.
- **Retrieval** (`retrieval.py`, no deps): sources > 12k chars are split into ~1000-char line chunks; BM25 (k1 1.5, b 0.75) over word tokens + CJK bigrams. Picks: chunk 0 (title/intro) + best hits + their neighbors, then fills leftover budget from the document head; output in document order with `[...]` gaps. No match -> head of document (old behavior). Index cached per source (`lru_cache`). Router query = new message + previous user message; patch query = instruction + target-slide digest. Full outline generation still gets the whole source.
- **Slide count**: `generate_smart_outline(enforce_count=True)` = prompt rule + up to 2 corrective retries + deterministic trim. `mentions_slide_count` (absolute, "6 slides") disables enforcement. `parse_slide_count_change` (relative: "add one more slide", "add a thank you slide", "remove the last slide", "加一页", "删掉最后一页"; clamp 4-20) sets target = current deck size + delta and takes precedence; the revision prompt then lists the current slide titles and says keep all existing slides. `render_composer` applies `opt_slides_sync` to the slider before it is created.
- **closing layout** (opt-in only): big centered phrase + accent bar + optional one-line presenter note; no cards/points/takeaway/contact details. The prompt forbids adding it unprompted (default decks end with a content slide). `sanitize_outline_dict` keeps one closing slide, moves it last, clears its takeaway; `_trim_slides_to_count` protects closing + the conclusion before it; closing is never auto-promoted to `image_split`; `pptx_build` uses role `cover` for it so custom templates give it the cover background.
- **timeline_flow**: 2-5 steps (`stage_label`, `title`, `description`); numbered nodes on a connector line, stage label above, card below; fonts shrink with step count. Sanitizer: model wrote `points` instead of steps -> steps from points; < 2 steps and no points -> `bullet_cards`.
- **comparison_table**: `table_headers` (2-5) + `table_rows` (List[List[str]], <= 6; Gemini-safe, no Dict). `normalize_table` pads/cuts rows to the header count, drops empty rows; unusable -> `bullet_cards`. Rendered as a NATIVE python-pptx table: accent header (white-contrast guard), zebra `card_bg`/`bg_content` rows, bold first column, fonts set in the renderer (`apply_slide_fonts` skips tables). Renderers fall back to bullet cards on bad data. Editor formats: `Stage | Title | Description` per line; table = header line + rows, cells split by `|`.
- **Remembered settings / API keys** (nothing on the server): per browser in `localStorage["chatppt.prefs.v1"]`. First run of a browser session: the component reads storage once (`setStateValue("loaded")`) -> `_on_loaded` callback applies it before widgets exist (settings, `opt_*` options, keys via setdefault so a key typed this session wins). No write happens before that load. Later runs: write only when `prefs_json(collect_prefs(...))` changed; the latest versioned op is re-sent every run and JS skips applied versions (per `_session_nonce`). Settings dialog section 3: "Remember my API keys on this browser" (default on; off = payload has no keys at all), "Forget this browser" (clear + pause until next Save). Storage blocked -> `_prefs_error` warning. Verified end to end in headless Chrome (restore after refresh, untick removes keys, forget clears and nothing is re-saved).
- **Chats kept in the browser** (IndexedDB, nothing on the server): first run of a browser session reads all records once -> `_on_loaded` callback restores them (replaces the untouched startup chat, opens the most recent). Later runs: only the ACTIVE chat is fingerprinted per part; changed parts are queued as versioned writes (latest per record). The browser applies them in the background WITHOUT reporting back (a value sent up from a component costs a full rerun): each write is re-sent for `RESEND_RUNS` (3) runs, the JS merges ops per session, skips applied versions and drains writes that arrive while busy. So an edit costs exactly one rerun (measured). The bridge MUST be mounted after the outline editor: mounted in the sidebar it fingerprinted the chat before the editor wrote that run's edit, and plain text edits were not saved (fixed, measured end to end). Deleted sidebar chats -> `delete_chat`; > 30 chats -> oldest dropped from storage. Not saved: empty chats, chats with a pending job; figures > 15 MB. Setting `remember_chats` (Settings section 3, default on; off or Forget -> `clear`). Verified end to end in headless Chrome (restore on reopen, slide delete persisted, chat delete erased).
- **History / projects**: `push_history` runs before every AI outline replacement (generation, patch) and every structure edit; text edits are not snapshotted on their own (they are inside the next snapshot). Undo/Redo and structure buttons are applied AFTER the editor read all fields (in-flight text kept), then `mark_slides_changed` + rerun. Save project = lazy `download_button(data=callable)` (settings captured first; no Session State in that thread). Open project (sidebar) = `on_click` callback -> new chat; restores safe settings + `opt_slides/language/tone/theme` (not template mode: template bytes are not saved). No server-side autosave on purpose: a deployed app has no stable per-user identity across a refresh, and a shared file would leak decks between users.
- **Structure edits** (editor, no tokens): per slide up / down / insert / delete icon buttons in a column beside the expander; cover pinned first, closing pinned last, insert never after closing, min 2 / max 20 slides; the moved / new slide is expanded (`chat["focus_slide"]`).
- **Preview**: each expander (`key`, `on_change="rerun"`) draws `render_slide_preview` only while `.open`, after its fields were read (live); theme = built-in palette or cached `template_theme_for_preview` (pptx_template). Expander icon `:material/warning:` when the overflow estimate fires; captions list the overflowing boxes. Keep `preview.py` numbers in sync with `pptx_render.py`.
- **Sanitizer**: equalizes chart categories/values, fills missing structures, forces slide 1 = cover, guarantees >=1 `image_split` on 4+ slide decks, sets `image_source_mode`.
- **Images**: Option 1 `Unsplash Photo First (Fallback to AI)` (default; 2-word retry; then Imagen 3 -> Pollinations); Option 2 `Reserve Photo Placeholder`; Option 3 `Report Images First (Fallback to Unsplash / AI)`. Always insert via `add_picture_fitted`.
- **Report figures**: `extract_images_from_report` keeps pictures >= 200 px short side and >= 52k px area (drops logos/icons), dedupes by SHA-1, converts non-JPEG/PNG to JPEG (PNG if alpha), downscales > 2400 px, max 30, document order. `assign_report_images` (pure): explicit per-slide `report_image_index` wins, other Option-3 `image_split` slides get the remaining figures in order; slides without one fall back to Option 1. `pptx_build` passes the assignment to the resolver; editor shows a `Report figure` select (`Auto` or a figure) + thumbnail of what will be used. `report_image_index` is a UI-only slide-dict key (not in `SlideSpec`); `patch_slides` keeps it.
- **Custom template**: inherits widescreen-scaled backgrounds, sample-slide fonts, theme/sampled colors; strips master/layout/slide decorations; WCAG white-text guard.
- **Token limits** (`_post_chat_completion`, used by both OpenAI-compatible calls): providers like Groq free tier count input + requested output cap against tokens-per-minute (gpt-oss-120b: 8000; our cap is 8192, fixed prompt + schema is ~3.1k tokens). On 413/429 with `Limit X, [Used Y,] Requested Z`: request too big only because of the cap -> retry with cap = room if >= 1500; minute budget used up -> wait once (<= 30 s); else `TokenBudgetError` (readable Markdown, shown as a ⚠️ chat message, never retried by `_request_outline`).
- **LLM quirks**: reasoning models (`gpt-oss`, Groq `qwen3`, `deepseek-reasoner`) use 8192 tokens, `reasoning_effort` low/none, `<think>` stripping, `message.reasoning` fallback, retry without optional params.
- **UI conventions** (Streamlit skill best practices): `width="stretch"` (never `use_container_width`); 2-5 options -> `st.segmented_control(required=True, format_func=...)` keeping the old stored values (template mode/colors constants and `IMAGE_MODE_LABELS` in `ui/common.py`); Material icons via `icon=":material/...:"` (not emoji; keeps composer width math on plain labels); sentence-case labels; `ASSISTANT_AVATAR` (= square logo), `EDIT_MARK` (generation) are the shared markers. Logo: `st.logo` in the sidebar (square icon when collapsed), `page_icon`, chat avatar, hero (data URI); CSS keeps it `image-rendering: pixelated`.
- **UI**: theme-neutral CSS; `.streamlit/config.toml` defines `[theme.light]`/`[theme.dark]` (accent #C96442; never put colors in the bare `[theme]` table, it locks one mode); 50 MB upload limit there.

## 4. Known constraints / gotchas
- Gemini structured output forbids arbitrary `Dict`: chart data stays parallel lists.
- `chat["slides"]` is `None` until the first deck exists: use `bool(...)` or `or []`, never bare `len()`.
- The `opt_slides` slider has NO default value in `st.slider`; the default 8 is set once via `ss.setdefault("opt_slides", 8)` (avoids Streamlit's "default value + Session State API" warning). Session state can only be set before the widget is created in a run.
- Session state is per browser session; refresh clears chats and keys.
- Patching files: files use CRLF; preserve it.

## 5. Pending
- [ ] User to verify on Streamlit Cloud: make a deck -> close the tab -> reopen -> chat is back; delete chat; Keep my chats off.
- [ ] User to verify on Streamlit Cloud: set keys -> refresh -> keys and settings remembered; Forget this browser.
- [ ] Known trade-off: keys in localStorage are readable by scripts on the app's own origin (no third-party JS is loaded); users on shared computers should untick Remember or use Forget.
- [ ] User to verify: Undo after a revise; up / down / insert / delete with the slider following; Save project -> refresh -> Open project; previews follow theme switch and a custom template.
- [ ] Preview is an estimate: browser fonts differ from PowerPoint's; table row growth and wrapping are approximations.
- [ ] Known limit: on small per-minute plans (Groq free gpt-oss-120b = 8000 TPM) long documents / big decks cannot fit one request. Options not built yet: compress the source to a token budget before the outline call, or generate large decks in batches.
- [ ] User to verify new layouts: a roadmap/process source -> `timeline_flow`; a multi-option comparison -> `comparison_table` (native table editable in PowerPoint, theme colors, also with a custom template and *Modern Dark*). Visual check was wireframe-only (no LibreOffice here).
- [ ] User to verify report figures: upload a PDF with charts, Settings -> image mode 3 (or per slide) -> figures appear on image slides; picker thumbnails; logos not offered.
- [ ] Idea: let the LLM see figure captions/nearby text to match figures to slides (currently document order).
- [ ] User to verify targeted patch: "把第 3 页改成三栏卡片" / "shorten the data on slide 5" -> only that slide changes (message "Updated slide N"); "make all titles shorter" and "add one more slide" still regenerate.
- [ ] User to verify retrieval: with a long PDF (> 12k chars) ask about a number from the second half -> correct answer.
- [ ] Known edge: a cross-language question (Chinese question, English source) gets few BM25 hits and falls back to the head of the document (same as before).
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
| 2026-10-17 | Feature: targeted slide patching (`ACTION: PATCH`, `patch_slides`, `SlidePatch`, `_sanitize_slide` extracted, schema-param JSON helpers) + lightweight BM25 chunk retrieval (`retrieval.py`) for router and patch context; smoke test adds `check_retrieval`, `check_patch`, patch-protocol asserts |
| 2026-10-17 | Feature: `timeline_flow` + `comparison_table` layouts (schemas, prompts, sanitizer, renderers, build, editor, router digest) and image mode 3 `Report Images First` (filtered extraction, `assign_report_images`, editor picker); README layouts updated; smoke test: 11 layouts, native table check, `check_new_layouts`, `check_report_images` |
| 2026-10-17 | Fix: Groq HTTP 413 'Request too large' (TPM) -> token-limit handling in `llm_client` (cap shrink, short wait, `TokenBudgetError` message); JSON recovery paths now validate against the requested schema (SlidePatch); smoke test `check_token_limits` |
| 2026-10-17 | UI: `use_container_width` -> `width="stretch"` (streamlit>=1.57); radios / 2-5 option selects -> `st.segmented_control`; emoji -> Material icons; sentence-case labels (verified with `AppTest`: no exceptions or deprecation warnings) |
| 2026-10-17 | Feature: Undo/Redo outline history + .json project save/open (`deck_ops`, `project_io`); no-token slide move/insert/delete in the editor; live HTML/CSS wireframe preview with overflow warnings (`preview.py`); fix: sidebar delete-chat icon was clipped |
| 2026-10-18 | Brand: pixel-art logo (`assets/`) as sidebar logo, favicon, assistant avatar and hero image |
| 2026-10-18 | Feature: settings and (opt-out) API keys remembered in the user's browser via an inline CCv2 localStorage bridge (`prefs.py`, `ui/browser_prefs.py`), Remember / Forget controls in Settings; smoke test `check_browser_prefs` |
| 2026-10-19 | Feature: chats kept in the user's browser via an inline CCv2 IndexedDB bridge (`chat_store.py`, `ui/browser_chats.py`), restored on reopen; Keep my chats setting; Forget also erases chats; project_io helpers made public; smoke test `check_chat_store` |
| 2026-10-19 | Perf/fix: chat store no longer acks (one rerun per edit instead of two); bridge moved after the editor so text edits are saved (they were missed when it ran in the sidebar) |