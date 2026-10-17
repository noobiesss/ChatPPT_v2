# ChatPPT V2 — Project State (Handoff Document for the Next AI)

> **How to use:** The user pastes (or uploads) this file at the start of a new chat, together with the latest `app.py`.
> Read it fully before touching any code. **Update the "Changelog", "Github Commit Log" and "Pending Tasks" sections at the end of every work session** and give the user the updated file.

Last updated: 2026-10-17 (Bug 4a-3 FIXED: 2-option image mode; Fixed Groq `openai/gpt-oss-120b` empty JSON response by setting `max_completion_tokens=8192`, `reasoning_effort="low"`, and empty-response fallback/retry)
Github Commit Log: `fix(llm,images): fix empty JSON on Groq reasoning models (gpt-oss-120b/qwen3) via max_completion_tokens=8192 & low reasoning_effort; simplify image mode to 2 options`
---

## 0. User Preferences (MUST FOLLOW)

- **Reply to the user in Chinese (华文 / 简体中文).**
- **All code and all code comments MUST be strictly in English.**
- Deliver changes as a complete, ready-to-run `app.py` (the user copies it over their local file). Preserve the existing CRLF line endings when patching programmatically.
- Whenever the project state changes, also provide an updated `project_state.md`.

---

## 1. Project Goal & Tech Stack

Build **"ChatPPT V2"**, an AI-powered presentation generator web app that turns uploaded reports (`.pdf`, `.docx`, `.txt`, `.md`) or pasted text into consulting-grade, visually diverse PowerPoint (`.pptx`) decks.

**Tech stack:** Python 3, Streamlit (web UI), Google GenAI SDK (`google-genai`), Pydantic v2 (structured outputs), `python-pptx` (slides + native charts), `pypdf`, `python-docx`, `Pillow`, and `urllib` (REST calls to OpenAI-compatible providers, Unsplash, Pollinations).

**Files:** `app.py` (single-file architecture), `requirements.txt` (needs `streamlit>=1.46.0`), `README.md`, `.gitignore`, `.streamlit/config.toml` (optional warm light theme; the app CSS is theme-neutral and works with light or dark).

---

## 2. Architecture (section order inside `app.py`)

1. **Pydantic schemas:** `PresentationOutline` -> `SlideSpec`.
   - `LayoutType` enum: `cover`, `section_divider`, `bullet_cards`, `three_cards`, `two_column`, `kpi_metrics`, `chart_slide`, `image_split`.
   - Supporting models: `StructuredPoint` (`heading`, `detail`), `ColumnContent`, `KPIMetric`, `ChartSpec` (`ChartType`: `bar`, `line`, `pie`), `ImageSpec` (`VisualStyle`: `stock_photo`, `ai_illustration`; `ImagePlacement`: `left`, `right`).
   - Two-option image mode constants: `IMAGE_MODE_UNSPLASH_FIRST = "Unsplash Photo First (Fallback to AI)"`, `IMAGE_MODE_PLACEHOLDER = "Reserve Photo Placeholder"`, `IMAGE_MODE_OPTIONS = [IMAGE_MODE_UNSPLASH_FIRST, IMAGE_MODE_PLACEHOLDER]`.
2. **Themes:** `BUILTIN_THEMES` — 4 widescreen 16:9 palettes (Corporate Blue, Modern Dark, Minimalist Light, Emerald Forest).
3. **Document & image extraction:** `extract_text_from_uploaded_file()`, `extract_images_from_report()`.
4. **Multi-provider LLM layer:**
   - `LLM_PROVIDERS` dict (static presets): `Google Gemini` (native SDK, `base_url=None`), `Groq (Fast & Free Tier)` (updated to active models: `llama-3.3-70b-versatile`, `openai/gpt-oss-120b`, `meta-llama/llama-4-scout-17b-16e-instruct`, `qwen/qwen3-32b`, `llama-3.1-8b-instant`), `OpenAI (Paid)`, `DeepSeek (Paid)`, `OpenRouter (All-in-One)`, `Custom OpenAI-Compatible API`. Each has `base_url`, `models` (ends with `"Custom Model ID..."`), `key_label`, `key_help`.
   - **Live model discovery:** `fetch_live_models()` (`@st.cache_data(ttl=600)`), `resolve_model_options()`, `_is_chat_model()`, `_derive_models_url()`, `NON_CHAT_KEYWORDS`.
   - **JSON handling & reasoning-model protection:** `extract_clean_json_str()` (strips closed/unclosed `<think>` blocks and code fences, handles `None`), `_parses_as_outline()`, `repair_json_str()`, `call_openai_compatible_json()` (sets `max_completion_tokens=8192` / `max_tokens=8192`, sets `reasoning_effort="low"` for `gpt-oss` models and `"none"` for Groq `qwen3`, checks `message.reasoning` if `message.content` is empty, and retries without `response_format` or optional params if needed).
5. **AI router & sanitizer:** `generate_smart_outline()` (uses `SYSTEM_INSTRUCTION`: McKinsey-style Action Titles + layout variety; returns `(slides, warning)` and verifies the slide count) built on `_request_outline()` (with automatic retry on transient failure), `_build_count_correction()`, `_trim_slides_to_count()`, `mentions_slide_count()`; and `sanitize_outline_dict()` (ensures `categories`/`values` equal length, guarantees at least one `image_split` slide on decks of 4+ slides, and defaults `image_source_mode` to `IMAGE_MODE_UNSPLASH_FIRST`).
6. **Two-option image engine:** `resolve_slide_image_bytes()` implements only two modes:
   - Option 1 (`IMAGE_MODE_UNSPLASH_FIRST`, the default if the user makes no selection): always queries `fetch_unsplash_image_bytes()` first (strips accidental `Client-ID ` prefixes and retries with 2-word core keywords if a long query returns 0 results), and falls back to `generate_ai_image_bytes()` (Imagen 3 `imagen-3.0-generate-002` if enabled, else Pollinations AI) only when no Unsplash key is configured or no photo is found.
   - Option 2 (`IMAGE_MODE_PLACEHOLDER`): returns `None` so `render_image_split_slide()` renders a clean reserved photo placeholder box.
   - `add_picture_fitted()` does aspect-ratio-safe centering and RGBA/WEBP -> RGB conversion.
7. **PPTX rendering & template inheritance engine:**
   - **Template inheritance helpers:** `_scale_template_to_widescreen()`, `extract_template_style()` + `_parse_theme_xml_colors_and_fonts()` + `_sample_slide_colors_and_fonts()`, `inspect_template_summary()`, `_capture_template_prototypes()` + `_create_slide_with_template()` + `_select_layout_for_role()` + `_is_full_bleed_background_shape()` + `_has_visible_fill()` + `_strip_template_decorations()`, and `apply_slide_fonts()`.
   - **Core renderers:** `build_presentation_pptx()`, `clear_existing_slides()`, `set_solid_background()`, `draw_card_shape()`, `draw_header_and_takeaway()`, plus 8 layout renderers (`render_cover_slide`, `render_section_divider_slide`, `render_bullet_cards_slide`, `render_three_cards_slide`, `render_two_column_slide`, `render_kpi_metrics_slide`, `render_chart_slide`, `render_image_split_slide`).
8. **Editor helpers:** `format_points_for_editor`, `parse_points_from_editor`, `format_kpis_for_editor`, `parse_kpis_from_editor`.
9. **Chat-style UI (Claude-like), section 8 of `app.py`:**
   - **State:** `st.session_state.chats` (dict id -> chat), `chat_order` (newest first), `active_chat`, `settings` (`api_keys` per provider, `provider`, `models`, `custom_base_url`, `image_mode`, `unsplash_key`, `use_imagen`), plus persistent `custom_tpl_bytes` and `custom_tpl_name` for uploaded `.pptx` templates.
   - **Settings:** `open_settings()` -> `settings_dialog()` (`@st.dialog`), ONE page and ONE Save button: (1) Provider selectbox -> API key input -> Model selectbox, (2) Images: 2-option `Slide image mode` radio (`Unsplash Photo First (Fallback to AI)` [default] vs `Reserve Photo Placeholder`), Unsplash Access Key, and Imagen 3 fallback checkbox. Saving a new global `image_mode` also updates existing slides in open chats.
   - **Editor & Export:** `render_outline_editor()` (for `image_split` slides, shows the same 2-option `Image Mode` radio defaulting to Option 1) and `render_export_block()`.

---

## 3. Completed Features & Modifications

- 8 consulting-grade visual layouts, conclusion-first Action Titles, bottom "Key Takeaway" banner (no speaker notes).
- Native editable PowerPoint charts (`COLUMN_CLUSTERED`, `LINE_MARKERS`, `PIE`).
- Two-option image mode (`Unsplash Photo First (Fallback to AI)` as the default, or `Reserve Photo Placeholder`).
- Multi-user safe, in-memory **BYOK** (Bring Your Own Key) architecture, ready for 24/7 cloud hosting.
- Multi-provider support (Gemini, Groq, OpenAI, DeepSeek, OpenRouter, custom OpenAI-compatible endpoint).
- **[2026-10-16] Live model list fix (verified by user).**
- **[2026-10-16] Malformed-JSON fix (verified by user).**
- **[2026-10-16] UI redesign & theme-neutral CSS fix (verified by user in light & dark mode).**
- **[2026-10-16] Bug 4a-1 fixed & verified by user (slide count respected).**
- **[2026-10-17] Bug 4a-2 fixed & verified by user (custom `.pptx` template inheritance: backgrounds + fonts/colors only, master/layout decorations stripped, white-text contrast guard on Key Takeaway banner).**
- **[2026-10-17] UI rework:** one-page Settings (provider -> API key -> model, plus Images, single Save); model button in composer opens Settings; tone moved to `⚙` options popover; composer buttons no longer truncated.
- **[2026-10-17] Bug 4a-3 fixed (Unsplash photos & 2-option image mode):**
  - Replaced the old 5-option image source + preferred style selectors with **only two options** (`IMAGE_MODE_UNSPLASH_FIRST = "Unsplash Photo First (Fallback to AI)"` and `IMAGE_MODE_PLACEHOLDER = "Reserve Photo Placeholder"`), defaulting everywhere to Option 1 if the user makes no selection; `resolve_slide_image_bytes()` always queries Unsplash first under Option 1 and falls back to AI generation only when Unsplash has no key or no match; `fetch_unsplash_image_bytes()` strips accidental `Client-ID ` prefixes and retries with the first 2 keywords when a longer query returns 0 results; `sanitize_outline_dict()` guarantees at least one `image_split` slide in decks of 4+ slides.
- **[2026-10-17] Fixed Groq `openai/gpt-oss-120b` empty JSON validation error (`input_value=''`):**
  - **Root cause:** Groq defaults `max_completion_tokens` to 1024 when omitted, and `openai/gpt-oss-120b` defaults `reasoning_effort` to `"medium"`. The model spent all 1024 tokens in `message.reasoning` (`finish_reason="length"`), returning `message.content = ""`.
  - **Fix:** `call_openai_compatible_json()` now sets `max_completion_tokens=8192` (for Groq/OpenAI) / `max_tokens=8192` (others), sets `reasoning_effort="low"` for `gpt-oss` models and `"none"` for Groq `qwen3`, falls back to checking `message.reasoning` or retrying without `response_format` if `content` is empty, and `_request_outline()` retries once on transient failures.

---

## 4. Known Issues / Constraints

- Gemini structured outputs do not support arbitrary `Dict` types; chart data must stay as parallel lists (`categories: List[str]`, `values: List[float]`).
- Images must always be inserted via `add_picture_fitted()` to avoid aspect-ratio stretching and WEBP/RGBA crashes in `python-pptx`.
- The repair logic only fixes stray closers and truncation; other malformations (missing commas, unquoted keys) are not handled.
- Chats are session-only (`st.session_state`): refreshing the page or restarting the server loses chats and API keys.
- Only one deck per chat; a new deck from a new report should be started with New chat.

---

## 4a. Open Bugs Reported by the User (2026-10-16)

1. **[CLOSED & VERIFIED 2026-10-16] Slide count not respected.**
2. **[CLOSED & VERIFIED 2026-10-17] Custom template not reflected in the output.**
3. **[FIXED 2026-10-17 — Awaiting User Verification] Unsplash photos missing & 2-option image mode.** Simplified to two options (`1. Unsplash Photo First (Fallback to AI)` [default] and `2. Reserve Photo Placeholder`), ensured Unsplash is always tried first under Option 1 with keyword fallback, and guaranteed at least one `image_split` slide in 4+ slide decks.

---

## 5. Pending Tasks

- [x] **Bug 4a-1 verified by user.**
- [x] **Bug 4a-2 verified by user (round 3).**
- [ ] **Verify bug 4a-3 fix & Groq `openai/gpt-oss-120b` fix:** test generating a deck with `openai/gpt-oss-120b` and rendering with an Unsplash Access Key (Option 1 default) and `Reserve Photo Placeholder` (Option 2).
- [ ] Optional: deploy to GitHub + Streamlit Community Cloud / Railway for 24/7 public access.
- [ ] Optional: new UI/rendering features.

---

## 6. Changelog

| Date | Change |
| :--- | :--- |
| (earlier) | V2 rewrite: 8 layouts, native charts, hybrid image engine, multi-provider BYOK, Gemini model selector |
| 2026-10-16 | Live model discovery per provider with refresh button and static fallback |
| 2026-10-16 | Groq `json_validate_failed` handling: `failed_generation` repair, JSON-mode retry, `repair_json_str()` |
| 2026-10-16 | Full UI redesign to a Claude-style chat interface with Settings dialog (all API keys), multi-chat sidebar, bottom composer |
| 2026-10-16 | Fixed invisible text: theme-neutral CSS, composer fade color from active theme; `streamlit>=1.46.0`; logged 3 open bugs (slide count, template fidelity, Unsplash photos) |
| 2026-10-16 | Bug 4a-1: slide count now verified (stronger prompt, corrective retries, deterministic trim, warning in chat); `generate_smart_outline()` returns `(slides, warning)`; revisions mentioning a slide count skip enforcement |
| 2026-10-16 | Bug 4a-2: implemented full `.pptx` custom template inheritance (widescreen scaling, prototype slide background/decor/logo replay, `theme1.xml` + slide color/font extraction, persistent session state & UI preview badge); updated Groq static presets & `README.md`; recorded user verification of 4a-1, UI, themes, live models, and 8 layouts |
| 2026-10-16 | Template inheritance reworked: only backgrounds (`<p:bg>` + full-bleed bg shapes), fonts and theme colors are inherited; frames, illustrations, header/footer widgets, logos and sample text are no longer replayed; no new UI options |
| 2026-10-17 | 4a-2 round 3: `_strip_template_decorations()` removes master/layout frames and ornaments; `_has_visible_fill()` stops outline frames counting as backgrounds; fonts now taken from sample-slide runs before `theme1.xml` |
| 2026-10-16 | Key Takeaway banner color now derived from the template accent with a white-text contrast guard (`_ensure_white_text_contrast`); same guard on bullet-card number badges and two-column right header |
| 2026-10-17 | UI rework: Settings is one page (provider -> API key -> model, plus Unsplash/Imagen) with a single Save; `render_model_picker()` replaced by `get_active_model_cfg()` + model button that opens Settings; tone moved into the ⚙ options popover; composer button text no longer truncated |
| 2026-10-17 | Bug 4a-3 fix: simplified slide image handling to 2 options (`Unsplash Photo First (Fallback to AI)` [default] and `Reserve Photo Placeholder`); Unsplash is always queried first under Option 1 with 2-word keyword retry before falling back to AI; guaranteed at least one `image_split` slide on 4+ slide decks |
| 2026-10-17 | Fixed Groq `openai/gpt-oss-120b` empty JSON response (`input_value=''`): set `max_completion_tokens=8192`, `reasoning_effort="low"` for `gpt-oss` models (`"none"` for Groq `qwen3`), fallback to `message.reasoning` / non-JSON-mode retry when `content` is empty, and automatic retry in `_request_outline()` |