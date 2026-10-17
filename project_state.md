# ChatPPT V2 — Project State (Handoff Document for the Next AI)

> **How to use:** Paste or upload this file at the start of a new chat, together with the latest `app.py`.
> Read it fully before touching any code. **Update "Changelog", "Github Commit Log", and "Pending Tasks" at the end of every work session** and return the updated file.

Last updated: 2026-10-17 (4a-1, 4a-2, 4a-3 ALL VERIFIED; 2-option image mode + reasoning model token/JSON fix)
Github Commit Log: `fix(llm,images): verify 2-option image mode (Unsplash first -> AI fallback, or placeholder) and fix empty JSON on Groq reasoning models via max_completion_tokens=8192 & low reasoning_effort`
---

## 0. User Preferences (MUST FOLLOW)

- **Reply to the user in Chinese (华文 / 简体中文).**
- **All code and all code comments MUST be strictly in English.**
- Deliver changes as a complete, ready-to-run `app.py` (preserve existing CRLF line endings when patching programmatically).
- Whenever the project state changes, also provide an updated `project_state.md`.

---

## 1. Project Goal & Tech Stack

Build **"ChatPPT V2"**, an AI-powered presentation generator web app that turns uploaded reports (`.pdf`, `.docx`, `.txt`, `.md`) or pasted text into consulting-grade, visually diverse PowerPoint (`.pptx`) decks.

- **Tech stack:** Python 3, Streamlit (`>=1.46.0`), Google GenAI SDK (`google-genai`), Pydantic v2, `python-pptx`, `pypdf`, `python-docx`, `Pillow`, `urllib`.
- **Files:** `app.py` (single-file architecture), `requirements.txt`, `README.md`, `.gitignore`, `.streamlit/config.toml`.

---

## 2. Architecture (section order inside `app.py`)

1. **Schemas & Constants:**
   - `PresentationOutline` -> `SlideSpec` (`LayoutType`: `cover`, `section_divider`, `bullet_cards`, `three_cards`, `two_column`, `kpi_metrics`, `chart_slide`, `image_split`).
   - Sub-models: `StructuredPoint`, `ColumnContent`, `KPIMetric`, `ChartSpec` (`bar`, `line`, `pie`), `ImageSpec` (`preferred_style` defaults to `stock_photo`; `placement`: `left`, `right`).
   - Two-option image mode constants: `IMAGE_MODE_UNSPLASH_FIRST = "Unsplash Photo First (Fallback to AI)"` (default), `IMAGE_MODE_PLACEHOLDER = "Reserve Photo Placeholder"`, `IMAGE_MODE_OPTIONS`.
2. **Presets & Themes:** `LLM_PROVIDERS` (Gemini, Groq, OpenAI, DeepSeek, OpenRouter, Custom OpenAI-Compatible) and `BUILTIN_THEMES` (Corporate Blue, Modern Dark, Minimalist Light, Emerald Forest).
3. **Extraction:** `extract_text_from_uploaded_file()`, `extract_images_from_report()`.
4. **LLM & JSON Layer:**
   - Live model discovery: `fetch_live_models()` (`@st.cache_data(ttl=600)`), `resolve_model_options()`, `_is_chat_model()`, `_derive_models_url()`.
   - JSON & reasoning protection: `extract_clean_json_str()` (strips closed/unclosed `<think>` blocks and code fences), `repair_json_str()`, `call_openai_compatible_json()` (sets `max_completion_tokens=8192` / `max_tokens=8192`, sets `reasoning_effort="low"` for `gpt-oss` and `"none"` for Groq `qwen3`, checks `message.reasoning` if `content` is empty, repairs `failed_generation`, and retries without `response_format` or optional params).
5. **Outline Router & Sanitizer:**
   - `generate_smart_outline()` -> `_request_outline()` (with retry), `_build_count_correction()` (up to 2 corrective retries), `_trim_slides_to_count()`, `mentions_slide_count()`.
   - `sanitize_outline_dict()` (equalizes chart `categories`/`values`, guarantees >=1 `image_split` slide on 4+ slide decks, defaults `image_source_mode` to `IMAGE_MODE_UNSPLASH_FIRST`).
6. **Two-Option Image Engine:**
   - `resolve_slide_image_bytes()`: Option 1 queries `fetch_unsplash_image_bytes()` first (strips `Client-ID ` prefix, retries with 2-word core query if needed) and falls back to `generate_ai_image_bytes()` (Imagen 3 `imagen-3.0-generate-002` -> Pollinations AI); Option 2 returns `None` to render a reserved placeholder box.
   - `add_picture_fitted()` centers images without stretching and converts RGBA/WEBP to RGB.
7. **PPTX Rendering & Custom Template Inheritance:**
   - Template helpers: `_scale_template_to_widescreen()`, `extract_template_style()` + `_parse_theme_xml_colors_and_fonts()` + `_sample_slide_colors_and_fonts()` (prefers sample-slide run fonts over `theme1.xml`), `_capture_template_prototypes()` + `_is_full_bleed_background_shape()` + `_has_visible_fill()`, `_strip_template_decorations()` (removes non-placeholder frames/ornaments from layouts/masters), `_create_slide_with_template()`, `apply_slide_fonts()`, `_ensure_white_text_contrast()`.
   - Renderers: `build_presentation_pptx()` + 8 layout renderers.
8. **Chat-Style UI:**
   - Session state: `chats`, `chat_order`, `active_chat`, `settings` (`api_keys`, `provider`, `models`, `custom_base_url`, `image_mode`, `unsplash_key`, `use_imagen`), `custom_tpl_bytes`, `custom_tpl_name`.
   - One-page `settings_dialog()` (Provider -> API key -> Model, plus 2-option `Slide image mode`, Unsplash key, Imagen toggle, single Save).
   - Sticky bottom composer (`📎 Attach`, `🎨 Template`, `🤖 Model`, `⚙ Options`, `Generate ➤`), per-slide `render_outline_editor()`, and `render_export_block()`.

---

## 3. Completed & Verified Features

- **[VERIFIED]** 8 consulting layouts, Action Titles, Key Takeaway banners, and native editable PowerPoint charts.
- **[VERIFIED]** Multi-provider BYOK with live `/models` discovery, caching, and manual refresh.
- **[VERIFIED]** Claude-style chat UI with theme-neutral CSS (works in both light and dark themes) and one-page Settings dialog.
- **[VERIFIED 4a-1]** Exact slide count enforcement (prompt constraint + corrective retries + deterministic trim + revision count detection).
- **[VERIFIED 4a-2]** Custom `.pptx` template inheritance: inherits widescreen-scaled backgrounds (`<p:bg>` + full-bleed filled/picture shapes), sample-slide fonts, and theme/sampled colors; strips master/layout/slide decorative frames and widgets; enforces WCAG white-text contrast on banners/badges.
- **[VERIFIED 4a-3]** Two-option slide image mode: `1. Unsplash Photo First (Fallback to AI)` (default everywhere) and `2. Reserve Photo Placeholder`, with 2-word Unsplash query retry and guaranteed `image_split` slide on 4+ slide decks.
- **[VERIFIED]** Reasoning model compatibility (`openai/gpt-oss-120b`, `qwen/qwen3-32b`, `deepseek-reasoner`): 8192 token budget, low/none reasoning effort, `<think>` and `message.reasoning` fallback.

---

## 4. Known Constraints

- Gemini structured outputs forbid arbitrary `Dict`; keep chart data as parallel lists (`categories: List[str]`, `values: List[float]`).
- Always insert images via `add_picture_fitted()` to avoid stretching and WEBP/RGBA crashes.
- Session state is in-memory per browser session; refreshing clears chats and keys.

---

## 5. Pending Tasks

- [x] Bug 4a-1 (slide count) — CLOSED & VERIFIED
- [x] Bug 4a-2 (custom `.pptx` template inheritance) — CLOSED & VERIFIED
- [x] Bug 4a-3 (2-option Unsplash/placeholder image mode + Groq `gpt-oss-120b` fix) — CLOSED & VERIFIED
- [ ] Optional: deploy to GitHub + Streamlit Community Cloud / Railway for 24/7 hosting.
- [ ] Optional: next UI or rendering enhancements requested by the user.

---

## 6. Changelog (Condensed)

| Date | Change |
| :--- | :--- |
| 2026-10-16 | V2 core: 8 layouts, native charts, multi-provider BYOK, live model discovery, JSON repair, Claude-style chat UI, theme-neutral CSS, Bug 4a-1 slide count enforcement |
| 2026-10-17 | Bug 4a-2 verified: custom `.pptx` template inheritance (backgrounds + sample-slide fonts/colors only, master/layout decorations stripped, WCAG contrast guard); one-page Settings rework |
| 2026-10-17 | Bug 4a-3 & reasoning model fix verified: 2-option image mode (`Unsplash Photo First (Fallback to AI)` [default] vs `Reserve Photo Placeholder`), 2-word Unsplash retry, guaranteed `image_split` slide, and `max_completion_tokens=8192` + `reasoning_effort` tuning for Groq `gpt-oss-120b` / `qwen3` |