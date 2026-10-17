# ChatPPT V2 — Project State (Handoff Document for the Next AI)

> **How to use:** Paste or upload this file at the start of a new chat, together with ONLY the module file(s) the task needs (see section 2a — you no longer need to upload or return the whole code base).
> Read it fully before touching any code. **Update "Changelog", "Github Commit Log", and "Pending Tasks" at the end of every work session** and return the updated file.

Last updated: 2026-10-17 (revision fix: relative slide-count edits like "add one more slide" now raise the target to current deck size + delta and sync the Number-of-slides slider)
Github Commit Log: `fix(outline): treat "add/remove N slide(s)" revisions as relative to the current deck, keep existing slides, sync the slide slider`
---

## 0. User Preferences (MUST FOLLOW)

- **Reply to the user in Chinese (华文 / 简体中文).**
- **All code and all code comments MUST be strictly in English.**
- Deliver changes as complete, ready-to-run **module files** (only the files that changed; keep their CRLF line endings when patching programmatically). Do not re-output untouched modules.
- Never rename or move a function/constant unless the task requires it; if you do, update every importer and section 2a.
- After every change run `python smoke_test.py` (offline, no API keys) and report the result.
- Whenever the project state changes, also provide an updated `project_state.md`.

---

## 1. Project Goal & Tech Stack

Build **"ChatPPT V2"**, an AI-powered presentation generator web app that turns uploaded reports (`.pdf`, `.docx`, `.txt`, `.md`) or pasted text into consulting-grade, visually diverse PowerPoint (`.pptx`) decks.

- **Tech stack:** Python 3, Streamlit (`>=1.46.0`), Google GenAI SDK (`google-genai`), Pydantic v2, `python-pptx`, `pypdf`, `python-docx`, `Pillow`, `urllib`.
- **Files:** `app.py` (thin entry point), `chatppt/` package (see 2a), `smoke_test.py`, `requirements.txt`, `README.md`, `.gitignore`, `.streamlit/config.toml`. Run with `streamlit run app.py`.

---

## 2a. File Map (what to give the AI for each task)

Dependencies only point DOWN this list (no cycles; core modules never import `ui/*`). Untouched modules do not need to be uploaded.

| File | Responsibility | Edit when... |
| :--- | :--- | :--- |
| `app.py` | Entry point: page config + page flow (`main`) | changing overall page layout |
| `chatppt/schemas.py` | Enums, Pydantic models (`SlideSpec`, `PresentationOutline`...), image-mode constants | adding a layout type or slide field |
| `chatppt/presets.py` | `LLM_PROVIDERS`, `BUILTIN_THEMES`, `TONE_OPTIONS`, `LANGUAGE_OPTIONS`, provider name constants | adding a provider/preset model/theme/tone/language |
| `chatppt/extraction.py` | PDF/DOCX/TXT/MD text + embedded image extraction | changing how reports are read |
| `chatppt/prompts.py` | `SYSTEM_INSTRUCTION` (LLM design prompt) | tuning outline quality/routing rules |
| `chatppt/llm_client.py` | Live model discovery, JSON cleaning/repair, `call_openai_compatible_json` | provider/model quirks, token limits, JSON failures |
| `chatppt/images.py` | Unsplash fetch, AI image fallback, `resolve_slide_image_bytes`, `add_picture_fitted` | image sources/queries/placement |
| `chatppt/outline.py` | `generate_smart_outline`, slide-count retries/trim, `sanitize_outline_dict` | outline requesting/validation/normalizing |
| `chatppt/pptx_template.py` | Custom `.pptx` inheritance (colors, fonts, backgrounds, decoration stripping, prototypes) | uploaded template not reproduced correctly |
| `chatppt/pptx_render.py` | Drawing helpers + the 8 layout renderers | changing how one layout looks |
| `chatppt/pptx_build.py` | `build_presentation_pptx` (deck assembly, layout dispatch) | deck-level flow, adding a layout dispatch entry |
| `chatppt/ui/state.py` | Session-state defaults, chat create/delete | new session key or default setting |
| `chatppt/ui/css.py` | `APP_CSS`, CSS injection (theme-independent: no hard-coded page colors) | look & feel |
| `chatppt/ui/settings.py` | Settings dialog (provider -> key -> model, image mode/keys), `get_active_model_cfg`, `get_model_options` | settings window, provider/model selection |
| `chatppt/ui/sidebar.py` | Sidebar (New chat, Settings, chat list) | sidebar |
| `chatppt/ui/messages.py` | Hero screen + chat messages | welcome screen/message look |
| `chatppt/ui/outline_editor.py` + `ui/editor_text.py` | Per-slide outline editor + text<->data helpers | what is editable per slide |
| `chatppt/ui/export.py` | Deck block: editor + render/download | export/download flow |
| `chatppt/ui/generation.py` | `handle_send`, `run_pending` (what happens on Generate) | send/validation/generation job |
| `chatppt/ui/composer.py` + `ui/common.py` | Bottom input bar and its popovers | input bar |

**Task -> files to upload (plus this file):**
- Outline quality / prompt wording -> `prompts.py` (+ `outline.py`)
- Slide count / outline validation -> `outline.py` (+ `schemas.py`)
- LLM provider, model list, empty/invalid JSON -> `llm_client.py` (+ `presets.py`)
- Uploaded template looks wrong (frames, fonts, colors) -> `pptx_template.py`
- A layout's visuals -> `pptx_render.py` (+ `schemas.py` if a new field)
- NEW layout type -> `schemas.py`, `outline.py`, `pptx_render.py`, `pptx_build.py`, `ui/outline_editor.py`
- Images (Unsplash/AI/placeholder) -> `images.py` (+ `ui/settings.py` for settings)
- Settings window -> `ui/settings.py` (+ `ui/state.py` for new defaults, `presets.py` for providers)
- Input bar / Generate behavior -> `ui/composer.py`, `ui/generation.py`
- Styling -> `ui/css.py`

**Rules for any AI editing the code:**
1. Keep layering: `schemas/presets` -> `extraction/prompts/llm_client` -> `outline/images` -> `pptx_template` -> `pptx_render` -> `pptx_build` -> `ui/*` -> `app.py`. A module may only import from modules above it. Core modules must not import `streamlit` UI code (they may use `st.cache_data`).
2. Streamlit session keys are created in `ui/state.py::init_state`; read/write them elsewhere via `st.session_state.settings[...]`.
3. If you add a new top-level function used by other modules, list it in the file map above.
4. Finish with `python smoke_test.py`; add a check there when you fix a bug that could regress (e.g. template frames/fonts).

---

## 2. Architecture (function-level overview; the file that holds each part is in 2a)

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
   - `generate_smart_outline()` -> `_request_outline()` (with retry), `_build_count_correction()` (up to 2 corrective retries), `_trim_slides_to_count()`, `mentions_slide_count()` (absolute counts, e.g. "6 slides" -> count not enforced), `parse_slide_count_change()` (relative edits, e.g. "add one more slide" / "remove the last slide" / "加一页" -> new target = current deck size + delta, clamped 4-20; takes precedence over `mentions_slide_count`).
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
- [x] Split code base verified by user (`smoke_test.py` passed, app runs).
- [ ] User to check the light/dark composer look after the CSS change (`ui/css.py`, `ui/composer.py`).
- [x] `config.toml` lives in `.streamlit/` (50 MB upload limit + themes apply).
- [ ] User to check Light / Dark / System after the theme change; tune colors only in `.streamlit/config.toml` (never put colors in the bare `[theme]` table: it locks the app to one mode).
- [ ] User to verify: generate 8 slides, then revise with "add one more slide conclusion" -> expect 9 slides (old 8 kept, Conclusion last) and the slider showing 9.
- [ ] Optional: deploy to GitHub + Streamlit Community Cloud / Railway for 24/7 hosting.
- [ ] Optional: next UI or rendering enhancements requested by the user.

---

## 6. Changelog (Condensed)

| Date | Change |
| :--- | :--- |
| 2026-10-16 | V2 core: 8 layouts, native charts, multi-provider BYOK, live model discovery, JSON repair, Claude-style chat UI, theme-neutral CSS, Bug 4a-1 slide count enforcement |
| 2026-10-17 | Bug 4a-2 verified: custom `.pptx` template inheritance (backgrounds + sample-slide fonts/colors only, master/layout decorations stripped, WCAG contrast guard); one-page Settings rework |
| 2026-10-17 | Bug 4a-3 & reasoning model fix verified: 2-option image mode (`Unsplash Photo First (Fallback to AI)` [default] vs `Reserve Photo Placeholder`), 2-word Unsplash retry, guaranteed `image_split` slide, and `max_completion_tokens=8192` + `reasoning_effort` tuning for Groq `gpt-oss-120b` / `qwen3` |
| 2026-10-17 | Refactor: monolithic `app.py` (112 top-level definitions, 3.4k lines) split into `chatppt/` package (schemas, presets, extraction, prompts, llm_client, outline, images, pptx_template, pptx_render, pptx_build, ui/*) + thin `app.py`; every definition verified AST-identical to the original; offline `smoke_test.py` added; `_derive_fallback_keywords` now lives in `images.py` to keep layering |
| 2026-10-17 | UI: composer dock no longer uses a Python-detected background color (it stayed dark after switching to Light); now a frosted blur strip + translucent composer works in both themes; textarea fully transparent (no dark outline); primary buttons use accent #C96442; footer note `ChatPPT can make mistakes. Please check responses.` under the composer; removed `_dock_background()` |
| 2026-10-17 | Theme: `.streamlit/config.toml` now defines `[theme.light]` and `[theme.dark]` (+ sidebar tables) so the menu offers System / Light / Dark; shared accent #C96442, warm cream light palette (#FAF9F5 / #F0EEE6 / text #1F1E1D), warm charcoal dark palette (#262624 / #30302E / text #ECEBE4) |
| 2026-10-17 | Fix: revision "add one more slide conclusion" replaced slide 8 because the slider (8) was still enforced and trimmed the 9th slide. Now `parse_slide_count_change()` (outline.py) computes target = current slides + delta, `handle_send` (generation.py) passes it as `num_slides`, appends the current slide titles + "keep all existing slides" to the LLM source, and queues `opt_slides_sync` which `render_composer` (composer.py) applies to the slider before it is created. Smoke test extended with 8 parser assertions |