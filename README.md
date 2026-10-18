<p align="center">
  <img src="assets/logo.png" alt="ChatPPT logo" width="104">
</p>

# ChatPPT V2

**Turn a report or your notes into a consulting-style PowerPoint deck by chatting.**

ChatPPT V2 is a Streamlit web app. You give it a PDF, Word file or pasted text; it designs a slide outline with conclusion-first titles, picks a fitting layout for every slide (cards, KPIs, charts, timelines, comparison tables, images…), lets you refine it in plain language or edit it by hand, and renders a fully editable `.pptx` with native PowerPoint charts and tables.

It works with your own API key for Google Gemini, Groq, OpenAI, DeepSeek, OpenRouter or any OpenAI-compatible endpoint.

---

## Contents

- [Features](#features)
- [Privacy and your API keys](#privacy-and-your-api-keys)
- [Quick start](#quick-start)
- [Deploy on Streamlit Community Cloud](#deploy-on-streamlit-community-cloud)
- [How to use](#how-to-use)
- [Slide layouts](#slide-layouts)
- [Project structure](#project-structure)
- [Development and tests](#development-and-tests)

---

## Features

### Chat that builds and edits the deck
- **One Send button.** Each message is routed automatically:
  - a **question** ("why three cards on slide 4?") gets an answer, and the deck stays untouched;
  - an **edit to specific slides** ("turn slide 3 into a comparison table") rewrites only those slides in seconds; every other slide stays exactly as it was;
  - a **deck-wide change** ("make it 6 slides, more charts", "translate to Chinese") regenerates the outline.
- **Relative slide-count edits** such as "add one more slide" or "remove the last slide" keep all existing slides and update the slide-count control.
- **Long documents:** questions about a long PDF or DOCX use the passages most relevant to the question (lightweight BM25 retrieval, no vector database), not just the first pages.

### Outline editor (no tokens needed)
- Edit every slide's layout, title, points, KPIs, chart data, timeline steps, table cells, image options and takeaway banner.
- **Move up / move down / insert / delete** slides with one click. The cover always stays first and a closing slide always stays last.
- **Live wireframe preview** of each open slide in the current theme, with a warning when text is likely to overflow its card or box.
- **Undo / Redo** across AI changes and structure edits (up to 20 versions).

### Visual quality
- **11 layouts**, including native editable charts (bar, line, pie) and native editable tables. See [Slide layouts](#slide-layouts).
- Conclusion-first **action titles**, short bold headings with concise details, and a "so what?" **takeaway banner**.
- **No invented numbers:** charts, KPIs and tables only use figures found in your source.

### Images
Three modes, chosen globally in Settings or per slide:
- **Unsplash / AI**: an Unsplash stock photo, falling back to an AI image (Google Imagen 3 if enabled, otherwise Pollinations).
- **Placeholder**: an empty photo box to fill in PowerPoint.
- **Report figures**: diagrams and photos extracted from your uploaded PDF/DOCX (logos, icons and duplicates are skipped), assigned in order or picked per slide with a thumbnail.

### Themes and templates
- 4 built-in 16:9 palettes: *Corporate Blue*, *Modern Dark*, *Minimalist Light*, *Emerald Forest*.
- **Your own `.pptx` template:** ChatPPT inherits its backgrounds, theme colors and fonts and scales non-16:9 templates to widescreen. Decorations on the template's sample slides and layouts are removed so they never collide with generated content.

### Save and resume
- **Your chats come back when you reopen the app.** Conversations, outlines, source text and report figures are kept in your own browser (IndexedDB) and the most recent chat opens automatically. The 30 most recent chats are kept; you can turn this off in Settings.
- **Settings are remembered in your browser** (provider, models, image mode, theme, slide count, language, tone), and so are your API keys unless you turn that off. See [Privacy and your API keys](#privacy-and-your-api-keys).
- **Save project** downloads a `.chatppt.json` file with the outline, conversation, source text and report figures. **Open project** in the sidebar restores it after a page refresh, on another computer, or for a colleague. Project files never contain API keys.

### Providers
- Google Gemini (native SDK), Groq, OpenAI, DeepSeek, OpenRouter, or any **custom OpenAI-compatible endpoint** (SiliconFlow, Together AI, Ollama, vLLM…).
- Live model lists from each provider, refreshed every 10 minutes or on demand.
- Robust output handling: JSON repair, reasoning-model support (`gpt-oss`, `qwen3`, `deepseek-reasoner`), automatic slide-count correction, and a readable explanation when a provider plan's tokens-per-minute limit is too small for the request.

---

## Privacy and your API keys

ChatPPT is a "bring your own key" app. Here is exactly what happens to your key:

| Question | Answer |
| :--- | :--- |
| Where is my key stored? | **Only in your own browser** (its local storage), and only if "Remember my API keys on this browser" is on. Nothing is stored on the server: no database, no files. |
| Where is my key sent? | To the app's server while you use it, because the server makes the call to the AI provider. It is held in memory for your session only and sent **only to the provider you selected** (and to Unsplash / Google Imagen for images, if you use them). |
| Is it logged? | No. The code never prints, logs or displays keys. |
| Can other users see my key? | No. Every browser session is separate. |
| Where are my chats and uploaded reports kept? | In your own browser (IndexedDB), only if "Keep my chats on this browser" is on. They are not stored on the server and not shared with anyone. Anyone using the same browser profile on that computer can open them. |
| Is it in saved projects? | No. `.chatppt.json` project files never contain API keys. |
| How do I remove it? | Settings → section 3 → **Forget this browser** erases everything saved in that browser (settings, keys and chats). Unticking "Remember my API keys" or "Keep my chats" and pressing **Save** removes only the keys or only the chats. |

**Please keep in mind:**

- As with every bring-your-own-key web app, you are trusting whoever runs the server, because the key passes through it. This repository is public, so you can check the code yourself; if you prefer, run ChatPPT locally (see [Quick start](#quick-start)).
- Protect yourself on the provider's side: create a **separate key** just for ChatPPT and set a **spending limit**, so you can revoke it at any time.
- On a **shared or public computer**, untick "Remember my API keys" and "Keep my chats", or click **Forget this browser** when you are done.

---

## Quick start

**Requirements:** Python **3.11 or newer**, and an API key for at least one provider (Google Gemini, Groq, OpenAI, DeepSeek, OpenRouter, or a custom OpenAI-compatible endpoint).

```bash
git clone https://github.com/noobiesss/ChatPPT_v2.git
cd ChatPPT_v2

python -m venv .venv
# Windows:      .venv\Scripts\activate
# macOS/Linux:  source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```

The app opens at http://localhost:8501. Click **Settings** in the sidebar, choose a provider, paste your API key and pick a model.

Optional:
- **Unsplash access key** (Settings → Images) for stock photos. Without it, image slides use AI images or placeholders.
- **Google Imagen 3** for AI images needs a Google Gemini key with Imagen access.

---

## Deploy on Streamlit Community Cloud

1. Fork or push this repository to your GitHub account.
2. On [share.streamlit.io](https://share.streamlit.io), click **Create app** and choose:
   - **Repository:** your copy of `ChatPPT_v2`
   - **Branch:** `ChatPPT_main`
   - **Main file path:** `app.py`
3. Under **Advanced settings**, pick **Python 3.12** (3.11 or newer is required).
4. Click **Deploy**. Users enter their own API keys in the app, so no secrets are needed.

Notes:
- Every push to the branch updates the app automatically. If a push adds new imports and the app shows an `ImportError`, open **Manage app → ⋮ → Reboot app**.
- The repository, branch and main file of a deployed app are fixed. If you rename the branch, delete the app and deploy it again (you can reuse the same app URL).
- The upload limit is 50 MB (`.streamlit/config.toml`).

---

## How to use

1. **Add your source:** attach a PDF, DOCX, TXT or MD file (**Attach** button) or paste text into the message box. Set the slide count, language and tone with the **"8 slides"** button, and the color palette or your own `.pptx` template with the **palette** button next to it.
2. **Send.** ChatPPT designs the outline, which appears below the chat.
3. **Refine:**
   - ask questions or request changes in the chat ("make slide 5 shorter", "add a thank-you slide");
   - or open a slide in the outline to edit it directly; the preview shows the result live.
   - Use **Undo** if you don't like a change.
4. **Render presentation (.pptx)** and download the file. Everything in it is editable in PowerPoint.
5. Optional: **Save project** to continue later with **Open project**.

---

## Slide layouts

| Layout | What it looks like | Best for |
| :--- | :--- | :--- |
| `cover` | Large title, accent bar, subtitle | Opening slide (always slide 1) |
| `section_divider` | Contrasting background with section title and overview | Transitions in decks of 8+ slides |
| `bullet_cards` | 3–4 stacked cards with numbered badges | Sequential explanations, general takeaways |
| `three_cards` | 3 parallel pillar cards | Three pillars, modules or recommendations |
| `two_column` | Two side-by-side columns with colored headers | Pros vs. cons, before vs. after, problem vs. solution |
| `kpi_metrics` | 2–4 big-number cards with label and context | Headline figures, growth, ROI, benchmarks |
| `chart_slide` | Native editable chart + 2–3 insight cards | Trends, comparisons, shares |
| `image_split` | Image with caption + 2–3 key points | Products, operations, workflows, scenarios |
| `timeline_flow` | 2–5 numbered steps on a line with stage labels and step cards | Roadmaps, phased plans, processes |
| `comparison_table` | Native editable table (up to 5 columns × 6 rows) | Options, vendors or scenarios across several criteria |
| `closing` | Large centered phrase + optional presenter line | Thank-you / Q&A slide, only when you ask for it |

---

## Project structure

```text
app.py                  Streamlit entry point
assets/                 Logo files
chatppt/                Core package (no UI code)
  schemas.py            Slide outline data model (Pydantic)
  prompts.py            Instructions for the AI (layout rules, Q&A routing)
  outline.py            Outline generation, targeted slide patches, validation
  chat_assistant.py     Question / patch / revise routing
  retrieval.py          BM25 retrieval over long documents
  llm_client.py         Provider calls, JSON repair, token-limit handling
  pptx_*.py             Rendering: layouts, custom templates, deck assembly
  images.py             Unsplash, AI images, report figures
  deck_ops.py           Undo / redo and slide move / insert / delete
  project_io.py         Save / open project files
  prefs.py              Settings and keys remembered in the browser
  chat_store.py         Chats kept in the browser (record layout, validation)
  preview.py            HTML wireframe preview and overflow estimate
  ui/                   Streamlit UI (sidebar, composer, editor, settings, ...)
smoke_test.py           Offline test suite (no API keys, no network)
project_state.md        Detailed developer notes and file map
```

---

## Development and tests

Run the offline smoke test after every change. It needs no API keys and no network:

```bash
python smoke_test.py
```

It builds a deck with all 11 layouts, checks custom-template inheritance, and covers routing, retrieval, targeted patches, token-limit handling, undo/redo, project files, slide moves, previews, browser preferences and the browser chat store.

`project_state.md` documents the architecture, the file map ("which file to edit for which task") and known constraints.

**Tech stack:** [Streamlit](https://streamlit.io/) (≥ 1.57), [Pydantic v2](https://docs.pydantic.dev/), [python-pptx](https://python-pptx.readthedocs.io/), [Google GenAI SDK](https://pypi.org/project/google-genai/), `pypdf`, `python-docx`, `Pillow`. OpenAI-compatible providers are called with the standard library (`urllib`), with no extra SDKs.
