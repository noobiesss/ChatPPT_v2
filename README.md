# 📊 ChatPPT V2 — Smart Visual Presentation Generator

**ChatPPT V2** is an AI-powered presentation engineering web app built with **Python**, **Streamlit**, **Pydantic v2**, and **`python-pptx`**. Instead of generating monotonous bullet-point slides, ChatPPT acts as a consulting-style presentation architect—featuring a **Claude-style conversational UI**, **multi-provider LLM support** with **live model discovery**, **11 distinct visual slide layouts**, **native editable PowerPoint charts**, **Smart Hybrid image routing**, and **deep `.pptx` custom template inheritance**.

---

## ✨ Key Features

* **Claude-Style Conversational Interface:**
  * Manage multiple presentation sessions in the left sidebar (`＋ New chat`, switch, or delete chats).
  * Sticky bottom composer bar with popovers for file attachment (`📎`), theme/template styling (`🎨`), tone selection, provider/model picker (`🤖`), and slide count/language (`⚙`).
  * **Conversational Revisions:** Send short follow-up prompts (e.g., *"make it 6 slides and add more charts"*) to regenerate and refine the deck outline from the same source material.
* **Multi-Provider BYOK (Bring Your Own Key) Architecture:**
  * Supports **Google Gemini** (native `google-genai` SDK), **Groq**, **OpenAI**, **DeepSeek**, **OpenRouter**, and any **Custom OpenAI-Compatible API** (SiliconFlow, Together AI, Ollama, vLLM).
  * **Live Model Discovery:** Automatically queries the provider's `/models` endpoint (and filters out non-chat models like Whisper/TTS/Embeddings) with a 10-minute cache and manual refresh button.
  * **Automatic JSON Repair & Slide-Count Verification:** Strips `<think>` reasoning blocks, repairs truncated or stray-brace JSON (`failed_generation`), and automatically retries or trims if a model misses the target slide count.
  * **In-Memory Session Privacy:** All API keys live strictly in the Settings dialog (`⚙️ Settings`) inside per-session server memory and are never written to disk.
* **Multi-Source Document & Embedded Figure Ingestion:**
  * Upload `.pdf`, `.docx`, `.txt`, or `.md` reports, or paste raw notes directly.
  * Automatically extracts embedded diagrams, charts, and images from `.pdf` and `.docx` files (logos, icons and duplicates are skipped). With the **Report Images First** image mode they fill `image_split` slides in document order, or you pick a figure per slide in the outline editor.
* **Consulting-Grade Copywriting & 11 Visual Layouts:**
  * Enforces conclusion-first **Action Titles**, scannable **Bold Heading + Concise Detail** anchor cards, and a **Bottom Takeaway ("So What?") Banner** on content slides.
* **Native Editable PowerPoint Charts:**
  * Extracts real numerical series from your report and renders native, editable PowerPoint Column, Line, and Pie charts (`XL_CHART_TYPE`).
* **Smart Hybrid Image Engine:**
  * Automatically routes `image_split` slides between **Unsplash Stock Photography** (`stock_photo` for real-world scenes, teams, hardware, cities) and **AI Concept Illustrations** (`ai_illustration` via **Google Imagen 3** or **Pollinations AI** fallback), plus **Extracted Report Figures**, **Custom Image Uploads**, and **Wireframe Placeholders**.
  * Uses `Pillow (PIL)` for aspect-ratio-safe bounding box fitting and automatic `WEBP`/`RGBA` to `RGB` conversion.
* **Built-in 16:9 Themes & Smart Custom `.pptx` Template Inheritance:**
  * Choose from 4 built-in widescreen palettes (*Corporate Blue*, *Modern Dark*, *Minimalist Light*, *Emerald Forest*) or upload your own `.pptx` template.
  * Custom `.pptx` mode automatically scales non-16:9 canvases (e.g., `10 × 5.625 in` Canva/Google Slides exports) to `13.333 × 7.5 in`, preserves Slide Master/layout backgrounds, replays decorative slide artwork/frames/corner logos, and extracts the template's `theme1.xml` / slide color scheme and fonts.

---

## 🎨 Supported Visual Slide Layouts

| Layout Archetype (`LayoutType`) | Visual Structure | Best Used For |
| :--- | :--- | :--- |
| **`cover`** | Widescreen title, accent bar, and subtitle | Slide 1 opening cover |
| **`section_divider`** | High-contrast background with section header & context | Theme transitions in decks with 8+ slides |
| **`bullet_cards`** | 3–4 stacked horizontal cards with numbered badges | Sequential explanations & general takeaways |
| **`three_cards`** | 3 parallel vertical pillar cards with top accent strips | 3 core pillars, architecture layers, or strategies |
| **`two_column`** | Side-by-side comparison containers with colored headers | Pros vs. Cons, Before vs. After, Problem vs. Solution |
| **`kpi_metrics`** | 2–4 big-number hero cards (`46pt` bold metric + context) | Financial highlights, latency drops, ROI, benchmarks |
| **`chart_slide`** | Native editable PPTX chart (left) + insight cards (right) | Time-series trends, category comparisons, market share |
| **`image_split`** | Framed image + caption on one side, key points on the other | Product visuals, system diagrams, workflows, case studies |
| **`timeline_flow`** | 3–5 numbered nodes on a connector line, stage label above and a step card below each | Roadmaps, phased plans, milestones, business processes |
| **`comparison_table`** | Native editable PowerPoint table: accent header row, zebra rows, bold criterion column | Options / vendors / scenarios compared across several criteria |
| **`closing`** | Large centered phrase + accent bar + optional presenter line (opt-in only) | Thank-you / Q&A slide when explicitly requested |

---

## 🛠️ Tech Stack

* **Frontend & Conversational UI:** [Streamlit](https://streamlit.io/) (`>=1.57.0`, theme-neutral CSS for light & dark mode)
* **LLM Providers:** [Google GenAI SDK](https://pypi.org/project/google-genai/) + zero-dependency OpenAI-compatible REST client (`urllib.request`) for Groq, OpenAI, DeepSeek, OpenRouter, and custom endpoints
* **Schema Validation & Structured Output:** [Pydantic v2](https://docs.pydantic.dev/)
* **Slide & Chart Rendering Engine:** [python-pptx](https://python-pptx.readthedocs.io/)
* **Document & Image Processing:** `pypdf`, `python-docx`, `Pillow (PIL)`

---

## 🚀 Installation & Quick Start

### 1. Prerequisites
* **Python 3.9+** installed on your system.
* At least one LLM API Key (**Google Gemini**, **Groq**, **OpenAI**, **DeepSeek**, or **OpenRouter**).

### 2. Install Dependencies
Clone or download this repository, navigate to the project folder, and install the required packages:

```bash
pip install -r requirements.txt