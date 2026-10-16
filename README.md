# 📊 ChatPPT V2 — Smart Visual Presentation Generator

**ChatPPT V2** is an AI-powered presentation engineering tool built with **Python**, **Streamlit**, **Google Gemini (`google-genai`)**, and **`python-pptx`**. Instead of generating monotonous bullet-point slides, ChatPPT acts as a consulting-style presentation architect—automatically routing raw reports into **8 distinct visual slide layouts**, generating **native editable PowerPoint charts**, extracting **embedded figures from PDFs/Word documents**, and synthesizing **AI illustrations**.

---

## ✨ Key Features

* **Multi-Source Document Ingestion:** Upload `.pdf`, `.docx`, `.txt`, or `.md` reports, or paste raw notes directly.
* **Automatic Embedded Image Extraction:** Automatically extracts existing diagrams, charts, and photos embedded inside uploaded `.pdf` and `.docx` files so you can reuse them on any slide.
* **Smart Layout Routing (8 Archetypes):** Powered by Gemini 2.5 Flash Structured Outputs (`Pydantic` schema validation) to prevent consecutive repetitive layouts and match visual form to content logic.
* **Consulting-Grade Copywriting:** Enforces conclusion-first **Action Titles**, scannable **Bold Heading + Concise Detail** anchor cards, and a **Bottom Takeaway ("So What?") Banner** on content slides.
* **Native Editable PowerPoint Charts:** Extracts real numerical series from your report and renders native, editable PowerPoint Column, Line, and Pie charts (`XL_CHART_TYPE`).
* **3-Tier Image & Visual Engine:**
  1. **AI Concept Illustration:** Generates 4:3 presentation graphics via **Google Imagen 3** (`imagen-3.0-generate-002`) with an automatic zero-key fallback to **Pollinations AI**.
  2. **Report Image Reuse:** Assign extracted PDF/DOCX figures to specific slides from a dropdown.
  3. **Custom Upload / Blueprint Placeholder:** Upload custom images per slide or render clean wireframe placeholders for offline/confidential workflows.
* **Aspect-Ratio Safe Image Fitting:** Uses `Pillow (PIL)` to calculate bounding boxes, convert `WEBP`/`RGBA` formats to `RGB`, and center images without stretching or distortion.
* **2-Step Interactive Editor:** Review and modify every slide's layout type, titles, structured points, chart numbers, and image sources in the browser before rendering the final `.pptx`.
* **Built-in 16:9 Themes & Custom `.pptx` Template Support:** Choose from 4 widescreen color palettes (*Corporate Blue*, *Modern Dark*, *Minimalist Light*, *Emerald Forest*) or upload your own `.pptx` template to inherit its Slide Master background.

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

---

## 🛠️ Tech Stack

* **Frontend & Workflow UI:** [Streamlit](https://streamlit.io/)
* **LLM & Structured Routing:** [Google GenAI SDK](https://pypi.org/project/google-genai/) (`gemini-2.5-flash` & `imagen-3.0-generate-002`)
* **Schema Validation:** [Pydantic v2](https://docs.pydantic.dev/)
* **Slide Rendering Engine:** [python-pptx](https://python-pptx.readthedocs.io/)
* **Document & Image Processing:** `pypdf`, `python-docx`, `Pillow (PIL)`

---

## 🚀 Installation & Quick Start

### 1. Prerequisites
* **Python 3.9+** installed on your system.
* A **Google Gemini API Key** (get one from [Google AI Studio](https://aistudio.google.com/)).

### 2. Install Dependencies
Clone or download this repository, navigate to the project folder, and install the required packages:

```bash
pip install -r requirements.txt
```

### 3. Run the Application
Launch the Streamlit web server:

```bash
streamlit run app.py
```

The app will automatically open in your default browser at `http://localhost:8501`.

---

## 📖 Usage Guide (3-Step Workflow)

### Step 1: Configure & Provide Source Material
1. Open the left sidebar and enter your **Gemini API Key**.
2. Set your desired **Number of Slides** (4–20), **Presentation Tone**, and **Output Language**.
3. Choose between a **Built-in 16:9 Color Theme** or **Upload Custom `.pptx` Template**.
4. Upload your report (`.pdf`, `.docx`, `.txt`, `.md`) and/or paste text into the main input area, then click **Analyze Report & Generate Smart Visual Outline**.

### Step 2: Review & Customize the Outline
Expand any slide card in the interactive editor to fine-tune its content or switch its **Slide Layout Type** dynamically:
* **Structured Points (`bullet_cards`, `three_cards`, `chart_slide`, `image_split`):** Edit one point per line using the `Heading :: Detail` format:
  ```text
  Automated Caching :: Reduced peak database read latency by 62% across Q3 traffic loads.
  ```
* **Hero KPIs (`kpi_metrics`):** Edit one metric per line using the `Value | Label | Context` format:
  ```text
  +62% | Throughput Gain | Measured across 10k concurrent requests
  ```
* **Native Charts (`chart_slide`):** Switch between `bar`, `line`, and `pie` charts, and edit comma-separated categories and numerical values.
* **Visuals (`image_split`):** Choose between **AI Generated Illustration**, **Use Image from Report** (if your PDF/DOCX contained embedded figures), **Upload Custom Image**, or **Visual Placeholder Only**.

### Step 3: Render & Download
Click **Render Complete Presentation (.pptx)** to build the slide deck and download `ChatPPT_V2_Presentation.pptx`. All shapes, text boxes, and charts remain 100% natively editable inside Microsoft PowerPoint, Keynote, or Google Slides.

---

## 📁 Project Structure

```text
ChatPPT/
├── app.py              # Main Streamlit application, Gemini router, and python-pptx rendering engine
├── requirements.txt    # Python package dependencies
└── README.md           # Project documentation
```