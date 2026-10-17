"""Static presets: LLM providers, built-in color themes, tone/language options.

Edit when: adding an LLM provider/preset model list, a color theme, a tone, or an output language.
"""

from typing import Dict

from pptx.dml.color import RGBColor


LLM_PROVIDERS: Dict[str, Dict] = {
    "Google Gemini": {
        "base_url": None,  # Handled natively via google-genai SDK
        "models": [
            "gemini-3.8-flash",
            "gemini-3.5-flash",
            "gemini-3.1-flash-lite",
            "gemini-2.5-flash",
            "Custom Model ID...",
        ],
        "key_label": "Google Gemini API Key",
        "key_help": "Enter your Google AI Studio / Vertex Gemini API Key.",
    },
    "Groq (Fast & Free Tier)": {
        "base_url": "https://api.groq.com/openai/v1/chat/completions",
        "models": [
            "llama-3.3-70b-versatile",
            "openai/gpt-oss-120b",
            "meta-llama/llama-4-scout-17b-16e-instruct",
            "qwen/qwen3-32b",
            "llama-3.1-8b-instant",
            "Custom Model ID...",
        ],
        "key_label": "Groq API Key (gsk_...)",
        "key_help": "Get a high-speed key from console.groq.com.",
    },
    "OpenAI (Paid)": {
        "base_url": "https://api.openai.com/v1/chat/completions",
        "models": [
            "gpt-4o",
            "gpt-4o-mini",
            "gpt-4.1",
            "gpt-4.1-mini",
            "Custom Model ID...",
        ],
        "key_label": "OpenAI API Key (sk-...)",
        "key_help": "Enter your OpenAI Platform API Key.",
    },
    "DeepSeek (Paid)": {
        "base_url": "https://api.deepseek.com/chat/completions",
        "models": [
            "deepseek-chat",
            "deepseek-reasoner",
            "Custom Model ID...",
        ],
        "key_label": "DeepSeek API Key",
        "key_help": "Enter your DeepSeek Platform API Key.",
    },
    "OpenRouter (All-in-One)": {
        "base_url": "https://openrouter.ai/api/v1/chat/completions",
        "models": [
            "anthropic/claude-3.7-sonnet",
            "google/gemini-2.5-pro",
            "meta-llama/llama-3.3-70b-instruct",
            "deepseek/deepseek-chat",
            "Custom Model ID...",
        ],
        "key_label": "OpenRouter API Key (sk-or-...)",
        "key_help": "Access Claude, Gemini, Llama, and DeepSeek with one OpenRouter key.",
    },
    "Custom OpenAI-Compatible API": {
        "base_url": "https://api.openai.com/v1/chat/completions",
        "models": ["Custom Model ID..."],
        "key_label": "Custom Provider API Key",
        "key_help": "Works with SiliconFlow, Together AI, Fireworks, vLLM, or Ollama.",
    },
}


BUILTIN_THEMES: Dict[str, Dict[str, RGBColor]] = {
    "Corporate Blue": {
        "bg_cover": RGBColor(15, 37, 55),
        "bg_section": RGBColor(22, 52, 77),
        "bg_content": RGBColor(244, 247, 251),
        "card_bg": RGBColor(255, 255, 255),
        "card_border": RGBColor(215, 225, 238),
        "accent": RGBColor(0, 114, 206),
        "title_cover": RGBColor(255, 255, 255),
        "subtitle_cover": RGBColor(185, 215, 240),
        "title_content": RGBColor(15, 37, 55),
        "body_content": RGBColor(55, 65, 80),
        "takeaway_bg": RGBColor(15, 37, 55),
        "takeaway_text": RGBColor(255, 255, 255),
    },
    "Modern Dark": {
        "bg_cover": RGBColor(16, 16, 22),
        "bg_section": RGBColor(24, 24, 34),
        "bg_content": RGBColor(22, 22, 30),
        "card_bg": RGBColor(34, 35, 48),
        "card_border": RGBColor(58, 60, 82),
        "accent": RGBColor(99, 102, 241),
        "title_cover": RGBColor(255, 255, 255),
        "subtitle_cover": RGBColor(165, 170, 195),
        "title_content": RGBColor(242, 244, 255),
        "body_content": RGBColor(195, 200, 220),
        "takeaway_bg": RGBColor(48, 50, 78),
        "takeaway_text": RGBColor(240, 242, 255),
    },
    "Minimalist Light": {
        "bg_cover": RGBColor(245, 245, 243),
        "bg_section": RGBColor(235, 235, 232),
        "bg_content": RGBColor(252, 252, 250),
        "card_bg": RGBColor(255, 255, 255),
        "card_border": RGBColor(220, 220, 218),
        "accent": RGBColor(35, 35, 35),
        "title_cover": RGBColor(20, 20, 20),
        "subtitle_cover": RGBColor(95, 95, 95),
        "title_content": RGBColor(20, 20, 20),
        "body_content": RGBColor(65, 65, 65),
        "takeaway_bg": RGBColor(35, 35, 35),
        "takeaway_text": RGBColor(255, 255, 255),
    },
    "Emerald Forest": {
        "bg_cover": RGBColor(18, 43, 32),
        "bg_section": RGBColor(26, 58, 44),
        "bg_content": RGBColor(243, 248, 245),
        "card_bg": RGBColor(255, 255, 255),
        "card_border": RGBColor(205, 225, 212),
        "accent": RGBColor(40, 135, 85),
        "title_cover": RGBColor(255, 255, 255),
        "subtitle_cover": RGBColor(185, 220, 200),
        "title_content": RGBColor(18, 43, 32),
        "body_content": RGBColor(45, 65, 55),
        "takeaway_bg": RGBColor(18, 43, 32),
        "takeaway_text": RGBColor(255, 255, 255),
    },
}


CUSTOM_PROVIDER = "Custom OpenAI-Compatible API"


OPENROUTER_PROVIDER = "OpenRouter (All-in-One)"


TONE_OPTIONS = {
    "Executive (McKinsey)": "Executive & Consulting (McKinsey Style)",
    "Professional": "Professional & Corporate",
    "Academic": "Academic & Analytical",
    "Venture Pitch": "Persuasive & Venture Pitch",
    "Technical": "Technical Architecture & Engineering",
}


LANGUAGE_OPTIONS = ["English", "Simplified Chinese", "Traditional Chinese", "Malay", "Japanese"]


REVISION_MAX_CHARS = 600  # Short follow-up messages are treated as revision instructions
