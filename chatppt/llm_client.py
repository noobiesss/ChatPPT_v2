"""LLM transport layer: live model discovery and JSON-safe OpenAI-compatible calls.

Edit when: fixing provider/model quirks (reasoning models, token limits), JSON repair, or model listing.
"""

import json
import re
import urllib.error
import urllib.request
from typing import List, Dict, Optional

from google import genai
from google.genai import types
import streamlit as st

from chatppt.schemas import PresentationOutline


# Substrings that identify non-chat models (speech, embeddings, images, moderation...)
NON_CHAT_KEYWORDS = (
    "whisper", "tts", "audio", "speech", "transcribe", "embed", "embedding",
    "guard", "moderation", "dall-e", "imagen", "veo", "image", "realtime",
    "orpheus", "aqa", "rerank", "babbage", "davinci", "live", "robotics",
    "computer-use", "search-preview", "omni-moderation",
)


def _is_chat_model(model_id: str) -> bool:
    lowered = model_id.lower()
    return not any(kw in lowered for kw in NON_CHAT_KEYWORDS)


def _derive_models_url(chat_url: str) -> str:
    """Turn a '.../chat/completions' endpoint into its sibling '.../models' endpoint."""
    url = chat_url.strip().rstrip("/")
    if url.endswith("/chat/completions"):
        return url[: -len("/chat/completions")] + "/models"
    return url + "/models"


@st.cache_data(ttl=600, show_spinner=False)
def fetch_live_models(provider: str, api_key: str, base_url: str) -> List[str]:
    """
    Query the provider's model-listing endpoint and return chat-capable model IDs.
    Raises on any failure (st.cache_data never caches exceptions), so callers can fall back.
    """
    api_key = (api_key or "").strip()

    if provider == "Google Gemini":
        client = genai.Client(api_key=api_key)
        ids = []
        for m in client.models.list():
            actions = getattr(m, "supported_actions", None) or []
            name = (getattr(m, "name", "") or "").replace("models/", "")
            if "generateContent" in actions and name.startswith("gemini") and _is_chat_model(name):
                ids.append(name)
        return sorted(set(ids), reverse=True)

    req = urllib.request.Request(
        _derive_models_url(base_url),
        headers={
            "Authorization": f"Bearer {api_key}",
            "User-Agent": "ChatPPT-V2/2.0",
        },
        method="GET",
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        data = json.loads(resp.read().decode("utf-8"))

    items = data.get("data", data) if isinstance(data, dict) else data
    ids = []
    for item in items:
        if not isinstance(item, dict) or not item.get("id"):
            continue
        # Groq flags retired/inactive models via an "active" field
        if item.get("active") is False:
            continue
        model_id = item["id"]
        if not _is_chat_model(model_id):
            continue
        # OpenAI lists many legacy/non-chat families; keep only chat-capable prefixes
        if provider == "OpenAI (Paid)" and not model_id.startswith(("gpt-", "chatgpt-", "o1", "o3", "o4")):
            continue
        ids.append(model_id)
    return sorted(set(ids))


def resolve_model_options(provider: str, api_key: str, base_url: str, static_models: List[str]):
    """
    Return (model_list, status_message). Prefers the live list; falls back to the static preset.
    The static list's order is kept for entries that still exist so sensible defaults stay on top.
    """
    static_clean = [m for m in static_models if m != "Custom Model ID..."]
    needs_key = provider != "OpenRouter (All-in-One)"
    if needs_key and not (api_key or "").strip():
        return static_models, "Enter your API key to load the live model list."
    if provider == "Custom OpenAI-Compatible API" and not (base_url or "").strip():
        return static_models, "Enter an endpoint URL to load models."

    try:
        live = fetch_live_models(provider, api_key, base_url or "")
    except Exception as exc:  # noqa: BLE001 - any network/auth/parse failure falls back
        return static_models, f"Could not fetch live models ({type(exc).__name__}); showing presets."
    if not live:
        return static_models, "Provider returned no chat models; showing presets."

    preferred = [m for m in static_clean if m in live]
    ordered = preferred + [m for m in live if m not in preferred]
    return ordered + ["Custom Model ID..."], f"Loaded {len(live)} live models from {provider}."


def extract_clean_json_str(raw_text: Optional[str]) -> str:
    """
    Strip <think>...</think> reasoning blocks (including unclosed <think> tags) and Markdown
    code fences from LLM responses so Pydantic can reliably parse JSON across DeepSeek, Qwen,
    GPT-OSS, Llama, and OpenAI models.
    """
    if not raw_text or not isinstance(raw_text, str):
        return ""

    cleaned = re.sub(r"<think>.*?</think>", "", raw_text, flags=re.DOTALL).strip()
    if "<think>" in cleaned:
        think_pos = cleaned.find("<think>")
        brace_after = cleaned.find("{", think_pos)
        if brace_after != -1:
            cleaned = (cleaned[:think_pos] + cleaned[brace_after:]).strip()
        else:
            cleaned = cleaned[:think_pos].strip()

    fence_match = re.search(r"```(?:json)?\s*(\{.*\})\s*```", cleaned, flags=re.DOTALL)
    if fence_match:
        return fence_match.group(1).strip()

    first_brace = cleaned.find("{")
    last_brace = cleaned.rfind("}")
    if first_brace != -1 and last_brace != -1 and last_brace > first_brace:
        return cleaned[first_brace:last_brace + 1]
    return cleaned


def _parses_as_outline(text: str) -> bool:
    """Return True if the text is valid JSON that satisfies the PresentationOutline schema."""
    try:
        PresentationOutline.model_validate_json(text)
        return True
    except Exception:  # noqa: BLE001
        return False


def repair_json_str(raw_text: str) -> str:
    """
    Best-effort repair for slightly malformed LLM JSON (the most common failures are a stray
    closing brace/bracket or missing closers at the end). Returns the original text unchanged
    if no repair produces a schema-valid outline.
    """
    if not raw_text or _parses_as_outline(raw_text):
        return raw_text

    # Scan once, ignoring characters inside string literals
    stack: List[str] = []
    closer_positions: List[int] = []
    in_string = False
    escaped = False
    for idx, ch in enumerate(raw_text):
        if in_string:
            if escaped:
                escaped = False
            elif ch == "\\":
                escaped = True
            elif ch == '"':
                in_string = False
            continue
        if ch == '"':
            in_string = True
        elif ch in "{[":
            stack.append("}" if ch == "{" else "]")
        elif ch in "}]":
            closer_positions.append(idx)
            if stack and stack[-1] == ch:
                stack.pop()

    # Strategy 1: truncated output -> close any unterminated string/containers
    candidate = raw_text + ('"' if in_string else "") + "".join(reversed(stack))
    if _parses_as_outline(candidate):
        return candidate

    # Strategy 2: a single stray closing brace/bracket -> try removing each one
    for pos in closer_positions:
        candidate = raw_text[:pos] + raw_text[pos + 1:]
        if _parses_as_outline(candidate):
            return candidate

    return raw_text


def call_openai_compatible_json(
    base_url: str,
    api_key: str,
    model_name: str,
    system_prompt: str,
    user_prompt: str,
) -> str:
    """
    Zero-dependency REST client for OpenAI-compatible chat completion endpoints
    (Groq, OpenAI, DeepSeek, OpenRouter, SiliconFlow, etc.).
    Explicitly budgets completion tokens (8192) and lowers reasoning effort on reasoning
    models (such as Groq's `openai/gpt-oss-120b` and `qwen/qwen3-32b`) so they never
    exhaust the default 1024-token ceiling on internal thinking and return `content: ""`.
    """
    schema_json = json.dumps(PresentationOutline.model_json_schema(), ensure_ascii=False)
    full_system_prompt = (
        f"{system_prompt}\n\n"
        "### OUTPUT FORMAT REQUIREMENT:\n"
        "You MUST respond ONLY with a valid JSON object that strictly conforms to the following JSON Schema. "
        "Do NOT wrap the JSON in markdown explanations or add extra text outside the JSON object.\n"
        f"JSON Schema:\n{schema_json}"
    )

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key.strip()}",
        "HTTP-Referer": "https://chatppt-v2.streamlit.app",
        "X-Title": "ChatPPT V2",
        "User-Agent": "ChatPPT-V2/2.0",
    }

    url_lower = base_url.lower()
    model_lower = model_name.lower()
    is_groq = "groq.com" in url_lower
    is_openai = "openai.com" in url_lower

    last_finish_reason = None

    for use_json_mode in (True, False):
        for include_extra_params in (True, False):
            payload: Dict[str, object] = {
                "model": model_name,
                "messages": [
                    {"role": "system", "content": full_system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                "temperature": 0.3,
            }
            if use_json_mode:
                payload["response_format"] = {"type": "json_object"}

            if include_extra_params:
                # Reasoning models (e.g. Groq openai/gpt-oss-120b) default to 1024 max_completion_tokens
                # and spend all of it inside `message.reasoning`, leaving `message.content` empty ("").
                if is_groq or is_openai or model_lower.startswith(("o1", "o3", "o4", "gpt-5")):
                    payload["max_completion_tokens"] = 8192
                else:
                    payload["max_tokens"] = 8192

                if "gpt-oss" in model_lower:
                    payload["reasoning_effort"] = "low"
                elif is_groq and "qwen3" in model_lower:
                    payload["reasoning_effort"] = "none"

            req = urllib.request.Request(
                base_url.strip(),
                data=json.dumps(payload).encode("utf-8"),
                headers=headers,
                method="POST",
            )
            try:
                with urllib.request.urlopen(req, timeout=90) as resp:
                    resp_data = json.loads(resp.read().decode("utf-8"))
                    choice0 = (resp_data.get("choices") or [{}])[0]
                    last_finish_reason = choice0.get("finish_reason")
                    msg = choice0.get("message") or {}
                    content = msg.get("content") or ""
                    cleaned = extract_clean_json_str(content)

                    # If content is empty, check if the model placed JSON inside a reasoning field
                    if not cleaned or "{" not in cleaned:
                        reasoning_raw = msg.get("reasoning") or msg.get("reasoning_content") or ""
                        cand = extract_clean_json_str(reasoning_raw)
                        if cand and _parses_as_outline(repair_json_str(cand)):
                            return repair_json_str(cand)

                    if cleaned and "{" in cleaned:
                        return cleaned

                    # Empty content returned (e.g. json_object mode rejected silently or token limit hit):
                    # break inner loop to retry without json_object mode.
                    break
            except urllib.error.HTTPError as http_err:
                err_body = http_err.read().decode("utf-8", errors="ignore")
                lowered_err = err_body.lower()
                if http_err.code == 400:
                    # 1. Groq-style "json_validate_failed": repair `failed_generation` if present
                    try:
                        failed_text = json.loads(err_body).get("error", {}).get("failed_generation")
                    except (ValueError, AttributeError):
                        failed_text = None
                    if isinstance(failed_text, str) and failed_text.strip():
                        repaired = repair_json_str(extract_clean_json_str(failed_text))
                        if _parses_as_outline(repaired):
                            return repaired

                    # 2. Endpoint rejected max_completion_tokens / max_tokens / reasoning_effort:
                    # retry immediately without those optional parameters
                    if include_extra_params and any(
                        k in lowered_err
                        for k in ("max_completion_tokens", "max_tokens", "reasoning_effort", "unrecognized", "unknown")
                    ):
                        continue

                    # 3. Model rejected json_object mode or failed strict JSON generation:
                    # break inner loop to retry with use_json_mode=False
                    if use_json_mode and (
                        "response_format" in lowered_err
                        or "json_validate_failed" in lowered_err
                        or "failed to generate json" in lowered_err
                        or "json" in lowered_err
                    ):
                        break
                raise RuntimeError(f"Provider API HTTP {http_err.code}: {err_body}") from http_err

    raise RuntimeError(
        f"The model `{model_name}` returned an empty response (finish_reason={last_finish_reason!r}). "
        "It may have exhausted its output token limit during internal reasoning. "
        "Please try again or switch to another model (such as `llama-3.3-70b-versatile`)."
    )


def strip_think_tags(raw_text: Optional[str]) -> str:
    """Remove <think>...</think> blocks (and an unclosed trailing <think>) from a plain-text LLM reply."""
    if not raw_text or not isinstance(raw_text, str):
        return ""
    cleaned = re.sub(r"<think>.*?</think>", "", raw_text, flags=re.DOTALL)
    if "<think>" in cleaned:
        cleaned = cleaned[: cleaned.find("<think>")]
    return cleaned.strip()


def call_gemini_text(
    api_key: str,
    model_name: str,
    system_prompt: str,
    user_prompt: str,
    temperature: float = 0.4,
) -> str:
    """Plain-text (non-JSON) Gemini call used for conversational answers."""
    client = genai.Client(api_key=(api_key or "").strip())
    response = client.models.generate_content(
        model=model_name,
        contents=user_prompt,
        config=types.GenerateContentConfig(system_instruction=system_prompt, temperature=temperature),
    )
    text = strip_think_tags(getattr(response, "text", None) or "")
    if not text:
        raise RuntimeError(f"Google Gemini (`{model_name}`) returned an empty response.")
    return text


def call_openai_compatible_text(
    base_url: str,
    api_key: str,
    model_name: str,
    system_prompt: str,
    user_prompt: str,
    max_tokens: int = 4096,
    temperature: float = 0.4,
) -> str:
    """
    Plain-text counterpart of call_openai_compatible_json for conversational answers.
    Applies the same reasoning-model quirks (token parameter name, reasoning effort) and retries
    once without the optional parameters if the endpoint rejects them.
    """
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key.strip()}",
        "HTTP-Referer": "https://chatppt-v2.streamlit.app",
        "X-Title": "ChatPPT V2",
        "User-Agent": "ChatPPT-V2/2.0",
    }
    url_lower = base_url.lower()
    model_lower = model_name.lower()
    is_groq = "groq.com" in url_lower
    is_openai = "openai.com" in url_lower
    last_finish_reason = None

    for include_extra_params in (True, False):
        payload: Dict[str, object] = {
            "model": model_name,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "temperature": temperature,
        }
        if include_extra_params:
            if is_groq or is_openai or model_lower.startswith(("o1", "o3", "o4", "gpt-5")):
                payload["max_completion_tokens"] = max_tokens
            else:
                payload["max_tokens"] = max_tokens
            if "gpt-oss" in model_lower:
                payload["reasoning_effort"] = "low"
            elif is_groq and "qwen3" in model_lower:
                payload["reasoning_effort"] = "none"

        req = urllib.request.Request(
            base_url.strip(), data=json.dumps(payload).encode("utf-8"), headers=headers, method="POST"
        )
        try:
            with urllib.request.urlopen(req, timeout=90) as resp:
                resp_data = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as http_err:
            err_body = http_err.read().decode("utf-8", errors="ignore")
            lowered_err = err_body.lower()
            if (
                http_err.code == 400
                and include_extra_params
                and any(k in lowered_err for k in ("max_completion_tokens", "max_tokens", "reasoning_effort", "unrecognized", "unknown"))
            ):
                continue
            raise RuntimeError(f"Provider API HTTP {http_err.code}: {err_body}") from http_err

        choice0 = (resp_data.get("choices") or [{}])[0]
        last_finish_reason = choice0.get("finish_reason")
        text = strip_think_tags((choice0.get("message") or {}).get("content") or "")
        if text:
            return text

    raise RuntimeError(
        f"The model `{model_name}` returned an empty answer (finish_reason={last_finish_reason!r}). "
        "Please try again or switch to another model."
    )