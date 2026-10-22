"""LLM transport layer: live model discovery and JSON-safe OpenAI-compatible calls.

Edit when: fixing provider/model quirks (reasoning models, token limits), JSON repair, or model listing.
"""

import json  # request / response bodies
import re  # cleaning model output and reading rate-limit error messages
import time  # sleeping before a rate-limit retry
import urllib.error  # HTTPError raised by urlopen on 4xx / 5xx responses
import urllib.request  # zero-dependency HTTP client
from typing import List, Dict, Optional, Type

from google import genai  # Google's SDK, used for Gemini instead of raw HTTP
from google.genai import types
import streamlit as st  # only for st.cache_data on the model list
from pydantic import BaseModel  # base class of the response schemas

from chatppt.schemas import PresentationOutline  # default schema a JSON reply must satisfy


# Substrings that identify non-chat models (speech, embeddings, images, moderation...)
NON_CHAT_KEYWORDS = (
    "whisper", "tts", "audio", "speech", "transcribe", "embed", "embedding",
    "guard", "moderation", "dall-e", "imagen", "veo", "image", "realtime",
    "orpheus", "aqa", "rerank", "babbage", "davinci", "live", "robotics",
    "computer-use", "search-preview", "omni-moderation",
)


def _is_chat_model(model_id: str) -> bool:
    lowered = model_id.lower()
    # A model is treated as a chat model unless its id contains one of the keywords above
    return not any(kw in lowered for kw in NON_CHAT_KEYWORDS)


def _derive_models_url(chat_url: str) -> str:
    """Turn a '.../chat/completions' endpoint into its sibling '.../models' endpoint."""
    url = chat_url.strip().rstrip("/")  # tolerate spaces and a trailing slash in user-typed URLs
    if url.endswith("/chat/completions"):
        # ".../v1/chat/completions" -> ".../v1/models"
        return url[: -len("/chat/completions")] + "/models"
    return url + "/models"  # the user gave a base URL such as ".../v1"


# Cached for 10 minutes per (provider, key, url), so reruns don't refetch the list
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
            actions = getattr(m, "supported_actions", None) or []  # what the model can do, e.g. "generateContent"
            name = (getattr(m, "name", "") or "").replace("models/", "")  # "models/gemini-x" -> "gemini-x"
            # Keep text-generating Gemini models only (no embeddings, Imagen, ...)
            if "generateContent" in actions and name.startswith("gemini") and _is_chat_model(name):
                ids.append(name)
        return sorted(set(ids), reverse=True)  # reverse: newer version numbers first

    # Every other provider: GET the OpenAI-style /models endpoint
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

    # OpenAI format is {"data": [...]}; some providers return the list directly
    items = data.get("data", data) if isinstance(data, dict) else data
    ids = []
    for item in items:
        if not isinstance(item, dict) or not item.get("id"):
            continue  # skip malformed entries
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
    return sorted(set(ids))  # set() removes duplicates, sorted() gives a stable order


def resolve_model_options(provider: str, api_key: str, base_url: str, static_models: List[str]):
    """
    Return (model_list, status_message). Prefers the live list; falls back to the static preset.
    The static list's order is kept for entries that still exist so sensible defaults stay on top.
    """
    static_clean = [m for m in static_models if m != "Custom Model ID..."]  # the preset without the menu entry
    needs_key = provider != "OpenRouter (All-in-One)"  # OpenRouter lists its models without a key
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

    preferred = [m for m in static_clean if m in live]  # recommended presets that the provider still offers
    ordered = preferred + [m for m in live if m not in preferred]  # then every other live model
    return ordered + ["Custom Model ID..."], f"Loaded {len(live)} live models from {provider}."


def extract_clean_json_str(raw_text: Optional[str]) -> str:
    """
    Strip <think>...</think> reasoning blocks (including unclosed <think> tags) and Markdown
    code fences from LLM responses so Pydantic can reliably parse JSON across DeepSeek, Qwen,
    GPT-OSS, Llama, and OpenAI models.
    """
    if not raw_text or not isinstance(raw_text, str):
        return ""

    # Remove complete reasoning blocks; DOTALL lets .*? span lines, the ? keeps each match to one block
    cleaned = re.sub(r"<think>.*?</think>", "", raw_text, flags=re.DOTALL).strip()
    # Unclosed <think> (reasoning cut off or never closed): drop the reasoning up to the first "{" after it
    if "<think>" in cleaned:
        think_pos = cleaned.find("<think>")
        brace_after = cleaned.find("{", think_pos)
        if brace_after != -1:
            cleaned = (cleaned[:think_pos] + cleaned[brace_after:]).strip()
        else:
            cleaned = cleaned[:think_pos].strip()  # no JSON after it: keep only what came before

    # ```json { ... } ``` (or a plain ``` fence): return the object inside the fence
    fence_match = re.search(r"```(?:json)?\s*(\{.*\})\s*```", cleaned, flags=re.DOTALL)
    if fence_match:
        return fence_match.group(1).strip()

    # Otherwise take everything from the first "{" to the last "}", dropping prose around the JSON
    first_brace = cleaned.find("{")
    last_brace = cleaned.rfind("}")
    if first_brace != -1 and last_brace != -1 and last_brace > first_brace:
        return cleaned[first_brace:last_brace + 1]
    return cleaned  # no braces at all: return as-is and let the caller fail on it


def _parses_as_outline(text: str, model: Type[BaseModel] = PresentationOutline) -> bool:
    """Return True if the text is valid JSON that satisfies the given schema (PresentationOutline by default)."""
    try:
        model.model_validate_json(text)  # parses the JSON and checks it against the schema in one step
        return True
    except Exception:  # noqa: BLE001
        return False


def repair_json_str(raw_text: str, model: Type[BaseModel] = PresentationOutline) -> str:
    """
    Best-effort repair for slightly malformed LLM JSON (the most common failures are a stray
    closing brace/bracket or missing closers at the end). Returns the original text unchanged
    if no repair produces a schema-valid outline.
    """
    if not raw_text or _parses_as_outline(raw_text, model):
        return raw_text  # nothing to repair

    # Scan once, ignoring characters inside string literals
    stack: List[str] = []  # closers still expected, innermost last ("}" for "{", "]" for "[")
    closer_positions: List[int] = []  # index of every } and ] outside strings (candidates for a stray one)
    in_string = False  # currently inside "..."
    escaped = False  # previous character was a backslash inside a string
    for idx, ch in enumerate(raw_text):
        if in_string:
            if escaped:
                escaped = False  # this character was escaped (e.g. \" ), so it cannot end the string
            elif ch == "\\":
                escaped = True
            elif ch == '"':
                in_string = False  # closing quote
            continue  # braces inside strings are text, not structure
        if ch == '"':
            in_string = True
        elif ch in "{[":
            stack.append("}" if ch == "{" else "]")  # remember which closer this opener needs
        elif ch in "}]":
            closer_positions.append(idx)
            if stack and stack[-1] == ch:
                stack.pop()  # matches the innermost open container

    # Strategy 1: truncated output -> close any unterminated string/containers
    # Close an open string first, then the containers from innermost to outermost
    candidate = raw_text + ('"' if in_string else "") + "".join(reversed(stack))
    if _parses_as_outline(candidate, model):
        return candidate

    # Strategy 2: a single stray closing brace/bracket -> try removing each one
    for pos in closer_positions:
        candidate = raw_text[:pos] + raw_text[pos + 1:]  # the text with this one character removed
        if _parses_as_outline(candidate, model):
            return candidate

    return raw_text  # no repair worked


def call_openai_compatible_json(
    base_url: str,
    api_key: str,
    model_name: str,
    system_prompt: str,
    user_prompt: str,
    response_model: Type[BaseModel] = PresentationOutline,
) -> str:
    """
    Zero-dependency REST client for OpenAI-compatible chat completion endpoints
    (Groq, OpenAI, DeepSeek, OpenRouter, SiliconFlow, etc.).
    Explicitly budgets completion tokens (8192) and lowers reasoning effort on reasoning
    models (such as Groq's `openai/gpt-oss-120b` and `qwen/qwen3-32b`) so they never
    exhaust the default 1024-token ceiling on internal thinking and return `content: ""`.
    """
    # The JSON Schema of the expected answer, pasted into the system prompt (not every provider
    # supports schema-enforced output, but all of them can read it)
    schema_json = json.dumps(response_model.model_json_schema(), ensure_ascii=False)
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
        "HTTP-Referer": "https://chatppt-v2.streamlit.app",  # OpenRouter shows these two as the calling app
        "X-Title": "ChatPPT V2",
        "User-Agent": "ChatPPT-V2/2.0",
    }

    url_lower = base_url.lower()
    model_lower = model_name.lower()
    is_groq = "groq.com" in url_lower
    is_openai = "openai.com" in url_lower

    last_finish_reason = None  # why the last reply ended ("stop", "length", ...), for the final error message

    # Up to 4 attempts, from most to least strict: JSON mode + token/reasoning params first, then drop
    # whatever the endpoint rejects. `continue` drops the extra params, `break` drops JSON mode.
    for use_json_mode in (True, False):
        for include_extra_params in (True, False):
            payload: Dict[str, object] = {
                "model": model_name,
                "messages": [
                    {"role": "system", "content": full_system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                "temperature": 0.3,  # low randomness: structured output, not creative writing
            }
            if use_json_mode:
                payload["response_format"] = {"type": "json_object"}  # ask the endpoint to return valid JSON only

            if include_extra_params:
                # Reasoning models (e.g. Groq openai/gpt-oss-120b) default to 1024 max_completion_tokens
                # and spend all of it inside `message.reasoning`, leaving `message.content` empty ("").
                # Groq / OpenAI / OpenAI reasoning models use the newer parameter name; others the classic one
                if is_groq or is_openai or model_lower.startswith(("o1", "o3", "o4", "gpt-5")):
                    payload["max_completion_tokens"] = 8192
                else:
                    payload["max_tokens"] = 8192

                # Less thinking leaves more of the token budget for the actual JSON
                if "gpt-oss" in model_lower:
                    payload["reasoning_effort"] = "low"
                elif is_groq and "qwen3" in model_lower:
                    payload["reasoning_effort"] = "none"

            try:
                resp_data = _post_chat_completion(base_url, headers, payload, model_name)
            except urllib.error.HTTPError as http_err:
                err_body = getattr(http_err, "chatppt_body", "")  # body text saved by _post_chat_completion
                lowered_err = err_body.lower()
                if http_err.code == 400:
                    # 1. Groq-style "json_validate_failed": repair `failed_generation` if present
                    try:
                        failed_text = json.loads(err_body).get("error", {}).get("failed_generation")
                    except (ValueError, AttributeError):
                        failed_text = None  # the body was not JSON, or not shaped like Groq's
                    if isinstance(failed_text, str) and failed_text.strip():
                        repaired = repair_json_str(extract_clean_json_str(failed_text), response_model)
                        if _parses_as_outline(repaired, response_model):
                            return repaired  # the rejected answer was fixable: no new request needed

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
                # Any other HTTP error (401 bad key, 404 unknown model, 5xx, ...) cannot be fixed by retrying
                raise RuntimeError(f"Provider API HTTP {http_err.code}: {err_body}") from http_err

            choice0 = (resp_data.get("choices") or [{}])[0]  # first (only) completion; {} if missing
            last_finish_reason = choice0.get("finish_reason")
            msg = choice0.get("message") or {}
            content = msg.get("content") or ""
            cleaned = extract_clean_json_str(content)

            # If content is empty, check if the model placed JSON inside a reasoning field
            if not cleaned or "{" not in cleaned:
                reasoning_raw = msg.get("reasoning") or msg.get("reasoning_content") or ""  # Groq / DeepSeek names
                cand = repair_json_str(extract_clean_json_str(reasoning_raw), response_model)
                if cand and _parses_as_outline(cand, response_model):
                    return cand

            if cleaned and "{" in cleaned:
                return cleaned  # looks like JSON; the caller repairs and validates it

            # Empty content returned (e.g. json_object mode rejected silently or token limit hit):
            # break inner loop to retry without json_object mode.
            break

    # All attempts returned nothing usable
    raise RuntimeError(
        f"The model `{model_name}` returned an empty response (finish_reason={last_finish_reason!r}). "
        "It may have exhausted its output token limit during internal reasoning. "
        "Please try again or switch to another model (such as `llama-3.3-70b-versatile`)."
    )


# Per-minute token limits (e.g. Groq free tier): the provider counts input + requested output cap
MIN_USEFUL_OUTPUT_TOKENS = 1500  # never shrink the output cap below this: a smaller outline would be cut off
TPM_RETRY_MAX_WAIT_S = 30  # longest "try again in Xs" the app is willing to sleep through
# "Limit 8000, Used 1200, Requested 16550" (Used is optional; numbers may contain thousands commas)
_LIMIT_RE = re.compile(r"Limit\s+([\d,]+)\s*,\s*(?:Used\s+([\d,]+)\s*,\s*)?Requested\s+([\d,]+)", re.IGNORECASE)
# "try again in 1m7.5s" or "try again in 7.5s" (group 1 = minutes, group 2 = seconds)
_RETRY_IN_RE = re.compile(r"try again in\s+(?:(\d+)m)?\s*([\d.]+)s", re.IGNORECASE)


class TokenBudgetError(RuntimeError):
    """The provider's tokens-per-minute limit cannot fit this request. The message is user-facing Markdown."""


def _parse_token_limit_error(body: str) -> Optional[Dict[str, Optional[float]]]:
    """Read 'Limit 8000, [Used 1200,] Requested 16550 ... try again in 7.5s' from a 413/429 error body."""
    m = _LIMIT_RE.search(body or "")
    if not m:
        return None  # not a token-limit message we understand
    nums = [int(g.replace(",", "")) if g else None for g in m.groups()]  # "16,550" -> 16550; missing Used -> None
    r = _RETRY_IN_RE.search(body)
    retry_after = (int(r.group(1) or 0) * 60 + float(r.group(2))) if r else None  # total seconds
    return {"limit": nums[0], "used": nums[1], "requested": nums[2], "retry_after": retry_after}


def _token_budget_message(model_name: str, info: Dict[str, Optional[float]], input_tokens: int) -> str:
    limit = int(info["limit"])
    if input_tokens < limit - MIN_USEFUL_OUTPUT_TOKENS:
        # The request itself fits: the minute's budget is just used up by earlier requests
        wait = info.get("retry_after")
        when = f"about {wait:.0f} s" if wait else "a minute"
        return (
            f"**Rate limit reached** for `{model_name}`: your plan allows {limit:,} tokens per minute and this "
            f"minute's budget is used up. Wait {when} and send again, ask for fewer slides, "
            "or switch to a model with a higher per-minute limit."
        )
    # The prompt alone leaves no room for an answer: waiting will not help
    return (
        f"**This request is too large for `{model_name}` on your current provider plan.** The plan allows "
        f"{limit:,} tokens per minute, but the instructions plus your document already need about "
        f"{input_tokens:,} tokens, before any room for the answer.\n\n"
        "Try one of these:\n"
        "- use a shorter document, or paste only the sections you need\n"
        "- pick a model with a higher per-minute limit (on Groq the limits differ per model), or use Google Gemini\n"
        "- upgrade your provider plan"
    )


def _post_chat_completion(url: str, headers: Dict[str, str], payload: Dict[str, object], model_name: str) -> dict:
    """
    POST one chat completion and return the parsed JSON. Handles per-minute token limits:
    - the request alone is over the limit only because of the output cap -> lower the cap and retry;
    - this minute's budget is temporarily used up -> wait once (<= TPM_RETRY_MAX_WAIT_S) and retry;
    - otherwise raise TokenBudgetError with a readable explanation.
    Other HTTP errors are re-raised as HTTPError with the body text in `.chatppt_body` (the stream is consumed).
    """
    # Name of the output-cap field in this payload, or None if the caller set no cap
    cap_key = next((k for k in ("max_completion_tokens", "max_tokens") if k in payload), None)
    waited = False  # the rate-limit wait is done at most once
    for _ in range(3):  # first try + at most two fixes (shrink the cap and/or wait)
        req = urllib.request.Request(
            url.strip(), data=json.dumps(payload).encode("utf-8"), headers=headers, method="POST"
        )
        try:
            with urllib.request.urlopen(req, timeout=90) as resp:
                return json.loads(resp.read().decode("utf-8"))  # success
        except urllib.error.HTTPError as http_err:
            body = http_err.read().decode("utf-8", errors="ignore")  # the body can only be read once
            # 413 = request too large, 429 = rate limited; only these carry token-limit details
            info = _parse_token_limit_error(body) if http_err.code in (413, 429) else None
            if not info:
                http_err.chatppt_body = body  # hand the already-read body to the caller
                raise
            # "Requested" = prompt tokens + the output cap we reserved, so the prompt size is the difference
            cap = int(payload.get(cap_key, 0)) if cap_key else 0
            input_tokens = max(int(info["requested"]) - cap, 0)
            limit = int(info["limit"])
            if int(info["requested"]) > limit:
                # The request alone is too big: shrink the output cap if what is left is still useful
                room = limit - input_tokens - 100  # output tokens that still fit, minus a 100-token margin
                if cap_key and cap > room >= MIN_USEFUL_OUTPUT_TOKENS:
                    payload[cap_key] = room  # the caller's payload is updated too, on purpose
                    continue
            elif not waited and info["retry_after"] is not None and info["retry_after"] <= TPM_RETRY_MAX_WAIT_S:
                # Fits in principle, but this minute's budget is spent: wait as told (+0.5 s), then resend
                time.sleep(info["retry_after"] + 0.5)
                waited = True
                continue
            raise TokenBudgetError(_token_budget_message(model_name, info, input_tokens)) from http_err
    # Three token-limit errors in a row
    raise TokenBudgetError(f"`{model_name}` kept hitting its per-minute token limit. Wait a minute and send again.")


def strip_think_tags(raw_text: Optional[str]) -> str:
    """Remove <think>...</think> blocks (and an unclosed trailing <think>) from a plain-text LLM reply."""
    if not raw_text or not isinstance(raw_text, str):
        return ""
    cleaned = re.sub(r"<think>.*?</think>", "", raw_text, flags=re.DOTALL)  # complete blocks
    if "<think>" in cleaned:
        cleaned = cleaned[: cleaned.find("<think>")]  # unclosed block: drop everything from the tag on
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
    # .text can be None (e.g. a blocked response); treat that as empty
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

    # Attempt 1 with the token-cap / reasoning parameters, attempt 2 without them
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
            # Same parameter-name and reasoning-effort rules as call_openai_compatible_json
            if is_groq or is_openai or model_lower.startswith(("o1", "o3", "o4", "gpt-5")):
                payload["max_completion_tokens"] = max_tokens
            else:
                payload["max_tokens"] = max_tokens
            if "gpt-oss" in model_lower:
                payload["reasoning_effort"] = "low"
            elif is_groq and "qwen3" in model_lower:
                payload["reasoning_effort"] = "none"

        try:
            resp_data = _post_chat_completion(base_url, headers, payload, model_name)
        except urllib.error.HTTPError as http_err:
            err_body = getattr(http_err, "chatppt_body", "")
            lowered_err = err_body.lower()
            # A 400 that names one of the optional parameters: retry without them
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
        # Empty answer: the loop tries once more without the extra parameters

    raise RuntimeError(
        f"The model `{model_name}` returned an empty answer (finish_reason={last_finish_reason!r}). "
        "Please try again or switch to another model."
    )
