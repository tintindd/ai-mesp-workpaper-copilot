from __future__ import annotations

import json
import os
import urllib.error
import urllib.request
from dataclasses import dataclass
from typing import Any


DEFAULT_DEEPSEEK_BASE_URL = "https://api.deepseek.com"
DEFAULT_DEEPSEEK_MODEL = "deepseek-chat"


@dataclass
class DeepSeekConfig:
    api_key: str
    base_url: str = DEFAULT_DEEPSEEK_BASE_URL
    model: str = DEFAULT_DEEPSEEK_MODEL


def load_deepseek_config(secrets: Any | None = None) -> DeepSeekConfig | None:
    api_key = os.getenv("DEEPSEEK_API_KEY")
    base_url = os.getenv("DEEPSEEK_BASE_URL", DEFAULT_DEEPSEEK_BASE_URL)
    model = os.getenv("DEEPSEEK_MODEL", DEFAULT_DEEPSEEK_MODEL)

    if secrets is not None:
        api_key = _secret_value(secrets, "DEEPSEEK_API_KEY", api_key)
        base_url = _secret_value(secrets, "DEEPSEEK_BASE_URL", base_url)
        model = _secret_value(secrets, "DEEPSEEK_MODEL", model)

    if not api_key:
        return None
    return DeepSeekConfig(api_key=api_key, base_url=base_url.rstrip("/"), model=model)


def _secret_value(secrets: Any, key: str, default: str | None = None) -> str | None:
    try:
        value = secrets.get(key, default)
    except Exception:
        return default
    return value or default


def deepseek_chat(
    messages: list[dict[str, str]],
    config: DeepSeekConfig,
    *,
    temperature: float = 0.1,
    max_tokens: int = 1200,
) -> str:
    payload = {
        "model": config.model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
        "stream": False,
    }
    request = urllib.request.Request(
        f"{config.base_url}/chat/completions",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {config.api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=45) as response:
            body = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"DeepSeek API HTTP {exc.code}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"DeepSeek API connection failed: {exc.reason}") from exc

    choices = body.get("choices") or []
    if not choices:
        raise RuntimeError("DeepSeek API returned no choices.")
    content = (choices[0].get("message") or {}).get("content")
    if not content:
        raise RuntimeError("DeepSeek API returned an empty response.")
    return content


def classify_filename_with_deepseek(filename: str, config: DeepSeekConfig) -> str:
    return deepseek_chat(
        [
            {
                "role": "system",
                "content": (
                    "You are helping normalize SAP audit evidence filenames. "
                    "Return concise JSON only. Extract sample_no, order_id, material_id, "
                    "report_type among CO03, KSBT, 3611, CKM3, and suggest a standard filename."
                ),
            },
            {"role": "user", "content": filename},
        ],
        config,
        temperature=0,
        max_tokens=500,
    )
