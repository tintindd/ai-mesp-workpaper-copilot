from __future__ import annotations

import json
import base64
import mimetypes
import os
import re
import time
import urllib.error
import urllib.request
import socket
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"}
DEFAULT_QWEN_OCR_API_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
DEFAULT_QWEN_OCR_MODEL = "qwen-vl-ocr-latest"
DEFAULT_ZHIPUAI_API_URL = "https://open.bigmodel.cn/api/paas/v4/chat/completions"
DEFAULT_ZHIPUAI_MODEL = "glm-4.6v-flash"
CKM3_OCR_PROMPT = (
    "你是 SAP CKM3 截图表格识别助手。请只从图片中识别 CKM3 明细表，"
    "先输出两行关键字段：订单编号: <识别值>、物料ID: <识别值>；识别不到填空。"
    "然后输出一个 HTML table，不要输出其他解释。表格第一行必须是表头。"
    "表头依次为：类别、交易数量、数量单位、初始评估、价格差异、工率差异、实际值、价格、"
    "公司间利润、直接材料、直接人工、间接人工、折旧与摊销、能耗、低值易耗、其他制造费用、成本构成总和。"
    "保留负号、尾随负号和千分位数字；识别不到的数字填 0。"
)


@dataclass
class OnlineOCRConfig:
    provider: str = "qwen"
    api_url: str = DEFAULT_QWEN_OCR_API_URL
    api_key: str | None = None
    auth_header: str = "Authorization"
    auth_scheme: str = "Bearer"
    model: str = DEFAULT_QWEN_OCR_MODEL
    retry_attempts: int = 3
    fallback_api_url: str | None = None
    fallback_api_key: str | None = None
    fallback_model: str | None = None


def is_image_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in IMAGE_EXTENSIONS


def load_online_ocr_config(secrets: Any | None = None) -> OnlineOCRConfig | None:
    provider = os.getenv("OCR_PROVIDER", "")
    qwen_api_url = os.getenv("QWEN_OCR_API_URL", DEFAULT_QWEN_OCR_API_URL)
    qwen_api_key = (
        os.getenv("QWEN_OCR_API_KEY")
        or os.getenv("DASHSCOPE_API_KEY")
        or os.getenv("BAILIAN_API_KEY")
    )
    qwen_model = os.getenv("QWEN_OCR_MODEL", DEFAULT_QWEN_OCR_MODEL)
    zhipu_api_url = os.getenv("ZHIPUAI_API_URL", DEFAULT_ZHIPUAI_API_URL)
    zhipu_api_key = os.getenv("ZHIPUAI_API_KEY") or os.getenv("BIGMODEL_API_KEY")
    zhipu_model = os.getenv("ZHIPUAI_MODEL", DEFAULT_ZHIPUAI_MODEL)

    if secrets is not None:
        provider = _secret_value(secrets, "OCR_PROVIDER", provider) or provider
        qwen_api_url = _secret_value(secrets, "QWEN_OCR_API_URL", qwen_api_url) or DEFAULT_QWEN_OCR_API_URL
        qwen_api_key = (
            _secret_value(secrets, "QWEN_OCR_API_KEY", qwen_api_key)
            or _secret_value(secrets, "DASHSCOPE_API_KEY", qwen_api_key)
            or _secret_value(secrets, "BAILIAN_API_KEY", qwen_api_key)
        )
        qwen_model = _secret_value(secrets, "QWEN_OCR_MODEL", qwen_model) or DEFAULT_QWEN_OCR_MODEL
        zhipu_api_url = _secret_value(secrets, "ZHIPUAI_API_URL", zhipu_api_url) or DEFAULT_ZHIPUAI_API_URL
        zhipu_api_key = (
            _secret_value(secrets, "ZHIPUAI_API_KEY", zhipu_api_key)
            or _secret_value(secrets, "BIGMODEL_API_KEY", zhipu_api_key)
        )
        zhipu_model = _secret_value(secrets, "ZHIPUAI_MODEL", zhipu_model) or DEFAULT_ZHIPUAI_MODEL

    provider_key = provider.lower().strip()
    if qwen_api_key and (not provider_key or provider_key in {"qwen", "qwen-ocr", "dashscope", "bailian"}):
        return OnlineOCRConfig(
            provider="qwen",
            api_url=qwen_api_url,
            api_key=qwen_api_key,
            model=qwen_model,
            fallback_api_url=zhipu_api_url if zhipu_api_key else None,
            fallback_api_key=zhipu_api_key,
            fallback_model=zhipu_model,
        )

    if zhipu_api_key and provider_key in {"zhipu", "zhipuai", "bigmodel"}:
        return OnlineOCRConfig(
            provider="zhipuai",
            api_url=zhipu_api_url,
            api_key=zhipu_api_key,
            model=zhipu_model,
        )

    return None


def _secret_value(secrets: Any, key: str, default: str | None = None) -> str | None:
    try:
        value = secrets.get(key, default)
    except Exception:
        return default
    return value or default


def online_ocr_available(config: OnlineOCRConfig | None) -> bool:
    return bool(config and config.api_url and config.api_key)


def recognize_uploaded_image(uploaded_file: Any, config: OnlineOCRConfig) -> dict[str, Any]:
    if not is_image_file(uploaded_file.name):
        raise ValueError(f"Not an image file: {uploaded_file.name}")

    response = call_online_ocr(
        filename=uploaded_file.name,
        content=bytes(uploaded_file.getbuffer()),
        config=config,
    )
    lines = extract_ocr_lines(response)
    return {
        "source_file": uploaded_file.name,
        "line_count": len(lines),
        "text": "\n".join(item["text"] for item in lines if item.get("text")),
        "lines": lines,
        "raw_response": response,
    }


def call_online_ocr(filename: str, content: bytes, config: OnlineOCRConfig) -> Any:
    if config.provider == "qwen":
        try:
            return call_openai_compatible_vision(
                filename=filename,
                content=content,
                config=config,
                default_model=DEFAULT_QWEN_OCR_MODEL,
                retry_label="Qwen OCR",
            )
        except Exception:
            fallback = load_zhipu_fallback_config(config)
            if not fallback:
                raise
            return call_openai_compatible_vision(
                filename=filename,
                content=content,
                config=fallback,
                default_model=DEFAULT_ZHIPUAI_MODEL,
                retry_label="ZhipuAI vision OCR fallback",
            )

    if config.provider == "zhipuai":
        return call_openai_compatible_vision(
            filename=filename,
            content=content,
            config=config,
            default_model=DEFAULT_ZHIPUAI_MODEL,
            retry_label="ZhipuAI vision OCR",
        )

    raise RuntimeError(f"Unsupported OCR provider: {config.provider}")


def load_zhipu_fallback_config(config: OnlineOCRConfig | None = None) -> OnlineOCRConfig | None:
    api_key = (config.fallback_api_key if config else None) or os.getenv("ZHIPUAI_API_KEY") or os.getenv("BIGMODEL_API_KEY")
    api_url = (config.fallback_api_url if config else None) or os.getenv("ZHIPUAI_API_URL", DEFAULT_ZHIPUAI_API_URL)
    model = (config.fallback_model if config else None) or os.getenv("ZHIPUAI_MODEL", DEFAULT_ZHIPUAI_MODEL)
    if not api_key:
        return None
    return OnlineOCRConfig(provider="zhipuai", api_url=api_url, api_key=api_key, model=model)


def call_openai_compatible_vision(
    *,
    filename: str,
    content: bytes,
    config: OnlineOCRConfig,
    default_model: str,
    retry_label: str,
) -> dict[str, Any]:
    content_type = mimetypes.guess_type(filename)[0] or "image/png"
    image_data = base64.b64encode(content).decode("ascii")
    payload = {
        "model": config.model or default_model,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": CKM3_OCR_PROMPT},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:{content_type};base64,{image_data}"},
                    },
                ],
            }
        ],
        "temperature": 0,
    }
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": f"Bearer {config.api_key}",
    }
    response = _request_json_with_retry(
        config.api_url,
        method="POST",
        headers=headers,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        timeout=180,
        attempts=max(config.retry_attempts, 1),
        retry_label=retry_label,
    )
    text = _extract_chat_completion_text(response)
    return {
        "text": text,
        "provider": config.provider,
        "model": config.model,
        "raw_response": response,
    }


def call_zhipuai_vision(filename: str, content: bytes, config: OnlineOCRConfig) -> dict[str, Any]:
    return call_openai_compatible_vision(
        filename=filename,
        content=content,
        config=config,
        default_model=DEFAULT_ZHIPUAI_MODEL,
        retry_label="ZhipuAI vision OCR",
    )


def recognize_ckm3_material_id(filename: str, content: bytes, config: OnlineOCRConfig) -> str:
    crop_bytes = _crop_ckm3_material_field(content)
    prompt = (
        "只读取图片中“物料”标签右侧第一个白色输入框里的 8 位数字。"
        "不要读取右侧描述文字。只返回严格 JSON：{\"material_id\":\"\"}"
    )
    response = call_openai_compatible_vision_with_prompt(
        filename=f"{Path(filename).stem}_material.png",
        content=crop_bytes,
        config=config,
        default_model=DEFAULT_QWEN_OCR_MODEL if config.provider == "qwen" else DEFAULT_ZHIPUAI_MODEL,
        retry_label=f"{config.provider} CKM3 material OCR",
        prompt=prompt,
    )
    data = _parse_jsonish(_extract_chat_completion_text(response.get("raw_response") or response))
    material_id = str(data.get("material_id") or "").strip()
    return _first_digit_code(material_id)


def recognize_co03_product_id(filename: str, content: bytes, config: OnlineOCRConfig) -> str:
    crop_bytes = _crop_co03_material_field(content)
    prompt = (
        "图片只有 SAP 顶部“材料”这一行。请读取“材料”标签右侧的第一个连续数字编码，"
        "遇到空格或字母就停止。只返回严格 JSON：{\"product_id\":\"\"}"
    )
    response = call_openai_compatible_vision_with_prompt(
        filename=f"{Path(filename).stem}_co03_material.png",
        content=crop_bytes,
        config=config,
        default_model=DEFAULT_QWEN_OCR_MODEL if config.provider == "qwen" else DEFAULT_ZHIPUAI_MODEL,
        retry_label=f"{config.provider} CO03 material OCR",
        prompt=prompt,
    )
    data = _parse_jsonish(_extract_chat_completion_text(response.get("raw_response") or response))
    product_id = str(data.get("product_id") or "").strip()
    return _first_digit_code(product_id)


def call_openai_compatible_vision_with_prompt(
    *,
    filename: str,
    content: bytes,
    config: OnlineOCRConfig,
    default_model: str,
    retry_label: str,
    prompt: str,
) -> dict[str, Any]:
    content_type = mimetypes.guess_type(filename)[0] or "image/png"
    image_data = base64.b64encode(content).decode("ascii")
    payload = {
        "model": config.model or default_model,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:{content_type};base64,{image_data}"},
                    },
                ],
            }
        ],
        "temperature": 0,
    }
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": f"Bearer {config.api_key}",
    }
    response = _request_json_with_retry(
        config.api_url,
        method="POST",
        headers=headers,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        timeout=120,
        attempts=max(config.retry_attempts, 1),
        retry_label=retry_label,
    )
    return {
        "text": _extract_chat_completion_text(response),
        "provider": config.provider,
        "model": config.model,
        "raw_response": response,
    }


def _crop_ckm3_material_field(content: bytes) -> bytes:
    try:
        from PIL import Image
    except ImportError as exc:
        raise RuntimeError("Pillow is required for CKM3 material field OCR.") from exc

    image = Image.open(BytesIO(content)).convert("RGB")
    width, height = image.size
    left = 0
    top = int(height * 0.134)
    right = int(width * 0.271)
    bottom = int(height * 0.169)
    crop = image.crop((left, top, right, bottom))
    crop = crop.resize((crop.width * 4, crop.height * 4))
    output = BytesIO()
    crop.save(output, format="PNG")
    return output.getvalue()


def _crop_co03_material_field(content: bytes) -> bytes:
    try:
        from PIL import Image
    except ImportError as exc:
        raise RuntimeError("Pillow is required for CO03 material field OCR.") from exc

    image = Image.open(BytesIO(content)).convert("RGB")
    width, height = image.size
    left = 0
    top = int(height * 0.213)
    right = int(width * 0.396)
    bottom = int(height * 0.246)
    crop = image.crop((left, top, right, bottom))
    crop = crop.resize((crop.width * 4, crop.height * 4))
    output = BytesIO()
    crop.save(output, format="PNG")
    return output.getvalue()


def _first_digit_code(text: str) -> str:
    match = re.search(r"\d{6,12}", str(text or ""))
    return match.group(0) if match else ""


def _parse_jsonish(text: str) -> dict[str, Any]:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end >= start:
        text = text[start : end + 1]
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return {}


def call_zhipuai_vision(filename: str, content: bytes, config: OnlineOCRConfig) -> dict[str, Any]:
    content_type = mimetypes.guess_type(filename)[0] or "image/png"
    image_data = base64.b64encode(content).decode("ascii")
    prompt = (
        "你是 SAP CKM3 截图表格识别助手。请只从图片中识别 CKM3 明细表，"
        "输出一个 HTML table，不要输出解释。第一行必须是表头。"
        "表头依次为：类别、交易数量、数量单位、初始评估、价格差异、工率差异、实际值、价格、"
        "公司间利润、直接材料、直接人工、间接人工、折旧与摊销、能耗、低值易耗、其他制造费用、成本构成总和。"
        "保留负号、尾随负号和千分位数字；识别不到的数字填 0。"
    )
    payload = {
        "model": config.model or DEFAULT_ZHIPUAI_MODEL,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:{content_type};base64,{image_data}"},
                    },
                ],
            }
        ],
        "temperature": 0,
    }
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": f"Bearer {config.api_key}",
    }
    response = _request_json_with_retry(
        config.api_url or DEFAULT_ZHIPUAI_API_URL,
        method="POST",
        headers=headers,
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        timeout=180,
        attempts=max(config.retry_attempts, 1),
        retry_label="ZhipuAI vision OCR",
    )
    text = _extract_chat_completion_text(response)
    return {
        "text": text,
        "provider": "zhipuai",
        "model": config.model,
        "raw_response": response,
    }


def _extract_chat_completion_text(response: Any) -> str:
    if not isinstance(response, dict):
        return str(response or "")
    choices = response.get("choices") or []
    if choices:
        message = (choices[0] or {}).get("message") or {}
        content = message.get("content")
        if isinstance(content, str):
            return content
        if isinstance(content, list):
            parts = []
            for item in content:
                if isinstance(item, dict):
                    parts.append(str(item.get("text") or item.get("content") or ""))
                else:
                    parts.append(str(item))
            return "\n".join(part for part in parts if part)
    return str(response.get("text") or response.get("content") or "")


def _request_json_with_retry(
    url: str,
    *,
    method: str,
    headers: dict[str, str],
    data: bytes | None = None,
    timeout: int = 45,
    attempts: int = 3,
    retry_label: str = "Online OCR",
) -> Any:
    last_error: Exception | None = None
    for attempt in range(1, attempts + 1):
        try:
            return _request_json(url, method=method, headers=headers, data=data, timeout=timeout)
        except RuntimeError as exc:
            message = str(exc)
            retryable = "HTTP 429" in message or "timed out" in message.lower() or "connection failed" in message.lower()
            if not retryable or attempt >= attempts:
                raise
            last_error = exc
        except (TimeoutError, socket.timeout) as exc:
            if attempt >= attempts:
                raise RuntimeError(f"{retry_label} timed out after {attempts} attempts.") from exc
            last_error = exc

        time.sleep(min(8 * attempt, 24))

    raise RuntimeError(f"{retry_label} failed after {attempts} attempts: {last_error}")


def _request_json(url: str, *, method: str, headers: dict[str, str], data: bytes | None = None, timeout: int = 45) -> Any:
    request = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            body = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Online OCR HTTP {exc.code}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Online OCR connection failed: {exc.reason}") from exc
    return json.loads(body)


def extract_ocr_lines(response: Any) -> list[dict[str, Any]]:
    lines: list[dict[str, Any]] = []

    def visit(node: Any) -> None:
        if isinstance(node, dict):
            text_value = _first_text_value(node)
            if text_value:
                lines.append(
                    {
                        "text": text_value,
                        "confidence": node.get("score") or node.get("confidence") or node.get("prob"),
                    }
                )
                return
            for value in node.values():
                visit(value)
        elif isinstance(node, list):
            for item in node:
                visit(item)
        elif isinstance(node, str) and node.strip():
            lines.append({"text": node.strip(), "confidence": None})

    visit(response)
    return _dedupe_lines(lines)


def _first_text_value(node: dict[str, Any]) -> str | None:
    for key in ("text", "words", "content", "rec_text", "ocr_text"):
        value = node.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
    return None


def _dedupe_lines(lines: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen: set[str] = set()
    result = []
    for line in lines:
        text = line.get("text")
        if not text or text in seen:
            continue
        seen.add(text)
        result.append(line)
    return result
