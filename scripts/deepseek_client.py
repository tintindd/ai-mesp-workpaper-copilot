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
                    "You normalize SAP audit evidence filenames for AI-MESP. "
                    "Return strict JSON only, no markdown. Extract sample_no, order_id, "
                    "material_id, report_type, evidence_kind, extension, confidence, and "
                    "standard_filename. report_type must be one of CO03, KSBT, 3611, CKM3. "
                    "evidence_kind must be one of 表格, 截图, 清单, 未知. "
                    "sample_no must be a single number when possible, not a path. "
                    "Use this standard filename pattern: "
                    "样本{sample_no}/{sample_no}.订单编号{order_id}-{report_type}-{evidence_kind}.{extension}. "
                    "For CKM3 with material_id, use: "
                    "样本{sample_no}/{sample_no}.订单编号{order_id}-物料ID-{material_id}-CKM3-{evidence_kind}.{extension}. "
                    "If a field is unknown, use null and still provide the best standard_filename."
                ),
            },
            {"role": "user", "content": filename},
        ],
        config,
        temperature=0,
        max_tokens=500,
    )


def parse_deepseek_json(content: str) -> dict[str, Any]:
    text = content.strip()
    if text.startswith("```"):
        lines = text.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        text = "\n".join(lines).strip()
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end >= start:
        text = text[start : end + 1]
    data = json.loads(text)
    if not isinstance(data, dict):
        raise ValueError("DeepSeek JSON response is not an object.")
    return data


def normalize_filename_with_deepseek(filename: str, config: DeepSeekConfig) -> dict[str, Any]:
    content = classify_filename_with_deepseek(filename, config)
    data = parse_deepseek_json(content)
    extension = data.get("extension")
    if not extension and "." in filename:
        extension = filename.rsplit(".", 1)[-1].lower()
    evidence_kind = data.get("evidence_kind") or _infer_evidence_kind(filename, extension)
    if evidence_kind == "未知":
        evidence_kind = _infer_evidence_kind(filename, extension)
    sample_no = data.get("sample_no")
    order_id = data.get("order_id")
    material_id = data.get("material_id")
    report_type = data.get("report_type")
    standard_filename = _build_standard_filename(
        sample_no=sample_no,
        order_id=order_id,
        material_id=material_id,
        report_type=report_type,
        evidence_kind=evidence_kind,
        extension=extension,
    )
    return {
        "source_file": filename,
        "sample_no": sample_no,
        "order_id": order_id,
        "material_id": material_id,
        "report_type": report_type,
        "evidence_kind": evidence_kind,
        "extension": extension,
        "standard_filename": standard_filename or data.get("standard_filename") or data.get("suggested_filename"),
        "confidence": data.get("confidence"),
        "raw_response": data,
    }


def _infer_evidence_kind(filename: str, extension: str | None) -> str:
    lowered = filename.lower()
    if "截图" in filename or extension in {"png", "jpg", "jpeg", "webp", "gif"}:
        return "截图"
    if "清单" in filename:
        return "清单"
    if "表格" in filename or extension in {"xlsx", "xlsm", "xls", "csv"}:
        return "表格"
    return "未知"


def _build_standard_filename(
    *,
    sample_no: Any,
    order_id: Any,
    material_id: Any,
    report_type: Any,
    evidence_kind: str,
    extension: str | None,
) -> str | None:
    if not sample_no or not order_id or not report_type or not extension:
        return None
    sample_text = str(sample_no)
    report_text = str(report_type).upper()
    extension_text = str(extension).lstrip(".")
    if report_text == "CKM3" and material_id:
        return (
            f"样本{sample_text}/{sample_text}.订单编号{order_id}-"
            f"物料ID-{material_id}-CKM3-{evidence_kind}.{extension_text}"
        )
    return f"样本{sample_text}/{sample_text}.订单编号{order_id}-{report_text}-{evidence_kind}.{extension_text}"


def test_deepseek_connection(config: DeepSeekConfig) -> str:
    return deepseek_chat(
        [
            {"role": "system", "content": "Return exactly: ok"},
            {"role": "user", "content": "connection test"},
        ],
        config,
        temperature=0,
        max_tokens=20,
    )


def clean_ocr_text_with_deepseek(filename: str, ocr_text: str, config: DeepSeekConfig) -> dict[str, Any]:
    content = deepseek_chat(
        [
            {
                "role": "system",
                "content": (
                    "You clean PaddleOCR text extracted from SAP audit evidence screenshots. "
                    "Return strict JSON only, no markdown. Identify sample_no, order_id, material_id, "
                    "report_type, evidence_kind, key_fields, standard_filename, confidence, and notes. "
                    "report_type must be one of CO03, KSBT, 3611, CKM3, unknown. "
                    "evidence_kind should usually be 截图. Use null for unknown values. "
                    "Use the AI-MESP filename convention when possible."
                ),
            },
            {
                "role": "user",
                "content": (
                    f"filename: {filename}\n\n"
                    "PaddleOCR text:\n"
                    f"{ocr_text[:8000]}"
                ),
            },
        ],
        config,
        temperature=0,
        max_tokens=1200,
    )
    data = parse_deepseek_json(content)
    extension = data.get("extension") or (filename.rsplit(".", 1)[-1].lower() if "." in filename else None)
    evidence_kind = data.get("evidence_kind") or "截图"
    sample_no = data.get("sample_no")
    order_id = data.get("order_id")
    material_id = data.get("material_id")
    report_type = data.get("report_type")
    standard_filename = _build_standard_filename(
        sample_no=sample_no,
        order_id=order_id,
        material_id=material_id,
        report_type=report_type,
        evidence_kind=evidence_kind,
        extension=extension,
    )
    return {
        "source_file": filename,
        "sample_no": sample_no,
        "order_id": order_id,
        "material_id": material_id,
        "report_type": report_type,
        "evidence_kind": evidence_kind,
        "standard_filename": standard_filename or data.get("standard_filename"),
        "confidence": data.get("confidence"),
        "key_fields": data.get("key_fields") or {},
        "notes": data.get("notes"),
        "raw_response": data,
    }

