from __future__ import annotations

import json
import mimetypes
import os
import time
import uuid
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"}
DEFAULT_PADDLEOCR_API_URL = "https://paddleocr.aistudio-app.com/api/v2/ocr/jobs"
DEFAULT_PADDLEOCR_MODEL = "PaddleOCR-VL-1.6"


@dataclass
class OnlineOCRConfig:
    api_url: str = DEFAULT_PADDLEOCR_API_URL
    api_key: str | None = None
    auth_header: str = "Authorization"
    auth_scheme: str = "bearer"
    file_field: str = "file"
    model: str = DEFAULT_PADDLEOCR_MODEL
    poll_interval_seconds: int = 5
    max_wait_seconds: int = 180


def is_image_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in IMAGE_EXTENSIONS


def load_online_ocr_config(secrets: Any | None = None) -> OnlineOCRConfig | None:
    api_url = os.getenv("PADDLEOCR_API_URL", DEFAULT_PADDLEOCR_API_URL)
    api_key = os.getenv("PADDLEOCR_API_KEY")
    auth_header = os.getenv("PADDLEOCR_AUTH_HEADER", "Authorization")
    auth_scheme = os.getenv("PADDLEOCR_AUTH_SCHEME", "bearer")
    file_field = os.getenv("PADDLEOCR_FILE_FIELD", "file")
    model = os.getenv("PADDLEOCR_MODEL", DEFAULT_PADDLEOCR_MODEL)

    if secrets is not None:
        api_url = _secret_value(secrets, "PADDLEOCR_API_URL", api_url)
        api_key = _secret_value(secrets, "PADDLEOCR_API_KEY", api_key)
        auth_header = _secret_value(secrets, "PADDLEOCR_AUTH_HEADER", auth_header) or "Authorization"
        auth_scheme = _secret_value(secrets, "PADDLEOCR_AUTH_SCHEME", auth_scheme) or "bearer"
        file_field = _secret_value(secrets, "PADDLEOCR_FILE_FIELD", file_field) or "file"
        model = _secret_value(secrets, "PADDLEOCR_MODEL", model) or DEFAULT_PADDLEOCR_MODEL

    if not api_url or not api_key:
        return None
    return OnlineOCRConfig(
        api_url=api_url,
        api_key=api_key,
        auth_header=auth_header,
        auth_scheme=auth_scheme,
        file_field=file_field,
        model=model,
    )


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
    if config.api_url.rstrip("/").endswith("/jobs"):
        return call_aistudio_paddleocr_job(filename=filename, content=content, config=config)

    boundary = f"----ai-mesp-{uuid.uuid4().hex}"
    body = _build_multipart_body(
        boundary=boundary,
        field_name=config.file_field,
        filename=filename,
        content=content,
    )
    headers = {
        "Content-Type": f"multipart/form-data; boundary={boundary}",
        "Accept": "application/json",
    }
    if config.api_key:
        headers[config.auth_header] = (
            config.api_key if not config.auth_scheme else f"{config.auth_scheme} {config.api_key}"
        )

    request = urllib.request.Request(config.api_url, data=body, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            response_body = response.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Online OCR HTTP {exc.code}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Online OCR connection failed: {exc.reason}") from exc

    try:
        return json.loads(response_body)
    except json.JSONDecodeError:
        return {"text": response_body}


def call_aistudio_paddleocr_job(filename: str, content: bytes, config: OnlineOCRConfig) -> Any:
    optional_payload = {
        "useDocOrientationClassify": False,
        "useDocUnwarping": False,
        "useChartRecognition": False,
    }
    boundary = f"----ai-mesp-{uuid.uuid4().hex}"
    body = _build_multipart_body(
        boundary=boundary,
        field_name=config.file_field,
        filename=filename,
        content=content,
        fields={
            "model": config.model,
            "optionalPayload": json.dumps(optional_payload),
        },
    )
    headers = _auth_headers(config)
    headers.update(
        {
            "Content-Type": f"multipart/form-data; boundary={boundary}",
            "Accept": "application/json",
        }
    )
    job_response = _request_json(config.api_url, method="POST", headers=headers, data=body, timeout=90)
    job_id = ((job_response.get("data") or {}).get("jobId")) if isinstance(job_response, dict) else None
    if not job_id:
        raise RuntimeError(f"Online OCR response missing jobId: {job_response}")

    job_result = _poll_aistudio_job(config, str(job_id))
    json_url = (((job_result.get("data") or {}).get("resultUrl") or {}).get("jsonUrl"))
    if not json_url:
        raise RuntimeError(f"Online OCR completed but jsonUrl is missing: {job_result}")
    return _download_jsonl_result(json_url)


def _poll_aistudio_job(config: OnlineOCRConfig, job_id: str) -> dict[str, Any]:
    deadline = time.time() + config.max_wait_seconds
    job_url = f"{config.api_url.rstrip('/')}/{job_id}"
    headers = _auth_headers(config)
    while time.time() < deadline:
        result = _request_json(job_url, method="GET", headers=headers, timeout=45)
        data = result.get("data") or {}
        state = data.get("state")
        if state == "done":
            return result
        if state == "failed":
            raise RuntimeError(data.get("errorMsg") or f"Online OCR job failed: {result}")
        if state not in {"pending", "running"}:
            raise RuntimeError(f"Unexpected Online OCR job state: {state}")
        time.sleep(config.poll_interval_seconds)
    raise RuntimeError("Online OCR job timed out.")


def _download_jsonl_result(json_url: str) -> dict[str, Any]:
    with urllib.request.urlopen(json_url, timeout=90) as response:
        text = response.read().decode("utf-8", errors="replace")
    pages = []
    markdown_texts = []
    for line in text.strip().splitlines():
        if not line.strip():
            continue
        payload = json.loads(line)
        pages.append(payload)
        result = payload.get("result") or {}
        for layout in result.get("layoutParsingResults") or []:
            markdown = layout.get("markdown") or {}
            md_text = markdown.get("text")
            if md_text:
                markdown_texts.append(md_text)
    return {
        "text": "\n\n".join(markdown_texts),
        "pages": pages,
    }


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


def _auth_headers(config: OnlineOCRConfig) -> dict[str, str]:
    if not config.api_key:
        return {}
    value = config.api_key if not config.auth_scheme else f"{config.auth_scheme} {config.api_key}"
    return {config.auth_header: value}


def _build_multipart_body(
    *,
    boundary: str,
    field_name: str,
    filename: str,
    content: bytes,
    fields: dict[str, str] | None = None,
) -> bytes:
    content_type = mimetypes.guess_type(filename)[0] or "application/octet-stream"
    parts = []
    for name, value in (fields or {}).items():
        parts.append(
            (
                f"--{boundary}\r\n"
                f'Content-Disposition: form-data; name="{name}"\r\n\r\n'
                f"{value}\r\n"
            ).encode("utf-8")
        )
    parts.append(
        (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="{field_name}"; filename="{Path(filename).name}"\r\n'
        f"Content-Type: {content_type}\r\n\r\n"
        ).encode("utf-8")
        + content
        + b"\r\n"
    )
    footer = f"\r\n--{boundary}--\r\n".encode("utf-8")
    return b"".join(parts) + footer


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
