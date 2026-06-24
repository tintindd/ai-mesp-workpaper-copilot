from __future__ import annotations

import json
import mimetypes
import os
import uuid
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"}


@dataclass
class OnlineOCRConfig:
    api_url: str
    api_key: str | None = None
    auth_header: str = "Authorization"
    auth_scheme: str = "Bearer"
    file_field: str = "file"


def is_image_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in IMAGE_EXTENSIONS


def load_online_ocr_config(secrets: Any | None = None) -> OnlineOCRConfig | None:
    api_url = os.getenv("PADDLEOCR_API_URL")
    api_key = os.getenv("PADDLEOCR_API_KEY")
    auth_header = os.getenv("PADDLEOCR_AUTH_HEADER", "Authorization")
    auth_scheme = os.getenv("PADDLEOCR_AUTH_SCHEME", "Bearer")
    file_field = os.getenv("PADDLEOCR_FILE_FIELD", "file")

    if secrets is not None:
        api_url = _secret_value(secrets, "PADDLEOCR_API_URL", api_url)
        api_key = _secret_value(secrets, "PADDLEOCR_API_KEY", api_key)
        auth_header = _secret_value(secrets, "PADDLEOCR_AUTH_HEADER", auth_header) or "Authorization"
        auth_scheme = _secret_value(secrets, "PADDLEOCR_AUTH_SCHEME", auth_scheme) or "Bearer"
        file_field = _secret_value(secrets, "PADDLEOCR_FILE_FIELD", file_field) or "file"

    if not api_url:
        return None
    return OnlineOCRConfig(
        api_url=api_url,
        api_key=api_key,
        auth_header=auth_header,
        auth_scheme=auth_scheme,
        file_field=file_field,
    )


def _secret_value(secrets: Any, key: str, default: str | None = None) -> str | None:
    try:
        value = secrets.get(key, default)
    except Exception:
        return default
    return value or default


def online_ocr_available(config: OnlineOCRConfig | None) -> bool:
    return bool(config and config.api_url)


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


def _build_multipart_body(*, boundary: str, field_name: str, filename: str, content: bytes) -> bytes:
    content_type = mimetypes.guess_type(filename)[0] or "application/octet-stream"
    header = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="{field_name}"; filename="{Path(filename).name}"\r\n'
        f"Content-Type: {content_type}\r\n\r\n"
    ).encode("utf-8")
    footer = f"\r\n--{boundary}--\r\n".encode("utf-8")
    return header + content + footer


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
