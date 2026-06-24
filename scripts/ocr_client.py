from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp"}


def is_image_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in IMAGE_EXTENSIONS


def paddleocr_available() -> bool:
    try:
        import paddleocr  # noqa: F401
    except Exception:
        return False
    return True


def recognize_uploaded_image(uploaded_file: Any, *, lang: str = "ch") -> dict[str, Any]:
    if not is_image_file(uploaded_file.name):
        raise ValueError(f"Not an image file: {uploaded_file.name}")

    suffix = Path(uploaded_file.name).suffix.lower()
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp:
        temp.write(uploaded_file.getbuffer())
        temp_path = Path(temp.name)

    try:
        return recognize_image_path(temp_path, source_name=uploaded_file.name, lang=lang)
    finally:
        temp_path.unlink(missing_ok=True)


def recognize_image_path(image_path: Path, *, source_name: str | None = None, lang: str = "ch") -> dict[str, Any]:
    try:
        from paddleocr import PaddleOCR
    except Exception as exc:
        raise RuntimeError("PaddleOCR is not installed in this environment.") from exc

    ocr = PaddleOCR(use_angle_cls=True, lang=lang, show_log=False)
    raw_result = ocr.ocr(str(image_path), cls=True)
    lines = _flatten_paddleocr_result(raw_result)
    return {
        "source_file": source_name or image_path.name,
        "line_count": len(lines),
        "text": "\n".join(item["text"] for item in lines if item.get("text")),
        "lines": lines,
    }


def _flatten_paddleocr_result(raw_result: Any) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    def visit(node: Any) -> None:
        if not node:
            return
        if _looks_like_ocr_line(node):
            text, score = node[1]
            rows.append({"text": str(text), "confidence": float(score)})
            return
        if isinstance(node, list):
            for child in node:
                visit(child)

    visit(raw_result)
    return rows


def _looks_like_ocr_line(node: Any) -> bool:
    if not isinstance(node, list) or len(node) < 2:
        return False
    payload = node[1]
    return (
        isinstance(payload, tuple)
        and len(payload) >= 2
        and isinstance(payload[0], str)
        and isinstance(payload[1], (int, float))
    )
