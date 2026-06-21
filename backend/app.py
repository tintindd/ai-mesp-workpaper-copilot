from __future__ import annotations

import json
import mimetypes
import os
import re
import shutil
import sys
import tempfile
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse


PROJECT_ROOT = Path(__file__).resolve().parents[1]
FRONTEND_ROOT = PROJECT_ROOT / "frontend"
SCRIPTS_ROOT = PROJECT_ROOT / "scripts"

if str(SCRIPTS_ROOT) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_ROOT))

from mesp_automation_engine import analyze_folder  # noqa: E402


MAX_UPLOAD_BYTES = 250 * 1024 * 1024


def safe_relative_path(value: str) -> Path:
    value = unquote(value or "").replace("\\", "/").strip("/")
    parts = []
    for part in value.split("/"):
        if not part or part in {".", ".."}:
            continue
        cleaned = re.sub(r'[<>:"|?*\x00-\x1f]', "_", part)
        parts.append(cleaned[:160])
    return Path(*parts) if parts else Path("uploaded_file")


def parse_content_disposition(value: str) -> dict[str, str]:
    result = {}
    for piece in value.split(";"):
        piece = piece.strip()
        if "=" not in piece:
            result.setdefault("type", piece.lower())
            continue
        key, raw = piece.split("=", 1)
        result[key.strip().lower()] = raw.strip().strip('"')
    return result


def parse_multipart(content_type: str, body: bytes) -> tuple[list[dict], dict[str, list[str]]]:
    match = re.search(r"boundary=(?P<boundary>[^;]+)", content_type)
    if not match:
        raise ValueError("Missing multipart boundary")

    boundary = match.group("boundary").strip().strip('"').encode("utf-8")
    delimiter = b"--" + boundary
    files = []
    fields: dict[str, list[str]] = {}

    for raw_part in body.split(delimiter):
        raw_part = raw_part.strip()
        if not raw_part or raw_part == b"--":
            continue
        if raw_part.endswith(b"--"):
            raw_part = raw_part[:-2].strip()
        header_blob, separator, content = raw_part.partition(b"\r\n\r\n")
        if not separator:
            continue
        if content.endswith(b"\r\n"):
            content = content[:-2]

        headers = {}
        for line in header_blob.decode("utf-8", errors="replace").split("\r\n"):
            if ":" in line:
                key, value = line.split(":", 1)
                headers[key.lower().strip()] = value.strip()

        disposition = parse_content_disposition(headers.get("content-disposition", ""))
        name = disposition.get("name", "")
        filename = disposition.get("filename")
        if filename:
            files.append({"field": name, "filename": filename, "content": content})
        elif name:
            fields.setdefault(name, []).append(content.decode("utf-8", errors="replace"))

    return files, fields


def save_uploads(files: list[dict], fields: dict[str, list[str]], target: Path) -> int:
    relative_paths = fields.get("relative_paths", [])
    saved = 0
    for index, item in enumerate(files):
        relative = relative_paths[index] if index < len(relative_paths) else item["filename"]
        destination = target / safe_relative_path(relative)
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(item["content"])
        saved += 1
    return saved


class AppHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(FRONTEND_ROOT), **kwargs)

    def end_headers(self) -> None:
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def send_json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False, indent=2, default=str).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path in {"", "/"}:
            self.path = "/index.html"
        if parsed.path == "/api/health":
            self.send_json({"ok": True, "mode": "upload-web-app"})
            return
        return super().do_GET()

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path != "/api/analyze":
            self.send_error(404, "Not found")
            return

        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length <= 0:
            self.send_json({"error": "No upload body received"}, status=400)
            return
        if content_length > MAX_UPLOAD_BYTES:
            self.send_json({"error": "Upload is too large"}, status=413)
            return

        content_type = self.headers.get("Content-Type", "")
        query = parse_qs(parsed.query)
        period = query.get("period", [""])[0]
        program = query.get("program", [""])[0]
        upload_dir = Path(tempfile.mkdtemp(prefix="mesp_upload_"))

        try:
            body = self.rfile.read(content_length)
            files, fields = parse_multipart(content_type, body)
            saved_count = save_uploads(files, fields, upload_dir)
            if saved_count == 0:
                self.send_json({"error": "No files found in upload"}, status=400)
                return
            result = analyze_folder(upload_dir, period=period, program=program)
            result["upload_summary"] = {"saved_file_count": saved_count}
            self.send_json(result)
        except Exception as exc:
            self.send_json({"error": str(exc)}, status=500)
        finally:
            shutil.rmtree(upload_dir, ignore_errors=True)


def main() -> None:
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "8766"))
    server = ThreadingHTTPServer((host, port), AppHandler)
    print("AI-MESP Workpaper Copilot upload web app")
    print(f"Listening on: http://{host}:{port}/")
    print(f"Project root: {PROJECT_ROOT}")
    server.serve_forever()


if __name__ == "__main__":
    main()
