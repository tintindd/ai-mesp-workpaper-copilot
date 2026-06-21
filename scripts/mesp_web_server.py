from __future__ import annotations

import json
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from mesp_automation_engine import analyze_folder


WORKSPACE = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = WORKSPACE / "training_files" / "训练文件"
DEFAULT_OUTPUT = WORKSPACE / "mesp_automation_result.json"


class MespHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(WORKSPACE), **kwargs)

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
        if parsed.path == "/":
            self.path = "/mesp_copilot_frontend.html"
            return super().do_GET()

        if parsed.path == "/api/analyze":
            query = parse_qs(parsed.query)
            period = query.get("period", [""])[0]
            program = query.get("program", [""])[0]
            input_dir = Path(query.get("input", [str(DEFAULT_INPUT)])[0])

            try:
                if not input_dir.exists():
                    self.send_json({"error": f"Input folder not found: {input_dir}"}, status=404)
                    return

                result = analyze_folder(input_dir, period=period, program=program)
                DEFAULT_OUTPUT.write_text(
                    json.dumps(result, ensure_ascii=False, indent=2, default=str),
                    encoding="utf-8",
                )
                self.send_json(result)
            except Exception as exc:  # Keep browser feedback useful for local demos.
                self.send_json({"error": str(exc)}, status=500)
            return

        if parsed.path == "/api/status":
            self.send_json(
                {
                    "workspace": str(WORKSPACE),
                    "default_input": str(DEFAULT_INPUT),
                    "default_input_exists": DEFAULT_INPUT.exists(),
                    "default_output": str(DEFAULT_OUTPUT),
                }
            )
            return

        return super().do_GET()


def main() -> None:
    server = ThreadingHTTPServer(("127.0.0.1", 8765), MespHandler)
    print("AI-MESP Workpaper Copilot local server")
    print("URL: http://127.0.0.1:8765/")
    print(f"Workspace: {WORKSPACE}")
    print(f"Default input: {DEFAULT_INPUT}")
    server.serve_forever()


if __name__ == "__main__":
    main()
