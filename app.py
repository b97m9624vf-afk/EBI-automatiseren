import base64
import importlib
import json
import os
import re
import sys
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

import docx_generator
import pdf_generator
import scraper

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")


def parse_pasted_schedule(raw_text: str) -> list[dict]:
    """Parse tab-separated or space-separated amortization schedule pasted from Excel/CCS."""
    rows = []
    for line in (raw_text or "").splitlines():
        line = line.strip()
        if not line:
            continue
        if line.lower().startswith("termijn"):
            continue
        parts = [p.strip() for p in line.split("\t") if p.strip() != ""]
        if len(parts) < 2:
            # Try splitting by 2+ spaces or € boundaries
            m0 = re.match(r"^0\s+(€\s*[\d.,]+)$", line)
            if m0:
                rows.append(
                    {
                        "termijn": "0",
                        "termijnbedrag": "",
                        "terugbetaling": "",
                        "rente_en_kosten": "",
                        "saldo": m0.group(1),
                    }
                )
                continue
            tokens = re.findall(r"^\d+|€\s*(?:[\d.,]+|-)", line)
            if len(tokens) >= 5:
                parts = tokens
        if not parts:
            continue
        termijn = parts[0].replace("€", "").strip()
        if termijn == "0" and len(parts) == 2:
            rows.append(
                {
                    "termijn": "0",
                    "termijnbedrag": "",
                    "terugbetaling": "",
                    "rente_en_kosten": "",
                    "saldo": parts[1] if "€" in parts[1] else f"€ {parts[1]}",
                }
            )
        elif len(parts) >= 5:
            def _fmt(v):
                v = v.strip()
                if not v:
                    return ""
                return v if v.startswith("€") else f"€ {v}"

            rows.append(
                {
                    "termijn": termijn,
                    "termijnbedrag": _fmt(parts[1]),
                    "terugbetaling": _fmt(parts[2]),
                    "rente_en_kosten": _fmt(parts[3]),
                    "saldo": _fmt(parts[4]),
                }
            )
    return rows


class EBIRequestHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        sys.stderr.write(f"[EBI Server] {self.address_string()} - {fmt % args}\n")

    def _send_json(self, payload: dict, status: int = 200):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json_body(self) -> dict:
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length) if length > 0 else b"{}"
        return json.loads(raw.decode("utf-8"))

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path in ("/", "/index.html"):
            index_path = os.path.join(TEMPLATES_DIR, "index.html")
            with open(index_path, "rb") as f:
                content = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
            return

        if parsed.path == "/collin_logo.svg":
            logo_path = os.path.join(BASE_DIR, "collin_logo.svg")
            if os.path.exists(logo_path):
                with open(logo_path, "rb") as f:
                    content = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "image/svg+xml")
                self.send_header("Content-Length", str(len(content)))
                self.end_headers()
                self.wfile.write(content)
                return

        if parsed.path == "/api/health":
            self._send_json({"status": "ok"})
            return

        self.send_error(404, "Not Found")

    def do_POST(self):
        parsed = urlparse(self.path)
        try:
            importlib.reload(scraper)
            importlib.reload(pdf_generator)
            importlib.reload(docx_generator)

            if parsed.path == "/api/scrape":
                body = self._read_json_body()
                url = (body.get("url") or "").strip()
                if not url:
                    self._send_json({"error": "Geef een geldige Collin Crowdfund pitch-URL op."}, status=400)
                    return
                html = scraper.fetch_html(url)
                data = scraper.parse_pitch_html(html, url)
                self._send_json({"data": data})
                return

            if parsed.path == "/api/recalculate-schedule":
                body = self._read_json_body()
                pasted = (body.get("pasted_schedule") or "").strip()
                if pasted:
                    rows = parse_pasted_schedule(pasted)
                    if not rows:
                        self._send_json({"error": "Kon geen geldig betaalschema herkennen uit de geplakte tekst."}, status=400)
                        return
                    self._send_json({"betaalschema": rows})
                    return

                params = body.get("loan_params") or {}
                rows = scraper.generate_amortization_schedule(
                    principal=float(params.get("principal", 100000)),
                    annual_rate_pct=float(params.get("annual_rate_pct", 7.5)),
                    months=int(params.get("months", 96)),
                    grace_months=int(params.get("grace_months", 0)),
                    balloon_amount=float(params.get("balloon_amount", 0.0)),
                    succes_fee_pct=float(params.get("succes_fee_pct", 2.0)),
                    admin_fee_pct=float(params.get("admin_fee_pct", 0.07)),
                    loan_type=str(params.get("loan_type", "Lineair")),
                )
                self._send_json({"betaalschema": rows})
                return

            if parsed.path == "/api/generate-pdf":
                body = self._read_json_body()
                data = body.get("data") or {}
                org_b64 = body.get("organigram_image_base64")
                org_bytes = None
                if org_b64:
                    if "," in org_b64:
                        org_b64 = org_b64.split(",", 1)[1]
                    org_bytes = base64.b64decode(org_b64)

                pdf_bytes = pdf_generator.generate_ebi_pdf(data, organigram_image_bytes=org_bytes)
                safe_name = re.sub(r"[^a-zA-Z0-9._-]+", "_", data.get("investeringskans") or "EBI").strip("_")
                filename = f"EBI_{data.get('leningnummer', '')}_{safe_name}.pdf".replace("__", "_")

                self.send_response(200)
                self.send_header("Content-Type", "application/pdf")
                self.send_header("Content-Disposition", f'inline; filename="{filename}"')
                self.send_header("Content-Length", str(len(pdf_bytes)))
                self.end_headers()
                self.wfile.write(pdf_bytes)
                return

            if parsed.path == "/api/generate-docx":
                body = self._read_json_body()
                data = body.get("data") or {}
                docx_bytes = docx_generator.generate_ebi_docx(data)
                safe_name = re.sub(r"[^a-zA-Z0-9._-]+", "_", data.get("investeringskans") or "EBI").strip("_")
                filename = f"EBI_{data.get('leningnummer', '')}_{safe_name}.docx".replace("__", "_")

                self.send_response(200)
                self.send_header(
                    "Content-Type",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
                self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                self.send_header("Content-Length", str(len(docx_bytes)))
                self.end_headers()
                self.wfile.write(docx_bytes)
                return

            self.send_error(404, "Not Found")
        except Exception as exc:
            import traceback

            traceback.print_exc()
            self._send_json({"error": str(exc)}, status=500)


def run_server(port: int = 8765, open_browser: bool = False):
    server = ThreadingHTTPServer(("127.0.0.1", port), EBIRequestHandler)
    url = f"http://127.0.0.1:{port}"
    print(f"Collin Crowdfund EBI Generator draait op: {url}")
    if open_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer gestopt.")
        server.server_close()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8765"))
    open_b = "--open" in sys.argv
    run_server(port=port, open_browser=open_b)
