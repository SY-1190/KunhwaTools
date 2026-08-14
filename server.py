from __future__ import annotations

import argparse
import base64
import json
import threading
from email import policy
from email.parser import BytesParser
from functools import partial
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, urlparse

import pdfplumber

from pdf_excel_converter import convert_pdf_bytes


ROOT = Path(__file__).resolve().parent
MAX_UPLOAD_BYTES = 60 * 1024 * 1024
CONVERSION_LIMIT = threading.BoundedSemaphore(2)


class KunhwaToolsHandler(SimpleHTTPRequestHandler):
    server_version = "KunhwaTools/0.1.22"

    def _send_json(self, status: HTTPStatus, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self) -> None:
        self.send_response(HTTPStatus.NO_CONTENT)
        self.send_header("Allow", "GET, POST, OPTIONS")
        self.end_headers()

    def do_GET(self) -> None:
        path = urlparse(self.path).path
        if path == "/api/pdf-to-excel/health":
            self._send_json(
                HTTPStatus.OK,
                {
                    "ok": True,
                    "service": "KunhwaTools PDF to Excel",
                    "engine": f"pdfplumber {pdfplumber.__version__}",
                    "mode": "local",
                },
            )
            return
        super().do_GET()

    def do_POST(self) -> None:
        if urlparse(self.path).path != "/api/pdf-to-excel/convert":
            self._send_json(HTTPStatus.NOT_FOUND, {"error": "요청한 변환 API를 찾을 수 없습니다."})
            return
        try:
            file_name, pdf_bytes = self._read_pdf_upload()
            with CONVERSION_LIMIT:
                workbook_bytes, report = convert_pdf_bytes(pdf_bytes, Path(file_name).stem)
            output_name = f"{Path(file_name).stem}_엑셀변환.xlsx"
            report_header = base64.urlsafe_b64encode(
                json.dumps(report, ensure_ascii=False).encode("utf-8")
            ).decode("ascii")
            self.send_response(HTTPStatus.OK)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            self.send_header("Content-Length", str(len(workbook_bytes)))
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(output_name)}")
            self.send_header("X-Kunhwa-Conversion-Report", report_header)
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(workbook_bytes)
        except ValueError as error:
            self._send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
        except Exception as error:
            self.log_error("PDF conversion failed: %s", error)
            self._send_json(
                HTTPStatus.INTERNAL_SERVER_ERROR,
                {"error": "PDF 정밀 변환 중 오류가 발생했습니다.", "detail": str(error)},
            )

    def _read_pdf_upload(self) -> tuple[str, bytes]:
        content_length = int(self.headers.get("Content-Length") or 0)
        if content_length <= 0:
            raise ValueError("업로드된 PDF 파일이 없습니다.")
        if content_length > MAX_UPLOAD_BYTES:
            raise ValueError("PDF 파일은 60MB 이하만 변환할 수 있습니다.")
        content_type = self.headers.get("Content-Type") or ""
        if "multipart/form-data" not in content_type.lower():
            raise ValueError("PDF 파일 업로드 형식이 올바르지 않습니다.")
        body = self.rfile.read(content_length)
        message = BytesParser(policy=policy.default).parsebytes(
            f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("ascii") + body
        )
        for part in message.iter_parts():
            if part.get_content_disposition() != "form-data" or part.get_param("name", header="content-disposition") != "file":
                continue
            file_name = part.get_filename() or "document.pdf"
            payload = part.get_payload(decode=True) or b""
            if not file_name.lower().endswith(".pdf") or not payload.startswith(b"%PDF-"):
                raise ValueError("올바른 PDF 파일을 선택해 주세요.")
            return Path(file_name).name, payload
        raise ValueError("업로드된 PDF 파일을 찾을 수 없습니다.")


def main() -> None:
    parser = argparse.ArgumentParser(description="KunhwaTools local web server")
    parser.add_argument("--bind", default="127.0.0.1")
    parser.add_argument("--port", default=8878, type=int)
    args = parser.parse_args()
    handler = partial(KunhwaToolsHandler, directory=str(ROOT))
    server = ThreadingHTTPServer((args.bind, args.port), handler)
    print(f"KunhwaTools: http://{args.bind}:{args.port}/")
    print(f"PDF engine: pdfplumber {pdfplumber.__version__}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
