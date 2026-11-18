"""예산 요구안 작성 웹앱 — 로컬 서버.

- 127.0.0.1:8789 에만 바인딩. 외부 통신·AI 호출 없음.
- 서버는 상태를 저장하지 않는다. 초안은 브라우저(자동 저장)와 사용자가 내려받는 .json 파일에만 있다.
- 다른 사이트에서 오는 요청(Origin/Host 불일치)은 거부한다.

API
  POST /api/parse   본문: .xlsx 바이트, 쿼리: header_row, map_비목, map_세부사업, map_배정액, map_집행액 (선택)
  POST /api/check   본문: 초안 JSON → 검사 결과 + 요약
  POST /api/export  본문: 초안 JSON → .xlsx (검사 실패 시 400)

실행: python budget_app/server.py
"""
from __future__ import annotations

import json
import sys
import urllib.parse
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(APP_DIR))
import budget_core as core  # noqa: E402

HOST, PORT = "127.0.0.1", 8789
STATIC = {"/": ("index.html", "text/html; charset=utf-8"),
          "/index.html": ("index.html", "text/html; charset=utf-8"),
          "/app.js": ("app.js", "text/javascript; charset=utf-8"),
          "/style.css": ("style.css", "text/css; charset=utf-8"),
          # 연습용 가상 샘플 (처음 쓰는 사람용)
          "/sample/budget.xlsx": ("sample/sample_budget.xlsx",
                                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
          "/sample/memo.md": ("sample/sample_memo.md", "text/plain; charset=utf-8")}
MAX_XLSX = 10 * 1024 * 1024
MAX_JSON = 5 * 1024 * 1024
ALLOWED_HOSTS = {f"127.0.0.1:{PORT}", f"localhost:{PORT}"}
ALLOWED_ORIGINS = {f"http://{h}" for h in ALLOWED_HOSTS}


class Handler(BaseHTTPRequestHandler):
    server_version = "BudgetApp"
    sys_version = ""

    # ── 공통 ──
    def log_message(self, fmt, *args):  # 경로·상태만 기록 (본문 기록 안 함)
        sys.stderr.write(f"[{datetime.now():%H:%M:%S}] {self.command} {self.path.split('?')[0]} {args[1] if len(args) > 1 else ''}\n")

    def _headers(self, code, ctype, length, extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(length))
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Security-Policy",
                         "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; "
                         "connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'none'")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()

    def _json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self._headers(code, "application/json; charset=utf-8", len(body))
        self.wfile.write(body)

    def _allowed(self) -> bool:
        if self.headers.get("Host", "") not in ALLOWED_HOSTS:
            return False
        origin = self.headers.get("Origin")
        return origin is None or origin in ALLOWED_ORIGINS

    def _body(self, limit):
        try:
            n = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            n = -1
        if n < 0 or n > limit:
            return None
        return self.rfile.read(n)

    # ── GET: 정적 파일만 ──
    def do_GET(self):
        if not self._allowed():
            return self._json(403, {"error": "허용되지 않은 요청입니다."})
        path = urllib.parse.urlsplit(self.path).path
        if path not in STATIC:
            return self._json(404, {"error": "없는 경로입니다."})
        name, ctype = STATIC[path]
        try:
            body = (APP_DIR / "static" / name).read_bytes()
        except OSError:
            return self._json(404, {"error": f"앱 파일이 없습니다: static/{name} — 배포 폴더를 다시 복사해 주세요."})
        self._headers(200, ctype, len(body))
        self.wfile.write(body)

    # ── POST: API ──
    def do_POST(self):
        if not self._allowed():
            return self._json(403, {"error": "허용되지 않은 요청입니다."})
        url = urllib.parse.urlsplit(self.path)
        try:
            if url.path == "/api/parse":
                return self._parse(urllib.parse.parse_qs(url.query))
            if url.path in ("/api/check", "/api/export"):
                raw = self._body(MAX_JSON)
                if raw is None:
                    return self._json(413, {"error": "초안 데이터가 너무 큽니다."})
                try:
                    data = json.loads(raw.decode("utf-8"))
                    assert isinstance(data, dict) and isinstance(data.get("rows"), list)
                except Exception:
                    return self._json(400, {"error": "초안 데이터 형식이 올바르지 않습니다."})
                return self._check(data) if url.path == "/api/check" else self._export(data)
            return self._json(404, {"error": "없는 경로입니다."})
        except Exception as e:  # 예상 못 한 오류: 내용은 서버 창에만
            sys.stderr.write(f"오류: {type(e).__name__}: {e}\n")
            return self._json(500, {"error": "서버 처리 중 오류가 났습니다. 서버 창의 메시지를 확인해 주세요."})

    def _parse(self, q):
        blob = self._body(MAX_XLSX)
        if blob is None:
            return self._json(413, {"error": "엑셀 파일은 10MB 이하만 올릴 수 있습니다."})
        header_row = None
        if q.get("header_row", [""])[0].isdigit():
            header_row = int(q["header_row"][0])
        mapping = {}
        for k in core.REQUIRED:
            v = q.get(f"map_{k}", [""])[0]
            if v.isdigit() and int(v) > 0:
                mapping[k] = int(v)
        res = core.parse_excel(blob, header_row, mapping)
        res["original"] = core.original_from_parse(res)
        return self._json(200, res)

    def _check(self, data):
        return self._json(200, core.check(data))

    def _export(self, data):
        result = core.check(data)
        if not result["ok"]:
            return self._json(400, {"error": "검사를 통과하지 못해 내보낼 수 없습니다.", **result})
        body = core.workbook_bytes(data, result)
        name = f"{core.year_of(data)}_예산요구안_초안_{datetime.now():%Y%m%d}.xlsx"
        self._headers(200, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", len(body), {
            "Content-Disposition": "attachment; filename*=UTF-8''" + urllib.parse.quote(name),
            "X-Check-Total": urllib.parse.quote(json.dumps(
                {k: result["summary"][k] for k in ("배정", "요구", "증감")}, ensure_ascii=False)),
        })
        self.wfile.write(body)


def main():
    srv = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"예산 요구안 작성 웹앱: http://{HOST}:{PORT}  (이 창을 닫으면 종료됩니다)")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
