#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
СТУ-Нормоконтроль — необязательный локальный сервер.

Основной сайт (index.html) работает сам по себе, в браузере, без сервера.
Этот сервер нужен только для двух вещей:

  1) отдать сайт по адресу http://127.0.0.1:8777/ ;
  2) попробовать преобразовать старый формат .doc в .docx, чтобы сайт мог
     проверить и его.

Преобразование выполняется внешним офисным пакетом (LibreOffice «soffice»),
если он установлен. Только на стандартной библиотеке Python 3.

Запуск:
    python3 server/server.py            # порт 8777
    python3 server/server.py 9000       # свой порт

Затем откройте http://127.0.0.1:8777/
"""

import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
from functools import partial
from http.server import HTTPServer, SimpleHTTPRequestHandler

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_PORT = 8777


def find_soffice():
    for name in ("soffice", "libreoffice", "soffice.bin"):
        path = shutil.which(name)
        if path:
            return path
    for path in (
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/opt/libreoffice/program/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if os.path.exists(path):
            return path
    return None


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, directory=None, **kwargs):
        super().__init__(*args, directory=directory, **kwargs)

    # ---------------------------------------------------------------- утилиты
    def _send_json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _read_body(self):
        length = int(self.headers.get("Content-Length", "0"))
        return self.rfile.read(length) if length else b""

    # ------------------------------------------------------------------ API
    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        if self.path.rstrip("/") == "/api/health":
            soffice = find_soffice()
            return self._send_json(200, {
                "ok": True,
                "soffice": soffice,
                "canConvertDoc": bool(soffice),
            })
        return super().do_GET()

    def do_POST(self):
        if self.path.rstrip("/") == "/api/convert-doc":
            return self.convert_doc()
        self._send_json(404, {"error": "Неизвестный адрес"})

    # ------------------------------------------------------------------ .doc
    def convert_doc(self):
        data = self._read_body()
        if not data:
            return self._send_json(400, {"error": "Пустой файл"})
        soffice = find_soffice()
        if not soffice:
            return self._send_json(501, {
                "error": "LibreOffice не найден",
                "hint": ("Для проверки .doc установите LibreOffice "
                         "(пакет libreoffice), затем перезапустите сервер. "
                         "Либо откройте документ в Word и сохраните как .docx."),
            })
        tmp = tempfile.mkdtemp(prefix="stu_doc_")
        try:
            src = os.path.join(tmp, "input.doc")
            with open(src, "wb") as fh:
                fh.write(data)
            cmd = [soffice, "--headless", "--norestore",
                   "--convert-to", "docx", "--outdir", tmp, src]
            proc = subprocess.run(cmd, capture_output=True, timeout=120)
            out = os.path.join(tmp, "input.docx")
            if not os.path.exists(out):
                err = (proc.stderr or b"").decode("utf-8", "replace").strip()
                return self._send_json(500, {
                    "error": "Преобразование не удалось",
                    "hint": err[-400:] or "Не удалось получить .docx из .doc.",
                })
            with open(out, "rb") as fh:
                result = fh.read()
            self.send_response(200)
            self.send_header("Content-Type",
                             "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Length", str(len(result)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(result)
        except Exception as exc:  # noqa: BLE001
            self._send_json(500, {"error": str(exc)})
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def log_message(self, fmt, *args):
        sys.stderr.write("[server] " + (fmt % args) + "\n")


def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_PORT
    handler = partial(Handler, directory=ROOT)
    httpd = HTTPServer(("127.0.0.1", port), handler)
    soffice = find_soffice()
    print("СТУ-Нормоконтроль — локальный сервер")
    print("Корень:", ROOT)
    print("Адрес:  http://127.0.0.1:%d/" % port)
    print("LibreOffice для .doc:", soffice or "не найден (проверка .doc недоступна)")
    print("Ctrl+C — остановить.")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nОстановлено.")
        httpd.server_close()


if __name__ == "__main__":
    main()
