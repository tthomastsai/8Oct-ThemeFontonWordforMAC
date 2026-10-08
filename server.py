#!/usr/bin/env python3
"""本機 HTTPS 伺服器：讓 Word 載入增益集 (https://localhost:3000)。

先執行一次（會要求輸入 Mac 密碼，用來信任憑證）：
    npx office-addin-dev-certs install
然後：
    python3 server.py        # 使用增益集期間，這個視窗要保持開著
"""
import http.server
import ssl
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CERT_DIR = Path.home() / ".office-addin-dev-certs"
CERT, KEY = CERT_DIR / "localhost.crt", CERT_DIR / "localhost.key"
PORT = 3000

if not (CERT.exists() and KEY.exists()):
    sys.exit("找不到開發憑證。請先執行：npx office-addin-dev-certs install")


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **k):
        super().__init__(*a, directory=str(ROOT), **k)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")   # 開發時避免 WebView 快取舊檔案
        super().end_headers()

    def log_message(self, fmt, *args):
        pass


ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
ctx.load_cert_chain(CERT, KEY)
httpd = http.server.ThreadingHTTPServer(("localhost", PORT), Handler)
httpd.socket = ctx.wrap_socket(httpd.socket, server_side=True)
print(f"增益集伺服器執行中：https://localhost:{PORT}/taskpane.html  （Ctrl+C 結束）")
try:
    httpd.serve_forever()
except KeyboardInterrupt:
    pass
