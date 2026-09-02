"""
server.py — Production WSGI entry point.
  - Lokal  : python server.py          (pakai Waitress, port 5000)
  - Render : otomatis dipanggil via Procfile
Set env var ADMIN_PIN untuk ganti PIN default (1234).
"""
import os
from waitress import serve
from main import app

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    host = "0.0.0.0"
    admin_pin = os.environ.get("ADMIN_PIN", "1234")
    print(f"[server] Starting on http://{host}:{port}")
    print(f"[server] Admin PIN: {admin_pin}  |  Upload: http://localhost:{port}/upload")
    serve(app, host=host, port=port, threads=4)
