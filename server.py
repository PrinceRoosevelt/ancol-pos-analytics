"""
server.py — Production WSGI entry point.
  - Lokal  : python server.py          (pakai Waitress, port 5000)
  - Render : otomatis dipanggil via Procfile
Set env var ADMIN_PIN untuk ganti PIN default (1234).
"""
import os
import threading
from waitress import serve
from main import app


def _prewarm_dashboard_cache() -> None:
    try:
        with app.test_client() as client:
            client.get("/", headers={"Accept-Encoding": "gzip"})
        print("[server] Dashboard RAM cache & Gzip payload pre-warmed (<1ms ready)")
    except Exception as exc:
        print(f"[server] Pre-warm skipped: {exc}")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    host = "0.0.0.0"
    admin_pin = os.environ.get("ADMIN_PIN", "1234")
    print(f"[server] Starting on http://{host}:{port}")
    print(f"[server] Admin PIN: {admin_pin}  |  Upload: http://localhost:{port}/upload")
    threading.Thread(target=_prewarm_dashboard_cache, daemon=True).start()
    serve(app, host=host, port=port, threads=4)
