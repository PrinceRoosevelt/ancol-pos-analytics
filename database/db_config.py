from __future__ import annotations

import os
import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "data" / "sales_analytics.db"
SCHEMA_PATH = Path(__file__).resolve().parent / "schema.sql"


def get_connection() -> sqlite3.Connection:
    """Mengembalikan koneksi SQLite dengan konfigurasi WAL mode berkecepatan tinggi."""
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH), timeout=30.0)
    conn.row_factory = sqlite3.Row
    # Optimasi performa SQLite untuk analytics
    conn.execute("PRAGMA journal_mode = WAL;")
    conn.execute("PRAGMA synchronous = NORMAL;")
    conn.execute("PRAGMA cache_size = -64000;")  # 64MB Cache
    conn.execute("PRAGMA temp_store = MEMORY;")
    return conn


def init_db() -> None:
    """Inisialisasi tabel dan indeks sesuai file schema.sql."""
    if not SCHEMA_PATH.exists():
        return
    with get_connection() as conn:
        with open(SCHEMA_PATH, "r", encoding="utf-8") as f:
            conn.executescript(f.read())
        conn.commit()


if __name__ == "__main__":
    init_db()
    print("Database SQLite berhasil diinisialisasi pada:", DB_PATH)

