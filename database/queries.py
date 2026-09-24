from __future__ import annotations

import sqlite3
from typing import Any

from database.db_config import get_connection


def fetch_all_sales(sync_if_needed: bool = True) -> list[dict[str, Any]]:
    """Mengambil seluruh data baris sales dari database SQLite."""
    conn = get_connection()
    try:
        cursor = conn.execute(
            """
            SELECT 
                year, date, month, hour, invoice, outlet, area,
                product, qty, item_net_sales, net_sales,
                transaction_total, invoice_discount, file_source
            FROM sales_items
            ORDER BY date ASC, hour ASC
            """
        )
        return [dict(row) for row in cursor.fetchall()]
    finally:
        conn.close()


def fetch_all_targets() -> list[dict[str, Any]]:
    """Mengambil master target harian dari database SQLite."""
    conn = get_connection()
    try:
        cursor = conn.execute(
            """
            SELECT date, month, outlet, area, target
            FROM budget_daily
            ORDER BY date ASC
            """
        )
        return [dict(row) for row in cursor.fetchall()]
    finally:
        conn.close()


def fetch_all_visitors() -> list[dict[str, Any]]:
    """Mengambil master pengunjung harian dari database SQLite."""
    conn = get_connection()
    try:
        cursor = conn.execute(
            """
            SELECT date, month, unit, area, visitors,
                   COALESCE(visitors_individu, 0) AS visitors_individu,
                   COALESCE(visitors_individu_bayar, 0) AS visitors_individu_bayar,
                   COALESCE(visitors_tidak_bayar, 0) AS visitors_tidak_bayar,
                   COALESCE(visitors_rombongan_langsung, 0) AS visitors_rombongan_langsung,
                   COALESCE(visitors_rombongan_agen, 0) AS visitors_rombongan_agen,
                   COALESCE(visitors_rombongan_total, 0) AS visitors_rombongan_total
            FROM visitor_actual
            ORDER BY date ASC
            """
        )
        return [dict(row) for row in cursor.fetchall()]
    finally:
        conn.close()


def save_visitor_db(
    entry_date: str,
    unit: str,
    count: int,
    area: str = "",
    individu: int = 0,
    rombongan_langsung: int = 0,
    rombongan_agen: int = 0,
    individu_bayar: int = 0,
    tidak_bayar: int = 0,
) -> None:
    """Simpan/Update pengunjung langsung ke tabel visitor_actual."""
    conn = get_connection()
    try:
        month_str = entry_date[:7]
        if individu == 0 and (individu_bayar > 0 or tidak_bayar > 0):
            individu = individu_bayar + tidak_bayar
        romb_total = rombongan_langsung + rombongan_agen
        if count == 0 and (individu > 0 or romb_total > 0):
            count = individu + romb_total

        conn.execute(
            """
            INSERT INTO visitor_actual (
                date, month, unit, area, visitors,
                visitors_individu, visitors_individu_bayar, visitors_tidak_bayar,
                visitors_rombongan_langsung, visitors_rombongan_agen, visitors_rombongan_total
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(date, unit) DO UPDATE SET
                visitors = excluded.visitors,
                area = CASE WHEN excluded.area != '' THEN excluded.area ELSE visitor_actual.area END,
                visitors_individu = excluded.visitors_individu,
                visitors_individu_bayar = excluded.visitors_individu_bayar,
                visitors_tidak_bayar = excluded.visitors_tidak_bayar,
                visitors_rombongan_langsung = excluded.visitors_rombongan_langsung,
                visitors_rombongan_agen = excluded.visitors_rombongan_agen,
                visitors_rombongan_total = excluded.visitors_rombongan_total
            """,
            (entry_date, month_str, unit, area, count, individu, individu_bayar, tidak_bayar, rombongan_langsung, rombongan_agen, romb_total),
        )
        conn.commit()
    finally:
        conn.close()


def save_visitors_bulk(records: list[dict[str, Any]]) -> int:
    """Simpan/Update banyak data pengunjung sekaligus dalam satu transaksi terisolasi."""
    if not records:
        return 0
    conn = get_connection()
    try:
        conn.execute("BEGIN TRANSACTION;")
        prepared = []
        for r in records:
            d = str(r["date"]).strip()
            m = str(r.get("month") or d[:7]).strip()
            u = str(r["unit"]).strip()
            a = str(r.get("area") or "").strip()
            ib = int(r.get("visitors_individu_bayar", 0) or 0)
            tb = int(r.get("visitors_tidak_bayar", 0) or 0)
            indiv = int(r.get("visitors_individu", 0) or 0)
            if indiv == 0 and (ib > 0 or tb > 0):
                indiv = ib + tb
            rl = int(r.get("visitors_rombongan_langsung", 0) or 0)
            ra = int(r.get("visitors_rombongan_agen", 0) or 0)
            rt = int(r.get("visitors_rombongan_total", 0) or 0)
            if rt == 0 and (rl > 0 or ra > 0):
                rt = rl + ra
            tot = int(r.get("visitors", 0) or 0)
            if tot == 0 and (indiv > 0 or rt > 0):
                tot = indiv + rt
            prepared.append((d, m, u, a, tot, indiv, ib, tb, rl, ra, rt))

        conn.executemany(
            """
            INSERT INTO visitor_actual (
                date, month, unit, area, visitors,
                visitors_individu, visitors_individu_bayar, visitors_tidak_bayar,
                visitors_rombongan_langsung, visitors_rombongan_agen, visitors_rombongan_total
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(date, unit) DO UPDATE SET
                visitors = excluded.visitors,
                area = CASE WHEN excluded.area != '' THEN excluded.area ELSE visitor_actual.area END,
                visitors_individu = excluded.visitors_individu,
                visitors_individu_bayar = excluded.visitors_individu_bayar,
                visitors_tidak_bayar = excluded.visitors_tidak_bayar,
                visitors_rombongan_langsung = excluded.visitors_rombongan_langsung,
                visitors_rombongan_agen = excluded.visitors_rombongan_agen,
                visitors_rombongan_total = excluded.visitors_rombongan_total
            """,
            prepared,
        )
        conn.commit()
        return len(prepared)
    finally:
        conn.close()


def execute_analytics_sql(query: str, params: tuple = (), max_rows: int = 50) -> list[dict[str, Any]]:
    """
    Mengeksekusi query SQL analitik secara aman dan terisolasi (Read-Only).
    Hanya query SELECT / WITH ... SELECT yang diizinkan.
    """
    import re
    q_stripped = query.strip()
    q_upper = q_stripped.upper()
    
    # Validasi awal: Harus SELECT atau WITH
    if not (q_upper.startswith("SELECT") or q_upper.startswith("WITH")):
        raise ValueError("Hanya query SELECT atau WITH yang diizinkan untuk analisis data.")
    
    # Blokir instruksi DDL/DML berbahaya
    FORBIDDEN_KEYWORDS = [
        "INSERT", "UPDATE", "DELETE", "DROP", "ALTER", "CREATE", 
        "TRUNCATE", "REPLACE", "PRAGMA", "ATTACH", "DETACH", "VACUUM",
        "GRANT", "REVOKE"
    ]
    for kw in FORBIDDEN_KEYWORDS:
        if re.search(rf"\b{kw}\b", q_upper):
            raise ValueError(f"Query mengandung instruksi terlarang: {kw}")

    # Enforce limit jika belum ada
    if "LIMIT" not in q_upper:
        q_stripped = f"{q_stripped.rstrip(';')} LIMIT {max_rows}"

    conn = get_connection()
    try:
        cursor = conn.execute(q_stripped, params)
        rows = cursor.fetchall()
        return [dict(row) for row in rows[:max_rows]]
    finally:
        conn.close()


