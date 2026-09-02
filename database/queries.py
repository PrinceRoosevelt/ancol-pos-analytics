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
            SELECT date, month, unit, area, visitors
            FROM visitor_actual
            ORDER BY date ASC
            """
        )
        return [dict(row) for row in cursor.fetchall()]
    finally:
        conn.close()


def save_visitor_db(entry_date: str, unit: str, count: int, area: str = "") -> None:
    """Simpan/Update pengunjung langsung ke tabel visitor_actual."""
    conn = get_connection()
    try:
        month_str = entry_date[:7]
        conn.execute(
            """
            INSERT INTO visitor_actual (date, month, unit, area, visitors)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(date, unit) DO UPDATE SET
                visitors = excluded.visitors,
                area = CASE WHEN excluded.area != '' THEN excluded.area ELSE visitor_actual.area END
            """,
            (entry_date, month_str, unit, area, count),
        )
        conn.commit()
    finally:
        conn.close()

