"""
inventory_engine/repository.py
Manages SQLite database storage for Inventory Snapshots (Gudang & Konter).
Reuses the central database connection from database.db_config.
"""
from __future__ import annotations

import sqlite3
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from database.db_config import get_connection


def init_inventory_tables() -> None:
    """Initialize inventory snapshot tables and performance indexes."""
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS inv_snapshots (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                filename_gudang TEXT,
                filename_konter TEXT,
                total_gudang_sku INTEGER DEFAULT 0,
                total_gudang_qty REAL DEFAULT 0,
                total_konter_sku INTEGER DEFAULT 0,
                total_konter_qty REAL DEFAULT 0,
                status TEXT DEFAULT 'ACTIVE'
            );
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS inv_gudang (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                snapshot_id INTEGER NOT NULL,
                sku_code TEXT NOT NULL,
                nama_barang TEXT,
                supplier TEXT,
                stok_akhir REAL DEFAULT 0,
                hpp REAL DEFAULT 0,
                harga_jual REAL DEFAULT 0,
                FOREIGN KEY (snapshot_id) REFERENCES inv_snapshots(id) ON DELETE CASCADE
            );
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS inv_konter (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                snapshot_id INTEGER NOT NULL,
                sku_code TEXT NOT NULL,
                nama_barang TEXT,
                supplier TEXT,
                kategori TEXT,
                hpp REAL DEFAULT 0,
                harga_jual REAL DEFAULT 0,
                outlet_code TEXT NOT NULL,
                stok_qty REAL DEFAULT 0,
                FOREIGN KEY (snapshot_id) REFERENCES inv_snapshots(id) ON DELETE CASCADE
            );
        """)
        # Create performance indices
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_inv_gudang_snap_sku ON inv_gudang(snapshot_id, sku_code);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_inv_konter_snap_sku ON inv_konter(snapshot_id, sku_code);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_inv_konter_snap_outlet ON inv_konter(snapshot_id, outlet_code);")
        conn.commit()


def save_snapshot(
    filename_gudang: Optional[str],
    filename_konter: Optional[str],
    gudang_items: List[Dict[str, Any]],
    konter_items: List[Dict[str, Any]],
) -> int:
    """
    Saves a parsed inventory snapshot into SQLite with high transaction throughput.
    Returns the created snapshot_id.
    """
    init_inventory_tables()
    now_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    total_gudang_sku = len(gudang_items)
    total_gudang_qty = sum(item.get("stok_akhir", 0.0) for item in gudang_items)

    # Count unique SKUs and total pieces across all counters
    unique_konter_skus = {item.get("sku_code") for item in konter_items if item.get("sku_code")}
    total_konter_sku = len(unique_konter_skus)
    total_konter_qty = sum(item.get("stok_qty", 0.0) for item in konter_items)

    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO inv_snapshots (
                created_at, filename_gudang, filename_konter,
                total_gudang_sku, total_gudang_qty,
                total_konter_sku, total_konter_qty, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, 'ACTIVE')
            """,
            (
                now_iso,
                filename_gudang or "",
                filename_konter or "",
                total_gudang_sku,
                total_gudang_qty,
                total_konter_sku,
                total_konter_qty,
            ),
        )
        snapshot_id = cursor.lastrowid

        # Batch insert gudang items
        if gudang_items:
            cursor.executemany(
                """
                INSERT INTO inv_gudang (
                    snapshot_id, sku_code, nama_barang, supplier,
                    stok_akhir, hpp, harga_jual
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    (
                        snapshot_id,
                        item.get("sku_code", ""),
                        item.get("nama_barang", ""),
                        item.get("supplier", ""),
                        float(item.get("stok_akhir") or 0.0),
                        float(item.get("hpp") or 0.0),
                        float(item.get("harga_jual") or 0.0),
                    )
                    for item in gudang_items
                ],
            )

        # Batch insert konter items
        if konter_items:
            cursor.executemany(
                """
                INSERT INTO inv_konter (
                    snapshot_id, sku_code, nama_barang, supplier,
                    kategori, hpp, harga_jual, outlet_code, stok_qty
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    (
                        snapshot_id,
                        item.get("sku_code", ""),
                        item.get("nama_barang", ""),
                        item.get("supplier", ""),
                        item.get("kategori", ""),
                        float(item.get("hpp") or 0.0),
                        float(item.get("harga_jual") or 0.0),
                        item.get("outlet_code", ""),
                        float(item.get("stok_qty") or 0.0),
                    )
                    for item in konter_items
                ],
            )

        conn.commit()
        return snapshot_id


def list_snapshots() -> List[Dict[str, Any]]:
    """Returns all snapshots sorted latest first."""
    init_inventory_tables()
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, created_at, filename_gudang, filename_konter,
                   total_gudang_sku, total_gudang_qty,
                   total_konter_sku, total_konter_qty, status
            FROM inv_snapshots
            ORDER BY id DESC
        """)
        rows = cursor.fetchall()
        return [dict(r) for r in rows]


def get_latest_snapshot_id() -> Optional[int]:
    """Returns the ID of the most recent snapshot, or None."""
    init_inventory_tables()
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM inv_snapshots ORDER BY id DESC LIMIT 1")
        row = cursor.fetchone()
        return row[0] if row else None


def get_snapshot_meta(snapshot_id: int) -> Optional[Dict[str, Any]]:
    """Returns metadata for a specific snapshot."""
    init_inventory_tables()
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM inv_snapshots WHERE id = ?", (snapshot_id,))
        row = cursor.fetchone()
        return dict(row) if row else None


def get_snapshot_gudang_data(snapshot_id: int) -> List[Dict[str, Any]]:
    """Fetches all warehouse stock rows for a snapshot."""
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT sku_code, nama_barang, supplier, stok_akhir, hpp, harga_jual
            FROM inv_gudang
            WHERE snapshot_id = ?
        """, (snapshot_id,))
        return [dict(r) for r in cursor.fetchall()]


def get_snapshot_konter_data(snapshot_id: int) -> List[Dict[str, Any]]:
    """Fetches all counter stock rows for a snapshot."""
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT sku_code, nama_barang, supplier, kategori, hpp, harga_jual, outlet_code, stok_qty
            FROM inv_konter
            WHERE snapshot_id = ?
        """, (snapshot_id,))
        return [dict(r) for r in cursor.fetchall()]

