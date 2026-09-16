# -*- coding: utf-8 -*-
"""
Google Sheets Live Auto-Sync Service for Ancol Store POS Analytics.
Supports Dual Methods:
1. Google Apps Script Webhook (Recommended - Zero Google Cloud setup, 1-click paste script)
2. Google Cloud Service Account (gspread)
"""
from __future__ import annotations

import json
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests

BASE_DIR = Path(__file__).resolve().parent.parent
CONFIG_PATH = BASE_DIR / "config" / "gsheet_config.json"
DEFAULT_CREDS_PATH = BASE_DIR / "config" / "google_credentials.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def load_config() -> Dict[str, Any]:
    """Membaca konfigurasi Google Sheets lokal."""
    if not CONFIG_PATH.exists():
        return {
            "mode": "webhook",
            "webhook_url": "",
            "spreadsheet_id": "",
            "sheet_url": "",
            "credentials_path": "config/google_credentials.json",
            "auto_sync_on_filter": False,
        }
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            cfg = json.load(f)
            if "webhook_url" not in cfg:
                cfg["webhook_url"] = ""
            return cfg
    except Exception:
        return {
            "mode": "webhook",
            "webhook_url": "",
            "spreadsheet_id": "",
            "sheet_url": "",
            "credentials_path": "config/google_credentials.json",
            "auto_sync_on_filter": False,
        }


def save_config(cfg: Dict[str, Any]) -> None:
    """Menyimpan konfigurasi Google Sheets."""
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)


def get_credentials_path() -> Path:
    cfg = load_config()
    rel_or_abs = cfg.get("credentials_path", "config/google_credentials.json")
    p = Path(rel_or_abs)
    if not p.is_absolute():
        p = BASE_DIR / p
    return p


def get_service_account_email() -> Optional[str]:
    """Mengambil email robot service account jika file credentials sudah ada."""
    creds_path = get_credentials_path()
    if not creds_path.exists():
        return None
    try:
        with open(creds_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data.get("client_email")
    except Exception:
        return None


def get_sync_status() -> Dict[str, Any]:
    """Mengecek kelengkapan webhook URL atau kredensial Service Account."""
    cfg = load_config()
    webhook_url = (cfg.get("webhook_url") or "").strip()
    has_webhook = bool(webhook_url and "script.google.com" in webhook_url)

    creds_path = get_credentials_path()
    creds_exist = creds_path.exists()
    client_email = get_service_account_email() if creds_exist else None
    spreadsheet_id = (cfg.get("spreadsheet_id") or "").strip()

    has_sa = bool(creds_exist and client_email and spreadsheet_id)
    is_ready = has_webhook or has_sa

    sheet_url = cfg.get("sheet_url") or (
        f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
        if spreadsheet_id
        else ""
    )

    return {
        "is_ready": is_ready,
        "mode": "webhook" if has_webhook else ("service_account" if has_sa else "none"),
        "webhook_url": webhook_url,
        "credentials_exist": creds_exist,
        "client_email": client_email,
        "spreadsheet_id": spreadsheet_id,
        "sheet_url": sheet_url,
        "credentials_file": str(creds_path.relative_to(BASE_DIR)) if creds_path.is_relative_to(BASE_DIR) else str(creds_path),
    }


def format_rupiah(val: Any) -> str:
    try:
        n = float(val or 0)
        return f"Rp {n:,.0f}".replace(",", ".")
    except Exception:
        return str(val)


def format_growth(val: Any) -> str:
    if val is None or val == "":
        return "-"
    try:
        return f"{float(val):+.1f}%"
    except Exception:
        return "-"


def format_int(val: Any) -> str:
    if val is None or val == "":
        return "0"
    try:
        return f"{int(float(val)):,}".replace(",", ".")
    except Exception:
        return str(val)


def format_diff_rupiah(val: Any) -> str:
    if val is None or val == "":
        return "-"
    try:
        n = float(val)
        sign = "+" if n > 0 else ("-" if n < 0 else "")
        return f"{sign}Rp {abs(n):,.0f}".replace(",", ".")
    except Exception:
        return str(val)


def build_tab_payloads(dashboard_data: Dict[str, Any], filter_desc: str = "Semua Data") -> Dict[str, Any]:
    """
    Menyiapkan data baris untuk 6 Tab Spreadsheet Eksekutif:
    1. 1_Ringkasan_KPI (Ringkasan Keuangan, YoY, Target, dan Model Bisnis)
    2. 2_Performa_Area (Performa Dufan, Atlantis, Beachpark)
    3. 3_Performa_Outlet (Rincian seluruh 18 konter kasir)
    4. 4_Rekanan_Supplier (Scorecard 64 supplier/vendor)
    5. 5_Kategori_Barang (Breakdown 29 kategori)
    6. 6_Top_50_Produk (Top 50 SKU terlaris & velocity)
    """
    timestamp_str = datetime.now().strftime("%d %b %Y %H:%M WIB")
    d = dashboard_data

    summary_2026 = d.get("summary", {}).get("2026", {})
    summary_2025 = d.get("summary", {}).get("2025", {})
    growth = d.get("summary", {}).get("growth", {}) or {}
    diff = d.get("summary", {}).get("diff", {}) or {}

    # -------------------------------------------------------------------------
    # 1. TAB KPI & RINGKASAN EKSEKUTIF
    # -------------------------------------------------------------------------
    kpi_rows: List[List[Any]] = [
        ["ANCOL STORE POS - EXECUTIVE SALES & GROWTH REPORT"],
        [f"Filter Aktif: {filter_desc or 'Semua Data'} | Sinkronisasi: {timestamp_str} | Sumber: SQLite v_sales_analytics"],
        [],
        ["RINGKASAN EKSEKUTIF BISNIS:"],
        [d.get("executive_narrative", "-")],
        [],
        ["Indikator / Metrik", "Tahun 2026 (Utama)", "Tahun 2025 (Pembanding)", "Pertumbuhan YoY (%)", "Selisih (+/-)"],
        [
            "Net Sales (Omset)",
            format_rupiah(summary_2026.get("net_sales")),
            format_rupiah(summary_2025.get("net_sales")),
            format_growth(growth.get("net_sales")),
            format_diff_rupiah(diff.get("net_sales")),
        ],
        [
            "Estimasi Laba Kotor",
            format_rupiah(d.get("gross_profit_2026", 0)),
            "-",
            "-",
            "-",
        ],
        [
            "Gross Margin (%)",
            f"{float(d.get('gross_margin_2026', 0) or 0):.1f}%",
            "-",
            "-",
            "-",
        ],
        [
            "Total HPP / Modal Barang",
            format_rupiah(d.get("total_cogs_2026", 0)),
            "-",
            "-",
            "-",
        ],
        [
            "Target Revenue 2026",
            format_rupiah(d.get("target_revenue", 0)),
            "-",
            "-",
            format_diff_rupiah(d.get("target_gap", 0)),
        ],
        [
            "Pencapaian Target (%)",
            f"{float(d.get('target_achievement', 0) or 0):.2f}%",
            "-",
            "-",
            "-",
        ],
        [
            "Total Transaksi (Struk)",
            format_int(summary_2026.get("transactions", 0)),
            format_int(summary_2025.get("transactions", 0)),
            format_growth(growth.get("transactions")),
            format_int(diff.get("transactions", 0)),
        ],
        [
            "Total Qty Terjual (Pcs)",
            format_int(summary_2026.get("qty", 0)),
            format_int(summary_2025.get("qty", 0)),
            format_growth(growth.get("qty")),
            format_int(diff.get("qty", 0)),
        ],
        [
            "Rata-rata Belanja / Struk (ATV)",
            format_rupiah(summary_2026.get("atv")),
            format_rupiah(summary_2025.get("atv")),
            format_growth(growth.get("atv")),
            format_diff_rupiah(diff.get("atv")),
        ],
        [
            "Total Pengunjung Rekreasi",
            format_int(d.get("total_visitors", 0)),
            "-",
            "-",
            "-",
        ],
        [
            "Spending Per Head (SPH)",
            format_rupiah(d.get("sph", 0)),
            "-",
            "-",
            "-",
        ],
        [
            "Capture Rate Pengunjung (%)",
            f"{float(d.get('capture_rate', 0) or 0):.2f}%",
            "-",
            "-",
            "-",
        ],
        [],
        ["BREAKDOWN MODEL BISNIS (KONSINYASI VS DAGANGAN):"],
        ["Model Bisnis", "Net Sales 2026", "Porsi Penjualan (%)", "Total HPP / Modal", "Estimasi Laba Kotor", "Gross Margin (%)", "Total Qty (Pcs)", "Jumlah SKU"],
    ]

    model_sum = d.get("model_summary", {})
    tot_mb_sales = 0.0
    tot_mb_cogs = 0.0
    tot_mb_profit = 0.0
    tot_mb_qty = 0.0
    tot_mb_skus = 0

    for m_key in ["KONSINYASI", "DAGANGAN"]:
        m = model_sum.get(m_key) or model_sum.get(m_key.lower()) or {}
        s_net = float(m.get("net_sales", 0) or 0.0)
        s_cogs = float(m.get("cogs", 0) or 0.0)
        s_profit = float(m.get("gross_profit", 0) or 0.0)
        s_qty = float(m.get("qty", 0) or 0.0)
        s_skus = int(m.get("products_count", 0) or 0)
        s_share = float(m.get("share", 0) or 0.0)
        s_margin = float(m.get("margin", 0) or 0.0)

        tot_mb_sales += s_net
        tot_mb_cogs += s_cogs
        tot_mb_profit += s_profit
        tot_mb_qty += s_qty
        tot_mb_skus += s_skus

        kpi_rows.append([
            f"Barang {m_key.capitalize()}",
            format_rupiah(s_net),
            f"{s_share:.1f}%",
            format_rupiah(s_cogs),
            format_rupiah(s_profit),
            f"{s_margin:.1f}%",
            format_int(s_qty),
            format_int(s_skus),
        ])

    tot_margin = (tot_mb_profit / tot_mb_sales * 100) if tot_mb_sales > 0 else 0.0
    kpi_rows.append([
        "TOTAL KESELURUHAN",
        format_rupiah(tot_mb_sales),
        "100.0%",
        format_rupiah(tot_mb_cogs),
        format_rupiah(tot_mb_profit),
        f"{tot_margin:.1f}%",
        format_int(tot_mb_qty),
        format_int(tot_mb_skus),
    ])

    # -------------------------------------------------------------------------
    # 2. TAB PERFORMA AREA
    # -------------------------------------------------------------------------
    area_rows: List[List[Any]] = [
        ["Rank", "Nama Area Rekreasi", "Net Sales 2026", "Kontribusi (%)", "Qty 2026", "Transaksi 2026", "Net Sales 2025", "Growth Sales (%)"]
    ]
    area_list = d.get("areas", [])
    if not area_list:
        area_rows.append(["-", "Tidak ada transaksi area untuk filter ini", "-", "-", "-", "-", "-", "-"])
    else:
        for idx, a in enumerate(area_list, 1):
            gw = a.get("growth_sales")
            area_rows.append([
                idx,
                a.get("name", "-"),
                format_rupiah(a.get("2026", {}).get("net_sales", 0)),
                f"{float(a.get('contrib_2026', 0) or 0):.1f}%",
                format_int(a.get("2026", {}).get("qty", 0)),
                format_int(a.get("2026", {}).get("transactions", 0)),
                format_rupiah(a.get("2025", {}).get("net_sales", 0)),
                format_growth(gw),
            ])

    # -------------------------------------------------------------------------
    # 3. TAB PERFORMA OUTLET (18 TITIK KASIR)
    # -------------------------------------------------------------------------
    outlet_to_area: Dict[str, str] = {}
    try:
        from main import read_outlet_mapping
        outlet_to_area = read_outlet_mapping()
    except Exception:
        pass

    outlet_rows: List[List[Any]] = [
        ["Rank", "Nama Konter / Outlet", "Unit Area Induk", "Net Sales 2026", "Kontribusi (%)", "Qty 2026", "Transaksi 2026", "Net Sales 2025", "Growth Sales (%)"]
    ]
    outlet_list = d.get("outlets", [])
    if not outlet_list:
        outlet_rows.append(["-", "Tidak ada transaksi outlet untuk filter ini", "-", "-", "-", "-", "-", "-", "-"])
    else:
        for idx, o in enumerate(outlet_list, 1):
            gw = o.get("growth_sales")
            outlet_rows.append([
                idx,
                o.get("name", "-"),
                outlet_to_area.get(o.get("name"), "LAINNYA"),
                format_rupiah(o.get("2026", {}).get("net_sales", 0)),
                f"{float(o.get('contrib_2026', 0) or 0):.1f}%",
                format_int(o.get("2026", {}).get("qty", 0)),
                format_int(o.get("2026", {}).get("transactions", 0)),
                format_rupiah(o.get("2025", {}).get("net_sales", 0)),
                format_growth(gw),
            ])

    # -------------------------------------------------------------------------
    # 4. TAB REKANAN / SUPPLIER (64 REKANAN)
    # -------------------------------------------------------------------------
    vendor_rows: List[List[Any]] = [
        ["Rank", "Nama Rekanan / Supplier", "Model Bisnis", "Net Sales 2026", "Kontribusi (%)", "Qty 2026 (Pcs)", "Estimasi Laba Kotor", "Gross Margin (%)", "Jumlah SKU"]
    ]
    vendor_list = d.get("suppliers", [])
    if not vendor_list:
        vendor_rows.append(["-", "Tidak ada data rekanan untuk filter ini", "-", "-", "-", "-", "-", "-", "-"])
    else:
        for idx, s in enumerate(vendor_list, 1):
            vendor_rows.append([
                idx,
                s.get("name", "LAINNYA"),
                s.get("jenis", "DAGANGAN"),
                format_rupiah(s.get("net_sales", 0)),
                f"{float(s.get('contrib', 0) or 0):.1f}%",
                format_int(s.get("qty", 0)),
                format_rupiah(s.get("gross_profit", 0)),
                f"{float(s.get('margin', 0) or 0):.1f}%",
                format_int(s.get("products_count", 0)),
            ])

    # -------------------------------------------------------------------------
    # 5. TAB KATEGORI BARANG (29 KATEGORI)
    # -------------------------------------------------------------------------
    cat_rows: List[List[Any]] = [
        ["Rank", "Kategori Produk", "Net Sales 2026", "Kontribusi (%)", "Qty 2026 (Pcs)", "Estimasi Laba Kotor", "Gross Margin (%)", "Jumlah SKU"]
    ]
    cat_list = d.get("categories", [])
    if not cat_list:
        cat_rows.append(["-", "Tidak ada data kategori untuk filter ini", "-", "-", "-", "-", "-", "-"])
    else:
        for idx, c in enumerate(cat_list, 1):
            cat_rows.append([
                idx,
                c.get("name", "LAINNYA"),
                format_rupiah(c.get("net_sales", 0)),
                f"{float(c.get('contrib', 0) or 0):.1f}%",
                format_int(c.get("qty", 0)),
                format_rupiah(c.get("gross_profit", 0)),
                f"{float(c.get('margin', 0) or 0):.1f}%",
                format_int(c.get("products_count", 0)),
            ])

    # -------------------------------------------------------------------------
    # 6. TAB TOP 50 PRODUK
    # -------------------------------------------------------------------------
    prod_rows: List[List[Any]] = [
        ["Rank", "Nama Produk / SKU", "Net Sales 2026", "Kontribusi (%)", "Qty 2026 (Pcs)", "Laju Harian", "Qty 2025", "Growth Qty (%)"]
    ]
    prod_list = d.get("products", [])
    if not prod_list:
        prod_rows.append(["-", "Tidak ada produk terjual untuk filter ini", "-", "-", "-", "-", "-", "-"])
    else:
        for idx, p in enumerate(prod_list[:50], 1):
            gq = p.get("growth_qty")
            d_avg = p.get("daily_avg", 0) or 0.0
            d_str = f"{d_avg:.0f} /hr" if d_avg >= 10 else f"{d_avg:.1f} /hr"
            prod_rows.append([
                idx,
                p.get("name", "-"),
                format_rupiah(p.get("2026", {}).get("net_sales", 0)),
                f"{float(p.get('contrib_2026', 0) or 0):.1f}%",
                format_int(p.get("2026", {}).get("qty", 0)),
                d_str,
                format_int(p.get("2025", {}).get("qty", 0)),
                format_growth(gq),
            ])

    return {
        "kpi_rows": kpi_rows,
        "area_rows": area_rows,
        "outlet_rows": outlet_rows,
        "vendor_rows": vendor_rows,
        "cat_rows": cat_rows,
        "prod_rows": prod_rows,
        "filter_desc": filter_desc,
        "timestamp": timestamp_str,
    }


def sync_via_webhook(webhook_url: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    """Mengirim payload data langsung ke Google Apps Script Web App."""
    try:
        resp = requests.post(
            webhook_url,
            json=payload,
            timeout=50,
            headers={"Content-Type": "application/json"},
            allow_redirects=True,
        )
        if resp.status_code == 200:
            try:
                res_data = resp.json()
                sheet_url = res_data.get("sheet_url")
                if sheet_url:
                    cfg = load_config()
                    cfg["sheet_url"] = sheet_url
                    save_config(cfg)
                return {
                    "success": True,
                    "message": res_data.get("message") or f"Sinkronisasi berhasil ke Google Sheets ({payload.get('timestamp')}).",
                    "sheet_url": sheet_url or resp.url,
                    "timestamp": payload.get("timestamp"),
                }
            except Exception:
                return {
                    "success": True,
                    "message": f"Data terkirim ke Google Sheets ({payload.get('timestamp')}).",
                    "sheet_url": None,
                    "timestamp": payload.get("timestamp"),
                }
        else:
            return {
                "success": False,
                "error_type": "WEBHOOK_HTTP_ERROR",
                "message": f"Google Apps Script merespons kode HTTP {resp.status_code}: {resp.text[:200]}",
                "sheet_url": None,
            }
    except Exception as e:
        return {
            "success": False,
            "error_type": "WEBHOOK_NETWORK_ERROR",
            "message": f"Gagal menghubungi webhook Google Apps Script: {str(e)}",
            "sheet_url": None,
        }


def sync_dashboard_to_sheets(
    dashboard_data: Dict[str, Any], filter_desc: str = "Semua Data"
) -> Dict[str, Any]:
    """Menyinkronkan data dashboard aktif Ancol POS ke Google Sheets via Webhook."""
    status = get_sync_status()
    payload = build_tab_payloads(dashboard_data, filter_desc)

    # METODE UTAMA: Webhook URL Google Apps Script
    if status["mode"] == "webhook" or (status.get("webhook_url") and "script.google.com" in status["webhook_url"]):
        return sync_via_webhook(status["webhook_url"], payload)

    return {
        "success": False,
        "error_type": "NOT_CONFIGURED",
        "message": "Webhook Google Apps Script belum dipasang. Klik tombol ⚙️ di dashboard untuk memasukkan URL Webhook.",
        "sheet_url": None,
    }
