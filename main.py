import gzip
import math
import os
import re
import sys
from collections import defaultdict
from datetime import datetime
from io import BytesIO
from pathlib import Path
from threading import Lock
from typing import Any

# Fix Windows console encoding for emoji characters
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

from flask import Flask, jsonify, render_template, request, send_file
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from werkzeug.utils import secure_filename

from database import (
    fetch_all_sales,
    fetch_all_targets,
    fetch_all_visitors,
    save_visitor_db,
    sync_database,
    execute_analytics_sql,
)


BASE_DIR = Path(__file__).resolve().parent
# Sales source root using the NEW POS export format.
# The parser automatically discovers every year folder named:
#   data/sales detail 2025
#   data/sales detail 2026
#   data/sales detail 2027
# etc.
# Add new monthly Excel files into the appropriate year folder and they
# will be picked up automatically. No code change is required per month.
SALES_ROOT_DIR = BASE_DIR / "data"
SALES_DETAIL_DIR_PATTERN = re.compile(r"^sales\s+detail\s+(\d{4})$", re.I)
REPORT_DIR = BASE_DIR / "reports"
MAPPING_FILE = BASE_DIR / "config" / "MASTER_OUTLET_MAPPING_V2.xlsx"
BUDGET_FILE = BASE_DIR / "data" / "target 2026" / "BUDGET_MERCH_ONLY_PYTHON_READY.xlsx"
LOOKUP_FILE = BASE_DIR / "config" / "DATA_LOOKUP.xlsx"
# New POS export format:
#   A = No. Invoice (invoice header) / Kode Barang (item)
#   B = Nama Barang
#   C = Tgl & Jam Invoice (date on invoice header)
#   E = time on invoice header
#   G = Jml
#   I = Penjualan / BKP (DPP) at item level
#   N = Disc. Inv. at "Total Struk" level only
#   O = PPN
#   R = Total at "Total Struk" level
#
# IMPORTANT:
# - Column I is preserved as the original item-level Net Sales source.
# - "Total Struk" is NEVER treated as an item.
# - Transaction Net Sales is calculated from column R / 1.11.
# - The transaction-level Net Sales is allocated proportionally back to
#   item rows so product/outlet totals remain accurate after invoice discount.
SALES_VALUE_COLUMN_INDEX = 8       # I
SALES_VALUE_COLUMN_HEADER = "Penjualan"
TRANSACTION_TOTAL_COLUMN_INDEX = 17  # R
TRANSACTION_TOTAL_TAX_FACTOR = 1.11
INVOICE_DISCOUNT_COLUMN_INDEX = 13   # N, informational only
app = Flask(__name__)
app.config["TEMPLATES_AUTO_RELOAD"] = True
app.jinja_env.auto_reload = True
_sales_cache: list[dict[str, Any]] | None = None
_sales_cache_signature: tuple[tuple[str, int, int], ...] = ()
_sales_cache_mapping_signature: tuple[int, int] | None = None
_sales_cache_lock = Lock()
_budget_cache: list[dict[str, Any]] | None = None
_budget_cache_signature: tuple[tuple[int, int], tuple[int, int]] | None = None
_budget_cache_lock = Lock()
_lookup_cache: dict[str, Any] | None = None
_lookup_cache_mtime: float | None = None
_lookup_cache_lock = Lock()
MONTH_NAMES = {
    "01": "Januari",
    "02": "Februari",
    "03": "Maret",
    "04": "April",
    "05": "Mei",
    "06": "Juni",
    "07": "Juli",
    "08": "Agustus",
    "09": "September",
    "10": "Oktober",
    "11": "November",
    "12": "Desember",
}
WEEKDAY_NAMES = {
    0: "Senin",
    1: "Selasa",
    2: "Rabu",
    3: "Kamis",
    4: "Jumat",
    5: "Sabtu",
    6: "Minggu",
}

OUTLET_TO_VISITOR_UNIT = {
    "SWIN Sea World Induk": "SeaWorld",
    "ODIN Samudra Induk": "Samudra",
    "AWIN Atlantis Induk": "Atlantis",
    "AWKL AWA Taman Kelapa 2": "Atlantis",
    "JBIN JBL Induk": "Samudra",
    "DFAR Dufan Arung Jeram": "Dufan",
    "DFGA Dufan Galactica": "Dufan",
    "DFIC Dufan Ice Age": "Dufan",
    "DFIL Dufan Induk Lama": "Dufan",
    "DFIN Dufan Induk": "Dufan",
    "DFKE Dufan Kereta Misteri": "Dufan",
    "DFOR Dufan Oriental": "Dufan",
    "DFSI Dufan Simulator": "Dufan",
    "DFTO Dufan Tornado": "Dufan",
    "DFWW Dufan WWN": "Dufan",
    "OL01 Online Shop": "Beachpark",
    "TJOM Merchandise Ombak Laut": "Beachpark",
    "TJSO Symphony Of The Sea": "Beachpark",
}

VISITOR_UNIT_TO_AREA = {
    "DUFAN": "DUFAN",
    "SEAWORLD": "AWAPARK",
    "SAMUDRA": "AWAPARK",
    "ATLANTIS": "AWAPARK",
    "BEACHPARK": "BEACHPARK",
}


def _number(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value) if math.isfinite(float(value)) else 0.0
    cleaned = re.sub(r"[^\d,.-]", "", str(value).strip())
    if not cleaned:
        return 0.0
    if "," in cleaned and "." in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    elif "," in cleaned:
        cleaned = cleaned.replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _date(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if value is None:
        return None
    for fmt in ("%d %b %Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(value).strip(), fmt)
        except ValueError:
            continue
    return None


def read_outlet_mapping() -> dict[str, str]:
    sheet = load_workbook(MAPPING_FILE, read_only=True, data_only=True)["OUTLET_MAPPING"]
    mapping: dict[str, str] = {}
    try:
        for outlet, area in sheet.iter_rows(min_row=2, values_only=True):
            if outlet and area:
                mapping[str(outlet).strip()] = str(area).strip()
    finally:
        sheet.parent.close()
    return mapping


def load_data_lookup() -> dict[str, Any]:
    """Membaca Master Data Lookup (Supplier, Kategori, Jenis, HPP) dengan file-mtime auto-reload."""
    global _lookup_cache, _lookup_cache_mtime
    if not LOOKUP_FILE.exists():
        return {"by_name": {}, "by_code": {}, "compact_lookup": {}}

    mtime = LOOKUP_FILE.stat().st_mtime
    with _lookup_cache_lock:
        if _lookup_cache is not None and _lookup_cache_mtime == mtime:
            return _lookup_cache

        wb = load_workbook(LOOKUP_FILE, data_only=True, read_only=True)
        sheet_name = "DATA JUAL 23-24-25-26"
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

        by_name: dict[str, dict[str, Any]] = {}
        by_code: dict[str, dict[str, Any]] = {}
        compact_lookup: dict[str, list[Any]] = {}

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not any(row):
                continue
            raw_code = row[0]
            raw_name = row[1]
            if not raw_name and not raw_code:
                continue

            clean_code = str(raw_code).strip() if raw_code is not None else ""
            clean_name = str(raw_name).strip() if raw_name is not None else ""
            supplier = str(row[2]).strip() if len(row) > 2 and row[2] is not None else "LAINNYA"
            kategori = str(row[3]).strip() if len(row) > 3 and row[3] is not None else "LAINNYA"
            jenis = str(row[4]).strip().upper() if len(row) > 4 and row[4] is not None else "DAGANGAN"
            if "KONSIN" in jenis:
                jenis = "KONSINYASI"
            elif "DAGANG" in jenis:
                jenis = "DAGANGAN"
            else:
                jenis = "DAGANGAN"

            hpp_val = 0.0
            if len(row) > 5 and row[5] is not None:
                try:
                    hpp_val = float(row[5])
                    if math.isnan(hpp_val):
                        hpp_val = 0.0
                except (ValueError, TypeError):
                    hpp_val = 0.0

            hj_val = 0.0
            if len(row) > 6 and row[6] is not None:
                try:
                    hj_val = float(row[6])
                    if math.isnan(hj_val):
                        hj_val = 0.0
                except (ValueError, TypeError):
                    hj_val = 0.0

            maskot = str(row[7]).strip() if len(row) > 7 and row[7] is not None else ""

            item_info = {
                "kode": clean_code,
                "nama": clean_name,
                "supplier": supplier,
                "kategori": kategori,
                "jenis": jenis,
                "hpp": hpp_val,
                "harga_jual": hj_val,
                "maskot": maskot,
            }

            if clean_code:
                by_code[clean_code] = item_info
            if clean_name:
                by_name[clean_name.casefold()] = item_info
                compact_lookup[clean_name] = [clean_code, supplier, kategori, jenis, hpp_val, hj_val]

        wb.close()
        _lookup_cache = {
            "by_name": by_name,
            "by_code": by_code,
            "compact_lookup": compact_lookup,
        }
        _lookup_cache_mtime = mtime
        return _lookup_cache


def read_budget_daily() -> list[dict[str, Any]]:
    """Membaca daily targets langsung dari database SQLite (sudah sinkron dengan Excel)."""
    return fetch_all_targets()


def read_target_daily(
    month: str | None = None,
    date: str | None = None,
    outlet: str | None = None,
    area: str | None = None,
    available_dates: set[str] | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
) -> dict[str, float]:
    daily: dict[str, float] = defaultdict(float)
    start_dm = start_date[5:] if start_date else ""
    end_dm = end_date[5:] if end_date else ""
    for target_row in read_budget_daily():
        target_date = target_row["date"]
        target_dm = target_date[5:]
        if start_date or end_date:
            if start_dm and target_dm < start_dm:
                continue
            if end_dm and target_dm > end_dm:
                continue
        elif month and target_row["month"] != month:
            continue
        if date and target_date != date:
            continue
        if available_dates is not None and target_date not in available_dates:
            continue
        if outlet and target_row["outlet"] != outlet:
            continue
        if area and target_row["area"] != area:
            continue
        daily[target_date] += target_row["target"]
    return dict(daily)


def read_all_targets() -> list[dict[str, Any]]:
    return fetch_all_targets()


def read_all_visitors() -> list[dict[str, Any]]:
    """Baca data master pengunjung harian langsung dari database SQLite."""
    return fetch_all_visitors()


def save_visitor_actual(entry_date: str, unit: str, count: int) -> bool:
    """Simpan/Update nilai pengunjung harian pada sheet Excel dan database SQLite."""
    if not BUDGET_FILE.exists():
        return False
    dt_val = _date(entry_date)
    if not dt_val:
        raise ValueError(f"Format tanggal tidak valid: {entry_date}")

    # 1. Update ke file Excel
    wb = load_workbook(BUDGET_FILE, data_only=False)
    try:
        if "VISITOR_ACTUAL" not in wb.sheetnames:
            ws = wb.create_sheet("VISITOR_ACTUAL")
            ws.append(["Date", "Visitor Unit", "Visitor Actual"])
        else:
            ws = wb["VISITOR_ACTUAL"]

        target_date_str = dt_val.strftime("%Y-%m-%d")
        updated = False

        for r in range(2, ws.max_row + 1):
            cell_d = ws.cell(r, 1).value
            cell_u = ws.cell(r, 2).value
            if cell_d is not None and cell_u is not None:
                parsed_cell_d = _date(cell_d)
                if parsed_cell_d and parsed_cell_d.strftime("%Y-%m-%d") == target_date_str:
                    if str(cell_u).strip().casefold() == unit.strip().casefold():
                        ws.cell(r, 3, int(count))
                        updated = True
                        break

        if not updated:
            ws.append([dt_val, unit.strip(), int(count)])

        wb.save(BUDGET_FILE)
    finally:
        wb.close()

    # 2. Update ke Database SQLite
    area_name = VISITOR_UNIT_TO_AREA.get(unit.upper(), "")
    save_visitor_db(target_date_str, unit.strip(), int(count), area_name)

    _invalidate_sales_cache()
    return True


def read_sales() -> list[dict[str, Any]]:
    """Baca data sales langsung dari SQLite Database yang terindeks dan cepat."""
    global _sales_cache
    with _sales_cache_lock:
        if _sales_cache is not None:
            return _sales_cache

        # Pastikan file Excel yang baru / diubah tersinkronisasi ke DB
        sync_database()
        rows = fetch_all_sales()

        _sales_cache = rows
        return rows


def summarize(rows: list[dict[str, Any]], year: int) -> dict[str, Any]:
    selected = [row for row in rows if row["year"] == year]
    qty = sum(row["qty"] for row in selected)
    sales = sum(row["net_sales"] for row in selected)
    item_net_sales = sum(row.get("item_net_sales", 0.0) for row in selected)
    unique_tx = len({(r["outlet"], r["date"], r.get("invoice", "")) for r in selected})
    transactions = unique_tx if unique_tx > 0 else len(selected)
    atv = (sales / transactions) if transactions > 0 else 0.0
    upt = (qty / transactions) if transactions > 0 else 0.0
    asp = (sales / qty) if qty > 0 else 0.0
    return {
        "qty": qty,
        "net_sales": sales,
        "item_net_sales": item_net_sales,
        "transactions": transactions,
        "atv": atv,
        "upt": upt,
        "asp": asp,
        "items_count": len(selected),
    }


def growth_percent(current: float, previous: float) -> float | None:
    if previous == 0:
        return None
    return ((current - previous) / previous) * 100


def build_dashboard(
    rows: list[dict[str, Any]],
    month: str | None = None,
    date: str | None = None,
    outlet: str | None = None,
    area: str | None = None,
    include_raw: bool = True,
    start_date: str | None = None,
    end_date: str | None = None,
    supplier: str | None = None,
    category: str | None = None,
    jenis: str | None = None,
) -> dict[str, Any]:
    # Master lookup integration
    lookup = load_data_lookup()
    compact_lk = lookup.get("compact_lookup", {})
    by_name_lookup = lookup.get("by_name", {})
    by_code_lookup = lookup.get("by_code", {})

    if supplier or category or jenis:
        rows = [
            r for r in rows
            if (not supplier or (compact_lk.get(r["product"]) and compact_lk[r["product"]][1] == supplier))
            and (not category or (compact_lk.get(r["product"]) and compact_lk[r["product"]][2] == category))
            and (not jenis or (compact_lk.get(r["product"]) and compact_lk[r["product"]][3] == jenis))
        ]

    # The 2025 comparison must cover exactly the calendar days already
    # available in 2026. This prevents, for example, comparing 1-29 Aug 2026
    # against a full 1-31 Aug 2025 period.
    start_dm = start_date[5:] if start_date else ""
    end_dm = end_date[5:] if end_date else ""

    current_period_rows = [
        row
        for row in rows
        if row["year"] == 2026
        and (not start_date or row["date"] >= start_date)
        and (not end_date or row["date"] <= end_date)
        and (start_date or end_date or not month or row["month"] == month)
        and (not date or row["date"] == date)
    ]
    current_period_dates = {row["date"] for row in current_period_rows}
    all_sales_dates_2026 = {row["date"] for row in rows if row["year"] == 2026}
    all_sales_day_months_2026 = {target_date[5:] for target_date in all_sales_dates_2026}
    comparison_day_months = {target_date[5:] for target_date in current_period_dates}

    filtered = [
        row
        for row in rows
        if (
            (
                (start_date or end_date)
                and (not start_dm or row["date"][5:] >= start_dm)
                and (not end_dm or row["date"][5:] <= end_dm)
            )
            or (
                not (start_date or end_date)
                and (not month or row["month"][5:] == month[5:])
            )
        )
        and (not date or row["date"][5:] == date[5:])
        and (not outlet or row["outlet"] == outlet)
        and (not area or row["area"] == area)
        and (row["year"] == 2026 or row["date"][5:] in comparison_day_months)
    ]
    summary = {str(year): summarize(filtered, year) for year in (2025, 2026)}
    summary["growth"] = {
        key: growth_percent(summary["2026"][key], summary["2025"][key])
        for key in ("net_sales", "qty", "transactions", "atv", "upt", "asp")
    }
    summary["diff"] = {
        "net_sales": summary["2026"]["net_sales"] - summary["2025"]["net_sales"],
        "transactions": summary["2026"]["transactions"] - summary["2025"]["transactions"],
        "qty": summary["2026"]["qty"] - summary["2025"]["qty"],
        "atv": summary["2026"]["atv"] - summary["2025"]["atv"],
    }

    by_outlet: dict[str, dict[str, dict[str, float]]] = defaultdict(
        lambda: defaultdict(lambda: {"qty": 0.0, "net_sales": 0.0, "transactions": 0})
    )
    by_product: dict[str, dict[str, dict[str, float]]] = defaultdict(
        lambda: defaultdict(lambda: {"qty": 0.0, "net_sales": 0.0, "transactions": 0})
    )
    by_month: dict[str, dict[str, dict[str, float]]] = defaultdict(
        lambda: defaultdict(lambda: {"qty": 0.0, "net_sales": 0.0, "transactions": 0})
    )
    by_day: dict[str, dict[str, dict[str, float]]] = defaultdict(
        lambda: defaultdict(lambda: {"qty": 0.0, "net_sales": 0.0, "transactions": 0})
    )
    by_hour: dict[str, dict[str, dict[str, float]]] = defaultdict(
        lambda: defaultdict(lambda: {"qty": 0.0, "net_sales": 0.0, "transactions": 0})
    )
    by_area: dict[str, dict[str, dict[str, float]]] = defaultdict(
        lambda: defaultdict(lambda: {"qty": 0.0, "net_sales": 0.0, "transactions": 0})
    )
    outlet_area_map: dict[str, str] = {}

    for row in filtered:
        year_key = str(row["year"])
        day_key = row["date"][5:]
        outlet_area_map[row["outlet"]] = row["area"]

        for target_group, key in (
            (by_outlet, row["outlet"]),
            (by_product, row["product"]),
            (by_month, row["month"]),
            (by_day, day_key),
            (by_hour, row["hour"]),
            (by_area, row["area"]),
        ):
            target_group[key][year_key]["qty"] += row["qty"]
            target_group[key][year_key]["net_sales"] += row["net_sales"]
            if row["invoice"]:
                target_group[key][year_key][f"_inv_{row['invoice']}"] = 1

    for target_group in (by_outlet, by_product, by_month, by_day, by_hour, by_area):
        for key in target_group:
            for year_key in ("2025", "2026"):
                inv_keys = [k for k in target_group[key][year_key] if k.startswith("_inv_")]
                target_group[key][year_key]["transactions"] = len(inv_keys)
                for k in inv_keys:
                    del target_group[key][year_key][k]

    def flatten(target_group: dict[str, dict[str, dict[str, float]]]) -> list[dict[str, Any]]:
        result = []
        for name, values in sorted(target_group.items()):
            row_2026 = values.get("2026", {"qty": 0.0, "net_sales": 0.0, "transactions": 0})
            row_2025 = values.get("2025", {"qty": 0.0, "net_sales": 0.0, "transactions": 0})
            sales_2026 = row_2026["net_sales"]
            sales_2025 = row_2025["net_sales"]
            qty_2026 = row_2026["qty"]
            qty_2025 = row_2025["qty"]
            tx_2026 = row_2026["transactions"]
            tx_2025 = row_2025["transactions"]

            result.append(
                {
                    "name": name,
                    "2026": row_2026,
                    "2025": row_2025,
                    "growth_sales": growth_percent(sales_2026, sales_2025),
                    "growth_qty": growth_percent(qty_2026, qty_2025),
                    "growth_tx": growth_percent(tx_2026, tx_2025),
                    "diff_sales": sales_2026 - sales_2025,
                    "diff_qty": qty_2026 - qty_2025,
                    "diff_tx": tx_2026 - tx_2025,
                    "contrib_2026": (
                        (sales_2026 / summary["2026"]["net_sales"] * 100)
                        if summary["2026"]["net_sales"] > 0
                        else 0.0
                    ),
                    "contrib_2025": (
                        (sales_2025 / summary["2025"]["net_sales"] * 100)
                        if summary["2025"]["net_sales"] > 0
                        else 0.0
                    ),
                }
            )
        return sorted(result, key=lambda item: item["2026"]["net_sales"], reverse=True)

    daily = flatten(by_day)
    for item in daily:
        date_value = datetime.strptime(f"2026-{item['name']}", "%Y-%m-%d")
        item["name"] = f"{date_value.day} {MONTH_NAMES[date_value.strftime('%m')]} 2026"

    def period(year: int) -> str:
        year_dates = sorted(row["date"] for row in filtered if row["year"] == year)
        if not year_dates:
            return f"{year}-01-01 -> {year}-01-01"
        return f"{year_dates[0]} -> {year_dates[-1]}"

    target_daily = read_target_daily(
        month,
        date,
        outlet,
        area,
        available_dates=current_period_dates,
        start_date=start_date,
        end_date=end_date,
    )

    has_date_range = bool(start_date or end_date)
    is_monthly_view = not month and not has_date_range

    if is_monthly_view:
        # YTD / Semua Bulan (tanpa custom date range) -> Tampilkan agregasi per Bulan
        m_keys = sorted({row["month"] for row in filtered if row["year"] == 2026})
        daily_chart = [
            {
                "date": m,
                "comparison_date": f"2025-{m[5:]}",
                "label": MONTH_NAMES.get(m[5:], m[5:]),
                "short_label": datetime.strptime(f"{m}-01", "%Y-%m-%d").strftime("%b"),
                "is_monthly": True,
                "2026": by_month.get(m, {}).get("2026", {"net_sales": 0})["net_sales"],
                "2025": by_month.get(m, {}).get("2025", {"net_sales": 0})["net_sales"],
            }
            for m in m_keys
        ]
        target_chart = [
            {
                "date": item["date"],
                "label": item["label"],
                "short_label": item["short_label"],
                "is_monthly": True,
                "target": sum(
                    v for d_str, v in target_daily.items() if d_str.startswith(item["date"])
                ),
                "actual": item["2026"],
            }
            for item in daily_chart
        ]
    else:
        # Bulan Tertentu atau Rentang Tanggal Dipilih -> Tampilkan agregasi per Hari / Tanggal
        daily_chart = [
            {
                "date": f"2026-{day_month}",
                "comparison_date": f"2025-{day_month}",
                "label": (
                    lambda date_value: f"{date_value.day} {date_value.strftime('%b')}"
                )(datetime.strptime(f"2026-{day_month}", "%Y-%m-%d")),
                "short_label": (
                    lambda date_value: f"{date_value.day} {date_value.strftime('%b')}"
                )(datetime.strptime(f"2026-{day_month}", "%Y-%m-%d")),
                "is_monthly": False,
                "2026": by_day.get(day_month, {}).get("2026", {"net_sales": 0})["net_sales"],
                "2025": by_day.get(day_month, {}).get("2025", {"net_sales": 0})["net_sales"],
            }
            for day_month in sorted(
                {row["date"][5:] for row in filtered if row["year"] == 2026}
            )
        ]
        target_chart = [
            {
                "date": item["date"],
                "label": item["label"],
                "short_label": item["short_label"],
                "is_monthly": False,
                "target": target_daily.get(item["date"], 0.0),
                "actual": item["2026"],
            }
            for item in daily_chart
        ]

    total_target = sum(target_daily.values())
    actual_revenue = summary["2026"]["net_sales"]
    target_achievement = (actual_revenue / total_target * 100) if total_target > 0 else 0.0
    target_gap = actual_revenue - total_target

    # Outlets and products lists
    flat_outlets = flatten(by_outlet)
    flat_products = flatten(by_product)
    flat_areas = flatten(by_area)
    flat_hourly = [item for item in sorted(flatten(by_hour), key=lambda item: int(item["name"])) if 6 <= int(item["name"]) <= 20]

    # Active outlets with 2026 sales
    active_outlets_2026 = [o for o in flat_outlets if o["2026"]["net_sales"] > 0]
    top_5_outlets = active_outlets_2026[:5]
    # Outlets needing evaluation: prioritize negative YoY growth, then lowest growth / sales
    negative_growth = [o for o in active_outlets_2026 if o.get("growth_sales") is not None and o["growth_sales"] < 0]
    negative_growth.sort(key=lambda x: x["growth_sales"])
    other_outlets = [o for o in active_outlets_2026 if o not in negative_growth]
    other_outlets.sort(key=lambda x: (x.get("growth_sales") if x.get("growth_sales") is not None else 999999, x["2026"]["net_sales"]))
    bottom_5_outlets = (negative_growth + other_outlets)[:5]
    active_days_2026 = len({row["date"] for row in rows if row["year"] == 2026}) or 1
    for p in flat_products:
        p["daily_avg"] = p["2026"]["qty"] / active_days_2026
        p["weekly_proj"] = round(p["daily_avg"] * 7)
    priority_products = sorted(flat_products, key=lambda x: x["2026"]["qty"], reverse=True)[:5]
    top_5_products = flat_products[:5]

    # Peak hour analysis
    peak_sales_hour = max(flat_hourly, key=lambda x: x["2026"]["net_sales"])["name"] if flat_hourly else "12"
    peak_tx_hour = max(flat_hourly, key=lambda x: x["2026"].get("transactions", 0))["name"] if flat_hourly else "12"

    # Master Lookup (Supplier, Kategori, Jenis, HPP, Gross Profit) Integration
    lookup = load_data_lookup()
    by_name_lookup = lookup.get("by_name", {})
    by_code_lookup = lookup.get("by_code", {})

    for p in flat_products:
        p_name = p["name"].strip()
        lk_meta = by_name_lookup.get(p_name.casefold())
        if lk_meta and lk_meta.get("harga_jual"):
            p["harga_jual"] = lk_meta["harga_jual"]
        elif p["2026"]["qty"] > 0 and p["2026"]["net_sales"] > 0:
            p["harga_jual"] = round((p["2026"]["net_sales"] / p["2026"]["qty"]) * 1.11)
        else:
            p["harga_jual"] = 0.0

    total_cogs_2026 = 0.0
    by_supplier: dict[str, dict[str, Any]] = defaultdict(
        lambda: {"qty": 0.0, "net_sales": 0.0, "cogs": 0.0, "gross_profit": 0.0, "jenis": "DAGANGAN", "products": set()}
    )
    by_category: dict[str, dict[str, Any]] = defaultdict(
        lambda: {"qty": 0.0, "net_sales": 0.0, "cogs": 0.0, "gross_profit": 0.0, "products": set()}
    )
    by_model: dict[str, dict[str, Any]] = {
        "KONSINYASI": {"qty": 0.0, "net_sales": 0.0, "cogs": 0.0, "gross_profit": 0.0, "products": set()},
        "DAGANGAN": {"qty": 0.0, "net_sales": 0.0, "cogs": 0.0, "gross_profit": 0.0, "products": set()},
    }
    unmapped_dict: dict[str, dict[str, Any]] = {}

    for row in filtered:
        if row["year"] != 2026:
            continue
        p_name = row["product"].strip()
        p_norm = p_name.casefold()
        q = row["qty"]
        s = row["net_sales"]

        item_meta = by_name_lookup.get(p_norm)
        if item_meta:
            hpp = item_meta.get("hpp", 0.0)
            cogs = q * hpp
            gp = s - cogs
            supplier = item_meta.get("supplier", "LAINNYA")
            kategori = item_meta.get("kategori", "LAINNYA")
            jenis = item_meta.get("jenis", "DAGANGAN")
            total_cogs_2026 += cogs

            by_supplier[supplier]["qty"] += q
            by_supplier[supplier]["net_sales"] += s
            by_supplier[supplier]["cogs"] += cogs
            by_supplier[supplier]["gross_profit"] += gp
            by_supplier[supplier]["jenis"] = jenis
            by_supplier[supplier]["products"].add(p_name)

            by_category[kategori]["qty"] += q
            by_category[kategori]["net_sales"] += s
            by_category[kategori]["cogs"] += cogs
            by_category[kategori]["gross_profit"] += gp
            by_category[kategori]["products"].add(p_name)

            if jenis in by_model:
                by_model[jenis]["qty"] += q
                by_model[jenis]["net_sales"] += s
                by_model[jenis]["cogs"] += cogs
                by_model[jenis]["gross_profit"] += gp
                by_model[jenis]["products"].add(p_name)
        else:
            if p_name not in unmapped_dict:
                unmapped_dict[p_name] = {"name": p_name, "qty": 0.0, "net_sales": 0.0, "outlets": set()}
            unmapped_dict[p_name]["qty"] += q
            unmapped_dict[p_name]["net_sales"] += s
            unmapped_dict[p_name]["outlets"].add(row["outlet"])

    gross_profit_2026 = actual_revenue - total_cogs_2026
    gross_margin_2026 = (gross_profit_2026 / actual_revenue * 100) if actual_revenue > 0 else 0.0

    flat_suppliers = []
    for s_name, s_data in by_supplier.items():
        s_sales = s_data["net_sales"]
        s_cogs = s_data["cogs"]
        s_gp = s_data["gross_profit"]
        flat_suppliers.append({
            "name": s_name,
            "jenis": s_data["jenis"],
            "qty": s_data["qty"],
            "net_sales": s_sales,
            "cogs": s_cogs,
            "gross_profit": s_gp,
            "margin": (s_gp / s_sales * 100) if s_sales > 0 else 0.0,
            "contrib": (s_sales / actual_revenue * 100) if actual_revenue > 0 else 0.0,
            "products_count": len(s_data["products"]),
        })
    flat_suppliers.sort(key=lambda x: x["net_sales"], reverse=True)

    flat_categories = []
    for c_name, c_data in by_category.items():
        c_sales = c_data["net_sales"]
        c_cogs = c_data["cogs"]
        c_gp = c_data["gross_profit"]
        flat_categories.append({
            "name": c_name,
            "qty": c_data["qty"],
            "net_sales": c_sales,
            "cogs": c_cogs,
            "gross_profit": c_gp,
            "margin": (c_gp / c_sales * 100) if c_sales > 0 else 0.0,
            "contrib": (c_sales / actual_revenue * 100) if actual_revenue > 0 else 0.0,
            "products_count": len(c_data["products"]),
        })
    flat_categories.sort(key=lambda x: x["net_sales"], reverse=True)

    model_summary = {}
    tot_model_sales = by_model["KONSINYASI"]["net_sales"] + by_model["DAGANGAN"]["net_sales"]
    for m_key in ("KONSINYASI", "DAGANGAN"):
        m_sales = by_model[m_key]["net_sales"]
        m_gp = by_model[m_key]["gross_profit"]
        model_summary[m_key] = {
            "net_sales": m_sales,
            "cogs": by_model[m_key]["cogs"],
            "gross_profit": m_gp,
            "qty": by_model[m_key]["qty"],
            "margin": (m_gp / m_sales * 100) if m_sales > 0 else 0.0,
            "share": (m_sales / tot_model_sales * 100) if tot_model_sales > 0 else 0.0,
            "products_count": len(by_model[m_key]["products"]),
        }

    flat_unmapped = [
        {
            "name": u["name"],
            "qty": u["qty"],
            "net_sales": u["net_sales"],
            "outlets": sorted(list(u["outlets"])),
        }
        for u in unmapped_dict.values()
    ]
    flat_unmapped.sort(key=lambda x: x["net_sales"], reverse=True)

    # Executive narrative generation
    growth_sales_val = summary["growth"]["net_sales"]
    growth_str = f"{growth_sales_val:+.2f}%" if growth_sales_val is not None else "0.0%"
    target_achieve_str = f"{target_achievement:.1f}%"
    
    status_text = "MELAMPAUI TARGET" if target_gap >= 0 else "DEFISIT DARI TARGET"
    executive_narrative = (
        f"Total penjualan 2026 tercatat sebesar Rp {actual_revenue:,.0f} ({target_achieve_str} dari target Rp {total_target:,.0f}) dengan estimasi laba kotor Rp {gross_profit_2026:,.0f} (Margin: {gross_margin_2026:.1f}%). "
        f"Performa penjualan tercatat {'tumbuh' if (growth_sales_val or 0) >= 0 else 'terkoreksi'} {growth_str} dibanding periode yang sama tahun 2025. "
        f"Rata-rata transaksi per pelanggan (ATV) adalah Rp {summary['2026']['atv']:,.0f} dengan puncak penjualan pada pukul {peak_sales_hour}:00."
    )

    # Visitor Actual Data Integration
    all_raw_visitors = read_all_visitors()
    total_visitors_2026 = 0
    matched_visitor_unit = None
    if outlet and outlet in OUTLET_TO_VISITOR_UNIT:
        matched_visitor_unit = OUTLET_TO_VISITOR_UNIT[outlet]

    for v in all_raw_visitors:
        v_date = v["date"]
        v_dm = v_date[5:]
        if start_date or end_date:
            if start_dm and v_dm < start_dm:
                continue
            if end_dm and v_dm > end_dm:
                continue
        elif month and v["month"] != month:
            continue

        if date and v_date != date:
            continue

        if matched_visitor_unit:
            if v["unit"].casefold() != matched_visitor_unit.casefold():
                continue
        elif area:
            if v["area"].casefold() != area.casefold():
                continue

        total_visitors_2026 += v["visitors"]

    sph_2026 = (actual_revenue / total_visitors_2026) if total_visitors_2026 > 0 else 0.0
    capture_rate = (summary["2026"]["transactions"] / total_visitors_2026 * 100) if total_visitors_2026 > 0 else 0.0

    return {
        "summary": summary,
        "total_cogs_2026": total_cogs_2026,
        "gross_profit_2026": gross_profit_2026,
        "gross_margin_2026": gross_margin_2026,
        "suppliers": flat_suppliers,
        "categories": flat_categories,
        "model_summary": model_summary,
        "unmapped_skus": flat_unmapped,
        "unmapped_count": len(flat_unmapped),
        "compact_lookup": lookup.get("compact_lookup", {}),
        "outlets": flat_outlets,
        "top_5_outlets": top_5_outlets,
        "bottom_5_outlets": bottom_5_outlets,
        "products": flat_products,
        "top_5_products": top_5_products,
        "monthly": flatten(by_month),
        "daily": daily,
        "areas": flat_areas,
        "hourly": flat_hourly,
        "hourly_chart": [
            {
                "hour": item["name"],
                "net_sales_2026": item["2026"]["net_sales"],
                "net_sales_2025": item["2025"]["net_sales"],
                "transactions_2026": item["2026"].get("transactions", 0),
                "transactions_2025": item["2025"].get("transactions", 0),
            }
            for item in flat_hourly
        ],
        "target_revenue": total_target,
        "target_achievement": target_achievement,
        "target_gap": target_gap,
        "target_status": status_text,
        "total_visitors": total_visitors_2026,
        "sph": sph_2026,
        "capture_rate": capture_rate,
        "peak_sales_hour": peak_sales_hour,
        "peak_tx_hour": peak_tx_hour,
        "priority_products": priority_products,
        "active_days_count": active_days_2026,
        "executive_narrative": executive_narrative,
        "months": sorted({row["month"] for row in rows if row["year"] == 2026}, reverse=True),
        "month_labels": {
            month: MONTH_NAMES.get(month[5:], month)
            for month in {row["month"] for row in rows if row["year"] == 2026}
        },
        "dates": sorted({row["date"] for row in rows if row["year"] == 2026}),
        "date_labels": {
            date: (
                lambda date_value: (
                    f"{WEEKDAY_NAMES[date_value.weekday()]}, "
                    f"{date_value.day} {MONTH_NAMES[date_value.strftime('%m')]} 2026"
                )
            )(datetime.strptime(date, "%Y-%m-%d"))
            for date in {row["date"] for row in rows if row["year"] == 2026}
        },
        "latest_date_label": (
            (
                lambda latest_dt: (
                    f"{latest_dt.day} {MONTH_NAMES[latest_dt.strftime('%m')]} {latest_dt.year}"
                )
            )(datetime.strptime(max({row["date"] for row in rows if row["year"] == 2026}), "%Y-%m-%d"))
            if any(row["year"] == 2026 for row in rows) else ""
        ),
        "outlet_options": sorted(
            {
                row["outlet"]
                for row in rows
                if not area or row["area"] == area
            }
        ),
        "area_options": sorted({row["area"] for row in rows}),
        "supplier_options": sorted({v[1] for v in lookup.get("compact_lookup", {}).values() if v[1]}),
        "category_options": sorted({v[2] for v in lookup.get("compact_lookup", {}).values() if v[2]}),
        "total_outlets": len({row["outlet"] for row in filtered}),
        "daily_chart": daily_chart,
        "target_chart": target_chart,
        "comparison_period": {
            "current": period(2026),
            "comparison": period(2025),
        },
        "rows_read": len(rows),
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "compact_sales": [
            [
                r["year"],
                r["date"],
                r["hour"],
                r["outlet"],
                r["area"],
                r["product"],
                r["qty"],
                r["net_sales"],
                r.get("invoice", ""),
                r.get("item_net_sales", 0.0),
            ]
            for r in rows
            if r["year"] == 2026 or r["date"][5:] in all_sales_day_months_2026
        ]
        if include_raw
        else [],
        # The browser-side filter engine receives only target days for which
        # 2026 sales data is currently available. This keeps its default and
        # filtered target totals aligned with the backend's MTD calculation.
        "compact_targets": (
            [
                target
                for target in read_all_targets()
                if target["date"] in all_sales_dates_2026
            ]
            if include_raw
            else []
        ),
        "compact_visitors": (
            all_raw_visitors
            if include_raw
            else []
        ),
    }


def build_excel_report(data: dict[str, Any], filter_desc: str = "") -> BytesIO:
    wb = Workbook()

    # Official Ancol Corporate Color Palette & Styling
    header_fill = PatternFill(start_color="0033A0", end_color="0033A0", fill_type="solid")       # Ocean Blue (Pantone 286 C)
    header_accent = PatternFill(start_color="00205B", end_color="00205B", fill_type="solid")     # Deep Blue (Pantone 281 C)
    total_fill = PatternFill(start_color="EAF1FC", end_color="EAF1FC", fill_type="solid")        # Brand Subtle Blue
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    title_font = Font(name="Segoe UI", size=14, bold=True, color="00205B")
    subtitle_font = Font(name="Segoe UI", size=9, italic=True, color="536B88")
    section_font = Font(name="Segoe UI", size=11, bold=True, color="0033A0")
    bold_font = Font(name="Segoe UI", size=10, bold=True, color="00205B")
    regular_font = Font(name="Segoe UI", size=10, color="1E293B")
    total_font = Font(name="Segoe UI", size=10, bold=True, color="00205B")

    thin_border = Border(
        left=Side(style="thin", color="DBE4F0"),
        right=Side(style="thin", color="DBE4F0"),
        top=Side(style="thin", color="DBE4F0"),
        bottom=Side(style="thin", color="DBE4F0"),
    )
    total_border = Border(
        left=Side(style="thin", color="DBE4F0"),
        right=Side(style="thin", color="DBE4F0"),
        top=Side(style="thin", color="0033A0"),
        bottom=Side(style="double", color="00205B"),
    )

    outlet_to_area = read_outlet_mapping()
    compact_lk = data.get("compact_lookup", {})

    # =========================================================================
    # 1. SHEET: RINGKASAN KPI & MODEL BISNIS
    # =========================================================================
    ws_sum = wb.active
    ws_sum.title = "Ringkasan KPI"
    ws_sum.views.sheetView[0].showGridLines = True

    ws_sum["A1"] = "SALES PERFORMANCE & GROWTH REPORT"
    ws_sum["A1"].font = title_font

    cut_off_info = data.get("latest_date_label") or "Terbaru"
    ws_sum["A2"] = f"Filter: {filter_desc or 'Semua Data'} | Diunduh: {datetime.now().strftime('%d %b %Y %H:%M')} | Data Cut-Off: {cut_off_info}"
    ws_sum["A2"].font = subtitle_font

    ws_sum["A4"] = "RINGKASAN EKSEKUTIF BISNIS:"
    ws_sum["A4"].font = section_font

    ws_sum.merge_cells("A5:E5")
    ws_sum["A5"] = data.get("executive_narrative", "")
    ws_sum["A5"].font = regular_font
    ws_sum["A5"].alignment = Alignment(wrap_text=True, vertical="top")
    ws_sum.row_dimensions[5].height = 54

    headers_kpi = [
        "Indikator / Metrik",
        "Tahun 2026 (Utama)",
        "Tahun 2025 (Pembanding)",
        "Pertumbuhan YoY (%)",
        "Selisih (+/-)",
    ]
    for col_num, h in enumerate(headers_kpi, 1):
        cell = ws_sum.cell(row=7, column=col_num, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[7].height = 24

    kpi_rows = [
        ("Net Sales (Omset Bersih)", data["summary"]["2026"]["net_sales"], data["summary"]["2025"]["net_sales"], data["summary"]["growth"]["net_sales"], data["summary"]["diff"]["net_sales"], "currency"),
        ("Estimasi Laba Kotor (Gross Profit)", data.get("gross_profit_2026", 0), None, None, None, "currency"),
        ("Gross Profit Margin (%)", (data.get("gross_margin_2026", 0) / 100) if data.get("gross_margin_2026") else 0, None, None, None, "percent"),
        ("Total HPP / Modal Barang (COGS)", data.get("total_cogs_2026", 0), None, None, None, "currency"),
        ("Target Revenue 2026", data.get("target_revenue", 0), None, None, data.get("target_gap", 0), "currency"),
        ("Pencapaian Target (%)", (data.get("target_achievement", 0) / 100) if data.get("target_revenue") else 0, None, None, None, "percent"),
        ("Total Transaksi (Struk)", data["summary"]["2026"]["transactions"], data["summary"]["2025"]["transactions"], data["summary"]["growth"]["transactions"], data["summary"]["diff"]["transactions"], "number"),
        ("Total Qty Terjual (Pcs)", data["summary"]["2026"]["qty"], data["summary"]["2025"]["qty"], data["summary"]["growth"]["qty"], data["summary"]["diff"]["qty"], "number"),
        ("ATV (Average Transaction Value)", data["summary"]["2026"]["atv"], data["summary"]["2025"]["atv"], data["summary"]["growth"]["atv"], data["summary"]["diff"]["atv"], "currency"),
        ("UPT (Units Per Transaction)", data["summary"]["2026"]["upt"], data["summary"]["2025"]["upt"], data["summary"]["growth"]["upt"], None, "decimal"),
        ("ASP (Average Selling Price)", data["summary"]["2026"]["asp"], data["summary"]["2025"]["asp"], data["summary"]["growth"]["asp"], None, "currency"),
        ("Total Pengunjung Rekreasi", data.get("total_visitors", 0), None, None, None, "number"),
        ("Spending Per Head (SPH)", data.get("sph", 0), None, None, None, "currency"),
        ("Capture Rate Pengunjung (%)", (data.get("capture_rate", 0) / 100) if data.get("capture_rate") else 0, None, None, None, "percent"),
    ]

    for r_idx, (label, v26, v25, growth, diff, fmt) in enumerate(kpi_rows, 8):
        c_lbl = ws_sum.cell(row=r_idx, column=1, value=label)
        c_lbl.font = bold_font if "Net Sales" in label or "Laba Kotor" in label else regular_font
        c26 = ws_sum.cell(row=r_idx, column=2, value=v26 if v26 is not None else "-")
        c25 = ws_sum.cell(row=r_idx, column=3, value=v25 if v25 is not None else "-")
        cg = ws_sum.cell(row=r_idx, column=4, value=(growth / 100) if growth is not None else "-")
        cd = ws_sum.cell(row=r_idx, column=5, value=diff if diff is not None else "-")

        for c in (c_lbl, c26, c25, cg, cd):
            c.border = thin_border
            if c != c_lbl: c.font = regular_font

        if fmt == "currency":
            if isinstance(v26, (int, float)): c26.number_format = "#,##0"
            if isinstance(v25, (int, float)): c25.number_format = "#,##0"
            if isinstance(diff, (int, float)): cd.number_format = "+#,##0;-#,##0;0"
        elif fmt == "number":
            if isinstance(v26, (int, float)): c26.number_format = "#,##0"
            if isinstance(v25, (int, float)): c25.number_format = "#,##0"
            if isinstance(diff, (int, float)): cd.number_format = "+#,##0;-#,##0;0"
        elif fmt == "percent":
            if isinstance(v26, (int, float)): c26.number_format = "0.00%"
        elif fmt == "decimal":
            if isinstance(v26, (int, float)): c26.number_format = "0.00"
            if isinstance(v25, (int, float)): c25.number_format = "0.00"

        if isinstance(growth, (int, float)):
            cg.number_format = "+0.00%;-0.00%;0.00%"

    # Sub-Tabel Model Bisnis: Konsinyasi vs Dagangan
    row_mb_title = len(kpi_rows) + 10
    ws_sum.cell(row=row_mb_title, column=1, value="BREAKDOWN MODEL BISNIS (KONSINYASI VS DAGANGAN):").font = section_font

    mb_headers = [
        "Model Bisnis",
        "Net Sales 2026",
        "Porsi Penjualan (%)",
        "Total HPP / Modal",
        "Estimasi Laba Kotor",
        "Gross Margin (%)",
        "Total Qty (Pcs)",
        "Jumlah SKU Aktif",
    ]
    row_mb_head = row_mb_title + 1
    for col_num, h in enumerate(mb_headers, 1):
        cell = ws_sum.cell(row=row_mb_head, column=col_num, value=h)
        cell.fill = header_accent
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[row_mb_head].height = 22

    model_summary = data.get("model_summary", {})
    start_mb_row = row_mb_head + 1
    for idx, m_key in enumerate(["KONSINYASI", "DAGANGAN"]):
        curr_row = start_mb_row + idx
        m_item = model_summary.get(m_key, {})
        ws_sum.cell(row=curr_row, column=1, value=f"Barang {m_key.capitalize()}").font = bold_font
        
        c_s = ws_sum.cell(row=curr_row, column=2, value=m_item.get("net_sales", 0))
        c_s.number_format = "#,##0"
        
        c_sh = ws_sum.cell(row=curr_row, column=3, value=(m_item.get("share", 0) / 100))
        c_sh.number_format = "0.00%"
        
        c_cg = ws_sum.cell(row=curr_row, column=4, value=m_item.get("cogs", 0))
        c_cg.number_format = "#,##0"
        
        c_gp = ws_sum.cell(row=curr_row, column=5, value=m_item.get("gross_profit", 0))
        c_gp.number_format = "#,##0"
        
        c_mg = ws_sum.cell(row=curr_row, column=6, value=(m_item.get("margin", 0) / 100))
        c_mg.number_format = "0.00%"
        
        c_qt = ws_sum.cell(row=curr_row, column=7, value=m_item.get("qty", 0))
        c_qt.number_format = "#,##0"
        
        c_sk = ws_sum.cell(row=curr_row, column=8, value=m_item.get("products_count", 0))
        c_sk.number_format = "#,##0"

        for col_idx in range(1, 9):
            ws_sum.cell(row=curr_row, column=col_idx).border = thin_border

    # Baris Total Model Bisnis dengan Formula
    tot_mb_row = start_mb_row + 2
    ws_sum.cell(row=tot_mb_row, column=1, value="TOTAL KESELURUHAN").font = total_font
    c_tot_s = ws_sum.cell(row=tot_mb_row, column=2, value=f"=SUM(B{start_mb_row}:B{tot_mb_row-1})")
    c_tot_s.number_format = "#,##0"
    c_tot_sh = ws_sum.cell(row=tot_mb_row, column=3, value=f"=SUM(C{start_mb_row}:C{tot_mb_row-1})")
    c_tot_sh.number_format = "0.00%"
    c_tot_cg = ws_sum.cell(row=tot_mb_row, column=4, value=f"=SUM(D{start_mb_row}:D{tot_mb_row-1})")
    c_tot_cg.number_format = "#,##0"
    c_tot_gp = ws_sum.cell(row=tot_mb_row, column=5, value=f"=SUM(E{start_mb_row}:E{tot_mb_row-1})")
    c_tot_gp.number_format = "#,##0"
    c_tot_mg = ws_sum.cell(row=tot_mb_row, column=6, value=f"=E{tot_mb_row}/B{tot_mb_row}")
    c_tot_mg.number_format = "0.00%"
    c_tot_qt = ws_sum.cell(row=tot_mb_row, column=7, value=f"=SUM(G{start_mb_row}:G{tot_mb_row-1})")
    c_tot_qt.number_format = "#,##0"
    c_tot_sk = ws_sum.cell(row=tot_mb_row, column=8, value=f"=SUM(H{start_mb_row}:H{tot_mb_row-1})")
    c_tot_sk.number_format = "#,##0"

    for col_idx in range(1, 9):
        c = ws_sum.cell(row=tot_mb_row, column=col_idx)
        c.fill = total_fill
        c.font = total_font
        c.border = total_border

    # =========================================================================
    # 2. SHEET: PERFORMA AREA
    # =========================================================================
    ws_ar = wb.create_sheet(title="Performa Area")
    ws_ar.views.sheetView[0].showGridLines = True
    area_headers = [
        "Rank", "Nama Area Rekreasi", "Net Sales 2026", "Kontribusi 2026 (%)",
        "Qty 2026", "Transaksi 2026", "Net Sales 2025", "Growth Sales (%)",
    ]
    for col_num, h in enumerate(area_headers, 1):
        cell = ws_ar.cell(row=1, column=col_num, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_ar.row_dimensions[1].height = 24

    areas_list = data.get("areas", [])
    for r_idx, a in enumerate(areas_list, 2):
        ws_ar.cell(row=r_idx, column=1, value=r_idx - 1).alignment = Alignment(horizontal="center")
        ws_ar.cell(row=r_idx, column=2, value=a["name"]).font = bold_font
        c_s26 = ws_ar.cell(row=r_idx, column=3, value=a["2026"]["net_sales"])
        c_s26.number_format = "#,##0"
        c_cb = ws_ar.cell(row=r_idx, column=4, value=(a.get("contrib_2026", 0) / 100))
        c_cb.number_format = "0.00%"
        c_q26 = ws_ar.cell(row=r_idx, column=5, value=a["2026"]["qty"])
        c_q26.number_format = "#,##0"
        c_t26 = ws_ar.cell(row=r_idx, column=6, value=a["2026"].get("transactions", 0))
        c_t26.number_format = "#,##0"
        c_s25 = ws_ar.cell(row=r_idx, column=7, value=a["2025"]["net_sales"])
        c_s25.number_format = "#,##0"
        gw = a.get("growth_sales")
        c_gw = ws_ar.cell(row=r_idx, column=8, value=(gw / 100) if gw is not None else "-")
        if gw is not None: c_gw.number_format = "+0.00%;-0.00%;0.00%"

        for col_idx in range(1, 9):
            c = ws_ar.cell(row=r_idx, column=col_idx)
            c.border = thin_border
            if col_idx != 2: c.font = regular_font

    if areas_list:
        last_ar_row = len(areas_list) + 1
        tot_ar_row = last_ar_row + 1
        ws_ar.cell(row=tot_ar_row, column=1, value="").alignment = Alignment(horizontal="center")
        ws_ar.cell(row=tot_ar_row, column=2, value="TOTAL KESELURUHAN").font = total_font
        ws_ar.cell(row=tot_ar_row, column=3, value=f"=SUM(C2:C{last_ar_row})").number_format = "#,##0"
        ws_ar.cell(row=tot_ar_row, column=4, value=f"=SUM(D2:D{last_ar_row})").number_format = "0.00%"
        ws_ar.cell(row=tot_ar_row, column=5, value=f"=SUM(E2:E{last_ar_row})").number_format = "#,##0"
        ws_ar.cell(row=tot_ar_row, column=6, value=f"=SUM(F2:F{last_ar_row})").number_format = "#,##0"
        ws_ar.cell(row=tot_ar_row, column=7, value=f"=SUM(G2:G{last_ar_row})").number_format = "#,##0"
        ws_ar.cell(row=tot_ar_row, column=8, value=f"=(C{tot_ar_row}-G{tot_ar_row})/G{tot_ar_row}").number_format = "+0.00%;-0.00%;0.00%"
        for col_idx in range(1, 9):
            c = ws_ar.cell(row=tot_ar_row, column=col_idx)
            c.fill = total_fill
            c.font = total_font
            c.border = total_border

    # =========================================================================
    # 3. SHEET: PERFORMA OUTLET
    # =========================================================================
    ws_out = wb.create_sheet(title="Performa Outlet")
    ws_out.views.sheetView[0].showGridLines = True
    out_headers = [
        "Rank", "Nama Konter / Outlet", "Unit Area Induk", "Net Sales 2026",
        "Kontribusi 2026 (%)", "Qty 2026", "Transaksi 2026", "Net Sales 2025", "Growth Sales (%)",
    ]
    for col_num, h in enumerate(out_headers, 1):
        cell = ws_out.cell(row=1, column=col_num, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_out.row_dimensions[1].height = 24

    outlets_list = data.get("outlets", [])
    for r_idx, o in enumerate(outlets_list, 2):
        ws_out.cell(row=r_idx, column=1, value=r_idx - 1).alignment = Alignment(horizontal="center")
        ws_out.cell(row=r_idx, column=2, value=o["name"]).font = bold_font
        ws_out.cell(row=r_idx, column=3, value=outlet_to_area.get(o["name"], "LAINNYA")).font = regular_font
        c_s26 = ws_out.cell(row=r_idx, column=4, value=o["2026"]["net_sales"])
        c_s26.number_format = "#,##0"
        c_cb = ws_out.cell(row=r_idx, column=5, value=(o.get("contrib_2026", 0) / 100))
        c_cb.number_format = "0.00%"
        c_q26 = ws_out.cell(row=r_idx, column=6, value=o["2026"]["qty"])
        c_q26.number_format = "#,##0"
        c_t26 = ws_out.cell(row=r_idx, column=7, value=o["2026"].get("transactions", 0))
        c_t26.number_format = "#,##0"
        c_s25 = ws_out.cell(row=r_idx, column=8, value=o["2025"]["net_sales"])
        c_s25.number_format = "#,##0"
        gw = o.get("growth_sales")
        c_gw = ws_out.cell(row=r_idx, column=9, value=(gw / 100) if gw is not None else "-")
        if gw is not None: c_gw.number_format = "+0.00%;-0.00%;0.00%"

        for col_idx in range(1, 10):
            c = ws_out.cell(row=r_idx, column=col_idx)
            c.border = thin_border
            if col_idx not in (2, 3): c.font = regular_font

    if outlets_list:
        last_out_row = len(outlets_list) + 1
        tot_out_row = last_out_row + 1
        ws_out.cell(row=tot_out_row, column=1, value="").alignment = Alignment(horizontal="center")
        ws_out.cell(row=tot_out_row, column=2, value="TOTAL KESELURUHAN").font = total_font
        ws_out.cell(row=tot_out_row, column=3, value="").font = total_font
        ws_out.cell(row=tot_out_row, column=4, value=f"=SUM(D2:D{last_out_row})").number_format = "#,##0"
        ws_out.cell(row=tot_out_row, column=5, value=f"=SUM(E2:E{last_out_row})").number_format = "0.00%"
        ws_out.cell(row=tot_out_row, column=6, value=f"=SUM(F2:F{last_out_row})").number_format = "#,##0"
        ws_out.cell(row=tot_out_row, column=7, value=f"=SUM(G2:G{last_out_row})").number_format = "#,##0"
        ws_out.cell(row=tot_out_row, column=8, value=f"=SUM(H2:H{last_out_row})").number_format = "#,##0"
        ws_out.cell(row=tot_out_row, column=9, value=f"=(D{tot_out_row}-H{tot_out_row})/H{tot_out_row}").number_format = "+0.00%;-0.00%;0.00%"
        for col_idx in range(1, 10):
            c = ws_out.cell(row=tot_out_row, column=col_idx)
            c.fill = total_fill
            c.font = total_font
            c.border = total_border

    # =========================================================================
    # 4. SHEET: REKANAN & SUPPLIER (VENDOR PERFORMANCE)
    # =========================================================================
    suppliers_list = data.get("suppliers", [])
    if suppliers_list:
        ws_supp = wb.create_sheet(title="Top Supplier")
        ws_supp.views.sheetView[0].showGridLines = True
        supp_headers = [
            "Rank", "Nama Rekanan / Supplier", "Model Bisnis", "Net Sales 2026",
            "Kontribusi (%)", "Qty 2026 (Pcs)", "Estimasi Laba Kotor", "Gross Margin (%)", "Jumlah SKU",
        ]
        for col_num, h in enumerate(supp_headers, 1):
            cell = ws_supp.cell(row=1, column=col_num, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws_supp.row_dimensions[1].height = 24

        for r_idx, s in enumerate(suppliers_list, 2):
            ws_supp.cell(row=r_idx, column=1, value=r_idx - 1).alignment = Alignment(horizontal="center")
            ws_supp.cell(row=r_idx, column=2, value=s["name"]).font = bold_font
            ws_supp.cell(row=r_idx, column=3, value=s.get("jenis", "DAGANGAN")).alignment = Alignment(horizontal="center")
            c_s = ws_supp.cell(row=r_idx, column=4, value=s.get("net_sales", 0))
            c_s.number_format = "#,##0"
            c_cb = ws_supp.cell(row=r_idx, column=5, value=(s.get("contrib", 0) / 100))
            c_cb.number_format = "0.00%"
            c_q = ws_supp.cell(row=r_idx, column=6, value=s.get("qty", 0))
            c_q.number_format = "#,##0"
            c_gp = ws_supp.cell(row=r_idx, column=7, value=s.get("gross_profit", 0))
            c_gp.number_format = "#,##0"
            c_mg = ws_supp.cell(row=r_idx, column=8, value=(s.get("margin", 0) / 100))
            c_mg.number_format = "0.00%"
            c_sk = ws_supp.cell(row=r_idx, column=9, value=s.get("products_count", 0))
            c_sk.number_format = "#,##0"

            for col_idx in range(1, 10):
                c = ws_supp.cell(row=r_idx, column=col_idx)
                c.border = thin_border
                if col_idx != 2: c.font = regular_font

        last_sup_row = len(suppliers_list) + 1
        tot_sup_row = last_sup_row + 1
        ws_supp.cell(row=tot_sup_row, column=1, value="").alignment = Alignment(horizontal="center")
        ws_supp.cell(row=tot_sup_row, column=2, value="TOTAL KESELURUHAN").font = total_font
        ws_supp.cell(row=tot_sup_row, column=3, value="").alignment = Alignment(horizontal="center")
        ws_supp.cell(row=tot_sup_row, column=4, value=f"=SUM(D2:D{last_sup_row})").number_format = "#,##0"
        ws_supp.cell(row=tot_sup_row, column=5, value=f"=SUM(E2:E{last_sup_row})").number_format = "0.00%"
        ws_supp.cell(row=tot_sup_row, column=6, value=f"=SUM(F2:F{last_sup_row})").number_format = "#,##0"
        ws_supp.cell(row=tot_sup_row, column=7, value=f"=SUM(G2:G{last_sup_row})").number_format = "#,##0"
        ws_supp.cell(row=tot_sup_row, column=8, value=f"=G{tot_sup_row}/D{tot_sup_row}").number_format = "0.00%"
        ws_supp.cell(row=tot_sup_row, column=9, value=f"=SUM(I2:I{last_sup_row})").number_format = "#,##0"
        for col_idx in range(1, 10):
            c = ws_supp.cell(row=tot_sup_row, column=col_idx)
            c.fill = total_fill
            c.font = total_font
            c.border = total_border

    # =========================================================================
    # 5. SHEET: KATEGORI MERCHANDISE
    # =========================================================================
    categories_list = data.get("categories", [])
    if categories_list:
        ws_cat = wb.create_sheet(title="Kategori Produk")
        ws_cat.views.sheetView[0].showGridLines = True
        cat_headers = [
            "Rank", "Kategori Merchandise", "Net Sales 2026", "Kontribusi (%)",
            "Qty 2026 (Pcs)", "Estimasi Laba Kotor", "Gross Margin (%)", "Jumlah SKU",
        ]
        for col_num, h in enumerate(cat_headers, 1):
            cell = ws_cat.cell(row=1, column=col_num, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws_cat.row_dimensions[1].height = 24

        for r_idx, c_item in enumerate(categories_list, 2):
            ws_cat.cell(row=r_idx, column=1, value=r_idx - 1).alignment = Alignment(horizontal="center")
            ws_cat.cell(row=r_idx, column=2, value=c_item["name"]).font = bold_font
            c_s = ws_cat.cell(row=r_idx, column=3, value=c_item.get("net_sales", 0))
            c_s.number_format = "#,##0"
            c_cb = ws_cat.cell(row=r_idx, column=4, value=(c_item.get("contrib", 0) / 100))
            c_cb.number_format = "0.00%"
            c_q = ws_cat.cell(row=r_idx, column=5, value=c_item.get("qty", 0))
            c_q.number_format = "#,##0"
            c_gp = ws_cat.cell(row=r_idx, column=6, value=c_item.get("gross_profit", 0))
            c_gp.number_format = "#,##0"
            c_mg = ws_cat.cell(row=r_idx, column=7, value=(c_item.get("margin", 0) / 100))
            c_mg.number_format = "0.00%"
            c_sk = ws_cat.cell(row=r_idx, column=8, value=c_item.get("products_count", 0))
            c_sk.number_format = "#,##0"

            for col_idx in range(1, 9):
                c = ws_cat.cell(row=r_idx, column=col_idx)
                c.border = thin_border
                if col_idx != 2: c.font = regular_font

        last_cat_row = len(categories_list) + 1
        tot_cat_row = last_cat_row + 1
        ws_cat.cell(row=tot_cat_row, column=1, value="").alignment = Alignment(horizontal="center")
        ws_cat.cell(row=tot_cat_row, column=2, value="TOTAL KESELURUHAN").font = total_font
        ws_cat.cell(row=tot_cat_row, column=3, value=f"=SUM(C2:C{last_cat_row})").number_format = "#,##0"
        ws_cat.cell(row=tot_cat_row, column=4, value=f"=SUM(D2:D{last_cat_row})").number_format = "0.00%"
        ws_cat.cell(row=tot_cat_row, column=5, value=f"=SUM(E2:E{last_cat_row})").number_format = "#,##0"
        ws_cat.cell(row=tot_cat_row, column=6, value=f"=SUM(F2:F{last_cat_row})").number_format = "#,##0"
        ws_cat.cell(row=tot_cat_row, column=7, value=f"=F{tot_cat_row}/C{tot_cat_row}").number_format = "0.00%"
        ws_cat.cell(row=tot_cat_row, column=8, value=f"=SUM(H2:H{last_cat_row})").number_format = "#,##0"
        for col_idx in range(1, 9):
            c = ws_cat.cell(row=tot_cat_row, column=col_idx)
            c.fill = total_fill
            c.font = total_font
            c.border = total_border

    # =========================================================================
    # 6. SHEET: TOP PRODUK & SKU DETAIL (ENRICHED WITH MASTER LOOKUP)
    # =========================================================================
    ws_pr = wb.create_sheet(title="Top Produk")
    ws_pr.views.sheetView[0].showGridLines = True
    prod_headers = [
        "Rank", "Kode SKU", "Nama Barang / Produk", "Rekanan / Supplier",
        "Kategori", "Model Bisnis", "Net Sales 2026", "Kontribusi 2026 (%)",
        "Qty 2026", "Qty 2025", "Growth Qty (%)",
    ]
    for col_num, h in enumerate(prod_headers, 1):
        cell = ws_pr.cell(row=1, column=col_num, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_pr.row_dimensions[1].height = 24

    products_list = data.get("products", [])[:300]
    for r_idx, p in enumerate(products_list, 2):
        p_name = p["name"]
        meta = compact_lk.get(p_name, ["-", "-", "-", "-", 0, 0])
        sku_code = meta[0] if meta[0] else "-"
        supplier_name = meta[1] if meta[1] else "-"
        category_name = meta[2] if meta[2] else "-"
        business_model = meta[3] if meta[3] else "-"

        ws_pr.cell(row=r_idx, column=1, value=r_idx - 1).alignment = Alignment(horizontal="center")
        ws_pr.cell(row=r_idx, column=2, value=sku_code).alignment = Alignment(horizontal="center")
        ws_pr.cell(row=r_idx, column=3, value=p_name).font = bold_font
        ws_pr.cell(row=r_idx, column=4, value=supplier_name)
        ws_pr.cell(row=r_idx, column=5, value=category_name)
        ws_pr.cell(row=r_idx, column=6, value=business_model).alignment = Alignment(horizontal="center")

        c_s26 = ws_pr.cell(row=r_idx, column=7, value=p["2026"]["net_sales"])
        c_s26.number_format = "#,##0"
        c_cb = ws_pr.cell(row=r_idx, column=8, value=(p.get("contrib_2026", 0) / 100))
        c_cb.number_format = "0.00%"
        c_q26 = ws_pr.cell(row=r_idx, column=9, value=p["2026"]["qty"])
        c_q26.number_format = "#,##0"
        c_q25 = ws_pr.cell(row=r_idx, column=10, value=p["2025"]["qty"])
        c_q25.number_format = "#,##0"
        gq = p.get("growth_qty")
        c_gq = ws_pr.cell(row=r_idx, column=11, value=(gq / 100) if gq is not None else "-")
        if gq is not None: c_gq.number_format = "+0.00%;-0.00%;0.00%"

        for col_idx in range(1, 12):
            c = ws_pr.cell(row=r_idx, column=col_idx)
            c.border = thin_border
            if col_idx != 3: c.font = regular_font

    if products_list:
        last_pr_row = len(products_list) + 1
        tot_pr_row = last_pr_row + 1
        ws_pr.cell(row=tot_pr_row, column=1, value="").alignment = Alignment(horizontal="center")
        ws_pr.cell(row=tot_pr_row, column=2, value="").alignment = Alignment(horizontal="center")
        ws_pr.cell(row=tot_pr_row, column=3, value="TOTAL TOP 300 PRODUK").font = total_font
        ws_pr.cell(row=tot_pr_row, column=4, value="")
        ws_pr.cell(row=tot_pr_row, column=5, value="")
        ws_pr.cell(row=tot_pr_row, column=6, value="")
        ws_pr.cell(row=tot_pr_row, column=7, value=f"=SUM(G2:G{last_pr_row})").number_format = "#,##0"
        ws_pr.cell(row=tot_pr_row, column=8, value=f"=SUM(H2:H{last_pr_row})").number_format = "0.00%"
        ws_pr.cell(row=tot_pr_row, column=9, value=f"=SUM(I2:I{last_pr_row})").number_format = "#,##0"
        ws_pr.cell(row=tot_pr_row, column=10, value=f"=SUM(J2:J{last_pr_row})").number_format = "#,##0"
        ws_pr.cell(row=tot_pr_row, column=11, value=f"=(I{tot_pr_row}-J{tot_pr_row})/J{tot_pr_row}").number_format = "+0.00%;-0.00%;0.00%"
        for col_idx in range(1, 12):
            c = ws_pr.cell(row=tot_pr_row, column=col_idx)
            c.fill = total_fill
            c.font = total_font
            c.border = total_border

    # =========================================================================
    # 7. SHEET: TREN PENJUALAN HARIAN YOY
    # =========================================================================
    daily_chart_list = data.get("daily_chart", [])
    if daily_chart_list:
        ws_dy = wb.create_sheet(title="Tren Penjualan Harian")
        ws_dy.views.sheetView[0].showGridLines = True
        daily_headers = [
            "Periode / Tanggal", "Net Sales 2026", "Net Sales 2025",
            "Pertumbuhan YoY (%)", "Selisih Omset (+/-)",
        ]
        for col_num, h in enumerate(daily_headers, 1):
            cell = ws_dy.cell(row=1, column=col_num, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws_dy.row_dimensions[1].height = 24

        for r_idx, d_item in enumerate(daily_chart_list, 2):
            d_label = d_item.get("label") or d_item.get("date") or f"Hari ke-{r_idx-1}"
            s26 = d_item.get("2026", 0.0)
            s25 = d_item.get("2025", 0.0)

            ws_dy.cell(row=r_idx, column=1, value=d_label).font = bold_font
            c_s26 = ws_dy.cell(row=r_idx, column=2, value=s26)
            c_s26.number_format = "#,##0"
            c_s25 = ws_dy.cell(row=r_idx, column=3, value=s25)
            c_s25.number_format = "#,##0"

            c_gw = ws_dy.cell(row=r_idx, column=4, value=f"=(B{r_idx}-C{r_idx})/C{r_idx}" if s25 > 0 else "-")
            if s25 > 0: c_gw.number_format = "+0.00%;-0.00%;0.00%"
            c_df = ws_dy.cell(row=r_idx, column=5, value=f"=B{r_idx}-C{r_idx}")
            c_df.number_format = "+#,##0;-#,##0;0"

            for col_idx in range(1, 6):
                c = ws_dy.cell(row=r_idx, column=col_idx)
                c.border = thin_border
                if col_idx != 1: c.font = regular_font

        last_dy_row = len(daily_chart_list) + 1
        tot_dy_row = last_dy_row + 1
        ws_dy.cell(row=tot_dy_row, column=1, value="TOTAL KESELURUHAN").font = total_font
        ws_dy.cell(row=tot_dy_row, column=2, value=f"=SUM(B2:B{last_dy_row})").number_format = "#,##0"
        ws_dy.cell(row=tot_dy_row, column=3, value=f"=SUM(C2:C{last_dy_row})").number_format = "#,##0"
        ws_dy.cell(row=tot_dy_row, column=4, value=f"=(B{tot_dy_row}-C{tot_dy_row})/C{tot_dy_row}").number_format = "+0.00%;-0.00%;0.00%"
        ws_dy.cell(row=tot_dy_row, column=5, value=f"=B{tot_dy_row}-C{tot_dy_row}").number_format = "+#,##0;-#,##0;0"
        for col_idx in range(1, 6):
            c = ws_dy.cell(row=tot_dy_row, column=col_idx)
            c.fill = total_fill
            c.font = total_font
            c.border = total_border

    # =========================================================================
    # 8. SHEET: PERFORMA JAM OPERASIONAL
    # =========================================================================
    hourly_list = data.get("hourly_chart", [])
    if hourly_list:
        ws_hr = wb.create_sheet(title="Performa Jam")
        ws_hr.views.sheetView[0].showGridLines = True
        hourly_headers = [
            "Jam Operasional", "Net Sales 2026", "Transaksi 2026",
            "Net Sales 2025", "Transaksi 2025", "Growth Sales (%)", "Growth Tx (%)",
        ]
        for col_num, h in enumerate(hourly_headers, 1):
            cell = ws_hr.cell(row=1, column=col_num, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws_hr.row_dimensions[1].height = 24

        for r_idx, hr in enumerate(hourly_list, 2):
            ws_hr.cell(row=r_idx, column=1, value=f"Pukul {hr['hour']}:00").alignment = Alignment(horizontal="center")
            c_s26 = ws_hr.cell(row=r_idx, column=2, value=hr["net_sales_2026"])
            c_s26.number_format = "#,##0"
            c_t26 = ws_hr.cell(row=r_idx, column=3, value=hr["transactions_2026"])
            c_t26.number_format = "#,##0"
            c_s25 = ws_hr.cell(row=r_idx, column=4, value=hr["net_sales_2025"])
            c_s25.number_format = "#,##0"
            c_t25 = ws_hr.cell(row=r_idx, column=5, value=hr["transactions_2025"])
            c_t25.number_format = "#,##0"

            s25_val = hr["net_sales_2025"]
            t25_val = hr["transactions_2025"]
            c_gs = ws_hr.cell(row=r_idx, column=6, value=f"=(B{r_idx}-D{r_idx})/D{r_idx}" if s25_val > 0 else "-")
            if s25_val > 0: c_gs.number_format = "+0.00%;-0.00%;0.00%"
            c_gt = ws_hr.cell(row=r_idx, column=7, value=f"=(C{r_idx}-E{r_idx})/E{r_idx}" if t25_val > 0 else "-")
            if t25_val > 0: c_gt.number_format = "+0.00%;-0.00%;0.00%"

            for col_idx in range(1, 8):
                c = ws_hr.cell(row=r_idx, column=col_idx)
                c.border = thin_border
                c.font = regular_font

        last_hr_row = len(hourly_list) + 1
        tot_hr_row = last_hr_row + 1
        ws_hr.cell(row=tot_hr_row, column=1, value="TOTAL KESELURUHAN").font = total_font
        ws_hr.cell(row=tot_hr_row, column=2, value=f"=SUM(B2:B{last_hr_row})").number_format = "#,##0"
        ws_hr.cell(row=tot_hr_row, column=3, value=f"=SUM(C2:C{last_hr_row})").number_format = "#,##0"
        ws_hr.cell(row=tot_hr_row, column=4, value=f"=SUM(D2:D{last_hr_row})").number_format = "#,##0"
        ws_hr.cell(row=tot_hr_row, column=5, value=f"=SUM(E2:E{last_hr_row})").number_format = "#,##0"
        ws_hr.cell(row=tot_hr_row, column=6, value=f"=(B{tot_hr_row}-D{tot_hr_row})/D{tot_hr_row}").number_format = "+0.00%;-0.00%;0.00%"
        ws_hr.cell(row=tot_hr_row, column=7, value=f"=(C{tot_hr_row}-E{tot_hr_row})/E{tot_hr_row}").number_format = "+0.00%;-0.00%;0.00%"
        for col_idx in range(1, 8):
            c = ws_hr.cell(row=tot_hr_row, column=col_idx)
            c.fill = total_fill
            c.font = total_font
            c.border = total_border

    # =========================================================================
    # 9. SHEET: RADAR SKU BARU (UNMAPPED PRODUCTS)
    # =========================================================================
    unmapped_list = data.get("unmapped_skus", [])
    if unmapped_list:
        ws_un = wb.create_sheet(title="Radar SKU Baru")
        ws_un.views.sheetView[0].showGridLines = True
        un_headers = [
            "No", "Nama Produk Belum Terdaftar di Master", "Qty Terjual",
            "Net Sales 2026", "Konter / Outlet Terdeteksi",
        ]
        for col_num, h in enumerate(un_headers, 1):
            cell = ws_un.cell(row=1, column=col_num, value=h)
            cell.fill = header_accent
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws_un.row_dimensions[1].height = 24

        for r_idx, u in enumerate(unmapped_list, 2):
            ws_un.cell(row=r_idx, column=1, value=r_idx - 1).alignment = Alignment(horizontal="center")
            ws_un.cell(row=r_idx, column=2, value=u["name"]).font = bold_font
            c_q = ws_un.cell(row=r_idx, column=3, value=u.get("qty", 0))
            c_q.number_format = "#,##0"
            c_s = ws_un.cell(row=r_idx, column=4, value=u.get("net_sales", 0))
            c_s.number_format = "#,##0"
            outlets_str = ", ".join(u["outlets"]) if isinstance(u.get("outlets"), list) else str(u.get("outlets", "-"))
            ws_un.cell(row=r_idx, column=5, value=outlets_str)

            for col_idx in range(1, 6):
                c = ws_un.cell(row=r_idx, column=col_idx)
                c.border = thin_border
                if col_idx != 2: c.font = regular_font

    # =========================================================================
    # ERGONOMICS: FREEZE PANES & ACCURATE COLUMN WIDTH BUDGETING
    # =========================================================================
    for ws in wb.worksheets:
        # Enable grid lines
        ws.views.sheetView[0].showGridLines = True

        # Freeze headers so columns stay visible during scroll
        if ws.title == "Ringkasan KPI":
            ws.freeze_panes = "A8"
        else:
            ws.freeze_panes = "A2"

        # Calculate optimal column widths (safely skipping merged narrative cell A5)
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            vals = []
            for cell in col:
                # Exclude merged narrative block in Ringkasan KPI to prevent 340-char blowout
                if ws.title == "Ringkasan KPI" and col_letter == "A" and cell.row in (4, 5):
                    continue
                if cell.value is not None:
                    # Limit formula strings from skewing width
                    s_val = str(cell.value)
                    if s_val.startswith("="):
                        vals.append(12)
                    else:
                        vals.append(len(s_val))
            max_len = max(vals) if vals else 10
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 13), 50)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def write_report(html: str) -> None:
    """Optionally persist a static snapshot.

    The interactive dashboard includes its source data for client-side filters,
    so saving every rendered page produced 90 MB HTML files. Static snapshots
    are therefore opt-in via WRITE_STATIC_REPORT=1.
    """
    if os.environ.get("WRITE_STATIC_REPORT") != "1":
        return
    REPORT_DIR.mkdir(exist_ok=True)
    (REPORT_DIR / "sales_dashboard.html").write_text(html, encoding="utf-8")


@app.after_request
def compress_large_response(response):
    """Reduce transfer size for the data-heavy legacy client without UI changes."""
    accepts_gzip = "gzip" in request.headers.get("Accept-Encoding", "").lower()
    compressible = response.mimetype in {"text/html", "application/json"}
    if (
        response.direct_passthrough
        or not accepts_gzip
        or not compressible
        or response.headers.get("Content-Encoding")
    ):
        return response

    payload = response.get_data()
    if len(payload) < 1024:
        return response
    response.set_data(gzip.compress(payload, compresslevel=6))
    response.headers["Content-Encoding"] = "gzip"
    response.headers["Vary"] = "Accept-Encoding"
    return response


@app.route("/")
def index():
    month = request.args.get("month") or None
    date = request.args.get("date") or None
    start_date = request.args.get("start_date") or None
    end_date = request.args.get("end_date") or None
    outlet = request.args.get("outlet") or None
    area = request.args.get("area") or None
    jenis = request.args.get("jenis") or None
    html = render_template(
        "index.html",
        data=build_dashboard(
            read_sales(),
            month,
            date,
            outlet,
            area,
            start_date=start_date,
            end_date=end_date,
            jenis=jenis,
        ),
        selected_month=month or "",
        selected_date=date or "",
        selected_outlet=outlet or "",
        selected_area=area or "",
        selected_jenis=jenis or "",
    )
    write_report(html)
    return html


@app.route("/download")
def download_excel():
    month = request.args.get("month") or None
    date = request.args.get("date") or None
    start_date = request.args.get("start_date") or None
    end_date = request.args.get("end_date") or None
    outlet = request.args.get("outlet") or None
    area = request.args.get("area") or None
    supplier = request.args.get("supplier") or None
    category = request.args.get("category") or None
    jenis = request.args.get("jenis") or None

    all_rows = read_sales()
    dashboard_data = build_dashboard(
        all_rows,
        month,
        date,
        outlet,
        area,
        include_raw=False,
        start_date=start_date,
        end_date=end_date,
        supplier=supplier,
        category=category,
        jenis=jenis,
    )

    filter_parts = []
    if month: filter_parts.append(f"Bulan_{month}")
    if start_date or end_date:
        filter_parts.append(f"Tgl_{start_date or 'Awal'}_sd_{end_date or 'Akhir'}")
    elif date:
        filter_parts.append(f"Tgl_{date}")
    if area: filter_parts.append(f"Area_{area}")
    if outlet: filter_parts.append(f"Outlet_{outlet}")
    if supplier: filter_parts.append(f"Supplier_{supplier}")
    if category: filter_parts.append(f"Kategori_{category}")
    if jenis: filter_parts.append(f"Model_{jenis}")

    filter_desc = " - ".join(filter_parts) if filter_parts else "Semua Data"
    file_name = f"Laporan_Sales_{'_'.join(filter_parts) if filter_parts else 'Semua'}.xlsx"

    excel_buffer = build_excel_report(dashboard_data, filter_desc)
    return send_file(
        excel_buffer,
        as_attachment=True,
        download_name=file_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )



@app.route("/api/dashboard")
def dashboard_api():
    return jsonify(
        build_dashboard(
            read_sales(),
            request.args.get("month") or None,
            request.args.get("date") or None,
            request.args.get("outlet") or None,
            request.args.get("area") or None,
            include_raw=False,
            start_date=request.args.get("start_date") or None,
            end_date=request.args.get("end_date") or None,
            supplier=request.args.get("supplier") or None,
            category=request.args.get("category") or None,
            jenis=request.args.get("jenis") or None,
        )
    )


# ─────────────────────────────────────────────────────────────────────────────
# ADMIN: Upload File Excel Sales & Target
# ─────────────────────────────────────────────────────────────────────────────
ADMIN_PIN = os.environ.get("ADMIN_PIN", "1234")
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}
app.config.setdefault("MAX_CONTENT_LENGTH", 64 * 1024 * 1024)  # 64 MB


def _allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def _invalidate_sales_cache() -> None:
    global _sales_cache, _sales_cache_signature, _sales_cache_mapping_signature
    global _budget_cache, _budget_cache_signature
    global _lookup_cache, _lookup_cache_mtime
    with _sales_cache_lock:
        _sales_cache = None
        _sales_cache_signature = ()
        _sales_cache_mapping_signature = None
    with _budget_cache_lock:
        _budget_cache = None
        _budget_cache_signature = None
    with _lookup_cache_lock:
        _lookup_cache = None
        _lookup_cache_mtime = None


@app.route("/upload", methods=["GET", "POST"])
def upload_file():
    """
    GET  -> tampilkan halaman upload admin
    POST -> terima file Excel, validasi PIN, simpan ke folder yang sesuai,
           lalu invalidate cache agar data langsung ter-refresh.
    """
    if request.method == "GET":
        return render_template("upload.html", message=None, error=None)

    # Validasi PIN
    pin = (request.form.get("pin") or "").strip()
    if pin != ADMIN_PIN:
        return render_template("upload.html", message=None, error="❌ PIN salah. Akses ditolak."), 403

    # Validasi file yang di-upload
    files = request.files.getlist("files")
    if not files or all(f.filename == "" for f in files):
        return render_template("upload.html", message=None, error="❌ Tidak ada file yang dipilih."), 400

    saved_files: list[str] = []
    skipped_files: list[str] = []

    for f in files:
        if not f or f.filename == "":
            continue
        fname = secure_filename(f.filename)
        if not _allowed_file(fname):
            skipped_files.append(f"{fname} (bukan .xlsx/.xls)")
            continue

        category = request.form.get("category", "auto")
        raw_name = f.filename.lower()
        if category == "2025":
            dest_dir = SALES_ROOT_DIR / "sales detail 2025"
        elif category == "2026":
            dest_dir = SALES_ROOT_DIR / "sales detail 2026"
        elif category == "target":
            dest_dir = BUDGET_FILE.parent
        else:
            if "target" in raw_name or "budget" in raw_name:
                dest_dir = BUDGET_FILE.parent
            elif "2025" in raw_name:
                dest_dir = SALES_ROOT_DIR / "sales detail 2025"
            else:
                dest_dir = SALES_ROOT_DIR / "sales detail 2026"

        dest_dir.mkdir(parents=True, exist_ok=True)
        dest_path = dest_dir / fname

        f.save(str(dest_path))
        saved_files.append(f"{fname} → {dest_dir.name}/")

    if not saved_files:
        return render_template(
            "upload.html",
            message=None,
            error="❌ Tidak ada file valid yang berhasil disimpan. " + (", ".join(skipped_files) if skipped_files else ""),
        ), 400

    # Invalidate cache
    _invalidate_sales_cache()

    # Re-render static report
    try:
        updated_rows = read_sales()
        dashboard_data = build_dashboard(updated_rows)
        with app.app_context():
            write_report(
                render_template(
                    "index.html",
                    data=dashboard_data,
                    selected_month="",
                    selected_date="",
                    selected_outlet="",
                    selected_area="",
                )
            )
    except Exception:
        pass

    msg_parts = [f"✅ {len(saved_files)} file berhasil disimpan:"]
    for s in saved_files:
        msg_parts.append(f"  • {s}")
    if skipped_files:
        msg_parts.append(f"⚠️ Dilewati: {', '.join(skipped_files)}")
    msg_parts.append("🔄 Cache data di-reset. Dashboard sudah ter-refresh otomatis!")

    return render_template("upload.html", message="\n".join(msg_parts), error=None)


@app.route("/api/clear-cache", methods=["POST"])
def api_clear_cache():
    body = request.get_json(silent=True) or {}
    if str(body.get("pin", "")).strip() != ADMIN_PIN:
        return jsonify({"ok": False, "error": "PIN salah"}), 403
    _invalidate_sales_cache()
    return jsonify({"ok": True, "message": "Cache di-reset. Data akan di-reload pada request berikutnya."})


@app.route("/api/save-visitor", methods=["POST"])
def api_save_visitor():
    """Endpoint untuk input/update data pengunjung harian per unit wahana."""
    body = request.get_json(silent=True)
    if not body:
        # Fallback to form-data if submitted via regular form
        body = request.form

    pin = str(body.get("pin", "")).strip()
    if pin != ADMIN_PIN:
        return jsonify({"ok": False, "error": "PIN Admin salah. Akses ditolak."}), 403

    entry_date = str(body.get("date", "")).strip()
    unit = str(body.get("unit", "")).strip()
    count_val = body.get("count", 0)

    if not entry_date:
        return jsonify({"ok": False, "error": "Tanggal harus diisi."}), 400
    if not unit:
        return jsonify({"ok": False, "error": "Unit wahana harus dipilih."}), 400

    try:
        count_int = int(count_val)
        if count_int < 0:
            return jsonify({"ok": False, "error": "Jumlah pengunjung tidak boleh negatif."}), 400
    except (ValueError, TypeError):
        return jsonify({"ok": False, "error": "Format jumlah pengunjung harus berupa angka bulat."}), 400

    try:
        success = save_visitor_actual(entry_date, unit, count_int)
        if not success:
            return jsonify({"ok": False, "error": "Gagal menyimpan ke file budget."}), 500

        # Re-render static report
        try:
            updated_rows = read_sales()
            dashboard_data = build_dashboard(updated_rows)
            with app.app_context():
                write_report(
                    render_template(
                        "index.html",
                        data=dashboard_data,
                        selected_month="",
                        selected_date="",
                        selected_outlet="",
                        selected_area="",
                    )
                )
        except Exception:
            pass

        return jsonify({
            "ok": True,
            "message": f"✅ Berhasil menyimpan {count_int:,} pengunjung untuk {unit} pada tanggal {entry_date}."
        })
    except Exception as e:
        return jsonify({"ok": False, "error": f"Gagal menyimpan pengunjung: {str(e)}"}), 500


AI_KEY_FILE = BASE_DIR / ".gemini_key"

def get_gemini_api_key():
    key = os.environ.get("GEMINI_API_KEY")
    if key and key.strip():
        return key.strip()
    if AI_KEY_FILE.exists():
        try:
            with open(AI_KEY_FILE, "r", encoding="utf-8") as f:
                k = f.read().strip()
                if k:
                    return k
        except Exception:
            pass
    return None

@app.route("/api/ai-status", methods=["GET"])
def api_ai_status():
    key = get_gemini_api_key()
    masked = (key[:4] + "..." + key[-4:]) if key and len(key) > 8 else ("Set" if key else "")
    return jsonify({"has_key": bool(key), "masked_key": masked})

@app.route("/api/save-ai-key", methods=["POST"])
def api_save_ai_key():
    try:
        data = request.get_json(force=True) or {}
        key = data.get("key", "").strip()
        with open(AI_KEY_FILE, "w", encoding="utf-8") as f:
            f.write(key)
        return jsonify({"ok": True, "has_key": bool(key)})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/api/ask-ai", methods=["POST"])
def api_ask_ai():
    import json
    import urllib.request
    import urllib.error

    try:
        data = request.get_json(force=True) or {}
        user_prompt = (data.get("prompt") or data.get("question") or "").strip()
        if not user_prompt:
            return jsonify({"ok": False, "answer": "Mohon ketikkan pertanyaan Anda seputar data penjualan Ancol Store."})

        # Smart Dimension Extractor (Date, Month, Area, Outlet, Hour)
        import re
        from datetime import datetime
        p_lower = user_prompt.lower()

        MONTH_MAP = {
            "januari": "01", "jan": "01",
            "februari": "02", "feb": "02",
            "maret": "03", "mar": "03",
            "april": "04", "apr": "04",
            "mei": "05", "may": "05",
            "juni": "06", "jun": "06",
            "juli": "07", "jul": "07",
            "agustus": "08", "agu": "08", "aug": "08",
            "september": "09", "sep": "09",
            "oktober": "10", "okt": "10",
            "november": "11", "nov": "11",
            "desember": "12", "des": "12",
        }

        AREA_MAP = {
            "DUFAN": ["dufan"],
            "AWAPARK": ["atlantis", "awapark", "awa park", "kolam renang", "waterpark"],
            "BEACHPARK": ["beachpark", "beach park", "pantai", "beach", "pasir putih"],
        }

        # 1. Month detection
        detected_m_num = None
        for k, m_num in MONTH_MAP.items():
            if re.search(rf"\b{k}\b", p_lower):
                detected_m_num = m_num
                break
        detected_month = f"2026-{detected_m_num}" if detected_m_num else None

        # 2. Date detection
        rows_cache = read_sales()
        all_dates_2026 = sorted({r["date"] for r in rows_cache if r.get("year") == 2026})
        latest_date_in_db = all_dates_2026[-1] if all_dates_2026 else "2026-09-07"
        default_month_num = detected_m_num or latest_date_in_db[5:7]

        detected_date = None
        # Pattern A: "tanggal 7 september" or "7 sep" or "tgl 7"
        m_day = re.search(r"(?:tanggal|tgl)?\s*(\b[1-9]\b|\b[12][0-9]\b|\b3[01]\b)\s*(januari|februari|maret|april|mei|juni|juli|agustus|september|oktober|november|desember|jan|feb|mar|apr|may|jun|jul|agu|aug|sep|okt|oct|nov|des|dec)", p_lower)
        if m_day:
            d_val = int(m_day.group(1))
            m_word = m_day.group(2)
            m_num = MONTH_MAP.get(m_word, default_month_num)
            cand_date = f"2026-{m_num}-{d_val:02d}"
            if cand_date in all_dates_2026:
                detected_date = cand_date
        else:
            m_day_only = re.search(r"(?:tanggal|tgl)\s*(\b[1-9]\b|\b[12][0-9]\b|\b3[01]\b)", p_lower)
            if m_day_only:
                d_val = int(m_day_only.group(1))
                cand_date = f"2026-{default_month_num}-{d_val:02d}"
                if cand_date in all_dates_2026:
                    detected_date = cand_date

        # 2b. Date Range Detection (e.g. "1 sampai 10 agustus", "10-15 agustus", "antara tanggal 5 dan 20")
        detected_start_date = None
        detected_end_date = None
        m_range = re.search(r"(?:tanggal|tgl)?\s*(\b[1-9]\b|\b[12][0-9]\b|\b3[01]\b)\s*(?:sampai|s/d|sd|hingga|-)\s*(\b[1-9]\b|\b[12][0-9]\b|\b3[01]\b)\s*(januari|februari|maret|april|mei|juni|juli|agustus|september|oktober|november|desember|jan|feb|mar|apr|may|jun|jul|agu|aug|sep|okt|oct|nov|des|dec)?", p_lower)
        if m_range:
            d1_val = int(m_range.group(1))
            d2_val = int(m_range.group(2))
            m_range_word = m_range.group(3)
            m_range_num = MONTH_MAP.get(m_range_word, default_month_num) if m_range_word else default_month_num
            if d1_val <= d2_val:
                detected_start_date = f"2026-{m_range_num}-{d1_val:02d}"
                detected_end_date = f"2026-{m_range_num}-{d2_val:02d}"

        # 3. Area detection
        detected_area = None
        for area_code, kw_list in AREA_MAP.items():
            for kw in kw_list:
                if re.search(rf"\b{kw}\b", p_lower):
                    detected_area = area_code
                    break
            if detected_area:
                break

        # 4. Outlet detection
        all_outlets_list = sorted({r["outlet"] for r in rows_cache if r.get("outlet")})
        detected_outlet = None
        OUTLET_ALIASES = {
            "arung jeram": "DFAR Dufan Arung Jeram",
            "kereta misteri": "DFKE Dufan Kereta Misteri",
            "ice age": "DFIC Dufan Ice Age",
            "dufan induk lama": "DFIL Dufan Induk Lama",
            "dufan induk": "DFIN Dufan Induk",
            "wwn": "DFWW Dufan WWN",
            "treasure": "DFTR Dufan Treasure",
            "galactica": "DFGA Dufan Galactica",
            "sea world induk": "SWIN Sea World Induk",
            "sea world": "SWIN Sea World Induk",
            "seaworld": "SWIN Sea World Induk",
            "ocean bakery": "SWOB Sea World Ocean Bakery",
            "tug of war": "SWTG Sea World Tug of War",
            "ocean dream": "ODIN Ocean Dream Induk",
            "atlantis induk": "AWIN Atlantis Induk",
            "atlantis plaza": "AWCP Atlantis Plaza",
        }
        for alias, full_out in OUTLET_ALIASES.items():
            if alias in p_lower:
                detected_outlet = full_out
                break
        if not detected_outlet:
            for out_name in all_outlets_list:
                code_low = out_name.split()[0].lower()
                if re.search(rf"\b{code_low}\b", p_lower):
                    detected_outlet = out_name
                    break

        # 4b. Supplier Detection across all 86 suppliers
        detected_supplier = None
        lookup_data = load_data_lookup()
        compact_lk = lookup_data.get("compact_lookup", {})
        all_unique_suppliers = sorted({vals[1] for vals in compact_lk.values() if vals and len(vals) > 1 and vals[1] and vals[1] != "LAINNYA"})

        for supp in all_unique_suppliers:
            s_low = supp.lower()
            if len(s_low) >= 4 and s_low in p_lower:
                detected_supplier = supp
                break
        if not detected_supplier:
            for supp in all_unique_suppliers:
                s_clean = re.sub(r"^(cv|pt|ud|tb)\b\.?\s*", "", supp.lower()).strip()
                if len(s_clean) >= 4 and s_clean in p_lower:
                    detected_supplier = supp
                    break
        if not detected_supplier:
            tokens = set(re.findall(r"\w+", p_lower))
            for supp in all_unique_suppliers:
                s_words = [w for w in re.findall(r"\w+", supp.lower()) if len(w) >= 4 and w not in ("cv", "pt", "tb", "ud", "jaya", "indah", "lestari", "mandiri")]
                if s_words and all(w in tokens for w in s_words):
                    detected_supplier = supp
                    break

        # 4c. Category Detection across all 32 categories
        detected_category = None
        all_unique_categories = sorted({vals[2] for vals in compact_lk.values() if vals and len(vals) > 2 and vals[2] and vals[2] != "LAINNYA"})

        CATEGORY_ALIASES = {
            "AKSESORIES TUBUH": ["aksesoris tubuh", "aksesoris", "aksesori", "gelang", "kalung"],
            "ALAT TULIS": ["alat tulis", "atk", "pulpen", "pensil", "buku tulis", "notes"],
            "ATASAN ANAK": ["atasan anak", "baju anak", "kaos anak"],
            "BANDO": ["bando", "headband"],
            "BONEKA": ["boneka", "doll", "plush", "boneka lumba", "boneka otter"],
            "CELANA": ["celana", "pants"],
            "DRESS ANAK": ["dress anak", "dress", "gaun"],
            "FOTO / FRAME": ["foto", "frame", "bingkai foto", "pigura"],
            "GANT. KUNCI / PIN": ["gantungan kunci", "gant. kunci", "ganci", "pin", "gantungan"],
            "HANDUK": ["handuk", "towel"],
            "ICE CREAM": ["es krim", "ice cream"],
            "KIPAS": ["kipas", "fan"],
            "MAGNET": ["magnet", "tempelan kulkas"],
            "MAINAN": ["mainan", "toys", "toy", "balon"],
            "MAKANAN": ["makanan", "snack", "cemilan", "chiki"],
            "MINUMAN": ["minuman", "drink", "air mineral", "prima", "teh botol", "jus"],
            "MUG / GELAS/TUMBLER": ["mug", "gelas", "tumbler", "botol minum"],
            "OBAT": ["obat", "tolak angin", "minyak kayu putih", "paracetamol"],
            "PAKAIAN": ["pakaian", "kaos", "baju", "t-shirt", "jersey"],
            "PAKAIAN RENANG": ["pakaian renang", "baju renang", "swimwear"],
            "PAYUNG": ["payung", "jas hujan", "umbrella", "raincoat"],
            "PEMBALUT/PAMPERS/TISSUE": ["tissue", "tisu", "pampers", "pembalut", "popok"],
            "PERLENGKAPAN RENANG": ["kacamata renang", "pelampung", "alat renang", "ban renang"],
            "PERMEN": ["permen", "candy", "lollipop"],
            "PLASTIK/PACKAGING": ["plastik", "kantong plastik", "packaging"],
            "SANDAL": ["sandal", "slippers"],
            "SETELAN ANAK": ["setelan anak", "one set anak"],
            "STIKER": ["stiker", "sticker"],
            "TAS / BACK PACK/ BAG": ["tas", "backpack", "tote bag", "ransel", "sling bag"],
            "TEMPAT PENSIL/DOMPET": ["tempat pensil", "kotak pensil", "dompet", "pouch"],
            "TOPI": ["topi", "cap", "hat", "bucket hat"],
            "UMUM": ["umum", "general"],
        }
        for cat_name, aliases in CATEGORY_ALIASES.items():
            for alias in aliases:
                if re.search(rf"\b{re.escape(alias)}\b", p_lower):
                    detected_category = cat_name
                    break
            if detected_category:
                break
        if not detected_category:
            for c in all_unique_categories:
                if c.lower() in p_lower:
                    detected_category = c
                    break

        # 4d. Invoice / Struk / Diskon Query Detection
        is_invoice_query = bool(re.search(r"\b(?:struk|invoice|transaksi terbesar|transaksi tertinggi|nilai struk|diskon|keranjang|basket size)\b", p_lower))

        # Multi-turn Chat History
        history = data.get("history", [])
        if not isinstance(history, list):
            history = []

        # Context Carryover from previous questions in conversation
        if history:
            prev_user_prompts = [
                h.get("text", "") for h in history 
                if h.get("role") in ["user", "human"] and h.get("text")
            ]
            if prev_user_prompts:
                last_user_p = prev_user_prompts[-1].lower()
                if not detected_outlet:
                    for alias, full_out in OUTLET_ALIASES.items():
                        if alias in last_user_p:
                            detected_outlet = full_out
                            break
                    if not detected_outlet:
                        for out_name in all_outlets_list:
                            code_low = out_name.split()[0].lower()
                            if re.search(rf"\b{code_low}\b", last_user_p):
                                detected_outlet = out_name
                                break
                if not detected_area:
                    for area_code, kw_list in AREA_MAP.items():
                        for kw in kw_list:
                            if re.search(rf"\b{kw}\b", last_user_p):
                                detected_area = area_code
                                break
                        if detected_area:
                            break
                if not detected_m_num and not detected_date:
                    for k, m_num in MONTH_MAP.items():
                        if re.search(rf"\b{k}\b", last_user_p):
                            detected_m_num = m_num
                            detected_month = f"2026-{m_num}"
                            break
                if not detected_supplier:
                    for supp in all_unique_suppliers:
                        if supp.lower() in last_user_p:
                            detected_supplier = supp
                            break
                if not detected_category:
                    for cat_name, aliases in CATEGORY_ALIASES.items():
                        if any(re.search(rf"\b{re.escape(a)}\b", last_user_p) for a in aliases):
                            detected_category = cat_name
                            break

        # 5. Hour detection (e.g. "jam 15", "jam 3 sore", "pukul 16")
        detected_hour = None
        m_hour = re.search(r"(?:jam|pukul)\s*(\b[0-9]\b|\b1[0-9]\b|\b2[0-3]\b)", p_lower)
        if m_hour:
            h_val = int(m_hour.group(1))
            if any(s in p_lower for s in ["sore", "malam", "pm"]) and 1 <= h_val <= 11:
                h_val += 12
            detected_hour = f"{h_val:02d}"

        # 6. Comparative Head-to-Head Detection (e.g. "Dufan vs Atlantis", "Agustus vs September", "Weekend vs Weekday")
        is_comparative = bool(re.search(r"\b(?:vs|versus|dibanding|dibandingkan|bandingkan|banding|lawan|lebih (?:bagus|rame|tinggi|besar|banyak) mana|mana yang lebih)\b", p_lower))
        comp_areas = []
        for area_code, kw_list in AREA_MAP.items():
            for kw in kw_list:
                if re.search(rf"\b{kw}\b", p_lower):
                    if area_code not in comp_areas:
                        comp_areas.append(area_code)
                    break

        comp_months = []
        for k, m_num in MONTH_MAP.items():
            if re.search(rf"\b{k}\b", p_lower):
                cand_m = f"2026-{m_num}"
                if cand_m not in comp_months:
                    comp_months.append(cand_m)

        is_wknd_comp = (
            any(w in p_lower for w in ["weekend", "akhir pekan", "sabtu minggu"]) and 
            any(w in p_lower for w in ["weekday", "hari kerja", "senin jumat"])
        )

        comparative_text = ""
        if is_comparative or len(comp_areas) >= 2 or len(comp_months) >= 2 or is_wknd_comp:
            comp_blocks = []
            if len(comp_areas) >= 2:
                c_m = detected_month or (detected_date[:7] if detected_date else None)
                area_a, area_b = comp_areas[0], comp_areas[1]
                dash_a = build_dashboard(rows_cache, month=c_m, area=area_a, include_raw=False)
                dash_b = build_dashboard(rows_cache, month=c_m, area=area_b, include_raw=False)
                s_a, s_b = dash_a["summary"]["2026"], dash_b["summary"]["2026"]
                diff_sales = s_a["net_sales"] - s_b["net_sales"]
                pct_sales = ((s_a["net_sales"] / s_b["net_sales"]) - 1) * 100 if s_b["net_sales"] > 0 else 0
                winner = area_a if diff_sales >= 0 else area_b
                comp_blocks.append(
                    f"DATA KOMPARATIF HEAD-TO-HEAD ANTAR AREA ({area_a} VS {area_b}):\n"
                    f"- Area {area_a}: Omset Rp {s_a['net_sales']:,.0f} | Transaksi {s_a['transactions']:,} struk | ATV Rp {s_a['atv']:,.0f} | Qty {s_a['qty']:,} pcs\n"
                    f"- Area {area_b}: Omset Rp {s_b['net_sales']:,.0f} | Transaksi {s_b['transactions']:,} struk | ATV Rp {s_b['atv']:,.0f} | Qty {s_b['qty']:,} pcs\n"
                    f"- Perbedaan Omset: Rp {abs(diff_sales):,.0f} ({pct_sales:+.1f}% {'lebih tinggi' if diff_sales >= 0 else 'lebih rendah'} di {area_a})\n"
                    f"- Rekomendasi: Unggul di {winner}. Analisis apakah perbedaan karena volume pengunjung atau nilai belanja per struk (ATV)."
                )
            elif len(comp_months) >= 2:
                m_a, m_b = comp_months[0], comp_months[1]
                dash_a = build_dashboard(rows_cache, month=m_a, area=detected_area, outlet=detected_outlet, include_raw=False)
                dash_b = build_dashboard(rows_cache, month=m_b, area=detected_area, outlet=detected_outlet, include_raw=False)
                s_a, s_b = dash_a["summary"]["2026"], dash_b["summary"]["2026"]
                diff_sales = s_a["net_sales"] - s_b["net_sales"]
                pct_sales = ((s_a["net_sales"] / s_b["net_sales"]) - 1) * 100 if s_b["net_sales"] > 0 else 0
                name_a = MONTH_NAMES.get(m_a[5:], m_a)
                name_b = MONTH_NAMES.get(m_b[5:], m_b)
                comp_blocks.append(
                    f"DATA KOMPARATIF HEAD-TO-HEAD ANTAR BULAN ({name_a} VS {name_b}):\n"
                    f"- Bulan {name_a}: Omset Rp {s_a['net_sales']:,.0f} | Transaksi {s_a['transactions']:,} struk | ATV Rp {s_a['atv']:,.0f}\n"
                    f"- Bulan {name_b}: Omset Rp {s_b['net_sales']:,.0f} | Transaksi {s_b['transactions']:,} struk | ATV Rp {s_b['atv']:,.0f}\n"
                    f"- Selisih Performa: Rp {abs(diff_sales):,.0f} ({pct_sales:+.1f}%)"
                )
            elif is_wknd_comp:
                wknd_sales, wknd_tx, wknd_qty = 0, 0, 0
                wkdy_sales, wkdy_tx, wkdy_qty = 0, 0, 0
                c_m = detected_month or None
                for r in rows_cache:
                    if r.get("year") == 2026:
                        if c_m and r.get("month") != c_m:
                            continue
                        if detected_area and r.get("area") != detected_area:
                            continue
                        try:
                            dt = datetime.strptime(r["date"], "%Y-%m-%d")
                            if dt.weekday() in (5, 6):
                                wknd_sales += r["sales"]
                                wknd_tx += r["transactions"]
                                wknd_qty += r["qty"]
                            else:
                                wkdy_sales += r["sales"]
                                wkdy_tx += r["transactions"]
                                wkdy_qty += r["qty"]
                        except Exception:
                            pass
                wknd_atv = (wknd_sales / wknd_tx) if wknd_tx > 0 else 0
                wkdy_atv = (wkdy_sales / wkdy_tx) if wkdy_tx > 0 else 0
                tot_s = wknd_sales + wkdy_sales
                comp_blocks.append(
                    f"DATA KOMPARATIF WEEKEND VS WEEKDAY (AKHIR PEKAN VS HARI KERJA):\n"
                    f"- Weekend (Sabtu - Minggu): Omset Rp {wknd_sales:,.0f} | Transaksi {wknd_tx:,} struk | ATV Rp {wknd_atv:,.0f} | Qty {wknd_qty:,} pcs\n"
                    f"- Weekday (Senin - Jumat): Omset Rp {wkdy_sales:,.0f} | Transaksi {wkdy_tx:,} struk | ATV Rp {wkdy_atv:,.0f} | Qty {wkdy_qty:,} pcs\n"
                    f"- Porsi Penjualan: Weekend menyumbang {(wknd_sales / tot_s * 100) if tot_s > 0 else 0:.1f}% total omset"
                )
            if comp_blocks:
                comparative_text = "\n" + "\n\n".join(comp_blocks) + "\n"

        # 7. Calendar, Holidays, & Future Planning Intelligence
        is_future_or_holiday_query = any(k in p_lower for k in [
            "desember", "natal", "tahun baru", "tanggal merah", "libur", "liburan", "cuti bersama", 
            "persiapan", "kalender", "oktober", "november", "strategi jualan", "long weekend", "hari apa aja", "hari apa"
        ])
        
        holiday_calendar_text = ""
        if is_future_or_holiday_query:
            holiday_calendar_text = (
                "\nINFORMASI RESMI KALENDER NASIONAL, TANGGAL MERAH, & PROYEKSI HIGH SEASON 2026:\n"
                "1. KALENDER TANGGAL MERAH & HARI LIBUR NASIONAL AKHIR TAHUN 2026:\n"
                "   - Jumat, 25 Desember 2026: Hari Raya Natal (Hari Libur Nasional Resmi).\n"
                "   - Kamis, 24 Desember 2026: Potensi Cuti Bersama Hari Raya Natal.\n"
                "   - Sabtu-Minggu, 26-27 Desember 2026: Akhir Pekan Terusan (Menciptakan Long Weekend 4 hari beruntun).\n"
                "   - Kamis, 31 Desember 2026: Malam Pergantian Tahun Baru (Puncak Kerumunan Pengunjung Ancol & Pesta Kembang Api).\n"
                "   - Jumat, 1 Januari 2027: Libur Tahun Baru 2027 Masehi.\n"
                "2. KARAKTERISTIK TREN OPERASIONAL THEME PARK ANCOL SAAT LIBURAN (HIGH SEASON):\n"
                "   - Lonjakan pengunjung diproyeksikan melonjak 250% s/d 350% dibanding hari reguler.\n"
                "   - Jam sibuk bergeser lebih awal: Pukul 11:00 hingga 18:30 WIB terjadi lonjakan antrean kasir terus-menerus.\n"
                "   - Barang paling dicari: Minuman dingin (Prima 600ml & jus), jas hujan/payung (Desember musim hujan), serta merchandise boneka maskot Ancol.\n"
                "3. ROADMAP PERSIAPAN OPERASIONAL TOKO:\n"
                "   - H-14 (10-15 Desember): Kunci pengadaan buffer stock gudang merchandise & minuman minimal 3x lipat rata-rata mingguan.\n"
                "   - H-7: Susun jadwal shift staf kasir ekstra & siapkan perangkat EDC / POS mobile bantuan untuk mengurai antrean.\n"
                "   - Hari H (24 Des - 1 Jan): Buka seluruh loket kasir tanpa jeda (istirahat bergilir), display produk impulsif minuman & jas hujan di jalur antrean kasir.\n"
            )

        # 8. Targeted SQL Analytics (God Mode Engine)
        supplier_specific_text = ""
        supplier_analytics_data = None
        if detected_supplier:
            try:
                s_conds = ["year = 2026", "UPPER(supplier) = ?"]
                s_params = [detected_supplier.upper()]
                if detected_month:
                    s_conds.append("month = ?")
                    s_params.append(detected_month)
                if detected_outlet:
                    s_conds.append("outlet = ?")
                    s_params.append(detected_outlet)
                elif detected_area:
                    s_conds.append("area = ?")
                    s_params.append(detected_area)
                s_where = " AND ".join(s_conds)

                supp_sum = execute_analytics_sql(f"""
                    SELECT 
                        supplier, jenis,
                        COUNT(DISTINCT invoice) as total_tx,
                        SUM(qty) as total_qty,
                        SUM(net_sales) as total_sales,
                        SUM(gross_profit) as total_profit,
                        ROUND(SUM(gross_profit) * 100.0 / NULLIF(SUM(net_sales), 0), 1) as margin_pct
                    FROM v_sales_analytics
                    WHERE {s_where}
                    GROUP BY supplier, jenis
                """, tuple(s_params))

                supp_prods = execute_analytics_sql(f"""
                    SELECT product, SUM(qty) as qty, SUM(net_sales) as net_sales
                    FROM v_sales_analytics
                    WHERE {s_where}
                    GROUP BY product
                    ORDER BY net_sales DESC
                    LIMIT 5
                """, tuple(s_params))

                supp_outs = execute_analytics_sql(f"""
                    SELECT outlet, SUM(net_sales) as net_sales, SUM(qty) as qty
                    FROM v_sales_analytics
                    WHERE {s_where}
                    GROUP BY outlet
                    ORDER BY net_sales DESC
                    LIMIT 3
                """, tuple(s_params))

                if supp_sum:
                    ss = supp_sum[0]
                    supplier_analytics_data = {"summary": ss, "products": supp_prods, "outlets": supp_outs}
                    p_lines = [f"  * {p['product']}: Rp {p['net_sales']:,.0f} ({p['qty']:.0f} pcs)" for p in supp_prods]
                    o_lines = [f"  * {o['outlet']}: Rp {o['net_sales']:,.0f} ({o['qty']:.0f} pcs)" for o in supp_outs]
                    supplier_specific_text = (
                        f"\nDATA SPESIFIK REKANAN / SUPPLIER: {detected_supplier}\n"
                        f"- Model Bisnis: {ss.get('jenis', 'DAGANGAN')}\n"
                        f"- Total Penjualan (Net Sales 2026): Rp {ss.get('total_sales', 0):,.0f}\n"
                        f"- Total Kuantitas Terjual: {ss.get('total_qty', 0):,.0f} pcs\n"
                        f"- Total Transaksi (Struk): {ss.get('total_tx', 0):,} transaksi\n"
                        f"- Estimasi Laba Kotor: Rp {ss.get('total_profit', 0):,.0f}\n"
                        f"- Gross Margin: {ss.get('margin_pct', 0):.1f}%\n"
                        f"- Top 5 Produk Terlaris:\n" + "\n".join(p_lines) + "\n"
                        f"- Top 3 Cabang Penjual Terbesar:\n" + "\n".join(o_lines) + "\n"
                    )
            except Exception:
                pass

        category_specific_text = ""
        category_analytics_data = None
        if detected_category:
            try:
                c_conds = ["year = 2026", "UPPER(category) = ?"]
                c_params = [detected_category.upper()]
                if detected_month:
                    c_conds.append("month = ?")
                    c_params.append(detected_month)
                if detected_outlet:
                    c_conds.append("outlet = ?")
                    c_params.append(detected_outlet)
                elif detected_area:
                    c_conds.append("area = ?")
                    c_params.append(detected_area)
                c_where = " AND ".join(c_conds)

                cat_sum = execute_analytics_sql(f"""
                    SELECT 
                        category,
                        COUNT(DISTINCT invoice) as total_tx,
                        SUM(qty) as total_qty,
                        SUM(net_sales) as total_sales,
                        SUM(gross_profit) as total_profit,
                        ROUND(SUM(gross_profit) * 100.0 / NULLIF(SUM(net_sales), 0), 1) as margin_pct
                    FROM v_sales_analytics
                    WHERE {c_where}
                    GROUP BY category
                """, tuple(c_params))

                cat_prods = execute_analytics_sql(f"""
                    SELECT product, SUM(qty) as qty, SUM(net_sales) as net_sales
                    FROM v_sales_analytics
                    WHERE {c_where}
                    GROUP BY product
                    ORDER BY net_sales DESC
                    LIMIT 5
                """, tuple(c_params))

                cat_outs = execute_analytics_sql(f"""
                    SELECT outlet, SUM(net_sales) as net_sales, SUM(qty) as qty
                    FROM v_sales_analytics
                    WHERE {c_where}
                    GROUP BY outlet
                    ORDER BY net_sales DESC
                    LIMIT 3
                """, tuple(c_params))

                if cat_sum:
                    cs = cat_sum[0]
                    category_analytics_data = {"summary": cs, "products": cat_prods, "outlets": cat_outs}
                    cp_lines = [f"  * {p['product']}: Rp {p['net_sales']:,.0f} ({p['qty']:.0f} pcs)" for p in cat_prods]
                    co_lines = [f"  * {o['outlet']}: Rp {o['net_sales']:,.0f} ({o['qty']:.0f} pcs)" for o in cat_outs]
                    category_specific_text = (
                        f"\nDATA SPESIFIK KATEGORI BARANG: {detected_category}\n"
                        f"- Total Penjualan (Net Sales 2026): Rp {cs.get('total_sales', 0):,.0f}\n"
                        f"- Total Kuantitas Terjual: {cs.get('total_qty', 0):,.0f} pcs\n"
                        f"- Total Transaksi (Struk): {cs.get('total_tx', 0):,} transaksi\n"
                        f"- Estimasi Laba Kotor: Rp {cs.get('total_profit', 0):,.0f}\n"
                        f"- Gross Margin: {cs.get('margin_pct', 0):.1f}%\n"
                        f"- Top 5 Produk Terlaris:\n" + "\n".join(cp_lines) + "\n"
                        f"- Top 3 Cabang Terlaris:\n" + "\n".join(co_lines) + "\n"
                    )
            except Exception:
                pass

        invoice_specific_text = ""
        invoice_analytics_data = None
        if is_invoice_query:
            try:
                inv_conds = ["year = 2026"]
                inv_params = []
                if detected_month:
                    inv_conds.append("month = ?")
                    inv_params.append(detected_month)
                if detected_outlet:
                    inv_conds.append("outlet = ?")
                    inv_params.append(detected_outlet)
                elif detected_area:
                    inv_conds.append("area = ?")
                    inv_params.append(detected_area)
                inv_where = " AND ".join(inv_conds)

                top_invs = execute_analytics_sql(f"""
                    SELECT invoice, date, outlet, MAX(transaction_total) as total_nominal, MAX(invoice_discount) as discount, SUM(qty) as total_qty
                    FROM v_sales_analytics
                    WHERE {inv_where}
                    GROUP BY invoice, date, outlet
                    ORDER BY total_nominal DESC
                    LIMIT 5
                """, tuple(inv_params))

                inv_stats = execute_analytics_sql(f"""
                    SELECT 
                        COUNT(DISTINCT invoice) as total_invoices,
                        ROUND(AVG(transaction_total), 0) as avg_nominal,
                        MAX(transaction_total) as max_nominal,
                        ROUND(AVG(invoice_discount), 0) as avg_discount,
                        ROUND(SUM(invoice_discount), 0) as total_discount
                    FROM (
                        SELECT invoice, MAX(transaction_total) as transaction_total, MAX(invoice_discount) as invoice_discount
                        FROM v_sales_analytics
                        WHERE {inv_where}
                        GROUP BY invoice
                    )
                """, tuple(inv_params))

                if inv_stats:
                    istat = inv_stats[0]
                    invoice_analytics_data = {"stats": istat, "top_invoices": top_invs}
                    ti_lines = [f"  {idx}. No Struk {ti['invoice']} ({ti['date']} di {ti['outlet']}): Rp {ti['total_nominal']:,.0f} (Volume: {ti['total_qty']:.0f} pcs, Diskon: Rp {ti['discount']:,.0f})" for idx, ti in enumerate(top_invs, 1)]
                    invoice_specific_text = (
                        f"\nDATA ANALISIS STRUK TRANSAKSI & KERANJANG BELANJA (INVOICE ANALYTICS):\n"
                        f"- Total Struk Tercatat: {istat.get('total_invoices', 0):,} struk\n"
                        f"- Rata-rata Nilai per Struk (Basket Size / ATV): Rp {istat.get('avg_nominal', 0):,.0f}\n"
                        f"- Rekor Transaksi Struk Tertinggi: Rp {istat.get('max_nominal', 0):,.0f}\n"
                        f"- Rata-rata Diskon per Struk: Rp {istat.get('avg_discount', 0):,.0f}\n"
                        f"- Total Nilai Diskon Diberikan: Rp {istat.get('total_discount', 0):,.0f}\n"
                        f"- Top 5 Struk Transaksi Tertinggi:\n" + "\n".join(ti_lines) + "\n"
                    )
            except Exception:
                pass

        # Resolve Dashboard data based on extracted dimensions
        target_date = detected_date or data.get("date") or None
        target_month = (detected_date[:7] if detected_date else (detected_month or data.get("month") or None))
        target_area = detected_area or data.get("area") or None
        target_outlet = detected_outlet or data.get("outlet") or None
        target_start_date = detected_start_date or data.get("start_date") or None
        target_end_date = detected_end_date or data.get("end_date") or None
        target_supplier = detected_supplier or data.get("supplier") or None
        target_category = detected_category or data.get("category") or None
        target_jenis = (
            "KONSINYASI" if re.search(r"\bkonsin(?:yasi)?\b", p_lower)
            else ("DAGANGAN" if re.search(r"\bdagang(?:an)?\b|\bbeli putus\b", p_lower) else None)
        ) or data.get("jenis") or None
        summary_ctx = data.get("summary_context")

        if (detected_date or detected_month or detected_area or detected_outlet or 
            detected_supplier or detected_category or detected_start_date or detected_end_date or target_jenis):
            dash = build_dashboard(
                rows_cache,
                date=target_date,
                month=(target_month if not (target_date or target_start_date) else None),
                area=target_area,
                outlet=target_outlet,
                start_date=target_start_date,
                end_date=target_end_date,
                supplier=target_supplier,
                category=target_category,
                jenis=target_jenis,
                include_raw=False
            )
            label_parts = []
            if target_start_date and target_end_date:
                label_parts.append(f"Rentang {target_start_date} s/d {target_end_date}")
            elif target_date:
                dt_obj = datetime.strptime(target_date, "%Y-%m-%d")
                label_parts.append(f"Tanggal {dt_obj.day} {MONTH_NAMES.get(target_date[5:7], target_date[5:7])} {dt_obj.year}")
            elif target_month:
                label_parts.append(f"Bulan {MONTH_NAMES.get(target_month[5:], target_month)} {target_month[:4]}")
            if target_area:
                label_parts.append(f"Area {target_area}")
            if target_outlet:
                label_parts.append(f"Outlet {target_outlet}")
            if target_supplier:
                label_parts.append(f"Supplier {target_supplier}")
            if target_category:
                label_parts.append(f"Kategori {target_category}")
            if target_jenis:
                label_parts.append(f"Model {target_jenis.capitalize()}")
            filter_label = " - ".join(label_parts) if label_parts else (data.get("filter_label") or "Seluruh Data (YTD 2026)")
        elif summary_ctx and isinstance(summary_ctx, dict) and "summary" in summary_ctx:
            dash = summary_ctx
            filter_label = data.get("filter_label") or "Seluruh Data (YTD 2026)"
        else:
            dash = build_dashboard(
                rows_cache,
                month=data.get("month") or None,
                date=data.get("date") or None,
                outlet=data.get("outlet") or None,
                area=data.get("area") or None,
                start_date=data.get("start_date") or None,
                end_date=data.get("end_date") or None,
                supplier=data.get("supplier") or None,
                category=data.get("category") or None,
                jenis=target_jenis,
                include_raw=False
            )
            filter_label = data.get("filter_label") or "Seluruh Data (YTD 2026)"

        summary_26 = dash.get("summary", {}).get("2026", {"net_sales": 0, "transactions": 0, "qty": 0, "atv": 0})
        summary_25 = dash.get("summary", {}).get("2025", {"net_sales": 0, "transactions": 0, "qty": 0, "atv": 0})
        growth = dash.get("summary", {}).get("growth", {})
        target_rev = dash.get("target_revenue", 0)
        target_achieve = dash.get("target_achievement", 0)
        target_gap = dash.get("target_gap", 0)
        outlets_raw = dash.get("outlets", [])
        if outlets_raw:
            top_outlets = sorted(outlets_raw, key=lambda x: x.get("2026", {}).get("net_sales", 0), reverse=True)[:5]
            eval_cands = [o for o in outlets_raw if o.get("growth_sales") is not None]
            bottom_outlets = sorted(eval_cands, key=lambda x: x.get("growth_sales", 0))[:5] if eval_cands else []
        else:
            top_outlets = dash.get("top_5_outlets", [])
            bottom_outlets = dash.get("bottom_5_outlets", [])
        priority_products = dash.get("priority_products", [])
        products_raw = dash.get("products", [])
        products_by_sales = sorted(products_raw, key=lambda x: x.get("2026", {}).get("net_sales", 0), reverse=True)
        products_by_qty = sorted(products_raw, key=lambda x: x.get("2026", {}).get("qty", 0), reverse=True)

        peak_sales = dash.get("peak_sales_hour", "16")
        peak_tx = dash.get("peak_tx_hour", "16")

        # Theme Park Retail Metrics
        total_visitors_val = dash.get("total_visitors", 0)
        sph_val = dash.get("sph", 0.0)
        cap_rate_val = dash.get("capture_rate", 0.0)
        atv_val = summary_26.get("atv", 0.0)
        upt_val = (summary_26['qty'] / summary_26['transactions']) if summary_26.get('transactions', 0) > 0 else 0.0

        if total_visitors_val > 0:
            visitor_metric_text = (
                f"- Total Pengunjung Masuk Wahana (Visitors): {total_visitors_val:,.0f} orang\n"
                f"- Spending per Head (SpH): Rp {sph_val:,.0f} per orang (rata-rata belanja merchandise tiap pengunjung wahana)\n"
                f"- Capture Rate: {cap_rate_val:.2f}% (persentase pengunjung wahana yang membeli barang di toko kita)\n"
            )
        else:
            visitor_metric_text = (
                "- Data Pengunjung Masuk Wahana (Visitors): Belum terpetakan spesifik untuk filter kombinasi ini\n"
            )

        retail_basket_text = (
            f"- Rata-rata Belanja per Struk (ATV / Basket Size): Rp {atv_val:,.0f}\n"
            f"- Rata-rata Jumlah Barang per Struk (UPT): {upt_val:.2f} unit per transaksi\n"
        )

        # Safe growth formatting
        g_sales = growth.get('net_sales')
        g_sales_str = f"{g_sales:+.2f}%" if g_sales is not None else "0.00%"
        g_tx = growth.get('transactions')
        g_tx_str = f"{g_tx:+.2f}%" if g_tx is not None else "0.00%"
        g_qty = growth.get('qty')
        g_qty_str = f"{g_qty:+.2f}%" if g_qty is not None else "0.00%"

        top_o_lines = []
        for i, o in enumerate(top_outlets[:5]):
            gw = o.get('growth_sales')
            gw_str = f"{gw:+.1f}%" if gw is not None else "0.0%"
            top_o_lines.append(f"  {i+1}. {o['name']}: Rp {o['2026']['net_sales']:,.0f} (Kontribusi: {o.get('contrib_2026', 0):.1f}%, Growth: {gw_str})")
        top_o_text = "\n".join(top_o_lines) or "  - Data outlet tidak tersedia untuk filter ini"

        eval_o_lines = []
        for o in bottom_outlets[:5]:
            gw = o.get('growth_sales')
            gw_str = f"{gw:+.1f}%" if gw is not None else "0.0%"
            status_gw = 'Terkoreksi' if (gw or 0) < 0 else 'Skala Kecil'
            eval_o_lines.append(f"  - {o['name']}: Rp {o['2026']['net_sales']:,.0f} (Growth: {gw_str} YoY - {status_gw})")
        eval_o_text = "\n".join(eval_o_lines) or "  - Seluruh cabang mencatatkan performa positif"

        top_p_sales_text = "\n".join([
            f"  {i+1}. {p['name']}: Rp {p['2026']['net_sales']:,.0f} (Volume: {p['2026']['qty']:,} pcs, Kontribusi: {p.get('contrib_2026', 0):.1f}%)"
            for i, p in enumerate(products_by_sales[:10])
        ]) or "  - Data produk tidak tersedia untuk filter ini"

        top_p_qty_text = "\n".join([
            f"  {i+1}. {p['name']}: {p['2026']['qty']:,} pcs (Omset: Rp {p['2026']['net_sales']:,.0f}, Laju: {p.get('daily_avg', 0):.1f}/hari)"
            for i, p in enumerate(products_by_qty[:10])
        ]) or "  - Data produk tidak tersedia untuk filter ini"

        top_p_text = "\n".join([
            f"  - {p['name']}: Laju {p.get('daily_avg', 0):.1f} pcs/hari (Saran Restock Mingguan: ~{p.get('weekly_proj', 0):,} pcs)"
            for p in priority_products[:5]
        ]) or "  - Data produk prioritas tidak tersedia untuk filter ini"

        # Outlets sorted by Qty (Volume)
        outlets_by_qty = sorted(dash.get("outlets", []), key=lambda x: x.get("2026", {}).get("qty", 0), reverse=True)
        top_o_qty_text = "\n".join([
            f"  {i+1}. {o['name']}: {o['2026']['qty']:,} pcs (Omset: Rp {o['2026']['net_sales']:,.0f}, Kontribusi: {o.get('contrib_2026', 0):.1f}%)"
            for i, o in enumerate(outlets_by_qty[:8])
        ]) or "  - Data outlet tidak tersedia untuk filter ini"

        # Master Lookup (Gross Profit, Margin, Suppliers, Categories, Model Bisnis, Unmapped)
        gross_profit_val = dash.get("gross_profit_2026", 0.0)
        cogs_val = dash.get("total_cogs_2026", 0.0)
        gross_margin_val = dash.get("gross_margin_2026", 0.0)
        model_sum = dash.get("model_summary", {})
        konsin = model_sum.get("KONSINYASI", {})
        dagang = model_sum.get("DAGANGAN", {})
        suppliers_raw = dash.get("suppliers", [])
        categories_raw = dash.get("categories", [])
        unmapped_raw = dash.get("unmapped_skus", [])

        top_supp_lines = []
        for idx, s in enumerate(suppliers_raw[:8], 1):
            top_supp_lines.append(
                f"  {idx}. {s['name']} ({s.get('jenis', 'DAGANGAN')}): Omset Rp {s['net_sales']:,.0f} | Laba Rp {s['gross_profit']:,.0f} (Margin: {s.get('margin', 0):.1f}%, Qty: {s.get('qty', 0):,} pcs)"
            )
        top_supp_text = "\n".join(top_supp_lines) or "  - Data rekanan supplier tidak tersedia untuk filter ini"

        top_cat_lines = []
        for idx, c in enumerate(categories_raw[:8], 1):
            top_cat_lines.append(
                f"  {idx}. {c['name']}: Omset Rp {c['net_sales']:,.0f} | Laba Rp {c['gross_profit']:,.0f} (Margin: {c.get('margin', 0):.1f}%, Qty: {c.get('qty', 0):,} pcs)"
            )
        top_cat_text = "\n".join(top_cat_lines) or "  - Data kategori tidak tersedia untuk filter ini"

        if unmapped_raw:
            unmapped_sample = ", ".join([u['name'] for u in unmapped_raw[:3]])
            unmapped_status_text = (
                f"- Radar SKU Baru: Ditemukan {len(unmapped_raw)} produk terjual yang belum terdaftar di DATA_LOOKUP.xlsx (contoh: {unmapped_sample})."
            )
        else:
            unmapped_status_text = "- Radar SKU Baru: Seluruh 100% SKU produk yang terjual telah terdaftar lengkap di Master Data DATA_LOOKUP.xlsx."

        # Smart Product / SKU Keyword Search
        matched_products_text = ""
        stop_words = {
            "yang", "di", "dan", "ke", "dari", "pada", "untuk", "dengan", "ini", "itu",
            "bulan", "hari", "tahun", "tgl", "tanggal", "jam", "pukul", "area",
            "september", "agustus", "juli", "juni", "mei", "april", "maret", "februari", "januari",
            "tolong", "cek", "dong", "min", "bro", "konter", "outlet", "cabang", "toko",
            "spesifik", "barang", "produk", "berapa", "mana", "siapa", "gimana", "bagaimana",
            "paling", "tinggi", "rendah", "total", "semua", "data", "target", "capai", "ada",
            "mau", "tanya", "nanya", "kasih", "tau", "tahu", "detail", "detailnya",
            "laku", "terjual", "omsetnya", "harganya", "stok", "stoknya", "pcs", "unit", "rupiah"
        }
        import re
        query_words = [w for w in re.findall(r"\w+", p_lower) if len(w) >= 3 and w not in stop_words]
        top_matched_prods = []
        if query_words:
            search_pool = products_raw
            matched_items = []
            for p in search_pool:
                p_name_lower = p["name"].lower()
                score = sum(1 for w in query_words if w in p_name_lower)
                if score > 0:
                    matched_items.append((score, p))
            if matched_items:
                matched_items.sort(key=lambda x: (x[0], x[1]["2026"]["net_sales"]), reverse=True)
                top_matched_prods = [item[1] for item in matched_items[:5]]
                
                # Cari konter terlaris untuk top matched products
                top_outlets_for_matched = {}
                rows_for_search = read_sales()
                for r in rows_for_search:
                    if r["year"] == 2026 and (not detected_month or r["month"] == detected_month):
                        pname = r["product"]
                        if pname in [mp["name"] for mp in top_matched_prods[:3]]:
                            if pname not in top_outlets_for_matched:
                                top_outlets_for_matched[pname] = {}
                            o_name = r["outlet"]
                            top_outlets_for_matched[pname][o_name] = top_outlets_for_matched[pname].get(o_name, 0) + r["qty"]

                lines = []
                for p in top_matched_prods:
                    best_o_str = ""
                    if p["name"] in top_outlets_for_matched and top_outlets_for_matched[p["name"]]:
                        best_o, best_q = max(top_outlets_for_matched[p["name"]].items(), key=lambda x: x[1])
                        best_o_str = f" | Terlaris di: {best_o} ({best_q:,.0f} pcs)"
                    lines.append(
                        f"  * **{p['name']}**: Total Omset Rp {p['2026']['net_sales']:,.0f} | Terjual {p['2026']['qty']:,} pcs{best_o_str}"
                    )
                matched_products_text = f"\nHASIL PENCARIAN NAMA BARANG SPESIFIK:\n" + "\n".join(lines) + "\n"

        # Hourly Distribution Breakdown for Peak & Specific Hour
        hourly_chart = dash.get("hourly_chart", [])
        active_hours = [h for h in hourly_chart if h.get("net_sales_2026", 0) > 0]
        hourly_text = ""
        if active_hours:
            top_h_sales = sorted(active_hours, key=lambda x: x["net_sales_2026"], reverse=True)[:3]
            top_h_tx = sorted(active_hours, key=lambda x: x.get("transactions_2026", 0), reverse=True)[:3]
            hourly_lines = [
                "DISTRIBUSI JAM OPERASIONAL & JAM SIBUK:",
                f"- Puncak Penjualan (Revenue Terbesar): Pukul {peak_sales}:00 WIB",
                f"- Puncak Kepadatan Pengunjung (Struk Terbanyak): Pukul {peak_tx}:00 WIB",
                "  * Top 3 Jam Penjualan Terbesar: " + ", ".join([f"Pukul {h['hour']}:00 (Rp {h['net_sales_2026']:,.0f}, {h.get('transactions_2026', 0)} struk)" for h in top_h_sales]),
                "  * Top 3 Jam Transaksi Terpadat: " + ", ".join([f"Pukul {h['hour']}:00 ({h.get('transactions_2026', 0)} struk, Rp {h['net_sales_2026']:,.0f})" for h in top_h_tx]),
            ]
            if detected_hour:
                target_h = next((h for h in active_hours if str(int(h['hour'])) == str(int(detected_hour))), None)
                if target_h:
                    hourly_lines.append(f"- DATA SPESIFIK PUKUL {detected_hour}:00 WIB: Omset Rp {target_h['net_sales_2026']:,.0f} | Total {target_h.get('transactions_2026', 0)} struk transaksi")
                else:
                    hourly_lines.append(f"- DATA SPESIFIK PUKUL {detected_hour}:00 WIB: Tidak tercatat adanya transaksi penjualan pada jam ini.")
            hourly_text = "\n".join(hourly_lines) + "\n"

        system_context = f"""Anda adalah "Ancol AI Executive Business Analyst", asisten kecerdasan buatan resmi untuk jajaran direksi, general manager, kepala toko, dan tim operasional Ancol Store (PT Pembangunan Jaya Ancol Tbk).

GAYA JAWABAN & TATA ATURAN:
- Bahasa Indonesia yang santun, profesional, lugas, percaya diri, dan mudah dipahami oleh manajemen senior (boomer-friendly).
- Format jawaban dengan poin-poin bernomor (1, 2, 3), cetak tebal angka finansial penting (Rp, %, unit, jam, tanggal), dan berikan rekomendasi aksi konkret yang aplikatif di lapangan.
- PANDUAN UTAMA MENJAWAB:
  * PERTANYAAN MASA DEPAN, KALENDER LIBUR, TANGGAL MERAH, & PERSIAPAN JUALAN (Misal: Desember 2026, Natal, Tahun Baru, Musim Liburan):
    Gunakan bagian 'INFORMASI RESMI KALENDER NASIONAL' di bawah!
    Sebutkan tanggal merah secara spesifik beserta nama hari dan potensi long weekend-nya.
    Jelaskan proyeksi lonjakan pengunjung theme park (+250% s/d +350%) dan berikan 'Action Plan Persiapan Operasional' (kapan harus kunci buffer stock gudang H-14, jam sibuk mulai 11:00 WIB, pengaturan shift kasir penuh / mobile EDC).
    JANGAN MENOLAK PERTANYAAN atau sekadar berkata data tidak ada hanya karena bulan tersebut belum lewat!
  * PERBANDINGAN KOMPARATIF (HEAD-TO-HEAD, misal: Dufan vs Atlantis, Weekend vs Weekday, Agustus vs September):
    Gunakan bagian 'DATA KOMPARATIF HEAD-TO-HEAD' di bawah!
    Sajikan perbandingan angka omset, struk, dan ATV secara berdampingan (side-by-side) serta jelaskan entitas mana yang lebih unggul.
  * SPENDING PER HEAD (SPH) & CAPTURE RATE:
    Gunakan jika relevan untuk menganalisis apakah perubahan omset dipengaruhi oleh jumlah pengunjung atau efektivitas kasir toko dalam menggaet pembeli.
  * AVERAGE TRANSACTION VALUE (ATV) & UPT:
    Gunakan untuk menganalisis ukuran belanja keranjang per transaksi.
  * JAM OPERASIONAL & SIBUK:
    Jika user bertanya tentang jam sibuk/ramai atau jam tertentu, paparkan data nominal rupiah, struk, dan arahan kesiapan loket kasir 30 menit sebelum jam puncak.
  * BARANG / PRODUK:
    Jika user bertanya tentang barang/produk terlaris, WAJIB JAWAB DENGAN NAMA BARANG FISIK SPESIFIK (seperti Jas Hujan, Prima 600ml, Boneka, dll.), BUKAN nama toko/cabang!
  * REKANAN SUPPLIER SPESIFIK:
    Jika user bertanya tentang supplier tertentu (mencakup 86 supplier), WAJIB gunakan data dari bagian 'DATA SPESIFIK REKANAN / SUPPLIER' di bawah!
    Sajikan rincian omset, unit terjual, struk, laba kotor, margin %, produk terlarisnya, dan cabang kontributor utamanya.
  * KATEGORI BARANG SPESIFIK:
    Jika user bertanya tentang kategori barang tertentu (mencakup 32 kategori), WAJIB gunakan data dari bagian 'DATA SPESIFIK KATEGORI BARANG' di bawah!
    Sajikan rincian omset, unit terjual, margin laba kotor, produk unggulan, dan cabang penjual terbesarnya.
  * ANALISIS STRUK / TRANSAKSI / DISKON / BASKET SIZE:
    Jika user bertanya seputar struk terbesar, transaksi tertinggi, diskon faktur, atau rata-rata belanja per struk, WAJIB gunakan data dari bagian 'DATA ANALISIS STRUK TRANSAKSI & KERANJANG BELANJA' di bawah!
  * SUPPLIER / VENDOR & MARGIN KEUNTUNGAN:
    Jika user bertanya tentang supplier, margin, laba kotor, atau rekanan yang paling menguntungkan secara umum:
    Paparkan nama supplier, model bisnisnya (Konsinyasi atau Dagangan Beli Putus), omset, estimasi laba kotor, dan persentase margin laba kotor.
    Sebutkan rekanan vendor eksternal unggulan (seperti Johnson Hutahuruk, Yosi Detavian, PT Agape Indah Jaya, dll.) dan kontribusi internal (Merchandise).
  * MODEL BISNIS (KONSINYASI VS DAGANGAN):
    Sajikan perbandingan porsi omset (Konsinyasi ~67% vs Dagangan ~33%), nominal laba kotor, dan margin keuntungan kedua model bisnis.
  * RADAR SKU BARU / MASTER DATA:
    Jika user bertanya tentang SKU baru yang belum terdaftar atau status master data, paparkan informasinya secara lugas.
- Hindari istilah teknis IT/database/coding. Bersikaplah seperti konsultan bisnis strategis yang solutif.

{holiday_calendar_text}
{comparative_text}
{supplier_specific_text}
{category_specific_text}
{invoice_specific_text}
{matched_products_text}

KONTEKS FILTER AKTIF: {filter_label}
DATA RINGKASAN ANCOL STORE:
- Total Net Sales 2026: Rp {summary_26['net_sales']:,.0f}
- Total Net Sales 2025: Rp {summary_25['net_sales']:,.0f}
- Pertumbuhan (Growth YoY): {g_sales_str}
- Target Revenue: Rp {target_rev:,.0f} (Pencapaian: {target_achieve:.1f}%, Status: {'Melampaui Target' if target_gap >= 0 else 'Defisit'} sebesar Rp {abs(target_gap):,.0f})
- Estimasi Laba Kotor (Gross Profit): Rp {gross_profit_val:,.0f} (Gross Margin: {gross_margin_val:.1f}%)
- Total Modal Dasar Barang (HPP / COGS): Rp {cogs_val:,.0f}
- Model Bisnis Konsinyasi: Omset Rp {konsin.get('net_sales', 0):,.0f} (Porsi: {konsin.get('share', 0):.1f}%, Margin: {konsin.get('margin', 0):.1f}%, Laba: Rp {konsin.get('gross_profit', 0):,.0f})
- Model Bisnis Dagangan: Omset Rp {dagang.get('net_sales', 0):,.0f} (Porsi: {dagang.get('share', 0):.1f}%, Margin: {dagang.get('margin', 0):.1f}%, Laba: Rp {dagang.get('gross_profit', 0):,.0f})
{unmapped_status_text}
- Total Transaksi (Struk): {summary_26['transactions']:,} struk (YoY: {g_tx_str})
- Total Barang Terjual: {summary_26['qty']:,} unit (YoY: {g_qty_str})
{retail_basket_text}
{visitor_metric_text}
- Jam Sibuk Penjualan & Transaksi: Pukul {peak_sales}:00 WIB
- Jam Operasional Aktif: Pukul 06:00 s/d 20:00 WIB

{hourly_text}
TOP 10 BARANG / PRODUK BERDASARKAN OMSET (REVENUE):
{top_p_sales_text}

TOP 10 BARANG / PRODUK BERDASARKAN KUANTITAS TERJUAL (QTY):
{top_p_qty_text}

TOP CABANG BERDASARKAN OMSET (REVENUE):
{top_o_text}

TOP CABANG BERDASARKAN VOLUME TERJUAL (QTY):
{top_o_qty_text}

TOP 8 REKANAN / SUPPLIER VENDOR (BERDASARKAN OMSET & LABA KOTOR):
{top_supp_text}

TOP 8 KATEGORI BARANG (OMSET & ESTIMASI MARGIN):
{top_cat_text}

CABANG PERLU PERHATIAN / EVALUASI (PENURUNAN YoY):
{eval_o_text}

TOP PRODUK PRIORITAS (FAST-MOVING VELOCITY):
{top_p_text}
"""

        executive_context_prompt = system_context
        api_key = get_gemini_api_key()

        if api_key:
            try:
                candidate_models = [
                    "gemini-3.1-flash-lite",
                    "gemini-flash-lite-latest",
                ]

                # Build multi-turn conversational contents
                contents = []
                if history and isinstance(history, list):
                    valid_history = []
                    expected_role = "user"
                    for h in history[-4:]:
                        if not isinstance(h, dict):
                            continue
                        r = h.get("role")
                        t = (h.get("text") or "").strip()
                        if not t:
                            continue
                        if r == "user" and expected_role == "user":
                            valid_history.append({"role": "user", "parts": [{"text": t}]})
                            expected_role = "model"
                        elif r in ("assistant", "model") and expected_role == "model":
                            valid_history.append({"role": "model", "parts": [{"text": t}]})
                            expected_role = "user"
                    if valid_history and valid_history[-1]["role"] == "user":
                        valid_history.pop()
                    contents.extend(valid_history)

                contents.append({
                    "role": "user",
                    "parts": [{"text": f"SISTEM ANCOL POS BUSINESS INTELLIGENCE:\n{executive_context_prompt}\n\nPERTANYAAN EKSEKUTIF:\n{user_prompt}"}]
                })

                for model_name in candidate_models:
                    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
                    payload = {
                        "contents": contents,
                        "generationConfig": {
                            "temperature": 0.2,
                            "maxOutputTokens": 4096,
                            "thinkingConfig": {"thinkingBudget": 0}
                        }
                    }

                    req = urllib.request.Request(
                        url,
                        data=json.dumps(payload).encode("utf-8"),
                        headers={"Content-Type": "application/json"}
                    )

                    try:
                        with urllib.request.urlopen(req, timeout=7) as response:
                            res_body = json.loads(response.read().decode("utf-8"))
                            text = res_body["candidates"][0]["content"]["parts"][0]["text"].strip()
                            return jsonify({"ok": True, "answer": text, "engine": "gemini-cloud", "model": model_name})
                    except Exception:
                        continue
            except Exception as cloud_err:
                print(f"[ask-ai] Gemini cloud error, falling back to rule engine: {cloud_err}")

        # Intelligent Built-in Ancol Business Rule Engine (Instant Fallback)
        p_lower = user_prompt.lower()
        top_o_name = top_outlets[0]['name'] if top_outlets else "Seluruh Toko"
        top_o_sales = top_outlets[0]['2026']['net_sales'] if top_outlets else summary_26['net_sales']
        top_o_contrib = top_outlets[0].get('contrib_2026', 0) if top_outlets else 100.0

        bot_list_text = ""
        if bottom_outlets:
            bot_items = []
            for idx, bo in enumerate(bottom_outlets[:3], 1):
                gw = bo.get('growth_sales', 0)
                bot_items.append(f"{idx}. **{bo['name']}**: Omset Rp {bo['2026']['net_sales']:,.0f} ({gw:+.1f}% YoY)")
            bot_list_text = "\n".join(bot_items)
        else:
            bot_list_text = "Seluruh cabang terpantau stabil dan tidak ada indikasi penurunan signifikan."

        prod_list_text = ""
        if priority_products:
            prod_items = []
            for idx, pr in enumerate(priority_products[:3], 1):
                prod_items.append(f"{idx}. **{pr['name']}**: Laju **{pr.get('daily_avg', 0):.0f} pcs/hari** $\\rightarrow$ Rekomendasi stok mingguan: **~{pr.get('weekly_proj', 0):,} pcs**.")
            prod_list_text = "\n".join(prod_items)
        else:
            prod_list_text = "Data produk tidak terdeteksi untuk parameter ini."

        top_prod_name = priority_products[0]['name'] if priority_products else "Merchandise Utama"
        top_prod_daily = priority_products[0].get('daily_avg', 0) if priority_products else 0
        top_prod_weekly = priority_products[0].get('weekly_proj', 0) if priority_products else 0

        if is_future_or_holiday_query:
            answer = (
                "**Panduan Kalender Libur Nasional & Strategi Persiapan Penjualan (Desember 2026):**\n\n"
                "1. **Rincian Tanggal Merah & Hari Libur Nasional Resmi:**\n"
                "   - **Jumat, 25 Desember 2026**: Hari Raya Natal (membentuk *Long Weekend* 4 hari berturut-turut pada 25-27 Desember).\n"
                "   - **Kamis, 24 Desember 2026**: Potensi Cuti Bersama Natal.\n"
                "   - **Kamis, 31 Desember 2026**: Malam Pergantian Tahun Baru (puncak kepadatan pengunjung & festival Ancol).\n"
                "   - **Jumat, 1 Januari 2027**: Libur Tahun Baru 2027 Masehi.\n\n"
                "2. **Proyeksi Lonjakan Pengunjung Toko (Theme Park Context):**\n"
                "   - Diproyeksikan kenaikan pengunjung **250% – 350%** dibanding hari kerja reguler.\n"
                "   - Jam sibuk bergeser lebih awal: Mulai **pukul 11:00 WIB hingga 18:30 WIB** terjadi kepadatan terus-menerus.\n\n"
                "3. **Rekomendasi Taktis & Persiapan Lapangan (Action Plan):**\n"
                "   - **Stok Gudang (H-14)**: Kunci pengadaan buffer stock produk fast-moving (air mineral, jas hujan/payung musim hujan, boneka ikonik Ancol) minimal 3x kapasitas reguler sebelum 15 Desember.\n"
                "   - **Staf Kasir**: Operasikan seluruh counter kasir tanpa jeda (sistem istirahat bergilir) dan siapkan mesin EDC / mobile POS cadangan untuk mengurai antrean panjang.\n"
                "   - **Display Impulsif**: Letakkan keranjang minuman dingin dan payung persis di jalur antrean menuju kasir."
            )
        elif comparative_text:
            answer = (
                f"**Hasil Analisis Komparatif:**\n\n"
                f"{comparative_text.strip()}\n\n"
                f"💡 **Saran Bisnis**: Evaluasi faktor penggerak omset apakah didorong oleh lonjakan jumlah pembeli (struk) atau nilai belanja rata-rata per transaksi (ATV)."
            )
        elif detected_supplier and supplier_analytics_data and supplier_analytics_data.get("summary"):
            ss = supplier_analytics_data["summary"]
            p_items = [f"  {idx}. **{p['product']}**: Omset **Rp {p['net_sales']:,.0f}** ({p['qty']:.0f} pcs)" for idx, p in enumerate(supplier_analytics_data["products"], 1)]
            p_text = "\n".join(p_items) if p_items else "Belum ada rincian produk."
            o_items = [f"  {idx}. **{o['outlet']}**: Omset **Rp {o['net_sales']:,.0f}** ({o['qty']:.0f} pcs)" for idx, o in enumerate(supplier_analytics_data["outlets"], 1)]
            o_text = "\n".join(o_items) if o_items else "Belum ada rincian cabang."
            answer = (
                f"**Laporan Performa Rekanan Supplier ({detected_supplier}) - {filter_label}:**\n\n"
                f"1. **Ringkasan Finansial & Profitabilitas:**\n"
                f"   - Model Kerjasama: **{ss.get('jenis', 'DAGANGAN')}**\n"
                f"   - Total Omset (Net Sales 2026): **Rp {ss.get('total_sales', 0):,.0f}**\n"
                f"   - Total Barang Terjual: **{ss.get('total_qty', 0):,.0f} unit / pcs**\n"
                f"   - Estimasi Laba Kotor (Gross Profit): **Rp {ss.get('total_profit', 0):,.0f}** (Gross Margin: **{ss.get('margin_pct', 0):.1f}%**)\n\n"
                f"2. **Produk Terlaris ({detected_supplier}):**\n"
                f"{p_text}\n\n"
                f"3. **Cabang Kontributor Terbesar:**\n"
                f"{o_text}\n\n"
                f"💡 **Rekomendasi Operasional**: Pastikan ketersediaan buffer stock untuk produk fast-moving di atas terutama menjelang akhir pekan."
            )
        elif detected_category and category_analytics_data and category_analytics_data.get("summary"):
            cs = category_analytics_data["summary"]
            cp_items = [f"  {idx}. **{p['product']}**: Omset **Rp {p['net_sales']:,.0f}** ({p['qty']:.0f} pcs)" for idx, p in enumerate(category_analytics_data["products"], 1)]
            cp_text = "\n".join(cp_items) if cp_items else "Belum ada rincian produk."
            co_items = [f"  {idx}. **{o['outlet']}**: Omset **Rp {o['net_sales']:,.0f}** ({o['qty']:.0f} pcs)" for idx, o in enumerate(category_analytics_data["outlets"], 1)]
            co_text = "\n".join(co_items) if co_items else "Belum ada rincian cabang."
            answer = (
                f"**Laporan Performa Kategori ({detected_category}) - {filter_label}:**\n\n"
                f"1. **Kinerja Penjualan & Margin:**\n"
                f"   - Total Penjualan (Net Sales): **Rp {cs.get('total_sales', 0):,.0f}**\n"
                f"   - Total Kuantitas Terjual: **{cs.get('total_qty', 0):,.0f} unit**\n"
                f"   - Total Transaksi: **{cs.get('total_tx', 0):,} struk**\n"
                f"   - Estimasi Laba Kotor: **Rp {cs.get('total_profit', 0):,.0f}** (Gross Margin: **{cs.get('margin_pct', 0):.1f}%**)\n\n"
                f"2. **Produk Terlaris Kategori {detected_category}:**\n"
                f"{cp_text}\n\n"
                f"3. **Cabang dengan Penjualan Tertinggi:**\n"
                f"{co_text}\n\n"
                f"💡 **Insight Merchandising**: Optimalkan penempatan produk kategori ini di area depan counter atau jalur antrean kasir."
            )
        elif is_invoice_query and invoice_analytics_data and invoice_analytics_data.get("stats"):
            istat = invoice_analytics_data["stats"]
            ti_items = [f"  {idx}. **No Struk {ti['invoice']}** ({ti['date']} di *{ti['outlet']}*): **Rp {ti['total_nominal']:,.0f}** (Volume: {ti['total_qty']:.0f} pcs, Diskon: Rp {ti['discount']:,.0f})" for idx, ti in enumerate(invoice_analytics_data["top_invoices"], 1)]
            ti_text = "\n".join(ti_items) if ti_items else "Belum ada data struk."
            answer = (
                f"**Laporan Analisis Struk & Keranjang Belanja ({filter_label}):**\n\n"
                f"1. **Metrik Transaksi Utama:**\n"
                f"   - Rata-rata Belanja per Struk (Basket Size / ATV): **Rp {istat.get('avg_nominal', 0):,.0f}**\n"
                f"   - Rekor Nilai Struk Tertinggi: **Rp {istat.get('max_nominal', 0):,.0f}**\n"
                f"   - Total Diskon Faktur Diberikan: **Rp {istat.get('total_discount', 0):,.0f}** (Rata-rata: Rp {istat.get('avg_discount', 0):,.0f} per struk)\n\n"
                f"2. **Top 5 Transaksi Struk Terbesar:**\n"
                f"{ti_text}\n\n"
                f"💡 **Rekomendasi Kasir**: Tingkatkan promosi add-on produk impulsif di meja kasir untuk mendongkrak rata-rata belanja per struk."
            )
        elif any(k in p_lower for k in ["supplier", "vendor", "pemasok", "rekanan"]):
            supp_lines = []
            for idx, s in enumerate(suppliers_raw[:5], 1):
                supp_lines.append(
                    f"  {idx}. **{s['name']}** (*{s.get('jenis', 'DAGANGAN')}*): Omset **Rp {s['net_sales']:,.0f}** | Laba Kotor **Rp {s['gross_profit']:,.0f}** (Margin: **{s.get('margin', 0):.1f}%**)"
                )
            supp_list_str = "\n".join(supp_lines) if supp_lines else "Data supplier tidak tersedia untuk filter ini."
            top_partner = next((s for s in suppliers_raw if s['name'] != "MERCHANDISE"), suppliers_raw[0] if suppliers_raw else None)
            partner_name = top_partner['name'] if top_partner else "Mitra Konsinyasi"
            partner_margin = f"{top_partner.get('margin', 0):.1f}%" if top_partner else "0.0%"
            partner_profit = f"Rp {top_partner['gross_profit']:,.0f}" if top_partner else "Rp 0"
            answer = (
                f"**Analisis Rekanan & Supplier Vendor ({filter_label}):**\n\n"
                f"1. **Peringkat Top 5 Supplier Terbesar (Berdasarkan Omset & Laba):**\n"
                f"{supp_list_str}\n\n"
                f"2. **Mitra Vendor Unggulan**: Rekanan eksternal paling menguntungkan adalah **{partner_name}** yang membukukan laba kotor **{partner_profit}** dengan margin sehat **{partner_margin}**.\n\n"
                f"3. **Porsi Model Kerjasama**: Konsinyasi menyumbang **{konsin.get('share', 0):.1f}%** total omset (Margin: **{konsin.get('margin', 0):.1f}%**), sedangkan Dagangan beli putus menyumbang **{dagang.get('share', 0):.1f}%** (Margin: **{dagang.get('margin', 0):.1f}%**).\n\n"
                f"💡 **Rekomendasi Bisnis**: Perkuat pasokan barang fast-moving dari supplier ber-margin di atas 60% dan lakukan evaluasi kuota pajang rak untuk vendor dengan perputaran lambat."
            )
        elif any(k in p_lower for k in ["laba", "profit", "margin", "keuntungan", "hpp", "cogs", "modal"]):
            answer = (
                f"**Analisis Laba Kotor & Gross Margin ({filter_label}):**\n\n"
                f"1. **Estimasi Laba Kotor (Gross Profit)**: **Rp {gross_profit_val:,.0f}**\n"
                f"2. **Rata-rata Gross Margin**: **{gross_margin_val:.1f}%** dari total omset Rp {summary_26['net_sales']:,.0f}\n"
                f"3. **Total Beban Modal Pokok Barang (HPP / COGS)**: **Rp {cogs_val:,.0f}**\n\n"
                f"**Rincian Berdasarkan Model Kerjasama:**\n"
                f"- **Konsinyasi (Bagi Hasil Vendor)**: Omset **Rp {konsin.get('net_sales', 0):,.0f}** ({konsin.get('share', 0):.1f}% porsi), Laba Kotor **Rp {konsin.get('gross_profit', 0):,.0f}** (Margin: **{konsin.get('margin', 0):.1f}%**)\n"
                f"- **Dagangan (Beli Putus)**: Omset **Rp {dagang.get('net_sales', 0):,.0f}** ({dagang.get('share', 0):.1f}% porsi), Laba Kotor **Rp {dagang.get('gross_profit', 0):,.0f}** (Margin: **{dagang.get('margin', 0):.1f}%**)\n\n"
                f"💡 **Catatan Manajemen**: Margin laba kotor toko berada pada posisi yang sangat sehat (>58%). Pertahankan proporsi konsinyasi untuk meminimalisasi risiko dead stock barang tidak laku."
            )
        elif any(k in p_lower for k in ["konsinyasi", "dagangan", "beli putus", "model bisnis"]):
            answer = (
                f"**Perbandingan Model Bisnis: Konsinyasi vs Dagangan Beli Putus ({filter_label}):**\n\n"
                f"1. **Konsinyasi (Bagi Hasil Vendor)**:\n"
                f"   - Porsi Omset: **{konsin.get('share', 0):.1f}%** (Rp {konsin.get('net_sales', 0):,.0f})\n"
                f"   - Laba Kotor: **Rp {konsin.get('gross_profit', 0):,.0f}**\n"
                f"   - Gross Margin: **{konsin.get('margin', 0):.1f}%**\n\n"
                f"2. **Dagangan (Beli Putus / Direct Buy)**:\n"
                f"   - Porsi Omset: **{dagang.get('share', 0):.1f}%** (Rp {dagang.get('net_sales', 0):,.0f})\n"
                f"   - Laba Kotor: **Rp {dagang.get('gross_profit', 0):,.0f}**\n"
                f"   - Gross Margin: **{dagang.get('margin', 0):.1f}%**\n\n"
                f"💡 **Insight Strategis**: Model konsinyasi mendominasi 2/3 penjualan toko Ancol, memberikan fleksibilitas cashflow tanpa modal mati, sementara dagangan beli putus difokuskan untuk produk maskot inti dan merchandise eksklusif."
            )
        elif any(k in p_lower for k in ["unmapped", "radar", "sku baru", "belum terdaftar", "lookup"]):
            if unmapped_raw:
                sample_lines = "\n".join([f"  - **{u['name']}**: Terjual {u['qty']:,} pcs (Rp {u['net_sales']:,.0f})" for u in unmapped_raw[:5]])
                answer = (
                    f"**Laporan Audit Radar SKU Baru ({filter_label}):**\n\n"
                    f"Ditemukan **{len(unmapped_raw)} produk baru terjual** yang belum terdaftar di master data `DATA_LOOKUP.xlsx`.\n\n"
                    f"**Contoh Produk Belum Terpetakan:**\n"
                    f"{sample_lines}\n\n"
                    f"📋 **Langkah Perbaikan**: Buka tab **🏷️ Supplier & Kategori** pada dashboard, lalu klik tombol **Salin Daftar SKU Baru** untuk menyalin daftar SKU ke clipboard dan menambahkannya ke master file Excel Anda."
                )
            else:
                answer = (
                    f"**Status Master Data SKU ({filter_label}):**\n\n"
                    f"✅ **100% Sempurna**: Seluruh SKU produk yang terjual pada periode ini telah terdaftar lengkap di `DATA_LOOKUP.xlsx` dengan HPP, Supplier, dan Kategori yang valid."
                )
        elif any(k in p_lower for k in ["rapat", "direksi", "ringkasan", "brief", "kesimpulan", "summary"]):
            gap_text = f"melampaui target sebesar Rp {target_gap:,.0f}" if target_gap >= 0 else f"defisit Rp {abs(target_gap):,.0f} dari target"
            answer = (
                f"**Ringkasan Eksekutif Penjualan ({filter_label}):**\n\n"
                f"1. **Kinerja Finansial**: Total omset mencapai **Rp {summary_26['net_sales']:,.0f}**, mencatat pertumbuhan **{growth.get('net_sales', 0):+.2f}% YoY** dibanding periode 2025. Realisasi target tercatat **{target_achieve:.1f}%** ({gap_text}).\n\n"
                f"2. **Lokomotif & Evaluasi Toko**: Kontributor terbesar dipimpin oleh **{top_o_name}** (**Rp {top_o_sales:,.0f}** / {top_o_contrib:.1f}% porsi omset). Untuk evaluasi operasional, perhatikan cabang dengan pertumbuhan terkoreksi.\n\n"
                f"3. **Operasional & Produk Unggulan**: Puncak kepadatan pembeli terjadi pukul **{peak_sales}:00 WIB**. Produk terlaris adalah **{top_prod_name}** dengan laju harian **{top_prod_daily:.0f} pcs/hari** (estimasi pasokan mingguan **~{top_prod_weekly:,} pcs**)."
            )
        elif any(k in p_lower for k in ["jeblok", "turun", "evaluasi", "perhatian", "koreksi", "lemah", "drop"]):
            answer = (
                f"**Laporan Cabang yang Membutuhkan Evaluasi ({filter_label}):**\n\n"
                f"Berdasarkan perbandingan pertumbuhan YoY terhadap 2025:\n"
                f"{bot_list_text}\n\n"
                f"💡 **Rekomendasi Tindakan**: Periksa display merchandise di dekat area antrean kasir dan pastikan kesiapan personil 30 menit sebelum lonjakan pembeli pukul {peak_sales}:00 WIB."
            )
        elif any(k in p_lower for k in ["restock", "barang", "produk", "item", "gudang", "stok", "laris", "terlaris"]) or (any(k in p_lower for k in ["qty", "volume", "uang", "omset"]) and "konter" not in p_lower and "cabang" not in p_lower and "toko" not in p_lower and "outlet" not in p_lower):
            target_p_sales = products_by_sales[:5]
            target_p_qty = products_by_qty[:5]
            period_title = filter_label

            if top_matched_prods:
                mp = top_matched_prods[0]
                answer = (
                    f"**Detail Penjualan Produk ({period_title}):**\n\n"
                    f"1. **Nama Produk**: **{mp['name']}**\n"
                    f"2. **Total Omset**: **Rp {mp['2026']['net_sales']:,.0f}**\n"
                    f"3. **Total Unit Terjual**: **{mp['2026']['qty']:,} pcs**\n\n"
                    f"💡 **Rekomendasi Operasional**: Pastikan ketersediaan stok fisik produk ini di rak etalase depan kasir."
                )
            else:
                sales_lines = "\n".join([f"  {idx}. **{p['name']}**: Omset **Rp {p['2026']['net_sales']:,.0f}** ({p['2026']['qty']:,} pcs)" for idx, p in enumerate(target_p_sales[:5], 1)])
                qty_lines = "\n".join([f"  {idx}. **{p['name']}**: Terjual **{p['2026']['qty']:,} pcs** (Omset: Rp {p['2026']['net_sales']:,.0f})" for idx, p in enumerate(target_p_qty[:5], 1)])
                answer = (
                    f"**Peringkat Barang / Produk Unggulan ({period_title}):**\n\n"
                    f"**1. Berdasarkan Nilai Uang (Omset Terbesar):**\n"
                    f"{sales_lines}\n\n"
                    f"**2. Berdasarkan Kuantitas (Volume Qty Terlaris):**\n"
                    f"{qty_lines}\n\n"
                    f"📦 **Rekomendasi Gudang & Toko**: Prioritaskan pasokan barang-barang di atas terutama sebelum jam sibuk {peak_sales}:00 WIB."
                )
        elif any(k in p_lower for k in ["target", "gap", "capai", "pencapaian"]):
            status_str = "Melampaui Target" if target_gap >= 0 else "Defisit dari Target"
            answer = (
                f"**Status Target Penjualan ({filter_label}):**\n\n"
                f"- **Target Anggaran**: **Rp {target_rev:,.0f}**\n"
                f"- **Realisasi Penjualan**: **Rp {summary_26['net_sales']:,.0f}**\n"
                f"- **Tingkat Pencapaian**: **{target_achieve:.1f}%**\n"
                f"- **Posisi Gap**: **{status_str} sebesar Rp {abs(target_gap):,.0f}**\n\n"
                f"Kinerja penjualan saat ini tercatat bertumbuh **{g_sales_str}** dibanding tahun sebelumnya."
            )
        elif any(k in p_lower for k in ["jam", "sibuk", "puncak", "antrian", "kasir", "ramai"]):
            active_h = [h for h in dash.get("hourly_chart", []) if h.get("net_sales_2026", 0) > 0]
            top_h_s = sorted(active_h, key=lambda x: x["net_sales_2026"], reverse=True)[:3] if active_h else []
            h_details = ""
            if detected_hour:
                th = next((h for h in active_h if str(int(h['hour'])) == str(int(detected_hour))), None)
                if th:
                    h_details = f"\n\n⏰ **Kinerja Pukul {detected_hour}:00 WIB**:\n- Total Omset: **Rp {th['net_sales_2026']:,.0f}**\n- Total Transaksi: **{th.get('transactions_2026', 0)} struk**"
            h_rank_str = "\n".join([f"  {idx}. **Pukul {h['hour']}:00 WIB**: Omset **Rp {h['net_sales_2026']:,.0f}** ({h.get('transactions_2026', 0)} struk)" for idx, h in enumerate(top_h_s, 1)])
            answer = (
                f"**Analisis Jam Sibuk Toko ({filter_label}):**\n\n"
                f"- **Puncak Penjualan (Revenue)**: Terjadi pada **pukul {peak_sales}:00 WIB**.\n"
                f"- **Puncak Kepadatan Struk**: Terjadi pada **pukul {peak_tx}:00 WIB**.\n"
                f"- **Jam Operasional Toko**: Terjadwal pukul **06:00 s/d 20:00 WIB**.\n\n"
                f"**Jam Penjualan Tertinggi:**\n{h_rank_str}{h_details}\n\n"
                f"💡 **Arahan Operasional**: Pastikan seluruh loket kasir POS beroperasi penuh 30 menit sebelum jam puncak pukul {peak_sales}:00 WIB."
            )
        elif any(k in p_lower for k in ["qty", "kuantitas", "banyak", "volume", "unit"]):
            o_top_q = outlets_by_qty[0] if outlets_by_qty else None
            o_name_q = o_top_q['name'] if o_top_q else "Seluruh Toko"
            o_qty_val = o_top_q['2026']['qty'] if o_top_q else 0
            answer = (
                f"**Peringkat Konter Berdasarkan Kontribusi Qty ({filter_label}):**\n\n"
                f"Konter dengan kontribusi barang terbanyak adalah **{o_name_q}** dengan total penjualan **{o_qty_val:,} pcs**."
            )
        else:
            answer = (
                f"**Analisis Kinerja Ancol Store ({filter_label}):**\n\n"
                f"Total penjualan tercatat sebesar **Rp {summary_26['net_sales']:,.0f}** ({target_achieve:.1f}% dari target Rp {target_rev:,.0f}), tumbuh **{g_sales_str} YoY**.\n\n"
                f"Toko kontributor utama adalah **{top_o_name}** ({top_o_contrib:.1f}% omset), dengan produk terdepan **{top_prod_name}** ({top_prod_daily:.0f} pcs/hari).\n\n"
                f"Silakan klik salah satu tombol pintasan di atas untuk mendapatkan rincian spesifik seputar rapat, evaluasi toko, atau rekomendasi stok gudang."
            )

        if not api_key:
            answer += "\n\n*(💡 Tip: Anda dapat mengaktifkan Cloud AI Gemini tak terbatas dengan memasukkan API Key pada tombol ⚙️ Pengaturan di atas.)*"

        return jsonify({"ok": True, "answer": answer, "engine": "local-bi-rule"})

    except Exception as e:
        return jsonify({"ok": False, "answer": f"Terjadi kendala saat memproses analisa: {str(e)}"}), 500




if __name__ == "__main__":
    initial_data = build_dashboard(read_sales())
    with app.app_context():
        write_report(
            render_template(
                "index.html",
                data=initial_data,
                selected_month="",
                selected_date="",
                selected_outlet="",
                selected_area="",
            )
        )
    app.run(debug=False, use_reloader=False, host="0.0.0.0", port=5000)
