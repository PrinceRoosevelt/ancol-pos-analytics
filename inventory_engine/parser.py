"""
inventory_engine/parser.py
Robust Excel parser for:
1. Stok Gudang Utama (Columns: B=Kode, G=Nama, H=Supplier, M=Stok Akhir bersih, N=HPP, S=Harga Jual)
2. Stok Gabungan Konter (Columns: A=Supplier, B=Nama, D=Kode, E=Kategori, F=HPP, G=HJual, H=Gudang, I-AD=22 Konter)
"""
from __future__ import annotations

import io
from typing import Any, Dict, List, Tuple, Union
import openpyxl


def _to_float(val: Any) -> float:
    """Safely converts numeric or formatted string values to float."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", "")
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def _clean_sku(val: Any) -> str:
    """Normalizes SKU code string."""
    if val is None:
        return ""
    s = str(val).strip()
    # If Excel parsed it as float like 3122092301.0, convert to int representation
    if s.endswith(".0"):
        s = s[:-2]
    return s


def parse_stok_gudang(file_source: Union[str, bytes, io.BytesIO]) -> List[Dict[str, Any]]:
    """
    Parses 'stok gudang utama' Excel workbook.
    Header is at Row 5, Data starts at Row 6.
    Key column: M (idx 12) = Stok Akhir (bersih tanpa rusak/waste).
    """
    if isinstance(file_source, (bytes, bytearray)):
        file_source = io.BytesIO(file_source)

    wb = openpyxl.load_workbook(file_source, read_only=True, data_only=True)
    ws = wb.active

    items: List[Dict[str, Any]] = []
    # Search for header row (usually row 5)
    header_row_idx = 5
    sku_col_idx = 1     # Col B: Kode Barang
    name_col_idx = 6    # Col G: Nama Barang
    supp_col_idx = 7    # Col H: Supplier
    stock_col_idx = 12  # Col M: Stok Akhir
    hpp_col_idx = 13    # Col N: HPP
    hj_col_idx = 18     # Col S: Harga Jual

    for row_idx, row in enumerate(ws.iter_rows(min_row=1, values_only=True), start=1):
        if row_idx < 5:
            continue
        if row_idx == 5:
            # Dynamic header column check if needed
            for c_idx, cell in enumerate(row):
                c_str = str(cell or "").strip().lower()
                if "kode barang" in c_str or "kode item" in c_str:
                    sku_col_idx = c_idx
                elif "nama barang" in c_str:
                    name_col_idx = c_idx
                elif "supplier" in c_str:
                    supp_col_idx = c_idx
                elif "stok akhir" in c_str:
                    stock_col_idx = c_idx
                elif c_str == "hpp":
                    hpp_col_idx = c_idx
                elif "harga jual" in c_str:
                    hj_col_idx = c_idx
            continue

        # Data rows (row_idx >= 6)
        if not row or len(row) <= max(sku_col_idx, stock_col_idx):
            continue

        raw_sku = row[sku_col_idx] if sku_col_idx < len(row) else None
        sku = _clean_sku(raw_sku)
        if not sku:
            continue

        nama = str(row[name_col_idx] or "").strip() if name_col_idx < len(row) else ""
        if nama.lower() in ("total", "subtotal", "grand total") or not nama:
            continue

        supp = str(row[supp_col_idx] or "").strip() if supp_col_idx < len(row) else ""
        stok_akhir = _to_float(row[stock_col_idx] if stock_col_idx < len(row) else 0)
        hpp = _to_float(row[hpp_col_idx] if hpp_col_idx < len(row) else 0)
        hj = _to_float(row[hj_col_idx] if hj_col_idx < len(row) else 0)

        items.append({
            "sku_code": sku,
            "nama_barang": nama,
            "supplier": supp,
            "stok_akhir": stok_akhir,
            "hpp": hpp,
            "harga_jual": hj,
        })

    wb.close()
    return items


def parse_stok_konter(file_source: Union[str, bytes, io.BytesIO]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    Parses 'stok konter gabungan' Excel workbook.
    Header is at Row 3:
      Col A: Supplier
      Col B: Keterangan (Nama Barang)
      Col D: Kode barang
      Col E: Kategori
      Col F: hpp
      Col G: hjual
      Col H: A.GUDANG UTAMA MERCHANDISE (Warehouse stock from this file)
      Col I to AD: Outlets (Cbg.AWIN to Cbg.TJSO)
      Col AE: TOTAL
    Returns:
      (konter_items, fallback_gudang_items)
    """
    if isinstance(file_source, (bytes, bytearray)):
        file_source = io.BytesIO(file_source)

    wb = openpyxl.load_workbook(file_source, read_only=True, data_only=True)
    ws = wb.active

    konter_items: List[Dict[str, Any]] = []
    gudang_fallback_items: List[Dict[str, Any]] = []

    outlet_cols: List[Tuple[int, str]] = []  # (col_index, outlet_code)
    header_found = False

    for row_idx, row in enumerate(ws.iter_rows(min_row=1, values_only=True), start=1):
        if not row:
            continue

        # Header detection at Row 3 (or any row containing 'Kode barang' / 'TOTAL')
        if not header_found:
            row_strs = [str(c or "").strip() for c in row]
            if any("kode barang" in s.lower() for s in row_strs):
                header_found = True
                # Identify outlet columns (columns after hjual, before TOTAL)
                start_outlets = False
                for c_idx, cell_val in enumerate(row):
                    s = str(cell_val or "").strip()
                    if s.upper() == "TOTAL":
                        break
                    if "A.GUDANG" in s.upper():
                        start_outlets = True
                        continue
                    if start_outlets and s:
                        # Normalize outlet code: 'Cbg.AWIN' -> 'AWIN', 'Cbg.TJSO' -> 'TJSO'
                        clean_code = s.replace("Cbg.", "").strip()
                        outlet_cols.append((c_idx, clean_code))
                continue
            else:
                continue

        # Process data rows
        # Col 0: Supplier, Col 1: Nama/Keterangan, Col 3: Kode, Col 4: Kategori, Col 5: HPP, Col 6: HJ
        if len(row) < 8:
            continue

        nama = str(row[1] or "").strip()
        if not nama or nama.lower() in ("total", "subtotal", "grand total"):
            continue

        sku = _clean_sku(row[3] if len(row) > 3 else None)
        if not sku:
            continue

        supp = str(row[0] or "").strip()
        kategori = str(row[4] or "").strip() if len(row) > 4 else ""
        hpp = _to_float(row[5] if len(row) > 5 else 0)
        hj = _to_float(row[6] if len(row) > 6 else 0)
        stok_gudang_file = _to_float(row[7] if len(row) > 7 else 0)

        # Fallback warehouse record in case gudang file wasn't uploaded
        gudang_fallback_items.append({
            "sku_code": sku,
            "nama_barang": nama,
            "supplier": supp,
            "stok_akhir": stok_gudang_file,
            "hpp": hpp,
            "harga_jual": hj,
        })

        # Process each outlet
        for c_idx, outlet_code in outlet_cols:
            if c_idx < len(row):
                qty = _to_float(row[c_idx])
                # We store every SKU per outlet so that 0-stock situations are tracked!
                konter_items.append({
                    "sku_code": sku,
                    "nama_barang": nama,
                    "supplier": supp,
                    "kategori": kategori,
                    "hpp": hpp,
                    "harga_jual": hj,
                    "outlet_code": outlet_code,
                    "stok_qty": qty,
                })

    wb.close()
    return konter_items, gudang_fallback_items

