"""
inventory_engine/analytics.py
Core Analytical Brain for Inventory Intelligence:
1. Store Priority Matrix (Store ranking by velocity, revenue, and throughput)
2. Restock Recommendations (Safety stock buffer, warehouse-to-counter allocation)
3. Critical & Stock-Out Alerts (Zero stock high velocity, <48h runway, new SKUs)
4. Overstock & Smart Cross-Transfer Rebalancing (Inter-outlet transfers + tactical retail advice)
"""
from __future__ import annotations

import math
from collections import defaultdict
from typing import Any, Dict, List, Optional, Tuple

from database.db_config import get_connection
from inventory_engine.repository import (
    get_snapshot_gudang_data,
    get_snapshot_konter_data,
    get_snapshot_meta,
)

# Canonical 22 counter codes with friendly metadata
OUTLET_MASTER: Dict[str, Dict[str, Any]] = {
    "AWIN": {"name": "AWIN Atlantis Induk", "area": "AWAPARK", "tier": 1},
    "AWKL": {"name": "AWKL AWA Taman Kelapa 2", "area": "AWAPARK", "tier": 2},
    "AWTK": {"name": "AWTK Toko Atlantis", "area": "AWAPARK", "tier": 3},
    "DFAR": {"name": "DFAR Dufan Arung Jeram", "area": "DUFAN", "tier": 1},
    "DFGA": {"name": "DFGA Dufan Galactica", "area": "DUFAN", "tier": 2},
    "DFIC": {"name": "DFIC Dufan Ice Age", "area": "DUFAN", "tier": 2},
    "DFIL": {"name": "DFIL Dufan Induk Lama", "area": "DUFAN", "tier": 1},
    "DFIN": {"name": "DFIN Dufan Induk", "area": "DUFAN", "tier": 1},
    "DFKE": {"name": "DFKE Dufan Kereta Misteri", "area": "DUFAN", "tier": 1},
    "DFOR": {"name": "DFOR Dufan Oriental", "area": "DUFAN", "tier": 2},
    "DFSI": {"name": "DFSI Dufan Simulator", "area": "DUFAN", "tier": 2},
    "DFTO": {"name": "DFTO Dufan Tornado", "area": "DUFAN", "tier": 2},
    "DFWW": {"name": "DFWW Dufan WWN", "area": "DUFAN", "tier": 1},
    "ICG":  {"name": "ICG Ice Cream Galactica", "area": "DUFAN", "tier": 3},
    "JBIN": {"name": "JBIN JBL Induk", "area": "SAMUDRA", "tier": 2},
    "ODIN": {"name": "ODIN Samudra Induk", "area": "SAMUDRA", "tier": 1},
    "OL01": {"name": "OL01 Online Shop", "area": "BEACHPARK", "tier": 2},
    "SWIN": {"name": "SWIN Sea World Induk", "area": "SEAWORLD", "tier": 1},
    "TJBK": {"name": "TJBK Toko Dermaga Beachpark", "area": "BEACHPARK", "tier": 3},
    "TJBM": {"name": "TJBM Toko Symphony Merchandise", "area": "BEACHPARK", "tier": 3},
    "TJOM": {"name": "TJOM Merchandise Ombak Laut", "area": "BEACHPARK", "tier": 2},
    "TJSO": {"name": "TJSO Symphony Of The Sea", "area": "BEACHPARK", "tier": 1},
}


def get_outlet_sales_performance() -> List[Dict[str, Any]]:
    """
    Ranks all 22 outlets by sales performance, daily unit velocity, and gross revenue.
    """
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT 
                SUBSTR(outlet, 1, 4) as outlet_code,
                outlet as full_name,
                COUNT(DISTINCT invoice) as total_invoices,
                COUNT(DISTINCT date) as active_days,
                SUM(qty) as total_units_sold,
                SUM(net_sales) as total_revenue,
                ROUND(SUM(qty) * 1.0 / MAX(1, COUNT(DISTINCT date)), 1) as avg_daily_units,
                ROUND(SUM(net_sales) * 1.0 / MAX(1, COUNT(DISTINCT date)), 0) as avg_daily_revenue
            FROM v_sales_analytics
            WHERE outlet IS NOT NULL AND outlet != ''
            GROUP BY SUBSTR(outlet, 1, 4)
            ORDER BY total_units_sold DESC
        """)
        rows = cursor.fetchall()
        perf_map = {r["outlet_code"]: dict(r) for r in rows}

    results = []
    for code, meta in OUTLET_MASTER.items():
        db_stat = perf_map.get(code, {})
        tot_units = float(db_stat.get("total_units_sold") or 0.0)
        tot_rev = float(db_stat.get("total_revenue") or 0.0)
        act_days = int(db_stat.get("active_days") or 0)
        daily_u = float(db_stat.get("avg_daily_units") or 0.0)
        daily_r = float(db_stat.get("avg_daily_revenue") or 0.0)

        # Priority tier scoring
        if daily_u >= 100 or tot_units >= 50000:
            tier_badge = "TIER 1 (ULTRA HUB)"
            priority_score = 95
            tier_color = "#ef4444"
        elif daily_u >= 30 or tot_units >= 15000:
            tier_badge = "TIER 2 (REGULER RAMAI)"
            priority_score = 75
            tier_color = "#3b82f6"
        elif daily_u > 0:
            tier_badge = "TIER 3 (MODERAT)"
            priority_score = 50
            tier_color = "#10b981"
        else:
            tier_badge = "TIER 4 (NON-AKTIF / STORAGE)"
            priority_score = 20
            tier_color = "#6b7280"

        results.append({
            "outlet_code": code,
            "outlet_name": meta["name"],
            "area": meta["area"],
            "total_units_sold": tot_units,
            "total_revenue": tot_rev,
            "active_days": act_days,
            "avg_daily_units": daily_u,
            "avg_daily_revenue": daily_r,
            "tier_badge": tier_badge,
            "priority_score": priority_score,
            "tier_color": tier_color,
        })

    # Sort descending by priority score and units sold
    results.sort(key=lambda x: (x["priority_score"], x["total_units_sold"]), reverse=True)
    for rank, item in enumerate(results, start=1):
        item["rank"] = rank

    return results


def get_sku_velocity_map() -> Dict[Tuple[str, str], float]:
    """
    Computes daily sales velocity for every (sku_code, outlet_code) pair across all history.
    Uses product lifespan to prevent diluting newly introduced SKUs.
    """
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT 
                code as sku_code,
                SUBSTR(outlet, 1, 4) as outlet_code,
                SUM(qty) as total_qty,
                MIN(date) as first_sale,
                MAX(date) as last_sale,
                MAX(14, CAST(ROUND(julianday(MAX(date)) - julianday(MIN(date)) + 1) AS INT)) as lifespan_days,
                ROUND(SUM(qty) * 1.0 / MAX(14, CAST(ROUND(julianday(MAX(date)) - julianday(MIN(date)) + 1) AS INT)), 2) as daily_velocity
            FROM v_sales_analytics
            WHERE code IS NOT NULL AND code != ''
            GROUP BY code, SUBSTR(outlet, 1, 4)
        """)
        rows = cursor.fetchall()
        velocity_map = {}
        for r in rows:
            sku = str(r["sku_code"]).strip()
            outlet = str(r["outlet_code"]).strip()
            velocity_map[(sku, outlet)] = float(r["daily_velocity"] or 0.0)
        return velocity_map


def get_sku_network_velocity_map() -> Dict[str, float]:
    """
    Computes aggregate network-wide daily sales velocity for each SKU.
    """
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT 
                code as sku_code,
                SUM(qty) as total_qty,
                MAX(14, CAST(ROUND(julianday(MAX(date)) - julianday(MIN(date)) + 1) AS INT)) as lifespan_days,
                ROUND(SUM(qty) * 1.0 / MAX(14, CAST(ROUND(julianday(MAX(date)) - julianday(MIN(date)) + 1) AS INT)), 2) as network_velocity
            FROM v_sales_analytics
            WHERE code IS NOT NULL AND code != ''
            GROUP BY code
        """)
        rows = cursor.fetchall()
        return {str(r["sku_code"]).strip(): float(r["network_velocity"] or 0.0) for r in rows}


def run_inventory_engine(snapshot_id: int, safety_days: int = 7) -> Dict[str, Any]:
    """
    Executes all 4 analytics engines on the specified snapshot:
    1. Store Priority Matrix
    2. Restock Recommendations (Warehouse -> Counter)
    3. Critical & Stock-Out Alerts
    4. Overstock & Smart Cross-Transfer Advisory
    """
    meta = get_snapshot_meta(snapshot_id)
    if not meta:
        raise ValueError(f"Snapshot ID {snapshot_id} tidak ditemukan.")

    gudang_rows = get_snapshot_gudang_data(snapshot_id)
    konter_rows = get_snapshot_konter_data(snapshot_id)

    # 1. Master lookup maps
    # Map warehouse stock: sku_code -> {stok_akhir, hpp, harga_jual, nama_barang, supplier}
    gudang_stock_map: Dict[str, Dict[str, Any]] = {}
    for g in gudang_rows:
        sku = str(g["sku_code"]).strip()
        gudang_stock_map[sku] = {
            "stok_akhir": float(g["stok_akhir"] or 0.0),
            "hpp": float(g["hpp"] or 0.0),
            "harga_jual": float(g["harga_jual"] or 0.0),
            "nama_barang": g["nama_barang"] or "",
            "supplier": g["supplier"] or "",
        }

    # Outlet performance & tier ranking
    store_matrix = get_outlet_sales_performance()
    outlet_tier_map = {item["outlet_code"]: item["priority_score"] for item in store_matrix}
    outlet_name_map = {item["outlet_code"]: item["outlet_name"] for item in store_matrix}

    # SKU velocity data
    sku_outlet_velocity = get_sku_velocity_map()
    sku_network_velocity = get_sku_network_velocity_map()

    # Track remaining warehouse stock during allocation
    warehouse_remaining: Dict[str, float] = {
        sku: data["stok_akhir"] for sku, data in gudang_stock_map.items()
    }

    # Data structures for 4 engines
    restock_candidates = []
    critical_alerts = []
    stockout_alerts = []
    new_untracked_skus = []
    overstocked_items = []
    cross_transfers = []

    # Pre-organize counter items by SKU for cross-counter matching
    # sku -> list of {outlet_code, stok_qty, velocity, days_cov, nama_barang, supplier, etc.}
    sku_counter_matrix: Dict[str, List[Dict[str, Any]]] = defaultdict(list)

    for k in konter_rows:
        sku = str(k["sku_code"]).strip()
        outlet = str(k["outlet_code"]).strip()
        stok_qty = float(k["stok_qty"] or 0.0)
        nama = k["nama_barang"] or ""
        supp = k["supplier"] or ""
        kat = k["kategori"] or ""
        hpp = float(k["hpp"] or 0.0)
        hj = float(k["harga_jual"] or 0.0)

        # Sales velocity for this specific counter
        vel = sku_outlet_velocity.get((sku, outlet), 0.0)
        net_vel = sku_network_velocity.get(sku, 0.0)

        # Days of coverage left
        if vel > 0:
            days_coverage = round(stok_qty / vel, 1)
        else:
            days_coverage = 999.0 if stok_qty > 0 else 0.0

        item_state = {
            "sku_code": sku,
            "nama_barang": nama,
            "supplier": supp,
            "kategori": kat,
            "hpp": hpp,
            "harga_jual": hj,
            "outlet_code": outlet,
            "outlet_name": outlet_name_map.get(outlet, f"Konter {outlet}"),
            "stok_qty": stok_qty,
            "daily_velocity": vel,
            "network_velocity": net_vel,
            "days_coverage": days_coverage,
            "priority_score": outlet_tier_map.get(outlet, 50),
        }
        sku_counter_matrix[sku].append(item_state)

        # --- ENGINE 3: CRITICAL & STOCKOUT ALERTS ---
        if vel > 0.05:  # Item is actively selling at this counter
            if stok_qty == 0:
                stockout_alerts.append({
                    **item_state,
                    "urgency": "STOCKOUT",
                    "urgency_label": "HABIS TOTAL (Kehilangan Omset)",
                    "lost_daily_sales_est": round(vel * hj, 0),
                    "badge_color": "#ef4444",
                })
            elif days_coverage < 2.5:
                critical_alerts.append({
                    **item_state,
                    "urgency": "CRITICAL",
                    "urgency_label": f"KRITIS (< {days_coverage} Hari)",
                    "runway_hours": round(days_coverage * 24, 0),
                    "badge_color": "#f59e0b",
                })
        elif net_vel == 0 and vel == 0:
            # Check if this SKU has never been recorded in sales
            if stok_qty > 0 or gudang_stock_map.get(sku, {}).get("stok_akhir", 0) > 0:
                # Flag as new SKU once per SKU
                if not any(x["sku_code"] == sku for x in new_untracked_skus):
                    new_untracked_skus.append({
                        "sku_code": sku,
                        "nama_barang": nama,
                        "supplier": supp,
                        "stok_gudang": gudang_stock_map.get(sku, {}).get("stok_akhir", 0.0),
                        "harga_jual": hj,
                        "status": "BARANG BARU / BELUM ADA DATA JUAL",
                    })

        # --- ENGINE 2 CANDIDATE: RESTOCK TARGET ---
        if vel > 0.05:
            # Target stock = ceil(velocity * safety_days), buffer minimum 2 pcs
            target_stock = max(2, math.ceil(vel * safety_days))
            deficit = max(0, target_stock - stok_qty)
            if deficit > 0:
                restock_candidates.append({
                    **item_state,
                    "target_stock": target_stock,
                    "deficit": deficit,
                    "urgency_ratio": deficit / target_stock,
                })

        # --- ENGINE 4 CANDIDATE: OVERSTOCK DETECTION ---
        # Overstock defined as > 21 days coverage (with at least 5 pcs) OR > 10 pcs with zero sales
        if (vel > 0 and days_coverage > 21.0 and stok_qty >= 5) or (vel == 0 and stok_qty >= 8):
            overstocked_items.append({
                **item_state,
                "excess_units": max(1, math.ceil(stok_qty - (vel * safety_days))) if vel > 0 else stok_qty,
            })

    # --- PROCESS ENGINE 2: RESTOCK RECOMMENDATION ALLOCATION ---
    # Sort candidates by:
    # 1. Highest priority score of outlet (Tier 1 first)
    # 2. Highest deficit urgency ratio
    # 3. Highest daily velocity
    restock_candidates.sort(
        key=lambda x: (x["priority_score"], x["urgency_ratio"], x["daily_velocity"]),
        reverse=True,
    )

    restock_recommendations = []
    for cand in restock_candidates:
        sku = cand["sku_code"]
        deficit = cand["deficit"]
        avail_gudang = warehouse_remaining.get(sku, 0.0)

        if avail_gudang > 0:
            send_qty = min(deficit, avail_gudang)
            warehouse_remaining[sku] = avail_gudang - send_qty
            status = "DAPAT DIPENUHI PENUH" if send_qty == deficit else "DIPENUHI SEBAGIAN (STOK GUDANG TERBATAS)"
        else:
            send_qty = 0
            status = "STOK GUDANG KOSONG (PERLU ORDER KE SUPPLIER)"

        restock_recommendations.append({
            "sku_code": cand["sku_code"],
            "nama_barang": cand["nama_barang"],
            "supplier": cand["supplier"],
            "outlet_code": cand["outlet_code"],
            "outlet_name": cand["outlet_name"],
            "current_stock": cand["stok_qty"],
            "daily_velocity": cand["daily_velocity"],
            "target_stock": cand["target_stock"],
            "deficit": deficit,
            "recommended_send_qty": int(send_qty),
            "gudang_stock_initial": gudang_stock_map.get(sku, {}).get("stok_akhir", 0.0),
            "gudang_stock_left": warehouse_remaining.get(sku, 0.0),
            "status": status,
            "priority_score": cand["priority_score"],
        })

    # Filter restock recommendations: prioritize items with actual send_qty > 0 or critical deficit
    restock_actionable = [r for r in restock_recommendations if r["recommended_send_qty"] > 0]
    restock_out_of_warehouse = [r for r in restock_recommendations if r["recommended_send_qty"] == 0 and r["deficit"] > 0]

    # --- PROCESS ENGINE 4: CROSS-TRANSFER REBALANCING ---
    # Match overstocked counters with deficit counters for the exact same SKU
    for over in overstocked_items:
        sku = over["sku_code"]
        excess = over["excess_units"]
        if excess <= 0:
            continue

        # Look at all counters for this SKU that have deficit or 0 stock
        other_counters = [
            c for c in sku_counter_matrix.get(sku, [])
            if c["outlet_code"] != over["outlet_code"] and c["daily_velocity"] > 0.1
        ]
        # Sort other counters: lowest coverage and highest priority outlet first
        other_counters.sort(key=lambda x: (x["stok_qty"] == 0, x["priority_score"]), reverse=True)

        for target in other_counters:
            if excess <= 0:
                break
            target_stock_need = max(2, math.ceil(target["daily_velocity"] * safety_days))
            target_deficit = max(0, target_stock_need - target["stok_qty"])

            if target_deficit > 0:
                transfer_qty = min(excess, target_deficit)
                excess -= transfer_qty
                cross_transfers.append({
                    "sku_code": sku,
                    "nama_barang": over["nama_barang"],
                    "supplier": over["supplier"],
                    "from_outlet_code": over["outlet_code"],
                    "from_outlet_name": over["outlet_name"],
                    "from_current_stock": over["stok_qty"],
                    "from_velocity": over["daily_velocity"],
                    "to_outlet_code": target["outlet_code"],
                    "to_outlet_name": target["outlet_name"],
                    "to_current_stock": target["stok_qty"],
                    "to_velocity": target["daily_velocity"],
                    "transfer_qty": int(transfer_qty),
                    "reason": f"Konter {over['outlet_code']} overstock ({over['days_coverage']} hari), sedangkan Konter {target['outlet_code']} butuh {target_deficit} pcs.",
                })

    # --- OVERSTOCK STRATEGIC RETAIL ADVICE ---
    # Aggregate overstocked SKUs and provide concrete commercial action
    overstock_advice_list = []
    # Deduplicate overstock by SKU
    overstock_by_sku: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for o in overstocked_items:
        overstock_by_sku[o["sku_code"]].append(o)

    for sku, records in overstock_by_sku.items():
        total_excess = sum(r["excess_units"] for r in records)
        first_r = records[0]
        supp = first_r["supplier"]
        kat = first_r["kategori"]
        nama = first_r["nama_barang"]
        net_vel = first_r["network_velocity"]
        hj = first_r["harga_jual"]
        hpp = first_r["hpp"]
        affected_outlets = ", ".join(r["outlet_code"] for r in records[:5])
        if len(records) > 5:
            affected_outlets += f" (+{len(records)-5} konter lain)"

        # Determine tactical action
        if "konsinyasi" in kat.lower():
            action_code = "RETUR_KONSINYASI"
            action_title = "📦 Opsi Retur Supplier (Konsinyasi)"
            action_detail = f"Produk Konsinyasi perputaran lambat ({net_vel} pcs/hari). Ajukan retur ke supplier {supp} untuk mengosongkan rak dan etalase."
            action_color = "#8b5cf6"
        elif net_vel == 0:
            action_code = "FREEZE_PO_HOTZONE"
            action_title = "🛑 Freeze PO & Pajang di Hot-Zone Kasir"
            action_detail = f"Tidak ada penjualan tercatat. Bekukan pembelian ke supplier {supp}. Pindahkan sisa {total_excess} pcs ke kasir untuk memicu impulse purchase."
            action_color = "#ef4444"
        elif hpp > 0 and (hj - hpp) / max(1, hj) >= 0.40:
            action_code = "PROMO_BUNDLING"
            action_title = "🏷️ Rekomendasi Bundling / GWP"
            action_detail = f"Margin tebal ({round((hj-hpp)/hj*100)}%). Bundling dengan produk favorit kategori sejenis untuk mempercepat perputaran barang."
            action_color = "#10b981"
        else:
            action_code = "CLEARANCE_SALE"
            action_title = "📉 Clearance / Diskon Bertahap"
            action_detail = f"Stok menumpuk {total_excess} pcs. Pertimbangkan potongan harga 10-20% pada akhir pekan untuk likuidasi aset."
            action_color = "#f59e0b"

        overstock_advice_list.append({
            "sku_code": sku,
            "nama_barang": nama,
            "supplier": supp,
            "kategori": kat,
            "total_excess_units": total_excess,
            "affected_outlets": affected_outlets,
            "network_velocity": net_vel,
            "action_code": action_code,
            "action_title": action_title,
            "action_detail": action_detail,
            "action_color": action_color,
        })

    overstock_advice_list.sort(key=lambda x: x["total_excess_units"], reverse=True)

    # Sort alerts
    stockout_alerts.sort(key=lambda x: (x["priority_score"], x["daily_velocity"]), reverse=True)
    critical_alerts.sort(key=lambda x: (x["priority_score"], x["runway_hours"]), reverse=False)

    # Calculate executive KPIs
    total_restock_pcs = sum(r["recommended_send_qty"] for r in restock_actionable)
    total_transfer_pcs = sum(t["transfer_qty"] for t in cross_transfers)
    total_stockout_skus = len(stockout_alerts)
    total_critical_skus = len(critical_alerts)
    total_overstocked_skus = len(overstock_advice_list)

    return {
        "meta": meta,
        "kpis": {
            "total_stockout_alerts": total_stockout_skus,
            "total_critical_alerts": total_critical_skus,
            "total_restock_pcs": total_restock_pcs,
            "total_restock_skus": len(restock_actionable),
            "total_transfer_pcs": total_transfer_pcs,
            "total_transfer_recommendations": len(cross_transfers),
            "total_overstocked_skus": total_overstocked_skus,
            "total_new_skus": len(new_untracked_skus),
        },
        "store_priority_matrix": store_matrix,
        "stockout_alerts": stockout_alerts[:150],
        "critical_alerts": critical_alerts[:150],
        "restock_actionable": restock_actionable[:200],
        "restock_supplier_needed": restock_out_of_warehouse[:100],
        "cross_transfers": cross_transfers[:150],
        "overstock_advisory": overstock_advice_list[:150],
        "new_untracked_skus": new_untracked_skus[:50],
    }

