# PRODUCT.md — Ancol Store POS Analytics

## 1. Product Truth & Mission
**Ancol Store POS Analytics** is the mission-critical business intelligence and revenue analytics platform for **PT Pembangunan Jaya Ancol, Tbk** retail merchandise operations across all recreational units (Dufan, SeaWorld, Ocean Dream Samudra, Atlantis Water Adventures).

The system serves as the single source of truth for executive revenue monitoring, supplier revenue-share reconciliation, inventory velocity, and frontline point-of-sale operational decisions.

---

## 2. Surface Mode: `Operate`
Under the Impeccable framework, this surface is strictly classified as **Operate**:
- **Primary Goal**: The user completes critical analytical, auditing, and operational tasks.
- **Priority**: High glanceability, scanability, consistency, tabular numerical stability, and tactile feedback outrank decorative or expressive elements.
- **Brand Expression**: Subtle, authoritative executive restraint using official Ancol corporate colors and theme park accents.

---

## 3. Target User Personas

| Persona | Primary Goal | Critical Needs |
|---|---|---|
| **C-Level Executive / Direksi** | Glanceable macro revenue & margin health monitoring | • Instant Gross vs Net Profit<br>• Apples-to-apples YoY growth (2026 vs 2025)<br>• 1-Click WhatsApp Executive Summary |
| **Store & Merchandising Managers** | Partner reconciliation & assortment optimization | • Supplier revenue share (Konsinyasi vs Dagangan)<br>• High-velocity vs dead stock identification<br>• Vendor concentration and margin mix |
| **POS Supervisors & Cashiers** | Frontline operational troubleshooting | • 1024x768 touchscreen ergonomic targets<br>• Fast searchable combobox filters (<5ms)<br>• Instant reset without lag |

---

## 4. Hardware & Operating Constraints

- **Touchscreen Monitor Standard**: Optimized for real-world POS cash register displays (1024x768 resolution, 4:3 square aspect ratio) up to 4K ultra-wide executive monitors.
- **Touch Target Ergonomics**: All interactive elements (filter chips, buttons, combobox triggers, date selectors) must have minimum touch targets of **≥ 38px** (ideal ≥ 44px).
- **Strict Zero Horizontal Overflow**: 
  `document.documentElement.scrollWidth === window.innerWidth` at all viewport sizes. No horizontal scrollbars under any circumstance.
- **Offline & Low-Bandwidth Resilience**:
  - Backend runs on local Waitress WSGI on `0.0.0.0:5000`.
  - When Cloud Gemini AI is offline/unreachable, the **Local BI Rule Engine** automatically provides instant, offline-resilient responses with guaranteed HTTP 200 OK.
  - Zero CDN dead-ends: Core CSS and JS fallbacks ensure dashboard renders even if external CDNs experience network drops.

---

## 5. Non-Negotiable Business Logic & Data Integrity

1. **Zero Discrepancy Commitment**:
   - Total Revenue = Konsinyasi Revenue + Dagangan Revenue (Strictly 0 Diff).
   - Any financial discrepancy between macro KPI cards, segment filters, and exported Excel sheets is considered a critical bug.
2. **Apples-to-Apples YoY Calendar Alignment**:
   - Comparison between 2026 and 2025 must strictly align on calendar day-matching (`comparison_day_months`) to eliminate monthly cutoff and leap-year distortions.
3. **Master Data Single Source of Truth**:
   - Master SKU catalog (~5,945 items) is anchored in `config/DATA_LOOKUP.xlsx` with clean SKU code, supplier name, category, business model (`KONSINYASI` vs `DAGANGAN`), HPP, and Harga Jual.
   - Database layer indexed view `v_sales_analytics` in SQLite `sales_data.db`.
4. **Client-Side High-Speed Engine**:
   - `RAW_SALES` array (~56,000 transaction line items) must filter in memory in `< 5ms`.
   - Filter state must remain centralized in `currentFilter` across all 7 dimensions.

---

## 6. Recreational Unit Accents
Each unit has an assigned official brand color:
- **Dufan (Dunia Fantasi)**: `#E0004D`
- **SeaWorld Ancol**: `#006EB3`
- **Ocean Dream Samudra**: `#00A497`
- **Atlantis Water Adventures**: `#5C068C`
- **Corporate Ancol Primary**: `#0033A0` (Ocean Blue) and `#00205B` (Deep Blue)
