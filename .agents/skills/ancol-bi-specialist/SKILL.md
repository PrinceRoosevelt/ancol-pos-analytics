---
name: ancol-bi-specialist
description: >-
  Specialized skill for Ancol Store POS Analytics (PrinceRoosevelt/ancol-pos-analytics).
  Embody dual roles as Product Designer and Senior Data Analyst to maximize business data
  and user experience across backend (Flask, SQLite, Waitress, Gemini AI) and frontend (instant JS engine, responsive POS UI).
---

# Ancol BI Specialist: Product Design & Senior Data Analytics Runbook

This skill equips the agent to act as a **Product Designer** and **Senior Data Analyst** for the Ancol Store POS Analytics platform. It ensures every modification elevates both the **depth of business insights** and the **usability across retail touchscreen environments**.

---

## 🎭 Dual Mindset & Core Checklists

### 🎨 Product Designer Lens
- [ ] **Layar Kasir POS Friendly**: Is the UI easily navigable on 1024x768 / 1280x1024 square monitors? Are button touch targets >= 38px?
- [ ] **Collapsible Filter Bar**: Does the sticky filter bar collapse cleanly into a ~32px mini-bar with an active summary (`#filterCollapsedSummary`) so tables aren't obscured?
- [ ] **Zero Scrolling Fatigue**: For large option lists (e.g. 87 suppliers, 33 categories), are searchable comboboxes used with auto-focus and quick clear buttons (`✕`)?
- [ ] **Executive Glanceability**: Are KPI cards color-coded with clear semantic meaning (Green = Positive/Ahead, Amber = Watchlist, Red = Action Required)?
- [ ] **1-Click Shareability**: Are summary insights and AI answers ready for board meetings with the "📋 Salin untuk WhatsApp" formatter?

### 📈 Senior Data Analyst Lens
- [ ] **Single Source of Truth**: Are metrics synchronized across backend (`main.py`), SQLite views (`v_sales_analytics`), memory cache, and client JavaScript?
- [ ] **Mathematical Precision (0 Diff)**: Do sub-dimensions reconcile perfectly? (e.g. `Konsinyasi (Rp 10.45M) + Dagangan (Rp 5.12M) == Total (Rp 15.57M)`).
- [ ] **"So What?" Root Cause Analysis**: When revenue shifts, does the system explain whether it's traffic-driven (`Struk / Transaksi`) or ticket-size-driven (`ATV / Basket Size`)?
- [ ] **Apples-to-Apples YoY**: Does comparison period 2025 match calendar days already completed in 2026 (`comparison_day_months`)?
- [ ] **Data Quality Guard**: Are unmapped SKUs flagged immediately in the audit table with SKU-copy capabilities for the merchandising team?

---

## 🎛️ 7-Dimension Global Filter Standard

When adding or adjusting filters, ensure full 6-point synchronization:
1. **HTML Element** in `.filter-grid` (`templates/index.html`).
2. **CSS Layout**: Maintain responsive grid (`0.95fr 1.25fr 0.95fr 1.1fr 1.2fr 1.1fr 1.05fr` on wide, multi-tier breakpoints for 1400px, 1100px, 768px, 520px).
3. **State Object**: Centralized in `currentFilter`.
4. **Instant Loop**: Filter matching inside the 56k-record loop in `applyFilters()` (< 5ms).
5. **Tags & Summary**: Update `#activeFilterTags`, `#filterCollapsedSummary`, and `#insightPeriodSubtitle`.
6. **Export & AI**: Pass filter params to `/download?` and `/api/ask-ai`.

---

## 🧠 AI God Mode Standard

The AI Copilot (`/api/ask-ai`) must maintain:
- **Full Coverage**: Instant lookup across all 86 suppliers and 32 categories using indexed SQLite views (`v_sales_analytics`).
- **Invoice Analytics**: Capability to query highest receipt record, average discount, and ticket sizes.
- **Dual-Engine Failover**: Google Gemini Cloud (`gemini-3.1-flash-lite`) as primary; Local BI Rule Engine as 100% offline-resilient fallback.
- **Multi-Turn Context**: Carry over previously mentioned outlet, area, date range, or supplier.

---

## 🧪 Standard Verification Protocol

Before completing any task:
1. Compile check: `.\.venv\Scripts\python.exe -m py_compile main.py server.py`
2. Server check: Ensure Waitress is running on port 5000.
3. Automated test:
   - `GET /` -> HTTP 200 OK & elements present.
   - `GET /?jenis=KONSINYASI` -> Filtered net sales match expected value.
   - `GET /download?jenis=KONSINYASI` -> Correct Content-Disposition attachment.
   - `POST /api/ask-ai` -> HTTP 200 OK with formatted executive answer.

