# QUALITY.md — Engineering & Web Quality Standard

## 1. Principles of Measurement-Led Quality
Inspired by Addy Osmani's `web-quality-skills` and Google Chrome performance engineering, quality in the **Ancol Store POS Analytics** platform is measured with objective evidence, not subjective feelings.

Every code change must pass performance, accessibility, ergonomic, and data integrity gates before release.

---

## 2. Core Web Vitals & Performance Benchmarks

| Metric | Target Floor | Verification Method |
|---|---|---|
| **INP (Interaction to Next Paint)** | `< 150ms` (Good < 200ms) | Instant tactile feedback on filter clicks and tab switches |
| **LCP (Largest Contentful Paint)** | `< 2.2s` (Good < 2.5s) | Fast template rendering and optimized CSS delivery |
| **CLS (Cumulative Layout Shift)** | `< 0.05` (Good < 0.1) | Fixed height containers for charts, KPI skeleton loaders, no shifting banners |
| **Client Array Filtering** | `< 5ms` for ~56k records | High-performance in-memory filter loops in `applyFilters()` |
| **View Aggregation Latency** | `< 50ms` | SQLite `v_sales_analytics` indexed view queries |

---

## 3. Accessibility & Touch Ergonomics (WCAG AA)

- **Color Contrast Ratio**:
  - Body text (`#0F172A` on `#FFFFFF` / `#F8FAFC`): Contrast ratio **≥ 14:1** (Far exceeds WCAG AA 4.5:1).
  - Secondary text (`#64748B` on `#FFFFFF`): Contrast ratio **≥ 4.6:1** (Meets WCAG AA).
  - Badges & Pills: Contrast ratio between badge text and tinted badge background must be **≥ 4.5:1**.
- **Touch Target Sizing**:
  - Minimum touch target dimension: **38px × 38px** (POS terminal touch standard).
  - Spacing between adjacent touch targets: Minimum **8px** gap to prevent accidental clicks.
- **Keyboard Navigation & Focus**:
  - All interactive elements must show a distinct `2px` focus outline (`--brand-sky`) when navigated via keyboard Tab or barcode scanner inputs.

---

## 4. Data Rigor & Zero Discrepancy Protocol

Before any release or commit affecting calculations:
1. **The 0 Diff Formula**:
   $$\text{Total Revenue} - (\text{Revenue Konsinyasi} + \text{Revenue Dagangan}) = 0$$
2. **Gross Margin Bounds**:
   Overall gross margin percentage must remain between expected retail margins (typically 25% – 60%). Any sudden spike to 0% or >90% indicates a broken cost-of-goods (HPP) join.
3. **Apples-to-Apples YoY**:
   Calendar days compared in 2026 and 2025 must strictly match using the `comparison_day_months` alignment logic.

---

## 5. Pre-Commit & Pre-Push Verification Checklist

Before running `git commit` and `git push origin main`:
- [ ] **Python Syntax Check**: `python -m py_compile main.py server.py` exits with code 0.
- [ ] **Flask Live Rendering**: `with app.test_client() as client: res = client.get('/')` returns `HTTP 200 OK`.
- [ ] **Zero Horizontal Overflow**: `document.documentElement.scrollWidth === window.innerWidth` verified.
- [ ] **Numerical Tabular Stability**: All currency, percentage, and count displays have `tabular-nums` active.
- [ ] **Zero Regression**: No existing working charts, filters, or WhatsApp sharing features are mutated or broken.
- [ ] **Git Push**: All committed code immediately pushed to `origin/main`.
