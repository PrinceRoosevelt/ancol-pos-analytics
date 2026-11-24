# DESIGN.md — Ancol Store POS Analytics Design System

## 1. Design Philosophy: Executive Craftsmanship, Zero AI Slop
This design system defines the authoritative visual tokens, interaction rules, and quality constraints for Ancol Store POS Analytics.

Every UI element must feel human-crafted, authoritative, and ergonomically suited for both C-level executive presentations and busy retail POS touchscreen environments.

---

## 2. Anti-AI-Slop Banned Patterns (Quality Floor)
The following design anti-patterns are **strictly forbidden** in this codebase:
- ❌ **Generic AI Gradients**: No generic purple-to-pink (`#8b5cf6` to `#ec4899`), neon cyan-to-fuchsia, or angled 135deg pastel fades.
- ❌ **Unlayered Floating Cards**: No borderless flat cards floating on ambiguous gray backgrounds without clear elevation or outline.
- ❌ **Indiscriminate Glassmorphism**: Do not put heavy `backdrop-filter: blur(...)` across multiple nested panels or data tables (causes severe GPU stutter on POS hardware). Glassmorphism is strictly confined to Top Hero KPI cards.
- ❌ **Number Jitter**: No proportional numbers in financial columns or KPI cards. Every currency, percentage, and metric must use `font-variant-numeric: tabular-nums`.
- ❌ **Emoji as Financial Icons**: Never use raw emojis (`💰`, `📈`, `🔥`) as interface indicators or table icons. Use crisp, semantic SVG icons.
- ❌ **Missing Interactive States**: No button, tab, or input may exist without all 5 states (Default, Hover, Active/Pressed, Focus-Visible, Disabled).
- ❌ **Non-Standard Font Weights**: No fractional CSS font-weights like `450`, `650`, `750`, or `850`. Use standard scale `400`, `500`, `600`, `700`, `800`.

---

## 3. Color Token System

### A. Corporate Identity & Neutral Foundation
```css
:root {
  /* Ancol Corporate Identity */
  --brand-primary:   #0033A0; /* Ancol Ocean Blue — primary actions, active tabs */
  --brand-deep:      #00205B; /* Deep Executive Navy — headers, floating controls */
  --brand-sky:       #00B5E2; /* Ancol Sky Blue — accents, secondary progress */
  --brand-light:     #EBF4FF; /* Subtle brand tint for badges and selected rows */
  
  /* Semantic Status Colors (WCAG AA Compliant) */
  --color-success:   #059669; /* Emerald Green — positive growth, healthy margins */
  --color-warning:   #D97706; /* Amber — low margin alerts, seasonal warning */
  --color-danger:    #DC2626; /* Crimson Red — negative growth, dead stock */
  --color-info:      #0284C7; /* Azure — informational tooltips */

  /* Neutral Surface & Hierarchy (Light Theme) */
  --bg-canvas:       #F8FAFC; /* Slate 50 — clean base background */
  --panel-bg:        #FFFFFF; /* Pure white solid surface for data tables & charts */
  --panel-border:    #E2E8F0; /* Slate 200 — crisp 1px structural dividing lines */
  --text-main:       #0F172A; /* Slate 900 — high contrast primary text */
  --text-muted:      #64748B; /* Slate 500 — secondary subtitles & captions */
  --text-subtle:     #94A3B8; /* Slate 400 — disabled & placeholder text */
}
```

### B. Recreational Unit Accents
Strictly use official recreational unit hex colors across filters, charts, and outlet badges:
- **Dufan (Dunia Fantasi)**: `#E0004D`
- **SeaWorld Ancol**: `#006EB3`
- **Ocean Dream Samudra**: `#00A497`
- **Atlantis Water Adventures**: `#5C068C`

---

## 4. Typography Scale & Numerical Stability

- **Primary Typeface**: `Plus Jakarta Sans`, `-apple-system`, `BlinkMacSystemFont`, `Segoe UI`, `sans-serif`.
- **Numerical Stability**:
  ```css
  .kpi-num, .stat-value, .metric-value, td.number, .tabular {
    font-variant-numeric: tabular-nums;
    letter-spacing: -0.02em;
  }
  ```
- **Type Hierarchy**:
  - `Hero KPI Figures`: `28px` to `32px` | Weight: `800` | Line-height: `1.1`
  - `Section Headers`: `18px` to `20px` | Weight: `700` | Line-height: `1.2`
  - `Card Titles / Subheaders`: `14px` to `15px` | Weight: `600`
  - `Body Text / Table Cells`: `13px` to `14px` | Weight: `400` / `500`
  - `Badges & Captions`: `11px` to `12px` | Weight: `600` | Uppercase letter-spacing `0.04em`

---

## 5. Stratified Elevation & Depth Architecture

To guarantee high GPU performance on low-spec POS machines without sacrificing executive polish:
1. **Layer 0 (Canvas)**: `#F8FAFC` background.
2. **Layer 1 (Data Panels & Tables)**: Solid `#FFFFFF` background, `border: 1px solid var(--panel-border)`, subtle shadow `0 1px 3px rgba(15, 23, 42, 0.05)`. **Zero blur filter**.
3. **Layer 2 (Hero KPI Cards & Filter Bar)**: Elevated with subtle stratified glassmorphism:
   ```css
   .glass-kpi-card {
     background: rgba(255, 255, 255, 0.88);
     backdrop-filter: blur(20px);
     -webkit-backdrop-filter: blur(20px);
     border: 1.5px solid rgba(255, 255, 255, 0.95);
     box-shadow: 0 4px 16px -2px rgba(0, 51, 160, 0.06), 0 2px 6px -1px rgba(0, 0, 0, 0.04);
   }
   ```
4. **Layer 3 (Modals & Sticky Drawers)**: High z-index surfaces with solid or Deep Navy header (`#00205B`) + Sky Blue accent bar (`#00B5E2`).

---

## 6. Micro-Physics & Tactile Interaction

All interactive elements must provide immediate physical confirmation to touch/click:
- **Spring Transition Curve**:
  ```css
  transition: all 200ms cubic-bezier(0.16, 1, 0.3, 1);
  ```
- **Tactile Active Feedback (Touch & Click)**:
  ```css
  button:active, .interactive-chip:active, .tab-item:active {
    transform: scale(0.98);
  }
  ```
- **Focus-Visible Requirement**:
  ```css
  :focus-visible {
    outline: 2px solid var(--brand-sky);
    outline-offset: 2px;
  }
  ```
- **Touch Target Floor**:
  All clickable pills, combobox triggers, and buttons must enforce `min-height: 38px` and `min-width: 38px`.

---

## 7. Responsive Breakpoint Rules
- **Desktop Executive (≥1440px)**: 4-column KPI grid, dual chart rows, full 7-dimension filter bar.
- **Laptop & Tablet Landscape (1024px – 1439px)**: 2-column or 4-column responsive grid, collapsible filter drawers, sticky summary chips.
- **POS Screen Standard (1024x768 4:3)**: Compact spacing, collapsible filter bar, maximum viewport real estate reserved for data tables.
- **Zero Horizontal Overflow Rule**:
  Containers must use `max-width: 100%`, flex wrapping with `gap`, and tables must use dedicated horizontal scroll wrappers (`.table-responsive`) with sticky headers, never causing root window overflow.
