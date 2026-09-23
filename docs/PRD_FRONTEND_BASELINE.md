# PRD: High-Standard Frontend Architecture & Anti-AI-Slop Specification
## Ancol Store POS Analytics Platform (`PrinceRoosevelt/ancol-pos-analytics`)

| Dokumen | Keterangan |
| :--- | :--- |
| **Versi** | 1.0 (Baseline Frontend Specification) |
| **Status** | Active / Official Project Standard |
| **Target Lingkungan** | POS Touchscreen (1024x768, 1280x1024), Tablet Operasional, Laptop, & Ultra-wide Desktop |
| **Tujuan Utama** | Menghilangkan output UI generik ("AI Slop"), memastikan interaktivitas modern, fluid responsiveness, physics alami, dan performa instan (< 5ms loop). |
| **Prinsip Data** | **Zero Data Regression**: 100% data transaksi, pengunjung, dan target tetap utuh dan presisi. |

---

## 1. Problem Statement & Latar Belakang (Anti-Slop Manifesto)

Model AI secara default sering kali menghasilkan antarmuka yang seragam dan tidak memiliki karakter ("AI Slop"):
1. **Palet Klise & Murahan**: Menggunakan gradient ungu-pink (`from-purple-500 to-pink-500`) tanpa identitas visual brand.
2. **Tipografi Monoton**: Font tunggal generic tanpa display pairing, tanpa penyesuaian `tabular-nums` untuk angka finansial.
3. **Card Kaku & Datar**: Card abu-abu polos tanpa depth layer, micro-texture, atau highlight interaktif.
4. **Ketiadaan Physics Alami**: Transisi kaku (`ease-in-out` bawaan) tanpa feedback tactile (active press).
5. **Layout Rapuh**: Nilai lebar/tinggi mati (`w-[500px]`, `h-[600px]`) yang pecah pada layar mobile atau monitor POS 4:3.

**Solusi Standar Ancol POS Analytics**:
Menerapkan identitas visual korporat resmi PT Pembangunan Jaya Ancol Tbk, memadukan estetika *modern editorial dashboard* dengan performa komputasi instan berbasis browser tanpa library JS eksternal yang lambat.

---

## 2. Core Design Principles (Anti-Slop Rules)

### 2.1. Rule 1: Banned Patterns (Dilarang Digunakan)
- ❌ **Dilarang**: Gradient violet-ke-pink generik atau neon glow berlebihan.
- ❌ **Dilarang**: Menggunakan font default tanpa pairing yang jelas atau tanpa `tabular-nums` pada baris angka omzet/kuantitas.
- ❌ **Dilarang**: Flat card tanpa tactile feedback, inner-highlight, atau depth layering.
- ❌ **Dilarang**: Dimensi mati (*hardcoded pixel width*) pada elemen kontainer yang memicu *horizontal scrollbar*.
- ❌ **Dilarang**: Touch target tombol/filter di bawah 38px yang menyulitkan operasional layar sentuh kasir POS.

### 2.2. Rule 2: Ancol Official Brand Identity & Dynamic Unit Theming
Seluruh warna tema wajib merujuk pada standar resmi Ancol & Unit Rekreasi:
- **Korporat Utama**:
  - Ocean Blue: `#0033A0` (Pantone 286 C)
  - Deep Blue: `#00205B` (Pantone 281 C)
  - Sky Blue: `#00B5E2` (Pantone 306 C)
  - Green Success: `#80E0A7` (Pantone 353 C) / `#059669`
  - Yellow Accent: `#FED141` (Pantone 122 C)
- **Dynamic Recreation Unit Brand Accents**:
  - **Dufan**: Pink `#E0004D` (Pantone 1925 C / Dark: `#ff2b74`)
  - **SeaWorld**: Ocean Blue `#006EB3` (Pantone 3553 C / Dark: `#1a96e8`)
  - **Samudra**: Tosca `#00A497` (Pantone 3560 C / Dark: `#1fc7b7`)
  - **Atlantis**: Royal Purple `#5C068C` (Pantone 2597 C / Dark: `#9b3cdb`)
  - **Pasar Seni**: Tangerine `#F68D2E` (Pantone 715 C)
  - **Putri Duyung**: Amber Warm `#ED9B33` (Pantone 2011 C)

### 2.3. Rule 3: Visual Rhythm, Depth & Micro-Textures
- **Typography Hierarchy**:
  - **Display / Heading**: `Plus Jakarta Sans` dengan tracking rapat (`-0.02em` hingga `-0.03em`) dan bobot `700` atau `800`.
  - **Numbers / Currencies**: Wajib menyertakan `font-variant-numeric: tabular-nums` agar angka tidak bergeser saat data terfilter.
  - **Fluid Scaling**: Gunakan CSS `clamp()` untuk heading utama (misal: `font-size: clamp(20px, 2.2vw, 28px)`).
- **Depth & Layering**:
  - Glassmorphism terukur pada sticky filter bar dan modal (`backdrop-filter: blur(12px)`).
  - Subtle top inner-shadow / border highlight (`inset 0 1px 0 rgba(255, 255, 255, 0.9)` pada Light Mode, `rgba(255, 255, 255, 0.1)` pada Dark Mode).
  - Ambient radial background glow dengan opasitas sangat halus (`< 0.05`).

### 2.4. Rule 4: Natural Motion & Tactile Physics
- **Durasi Mikro**: 180ms – 240ms untuk interaksi klik, hover, dan dropdown toggle.
- **Easing Curve**: Wajib kurva pegas natural: `cubic-bezier(0.16, 1, 0.3, 1)`.
- **Tactile Pressed State**:
  - Seluruh tombol interaktif, pill switcher, dan card header memiliki efek tekan: `:active { transform: scale(0.98); }`.

---

## 3. Frontend Architecture & Zero-Dependency Engine

```
┌────────────────────────────────────────────────────────────────────────┐
│                   ANCOL HIGH-PERFORMANCE FRONTEND                      │
├────────────────────────────────┬───────────────────────────────────────┤
│ Core Templating & SSR          │ Flask + Jinja2 (Zero-hydration cost)  │
├────────────────────────────────┼───────────────────────────────────────┤
│ Client-side Memory Engine      │ Native Pure JS (< 5ms for 56k items)  │
├────────────────────────────────┼───────────────────────────────────────┤
│ Data Visualization Layer       │ HTML5 Canvas 2D (Sub-2ms render)      │
├────────────────────────────────┼───────────────────────────────────────┤
│ Design System & Styling        │ Modern Vanilla CSS (Variables, Grid)  │
├────────────────────────────────┼───────────────────────────────────────┤
│ Sheet Export Engine            │ SheetJS (Client-side fast generation) │
└────────────────────────────────┴───────────────────────────────────────┘
```
> **Catatan Arsitektur**: Dashboard sengaja tidak menggunakan runtime framework berat (seperti React/Next.js/Tailwind bundler) guna mempertahankan waktu pemfilteran instan (< 5ms) langsung di memori RAM browser pada PC POS berspesifikasi entry-level.

---

## 4. Functional Requirements & Responsiveness Matrix

### 4.1. Responsiveness Matrix (Termasuk Layar POS Kasir)
| Breakpoint | Target Layar | Perilaku Layout |
| :--- | :--- | :--- |
| **Mobile (`< 520px`)** | Smartphone Portrait | Single column stack, sticky mini-bar, tombol >= 42px, padding 12px. |
| **Tablet (`521px – 768px`)** | iPad / Android Tablet | 2-kolom adaptif, sticky filter bar ringkas, font fluid clamp. |
| **POS Square Monitor (`1024x768`)** | Layar Kasir Touchscreen 4:3 | Filter bar dapat di-*collapse* ke 32px mini-summary agar tabel transaksi langsung terlihat tanpa tertutup. |
| **Desktop (`1024px – 1366px`)** | Laptop / Monitor Kantor | 7-kolom filter grid, 5 KPI cards 1 baris, 6 grafik split 2 kolom. |
| **Ultra-wide (`1440px+`)** | Layar Eksekutif / Monitor 4K | Batasan kontainer `max-width: 1720px` dengan auto margin tengah agar tidak meregang berlebihan. |

### 4.2. State Completeness (5 Visual States)
Setiap elemen interaktif (Filter Dropdown, Combobox, Button, KPI Card, Table Row, Chart Pill) wajib memiliki 5 state:
1. **Default**: Tampilan tenang dengan border transparan atau subtle border (`rgba(0,0,0,0.08)` / `rgba(255,255,255,0.12)`).
2. **Hover**: Elevasi halus `translateY(-1px)` atau border glow warna brand.
3. **Active / Pressed**: Tactile feedback `transform: scale(0.98)` dengan transisi instan 100ms.
4. **Focus-visible**: Outline aksesibel 2px dengan kontras tinggi (`outline: 2px solid var(--brand)`).
5. **Disabled / Loading**: Opacity 0.45 dengan cursor not-allowed dan shimmer skeleton jika memuat data.

---

## 5. Quality Gate & Acceptance Criteria

| Checkpoint | Standar Kelulusan |
| :--- | :--- |
| **No Horizontal Scroll** | `document.documentElement.scrollWidth === window.innerWidth` di semua resolusi (100% bebas overflow). |
| **Zero Data Discrepancy** | Seluruh kalkulasi sub-segmen (Konsinyasi + Dagangan) sama persis dengan total revenue (0 Diff). |
| **Touchscreen Target** | Seluruh elemen sentuh (tombol filter, combobox, icon copy) memiliki tinggi/area sentuh `>= 38px`. |
| **Color Contrast (WCAG AA)** | Rasio kontras teks terhadap latar belakang `>= 4.5:1` pada mode terang dan gelap. |
| **Sub-5ms Execution** | Siklus eksekusi `applyFilters()` tidak melebihi 5 milidetik untuk menjaga responsivitas instan 60 FPS. |

