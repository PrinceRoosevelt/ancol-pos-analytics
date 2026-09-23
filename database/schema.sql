-- Skema Database SQL untuk Sales & Traffic Analytics (SQLite WAL Mode)

CREATE TABLE IF NOT EXISTS sync_meta (
    file_path TEXT PRIMARY KEY,
    file_type TEXT NOT NULL,
    file_mtime INTEGER NOT NULL,
    file_size INTEGER NOT NULL,
    last_synced TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS sales_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    file_source TEXT NOT NULL,
    year INTEGER NOT NULL,
    date TEXT NOT NULL,         -- YYYY-MM-DD
    month TEXT NOT NULL,        -- YYYY-MM
    hour TEXT NOT NULL,         -- HH (00-23)
    invoice TEXT,
    outlet TEXT NOT NULL,
    area TEXT NOT NULL,
    product TEXT NOT NULL,
    qty REAL NOT NULL,
    item_net_sales REAL NOT NULL,
    net_sales REAL NOT NULL,
    transaction_total REAL NOT NULL,
    invoice_discount REAL NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_sales_year_date ON sales_items (year, date);
CREATE INDEX IF NOT EXISTS idx_sales_month ON sales_items (month);
CREATE INDEX IF NOT EXISTS idx_sales_outlet ON sales_items (outlet);
CREATE INDEX IF NOT EXISTS idx_sales_area ON sales_items (area);
CREATE INDEX IF NOT EXISTS idx_sales_invoice ON sales_items (invoice);

CREATE TABLE IF NOT EXISTS budget_daily (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT NOT NULL,         -- YYYY-MM-DD
    month TEXT NOT NULL,        -- YYYY-MM
    outlet TEXT NOT NULL,
    area TEXT NOT NULL,
    target REAL NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_budget_date ON budget_daily (date);
CREATE INDEX IF NOT EXISTS idx_budget_month ON budget_daily (month);
CREATE INDEX IF NOT EXISTS idx_budget_outlet ON budget_daily (outlet);
CREATE INDEX IF NOT EXISTS idx_budget_area ON budget_daily (area);

CREATE TABLE IF NOT EXISTS visitor_actual (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT NOT NULL,         -- YYYY-MM-DD
    month TEXT NOT NULL,        -- YYYY-MM
    unit TEXT NOT NULL,
    area TEXT NOT NULL,
    visitors INTEGER NOT NULL,
    visitors_individu INTEGER DEFAULT 0,
    visitors_rombongan_langsung INTEGER DEFAULT 0,
    visitors_rombongan_agen INTEGER DEFAULT 0,
    visitors_rombongan_total INTEGER DEFAULT 0,
    UNIQUE(date, unit) ON CONFLICT REPLACE
);

CREATE INDEX IF NOT EXISTS idx_visitor_date ON visitor_actual (date);
CREATE INDEX IF NOT EXISTS idx_visitor_month ON visitor_actual (month);
CREATE INDEX IF NOT EXISTS idx_visitor_unit ON visitor_actual (unit);
CREATE INDEX IF NOT EXISTS idx_visitor_area ON visitor_actual (area);

CREATE INDEX IF NOT EXISTS idx_sales_product ON sales_items (product);
CREATE INDEX IF NOT EXISTS idx_sales_year_month ON sales_items (year, month);

CREATE TABLE IF NOT EXISTS product_lookup (
    product TEXT PRIMARY KEY,
    code TEXT,
    supplier TEXT,
    category TEXT,
    jenis TEXT,
    hpp REAL,
    harga_jual REAL,
    maskot TEXT
);

CREATE INDEX IF NOT EXISTS idx_pl_supplier ON product_lookup (supplier);
CREATE INDEX IF NOT EXISTS idx_pl_category ON product_lookup (category);
CREATE INDEX IF NOT EXISTS idx_pl_jenis ON product_lookup (jenis);

CREATE VIEW IF NOT EXISTS v_sales_analytics AS
SELECT 
    s.id, s.year, s.date, s.month, s.hour, s.invoice, s.outlet, s.area,
    s.product, s.qty, s.item_net_sales, s.net_sales, s.transaction_total, s.invoice_discount,
    p.code, 
    COALESCE(p.supplier, 'LAINNYA') AS supplier,
    COALESCE(p.category, 'LAINNYA') AS category,
    COALESCE(p.jenis, 'DAGANGAN') AS jenis,
    COALESCE(p.hpp, 0.0) AS hpp,
    COALESCE(p.harga_jual, 0.0) AS harga_jual,
    ROUND(s.net_sales - (COALESCE(p.hpp, 0.0) * s.qty), 2) AS gross_profit
FROM sales_items s
LEFT JOIN product_lookup p ON s.product = p.product;


