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
    UNIQUE(date, unit) ON CONFLICT REPLACE
);

CREATE INDEX IF NOT EXISTS idx_visitor_date ON visitor_actual (date);
CREATE INDEX IF NOT EXISTS idx_visitor_month ON visitor_actual (month);
CREATE INDEX IF NOT EXISTS idx_visitor_unit ON visitor_actual (unit);
CREATE INDEX IF NOT EXISTS idx_visitor_area ON visitor_actual (area);

