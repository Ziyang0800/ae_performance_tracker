import sqlite3

DB_PATH = "performance_tracker.db"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS stations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        sort_order INTEGER DEFAULT 99,
        target_pi REAL DEFAULT 3.0,
        is_active BOOLEAN DEFAULT 1
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS engineers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        station_id INTEGER,
        username TEXT UNIQUE NOT NULL,
        formal_name TEXT NOT NULL,
        display_name TEXT,
        shift_hours REAL DEFAULT 8.0,
        is_active BOOLEAN DEFAULT 1,
        joined_date TEXT,
        activation_date TEXT,
        resigned_date TEXT,
        custom_target_pi REAL,
        FOREIGN KEY (station_id) REFERENCES stations (id)
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS engineer_station_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        engineer_id INTEGER NOT NULL,
        station_id INTEGER NOT NULL,
        effective_year INTEGER NOT NULL,
        effective_month INTEGER NOT NULL,
        FOREIGN KEY (engineer_id) REFERENCES engineers (id),
        FOREIGN KEY (station_id) REFERENCES stations (id),
        UNIQUE(engineer_id, effective_year, effective_month)
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS calibrations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        service_id INTEGER,
        model_id TEXT,
        serial_no TEXT,
        cal_date DATE,
        cal_day INTEGER,
        cal_month INTEGER,
        cal_year INTEGER,
        cal_by TEXT,
        category TEXT,
        mos INTEGER,
        sender_dept TEXT
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS attendance (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        formal_name TEXT,
        day INTEGER,
        month INTEGER,
        year INTEGER,
        raw_status TEXT,
        actual_hours REAL DEFAULT 0.0,
        capped_hours REAL DEFAULT 0.0
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS quality_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        engineer_id INTEGER NOT NULL,
        year INTEGER NOT NULL,
        month INTEGER NOT NULL,
        mistakes INTEGER DEFAULT 0,
        FOREIGN KEY (engineer_id) REFERENCES engineers (id),
        UNIQUE(engineer_id, year, month)
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS rejection_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        service_id TEXT,
        model_id TEXT,
        serial_no TEXT,
        cal_by TEXT,
        reject_date TEXT,
        reject_year INTEGER,
        reject_month INTEGER,
        failure_reason TEXT,
        remark TEXT
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS ignored_technicians (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        reason TEXT DEFAULT 'FTE / Untracked Tech'
    );
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS kpi_configs (
        kpi_key TEXT PRIMARY KEY,
        kpi_name TEXT NOT NULL,
        target_label TEXT DEFAULT 'Q2 FY''26 Target',
        target_value REAL NOT NULL,
        unit TEXT DEFAULT '%'
    );
    """)

    default_stations = [
        ("Signal Analyzer", 1, 3.0),
        ("PNA", 2, 2.5),
        ("X-TEST", 3, 3.0),
        ("POWER SUPPLY", 4, 4.5),
        ("ZCAL", 5, 6.0),
        ("MXG/VXG/XSG", 6, 2.5),
        ("STDS LAB", 7, 4.5)
    ]
    for s_name, s_order, s_target in default_stations:
        cur.execute("""
            INSERT INTO stations (name, sort_order, target_pi, is_active) 
            VALUES (?, ?, ?, 1) 
            ON CONFLICT(name) DO UPDATE SET sort_order = excluded.sort_order, target_pi = excluded.target_pi
        """, (s_name, s_order, s_target))

    default_kpis = [
        ("ete_priority", "ETE Priority Management System", "Q2 FY'26 Target", 90.0, "%"),
        ("ete_crd_tat", "Electrical ETE CRD TAT", "Q2 FY'26 Target", 85.0, "%"),
        ("trade_unit", "Trade Unit (CRD)", "Q2 FY'26 Target", 90.0, "%"),
        ("krs_unit", "KRS Unit (CRD)", "Q2 FY'26 Target", 92.0, "%"),
        ("kgso_ext", "KGSO External 3 Working Days", "Q2 FY'26 Target", 95.0, "%"),
        ("pi", "Productivity Index", "Q2 FY'26 Target", 3.1, "units")
    ]
    for k_key, k_name, k_label, k_val, k_unit in default_kpis:
        cur.execute("""
            INSERT OR IGNORE INTO kpi_configs (kpi_key, kpi_name, target_label, target_value, unit)
            VALUES (?, ?, ?, ?, ?)
        """, (k_key, k_name, k_label, k_val, k_unit))

    conn.commit()
    conn.close()
    print("Database initialized successfully.")

if __name__ == "__main__":
    init_db()