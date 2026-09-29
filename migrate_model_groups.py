import sys
import os
import sqlite3
import pandas as pd

DB_PATH = "performance_tracker.db"

def migrate_model_groups(file_path):
    if not os.path.exists(file_path):
        print(f"\n❌ Error: File '{file_path}' not found!")
        print("Usage: python migrate_model_groups.py <path_to_excel_file.xlsx>\n")
        return

    print(f"\n📂 Reading mapping workbook: {file_path}...")
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"❌ Failed to read Excel file: {e}")
        return

    cols = {str(c).strip().lower(): c for c in df.columns}
    
    # Detect Model ID and Group columns
    model_col = cols.get("model id", cols.get("model_id", cols.get("model", None)))
    group_col = cols.get("group", cols.get("segment", cols.get("business group", None)))

    if not model_col or not group_col:
        print("❌ Error: Could not identify required columns.")
        print(f"Columns detected: {list(df.columns)}")
        print("Please ensure Column A has 'Model ID' and Column B has 'Group'.")
        return

    # Clean and filter
    df = df.dropna(subset=[model_col, group_col])
    df["clean_model"] = df[model_col].astype(str).str.strip().str.upper()
    df["clean_group"] = df[group_col].astype(str).str.strip()
    df = df[df["clean_model"] != ""]
    df = df[df["clean_group"] != ""]

    total_rows = len(df)
    unique_models = df["clean_model"].nunique()
    unique_groups = df["clean_group"].unique().tolist()

    print(f"🔍 Found {total_rows} total rows ({unique_models} unique Model IDs).")
    print(f"🏷️  Detected Groups: {unique_groups}")

    # Database connection & schema
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS model_groups (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        model_id TEXT UNIQUE NOT NULL,
        group_name TEXT NOT NULL
    );
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_model_groups_model ON model_groups(model_id);")

    records = [(r["clean_model"], r["clean_group"]) for _, r in df.drop_duplicates(subset=["clean_model"]).iterrows()]

    cur.executemany("""
        INSERT INTO model_groups (model_id, group_name)
        VALUES (?, ?)
        ON CONFLICT(model_id) DO UPDATE SET group_name = excluded.group_name
    """, records)

    conn.commit()
    conn.close()

    print(f"\n✅ Migration successful! Stored {len(records)} model mappings in SQLite database '{DB_PATH}'.\n")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("\n❌ Error: Please provide the Excel file path.")
        print("Usage example:")
        print("  python migrate_model_groups.py models_to_groups.xlsx\n")
    else:
        migrate_model_groups(sys.argv[1])