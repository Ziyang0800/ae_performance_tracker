import sqlite3

conn = sqlite3.connect("performance_tracker.db")
cur = conn.cursor()

# 1. Clear any premature September 2026 errors
cur.execute("DELETE FROM rejection_logs WHERE reject_year = 2026 AND reject_month = 9")
cur.execute("DELETE FROM quality_logs WHERE year = 2026 AND month = 9")

conn.commit()
print(f"✅ Cleared premature September rejections from DB! ({cur.rowcount} rows affected)")
conn.close()