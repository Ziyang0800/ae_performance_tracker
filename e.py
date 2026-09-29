import sqlite3

conn = sqlite3.connect("performance_tracker.db")

# Find the exact string from your calibrations
cal_name = conn.execute("SELECT DISTINCT cal_by FROM calibrations WHERE cal_by LIKE '%firdaus%'").fetchone()

if cal_name:
    exact_username = cal_name[0]
    print(f"Updating username to match calibration export: {exact_username}")
    conn.execute("UPDATE engineers SET username = ? WHERE formal_name LIKE '%Firdaus%'", (exact_username,))
    conn.commit()
    print("Updated successfully! Refresh your dashboard.")
else:
    print("No records found with 'firdaus' in calibrations table.")

conn.close()