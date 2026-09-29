import sqlite3
import pandas as pd
import re
from datetime import date

DB_PATH = "performance_tracker.db"


def clean_attendance_cell(val):
    if pd.isna(val):
        return "-", 0.0
    val_str = str(val).strip()
    if "|" in val_str:
        parts = val_str.split("|")
        try:
            return val_str, float(parts[0].strip())
        except ValueError:
            return val_str, 0.0
    try:
        return val_str, float(val_str)
    except ValueError:
        return val_str, 0.0


def find_attendance_header_and_period(file_stream):
    file_stream.seek(0)
    df_raw = pd.read_excel(file_stream, header=None, nrows=25)
    header_idx = 13
    year, month = 2026, 8

    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(x) for x in row.values])
        match = re.search(r'(\d{4})-(\d{2})-\d{2}', row_str)
        if match:
            year, month = int(match.group(1)), int(match.group(2))

        if "emp code" in row_str.lower() and "name" in row_str.lower():
            header_idx = idx

    return header_idx, year, month


def ingest_attendance(file_stream):
    header_idx, year, month = find_attendance_header_and_period(file_stream)
    file_stream.seek(0)
    df = pd.read_excel(file_stream, skiprows=header_idx)
    df.columns = [str(c).strip() for c in df.columns]

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    shifts = dict(cur.execute("SELECT LOWER(TRIM(formal_name)), shift_hours FROM engineers").fetchall())

    day_cols = [c for c in df.columns if str(c).strip().isdigit() and 1 <= int(str(c).strip()) <= 31]
    records = []

    for _, row in df.iterrows():
        raw_name = str(row.get("Name", "")).strip().replace('"', '')
        if not raw_name or raw_name.lower() in ["nan", "none", ""]:
            continue

        std_name = " ".join(raw_name.lower().split())
        shift_base = shifts.get(std_name, 8.0)

        for d_col in day_cols:
            day_num = int(str(d_col).strip())
            raw_str, actual_hrs = clean_attendance_cell(row[d_col])

            try:
                cur_date = date(year, month, day_num)
                is_weekend = cur_date.weekday() >= 5
            except ValueError:
                is_weekend = False

            if actual_hrs > 0:
                capped_hrs = min(actual_hrs, 10.0) if shift_base == 8.0 else actual_hrs
                final_status = f"{capped_hrs:.1f}".rstrip('0').rstrip('.')
            else:
                capped_hrs = 0.0
                if is_weekend and raw_str in ["0", "-", "0.0", "ABS", "", "None"]:
                    final_status = "NW"
                else:
                    final_status = raw_str

            records.append((raw_name, day_num, month, year, final_status, actual_hrs, capped_hrs))

    cur.execute("DELETE FROM attendance WHERE month = ? AND year = ?", (month, year))
    cur.executemany("""
                    INSERT INTO attendance (formal_name, day, month, year, raw_status, actual_hours, capped_hours)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, records)

    conn.commit()
    conn.close()
    print(f"--> [Attendance] Loaded {len(records)} daily logs for period: {year}-{month:02d}")
    return len(records)


def ingest_calibrations(file_stream):
    file_stream.seek(0)
    df = pd.read_excel(file_stream)

    cols = {str(c).strip().lower(): c for c in df.columns}
    cal_date_col = cols.get("caldate", "CalDate")
    cal_by_col = cols.get("calby", "CalBy")
    service_id_col = cols.get("serviceid", "ServiceID")
    model_id_col = cols.get("modelid", "ModelID")
    serial_no_col = cols.get("serialno", "SerialNo")
    cat_col = cols.get("category", "Category")
    mos_col = cols.get("mos", "MOS")

    # Detect SenderDept column dynamically
    sender_col = None
    for candidate in ["senderdept", "sender dept", "sender_dept", "department", "sender"]:
        if candidate in cols:
            sender_col = cols[candidate]
            break

    df = df.dropna(subset=[cal_by_col, cal_date_col])
    df[cal_date_col] = pd.to_datetime(df[cal_date_col], errors="coerce")
    df = df.dropna(subset=[cal_date_col])

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    distinct_periods = df[[cal_date_col]].copy()
    distinct_periods["y"] = distinct_periods[cal_date_col].dt.year
    distinct_periods["m"] = distinct_periods[cal_date_col].dt.month
    unique_months = distinct_periods[["y", "m"]].drop_duplicates().values.tolist()

    for y, m in unique_months:
        cur.execute("DELETE FROM calibrations WHERE cal_year = ? AND cal_month = ?", (int(y), int(m)))

    records = []
    for _, r in df.iterrows():
        c_date = r[cal_date_col]
        sender_val = str(r[sender_col]).strip() if sender_col and pd.notna(r[sender_col]) else ""
        records.append((
            int(r[service_id_col]) if pd.notna(r[service_id_col]) else 0,
            str(r[model_id_col]) if pd.notna(r[model_id_col]) else "",
            str(r[serial_no_col]) if pd.notna(r[serial_no_col]) else "",
            c_date.strftime("%Y-%m-%d"),
            int(c_date.day),
            int(c_date.month),
            int(c_date.year),
            str(r[cal_by_col]).strip(),
            str(r[cat_col]) if pd.notna(r[cat_col]) else "",
            int(r[mos_col]) if pd.notna(r[mos_col]) else 0,
            sender_val
        ))

    cur.executemany("""
                    INSERT INTO calibrations (service_id, model_id, serial_no, cal_date, cal_day, cal_month, cal_year,
                                              cal_by, category, mos, sender_dept)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, records)

    conn.commit()
    conn.close()
    print(f"--> [Calibrations] Successfully stored {len(records)} units for {len(unique_months)} month(s).")
    return len(records)


def ingest_rejections(file_stream):
    file_stream.seek(0)
    df = pd.read_excel(file_stream)
    cols = {str(c).strip().lower(): c for c in df.columns}

    cal_by_col = cols.get("calby", "CalBy")
    model_col = cols.get("model_id", cols.get("modelid", "model_id"))
    serial_col = cols.get("serial no", cols.get("serialno", "Serial No"))
    date_col = cols.get("approved/reject date",
                        cols.get("approved/rejected date", cols.get("caldate", "Approved/Reject Date")))
    reason_col = cols.get("failure reason", cols.get("reason", "Failure Reason"))
    remark_col = cols.get("remark", cols.get("remarks by inspector", "remark"))
    service_col = cols.get("service id", cols.get("serviceid", "Service ID"))

    df = df.dropna(subset=[cal_by_col, date_col])
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=[date_col])

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    df["y"] = df[date_col].dt.year
    df["m"] = df[date_col].dt.month
    unique_periods = df[["y", "m"]].drop_duplicates().values.tolist()

    for y, m in unique_periods:
        cur.execute("DELETE FROM rejection_logs WHERE reject_year = ? AND reject_month = ?", (int(y), int(m)))

    records = []
    for _, r in df.iterrows():
        records.append((
            str(r.get(service_col, "")),
            str(r.get(model_col, "")),
            str(r.get(serial_col, "")),
            str(r[cal_by_col]).strip(),
            r[date_col].strftime("%Y-%m-%d"),
            int(r["y"]),
            int(r["m"]),
            str(r.get(reason_col, "Other")),
            str(r.get(remark_col, ""))
        ))

    cur.executemany("""
                    INSERT INTO rejection_logs (service_id, model_id, serial_no, cal_by, reject_date, reject_year,
                                                reject_month, failure_reason, remark)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, records)

    for y, m in unique_periods:
        month_df = df[(df["y"] == y) & (df["m"] == m)]
        counts = month_df[cal_by_col].astype(str).str.strip().value_counts().to_dict()

        for u, cnt in counts.items():
            eng = cur.execute("SELECT id FROM engineers WHERE LOWER(TRIM(username)) = LOWER(TRIM(?))", (u,)).fetchone()
            if eng:
                eng_id = eng[0]
                cur.execute("""
                            INSERT INTO quality_logs (engineer_id, year, month, mistakes)
                            VALUES (?, ?, ?, ?) ON CONFLICT(engineer_id, year, month)
                    DO
                            UPDATE SET mistakes = excluded.mistakes
                            """, (eng_id, int(y), int(m), int(cnt)))

    conn.commit()
    conn.close()
    print(f"--> [Rejections] Ingested {len(records)} outgoing inspection errors.")
    return len(records)