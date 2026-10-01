import io
import os
import calendar
import re
from datetime import date
import sqlite3
import pandas as pd

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_LINE
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_DATA_LABEL_POSITION, XL_MARKER_STYLE
from pptx.chart.data import CategoryChartData
from pptx.oxml import parse_xml

# =============================================================================
# COLOR PALETTE DEFINITIONS
# =============================================================================
BG_COLOR = RGBColor(43, 45, 47)             # Dark Charcoal #2B2D2F
PURE_BLACK = RGBColor(0, 0, 0)              # Pure Black #000000
CORAL_TITLE = RGBColor(255, 89, 118)        # #FF5976
PINK_SUBTITLE = RGBColor(255, 199, 208)     # #FFC7D0
PENTRAC_RED = RGBColor(214, 40, 57)         # Accent Red

# Table Headers
TEAL_HEADER = RGBColor(64, 139, 150)        # #408B96 (Weekdays)
ORANGE_HEADER = RGBColor(255, 103, 32)      # #FF6720 (Weekends)
PH_HEADER = RGBColor(242, 169, 0)           # #F2A900 (Public Holidays)
RUST_ORANGE = RGBColor(211, 84, 0)          # #D35400 (Summary Table Header)
BLUE_HEADER = RGBColor(46, 117, 182)        # #2E75B6 (Quality Index Table Header)
CREAM_BEIGE = RGBColor(253, 235, 208)       # #FDEBD0 (Average P.I. banner)

# Executive KPI Matrix Palette
PURPLE_HEADER = RGBColor(88, 44, 131)       # Deep Purple #582C83
LAVENDER_TARGET = RGBColor(215, 199, 232)   # Soft Lavender #D7C7E8
KPI_GREEN = RGBColor(0, 176, 80)            # Pass Green #00B050
KPI_RED = RGBColor(231, 76, 60)             # Fail Red #E74C3C

# Resource Level Report Palette (Slide 4)
RES_PINK = RGBColor(255, 107, 129)
RES_BLUE = RGBColor(0, 123, 255)
RES_GREEN = RGBColor(0, 176, 80)
RES_YELLOW = RGBColor(255, 193, 7)
RES_STEEL = RGBColor(122, 165, 210)
RES_ROW_PINK = RGBColor(245, 220, 225)
RES_ROW_BLUE = RGBColor(190, 218, 245)
RES_ROW_GREEN = RGBColor(168, 216, 168)
RES_ROW_YELLOW = RGBColor(255, 230, 153)
RES_ROW_STEEL = RGBColor(220, 234, 250)

# Odd Person (Rows 1 & 2)
ODD_WEEKDAY = RGBColor(255, 255, 255)
ODD_WEEKEND = RGBColor(255, 225, 210)
ODD_PH = RGBColor(255, 239, 201)

# Even Person (Rows 3 & 4)
EVEN_WEEKDAY = RGBColor(193, 224, 228)
EVEN_WEEKEND = RGBColor(255, 194, 166)
EVEN_PH = RGBColor(255, 223, 148)

CODE_SPECIAL = RGBColor(255, 89, 118)
TEXT_BLACK = RGBColor(0, 0, 0)
WHITE_TEXT = RGBColor(255, 255, 255)

CONTAINER_BG = RGBColor(255, 255, 255)
CARD_BORDER = RGBColor(74, 189, 196)
PILL_BG = RGBColor(216, 243, 245)
PILL_BORDER = RGBColor(74, 189, 196)
GREEN_DELTA = RGBColor(46, 204, 113)
RED_DELTA = RGBColor(231, 76, 60)

# Station Chart Colors (1st Yellow, 2nd Blue, 3rd Green, 4th Purple, 5th Orange)
CHART_SERIES_COLORS = [
    RGBColor(255, 192, 0),   # 1st person: Yellow / Gold
    RGBColor(55, 96, 146),   # 2nd person: Steel Blue
    RGBColor(0, 176, 80),    # 3rd person: Green
    RGBColor(155, 89, 182),  # 4th person: Purple
    RGBColor(230, 126, 34)   # 5th person: Orange
]

OUTPUT_PURPLE = RGBColor(91, 52, 138)
GOLD_PI = RGBColor(255, 192, 0)
BLACK_TARGET = RGBColor(0, 0, 0)
GREEN_ACTUAL = RGBColor(0, 176, 80)

TAT_CATEGORIES = [
    {"key": "CRD_ETE", "label": "ETE", "weight": 15.0},
    {"key": "CRD_NEW", "label": "NEW", "weight": 35.0},
    {"key": "PSC_EXT", "label": "PSC_EXT", "weight": 35.0},
    {"key": "PSC_INT", "label": "PSC_INT", "weight": 15.0}
]

def normalize_str(s):
    return re.sub(r'\s+', ' ', str(s).strip().replace('"', '')).lower()

def get_keysight_fiscal_info(year, month):
    if month in [11, 12]:
        q, fy = "Q1", year + 1
    elif month == 1:
        q, fy = "Q1", year
    elif month in [2, 3, 4]:
        q, fy = "Q2", year
    elif month in [5, 6, 7]:
        q, fy = "Q3", year
    else:
        q, fy = "Q4", year
    return q, fy, f"{q} FY'{str(fy)[-2:]}"

def get_image_path(filename):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    exact_path = os.path.join(base_dir, filename)
    if os.path.exists(exact_path):
        return exact_path
    try:
        for f in os.listdir(base_dir):
            if f.lower() == filename.lower():
                return os.path.join(base_dir, f)
    except Exception:
        pass
    return None

def set_slide_background(slide, prs, image_filename, fallback_rgb):
    resolved_path = get_image_path(image_filename)
    if resolved_path and os.path.exists(resolved_path):
        slide.shapes.add_picture(resolved_path, 0, 0, prs.slide_width, prs.slide_height)
    else:
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = fallback_rgb
        bg.line.fill.background()

def is_working_day_status(raw_status):
    if not raw_status: return False
    s = str(raw_status).strip().upper()
    if s in ["NW", "AL", "MC", "PH", "PL", "CP", "UPL", "ABS", "RES", "-", ""]: return False
    if "|" in s: s = s.split("|")[0].strip()
    try: return float(s) > 0
    except ValueError: return False

def format_cell_native(cell, fill_rgb, text, font_size=8, bold=False, text_color=TEXT_BLACK, align=PP_ALIGN.CENTER, margin_left=None):
    cell.fill.solid()
    cell.fill.fore_color.rgb = fill_rgb
    
    if margin_left is not None:
        cell.margin_left = Inches(margin_left) if isinstance(margin_left, (int, float)) else margin_left
    elif align == PP_ALIGN.LEFT:
        cell.margin_left = Inches(0.14)
    else:
        cell.margin_left = Inches(0.01)

    cell.margin_right = Inches(0.01)
    cell.margin_top = Inches(0.02)
    cell.margin_bottom = Inches(0.02)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.text = str(text) if text is not None else ""
    for p in cell.text_frame.paragraphs:
        p.alignment = align
        p.font.name = "Arial"
        p.font.size = Pt(font_size)
        p.font.bold = bold
        p.font.color.rgb = text_color

def set_native_table_grid_style(table):
    tblPr = table._tbl.tblPr
    style_elem = tblPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}tableStyleId')
    if style_elem is not None:
        style_elem.text = '{21E4AEA4-E5BE-42F6-9CE4-175374077B40}'

def format_delta(current_val, prev_val, is_float=False):
    if prev_val is None or prev_val == 0: return "(-)"
    diff = round(current_val - prev_val, 1) if is_float else int(current_val - prev_val)
    return f"(+{diff:.1f})" if (is_float and diff > 0) else (f"(+{diff})" if diff > 0 else (f"({diff:.1f})" if is_float else f"({diff})"))

def parse_iso_date(date_val):
    if pd.isna(date_val) or not date_val: return None
    try:
        ts = pd.to_datetime(str(date_val).strip(), errors="coerce")
        return None if pd.isna(ts) else ts.date()
    except Exception: return None

def resolve_engineer_station(conn, engineer_id, fallback_station_id, rep_year, rep_month):
    row = conn.execute("""
        SELECT h.station_id, s.name as station_name, s.sort_order, s.target_pi, s.is_active
        FROM engineer_station_history h
        JOIN stations s ON h.station_id = s.id
        WHERE h.engineer_id = ? 
          AND (h.effective_year < ? OR (h.effective_year = ? AND h.effective_month <= ?))
        ORDER BY h.effective_year DESC, h.effective_month DESC
        LIMIT 1
    """, (engineer_id, rep_year, rep_year, rep_month)).fetchone()

    if row:
        return row["station_id"], row["station_name"], row["sort_order"], row["target_pi"], row["is_active"]
    
    fb = conn.execute("SELECT s.id as station_id, s.name as station_name, s.sort_order, s.target_pi, s.is_active FROM stations s WHERE s.id = ?", (fallback_station_id,)).fetchone()
    if fb:
        return fb["station_id"], fb["station_name"], fb["sort_order"], fb["target_pi"], fb["is_active"]
    return None, "Unassigned", 99, 3.0, 1


def add_kpi_scorecard_slide(prs, conn, year, month):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    set_slide_background(slide, prs, "Grey Wave.png", BG_COLOR)

    t_box = slide.shapes.add_textbox(Inches(1.51), Inches(0.37), Inches(11.5), Inches(0.65))
    p_t = t_box.text_frame.paragraphs[0]
    p_t.text = "CAL LAB OPERATIONS KPI SCORECARD"
    p_t.font.name, p_t.font.bold, p_t.font.size, p_t.font.color.rgb = "Arial", True, Pt(26), CORAL_TITLE

    rolling_3 = []
    cy, cm = year, month
    for _ in range(3):
        rolling_3.append((cy, cm))
        cm -= 1
        if cm == 0: cm, cy = 12, cy - 1
    rolling_3.reverse()

    month_cols = [f"{calendar.month_abbr[m_num]}'{str(y_num)[-2:]}" for y_num, m_num in rolling_3]
    kpi_rows = conn.execute("SELECT kpi_key, kpi_name, target_label, target_value, unit FROM kpi_configs ORDER BY rowid").fetchall()
    kpi_map = {r["kpi_key"]: r for r in kpi_rows}
    target_header_label = kpi_rows[0]["target_label"] if kpi_rows else "Q2 FY'26 Target"

    all_engineers = conn.execute("SELECT id, formal_name, username, station_id, shift_hours, activation_date, resigned_date FROM engineers WHERE track_performance = 1").fetchall()
    muted_users = set([r[0].strip().lower() for r in conn.execute("SELECT username FROM ignored_technicians").fetchall()])

    actual_data = {m_lbl: {} for m_lbl in month_cols}

    for (y_num, m_num), m_lbl in zip(rolling_3, month_cols):
        _, m_days = calendar.monthrange(y_num, m_num)
        m_start, m_end = date(y_num, m_num, 1), date(y_num, m_num, m_days)

        active_ae_usernames = []
        station_monthly_data = {}

        for e in all_engineers:
            u_clean = str(e["username"]).strip().lower()
            if u_clean in muted_users: continue
            act_d, res_d = parse_iso_date(e["activation_date"]), parse_iso_date(e["resigned_date"])
            if not act_d or act_d > m_end: continue
            if res_d and res_d < m_start: continue

            st_id, st_name, _, _, st_active = resolve_engineer_station(conn, e["id"], e["station_id"], y_num, m_num)
            if not st_active: continue

            active_ae_usernames.append(e["username"])
            if st_name not in station_monthly_data:
                station_monthly_data[st_name] = {"units": 0, "days": 0.0}

        if active_ae_usernames:
            placeholders = ",".join(["?"] * len(active_ae_usernames))
            q_cals = f"SELECT cal_by, category, mos, sender_dept FROM calibrations WHERE cal_year = ? AND cal_month = ? AND cal_by IN ({placeholders})"
            df_c = pd.read_sql(q_cals, conn, params=[y_num, m_num] + active_ae_usernames)
        else:
            df_c = pd.DataFrame()

        actual_data[m_lbl]["ete_priority"] = 100.0

        if not df_c.empty:
            df_ete = df_c[df_c["category"].astype(str).str.strip().str.upper() == "CRD_ETE"]
            actual_data[m_lbl]["ete_crd_tat"] = round(((df_ete["mos"] == 1).sum() / len(df_ete)) * 100.0) if len(df_ete) > 0 else 100.0
        else:
            actual_data[m_lbl]["ete_crd_tat"] = 100.0

        if not df_c.empty:
            df_new = df_c[df_c["category"].astype(str).str.strip().str.upper() == "CRD_NEW"]
            mask_krs = df_new["sender_dept"].astype(str).str.upper().str.contains(r"AX63/RSD|AX63/RSD/AC", regex=True)
            df_trade, df_krs = df_new[~mask_krs], df_new[mask_krs]
            actual_data[m_lbl]["trade_unit"] = round(((df_trade["mos"] == 1).sum() / len(df_trade)) * 100.0) if len(df_trade) > 0 else 92.0
            actual_data[m_lbl]["krs_unit"] = round(((df_krs["mos"] == 1).sum() / len(df_krs)) * 100.0) if len(df_krs) > 0 else 97.0
        else:
            actual_data[m_lbl]["trade_unit"] = 92.0
            actual_data[m_lbl]["krs_unit"] = 97.0

        if not df_c.empty:
            df_psc_ext = df_c[df_c["category"].astype(str).str.strip().str.upper() == "PSC_EXT"]
            actual_data[m_lbl]["kgso_ext"] = round(((df_psc_ext["mos"] == 1).sum() / len(df_psc_ext)) * 100.0) if len(df_psc_ext) > 0 else 100.0
        else:
            actual_data[m_lbl]["kgso_ext"] = 100.0

        # P.I. ALIGNMENT FIX: Exactly matching the Dashboard and P.I. slide (Average of Station P.I.s)
        df_att_month = pd.read_sql("SELECT formal_name, capped_hours, year, month FROM attendance WHERE year = ? AND month = ?", conn, params=[y_num, m_num])
        if not df_att_month.empty:
            df_att_month["clean_name"] = df_att_month["formal_name"].apply(normalize_str)

        for e in all_engineers:
            st_id, st_name, _, _, st_active = resolve_engineer_station(conn, e["id"], e["station_id"], y_num, m_num)
            if not st_active or st_name not in station_monthly_data: continue
            
            u_clean = e["username"].strip().lower()
            if not df_c.empty:
                station_monthly_data[st_name]["units"] += len(df_c[df_c["cal_by"].astype(str).str.strip().str.lower() == u_clean])
            
            fn_norm = normalize_str(e["formal_name"])
            shift = e["shift_hours"] if e["shift_hours"] > 0 else 8.0
            if not df_att_month.empty:
                hrs = df_att_month[(df_att_month["clean_name"] == fn_norm) & (df_att_month["year"] == y_num) & (df_att_month["month"] == m_num)]["capped_hours"].sum()
                station_monthly_data[st_name]["days"] += (hrs / shift)

        station_pis = []
        for s_info in station_monthly_data.values():
            if s_info["days"] > 0:
                station_pis.append(s_info["units"] / s_info["days"])
        
        dept_avg_station_pi = round(sum(station_pis) / len(station_pis), 1) if station_pis else 3.2
        actual_data[m_lbl]["pi"] = dept_avg_station_pi

    t_shape = slide.shapes.add_table(8, 5, Inches(0.4), Inches(1.2), Inches(12.5), Inches(5.6))
    tbl = t_shape.table
    set_native_table_grid_style(tbl)
    tbl.columns[0].width, tbl.columns[1].width, tbl.columns[2].width, tbl.columns[3].width, tbl.columns[4].width = Inches(3.7), Inches(2.4), Inches(2.13), Inches(2.13), Inches(2.14)

    c_m = tbl.cell(0, 0); c_m.merge(tbl.cell(1, 0))
    format_cell_native(c_m, PURPLE_HEADER, "Metrics", font_size=14, bold=True, text_color=WHITE_TEXT)

    c_tgt = tbl.cell(0, 1); c_tgt.merge(tbl.cell(1, 1))
    format_cell_native(c_tgt, PURPLE_HEADER, target_header_label, font_size=14, bold=True, text_color=WHITE_TEXT)

    c_act = tbl.cell(0, 2); c_act.merge(tbl.cell(0, 4))
    format_cell_native(c_act, PURPLE_HEADER, "ACTUAL (Last 3 months)", font_size=14, bold=True, text_color=WHITE_TEXT)

    for i, lbl in enumerate(month_cols):
        format_cell_native(tbl.cell(1, i + 2), PURPLE_HEADER, lbl, font_size=13, bold=True, text_color=WHITE_TEXT)

    metric_keys = ["ete_priority", "ete_crd_tat", "trade_unit", "krs_unit", "kgso_ext", "pi"]
    for row_idx, key in enumerate(metric_keys, start=2):
        k_cfg = kpi_map.get(key, {"kpi_name": key, "target_value": 90.0, "unit": "%"})
        t_val = k_cfg["target_value"]
        t_unit = k_cfg["unit"]

        format_cell_native(tbl.cell(row_idx, 0), WHITE_TEXT, k_cfg["kpi_name"], font_size=11, bold=True, text_color=TEXT_BLACK, align=PP_ALIGN.CENTER)
        t_disp = f"≥{t_val:.1f}" if t_unit == "units" else f"{int(t_val)}%"
        format_cell_native(tbl.cell(row_idx, 1), LAVENDER_TARGET, t_disp, font_size=13, bold=False, text_color=TEXT_BLACK)

        for col_i, m_lbl in enumerate(month_cols, start=2):
            val = actual_data[m_lbl].get(key, 0.0)
            is_met = val >= t_val
            cell_bg = KPI_GREEN if is_met else KPI_RED
            disp_str = f"{val:.1f}" if t_unit == "units" else f"{int(round(val))}%"
            format_cell_native(tbl.cell(row_idx, col_i), cell_bg, disp_str, font_size=13, bold=False, text_color=TEXT_BLACK)


# =============================================================================
# GENERATE MONTHLY PRESENTATION DECK (DASHBOARD EXPORT)
# =============================================================================
def generate_monthly_deck(conn, year, month):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    _, num_days = calendar.monthrange(year, month)
    month_start, month_end = date(year, month, 1), date(year, month, num_days)
    month_name, month_title = calendar.month_name[month].upper(), calendar.month_name[month].capitalize()

    prev_m = 12 if month == 1 else month - 1
    prev_y = year - 1 if month == 1 else year

    stations = conn.execute("SELECT id, name, sort_order, target_pi FROM stations WHERE is_active = 1 AND name NOT LIKE '%MECH%' ORDER BY sort_order ASC, name ASC").fetchall()

    df_cal_cur = pd.read_sql("SELECT cal_by, cal_day FROM calibrations WHERE cal_year = ? AND cal_month = ?", conn, params=[year, month])
    df_cal_prev = pd.read_sql("SELECT cal_by, COUNT(*) as units FROM calibrations WHERE cal_year = ? AND cal_month = ? GROUP BY cal_by", conn, params=[prev_y, prev_m])
    prev_user_cal_map = dict(zip(df_cal_prev["cal_by"], df_cal_prev["units"])) if not df_cal_prev.empty else {}

    df_att_cur = pd.read_sql("SELECT formal_name, day, raw_status, capped_hours FROM attendance WHERE year = ? AND month = ?", conn, params=[year, month])
    if not df_att_cur.empty: df_att_cur["clean_name"] = df_att_cur["formal_name"].str.strip().str.lower()

    df_att_prev = pd.read_sql("SELECT formal_name, SUM(capped_hours) as tot_hrs FROM attendance WHERE year = ? AND month = ? GROUP BY formal_name", conn, params=[prev_y, prev_m])
    prev_user_hrs_map = dict(zip(df_att_prev["formal_name"].str.strip().str.lower(), df_att_prev["tot_hrs"])) if not df_att_prev.empty else {}

    all_engineers = conn.execute("SELECT id, formal_name, username, display_name, station_id, shift_hours, activation_date, resigned_date, is_active, track_performance FROM engineers ORDER BY formal_name").fetchall()
    user_cur_units_map = df_cal_cur["cal_by"].value_counts().to_dict() if not df_cal_cur.empty else {}
    username_to_name = {e["username"].strip().lower(): (e["display_name"] or e["formal_name"].split()[0]) for e in all_engineers}

    roster_users = set([e["username"].strip().lower() for e in all_engineers if e["is_active"] and e["track_performance"]])
    muted_users = set([r[0].strip().lower() for r in conn.execute("SELECT username FROM ignored_technicians").fetchall()])

    # SLIDE 1: COVER
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide1, prs, "title with watermark.png", BG_COLOR)

    t_box1 = slide1.shapes.add_textbox(Inches(1.0), Inches(2.6), Inches(11.33), Inches(2.5))
    tf1 = t_box1.text_frame
    p_main = tf1.paragraphs[0]
    p_main.text = f"Cal Lab {month_title} {year} Review"
    p_main.font.name, p_main.font.bold, p_main.font.size, p_main.font.color.rgb = "Aptos", True, Pt(56), WHITE_TEXT
    p_main.alignment = PP_ALIGN.CENTER

    p_sub1 = tf1.add_paragraph()
    p_sub1.text = "Brought to you by Pentrac Itron"
    p_sub1.font.name, p_sub1.font.size, p_sub1.font.italic, p_sub1.font.color.rgb = "Arial", Pt(11), True, CORAL_TITLE
    p_sub1.alignment = PP_ALIGN.CENTER

    # SLIDE 2: AGENDA
    slide2 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide2, prs, "Black wave.png", PURE_BLACK)

    t_box2 = slide2.shapes.add_textbox(Inches(0.63), Inches(1.15), Inches(11.5), Inches(5.5))
    tf2 = t_box2.text_frame
    p2_title = tf2.paragraphs[0]
    p2_title.text = "Meeting Agenda:"
    p2_title.font.name, p2_title.font.bold, p2_title.font.size, p2_title.font.underline, p2_title.font.color.rgb = "Arial", True, Pt(42), True, CORAL_TITLE

    agenda_items = ["1. Resource Level Report", "2. Daily Output & Productivity Index (P.I.)", "3. Quality Index & KPIs", "4. Other updates"]
    for item in agenda_items:
        p_item = tf2.add_paragraph()
        p_item.text = item
        p_item.font.name, p_item.font.size, p_item.font.color.rgb = "Arial", Pt(20), WHITE_TEXT
        p_item.space_before = Pt(16)

    # SLIDE 3: DIVIDER 1
    slide3 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide3, prs, "Black.png", PURE_BLACK)
    p3 = slide3.shapes.add_textbox(Inches(1.0), Inches(3.0), Inches(11.33), Inches(1.5)).text_frame.paragraphs[0]
    p3.text = "Resource Level Report"
    p3.font.name, p3.font.bold, p3.font.size, p3.font.color.rgb = "Arial", True, Pt(48), WHITE_TEXT
    p3.alignment = PP_ALIGN.CENTER

    # SLIDE 4: RESOURCE LEVEL REPORT
    slide4 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide4, prs, "Grey Wave.png", BG_COLOR)

    p4 = slide4.shapes.add_textbox(Inches(1.51), Inches(0.37), Inches(11.5), Inches(0.65)).text_frame.paragraphs[0]
    p4.text = "RESOURCE LEVEL REPORT"
    p4.font.name, p4.font.bold, p4.font.size, p4.font.color.rgb = "Arial", True, Pt(26), CORAL_TITLE

    target_hc_row = conn.execute("SELECT setting_val FROM system_settings WHERE setting_key = 'target_headcount'").fetchone()
    target_headcount = int(target_hc_row[0]) if target_hc_row else 23

    rolling_5 = []
    cy, cm = year, month
    for _ in range(5):
        rolling_5.append((cy, cm))
        cm -= 1
        if cm == 0: cm, cy = 12, cy - 1
    rolling_5.reverse()

    all_staff_hc = conn.execute("SELECT formal_name, display_name, joined_date, activation_date, resigned_date FROM engineers").fetchall()

    t_res_shape = slide4.shapes.add_table(6, 5, Inches(0.4), Inches(1.2), Inches(12.5), Inches(5.6))
    tbl_res = t_res_shape.table
    set_native_table_grid_style(tbl_res)
    tbl_res.columns[0].width, tbl_res.columns[1].width, tbl_res.columns[2].width, tbl_res.columns[3].width, tbl_res.columns[4].width = Inches(2.2), Inches(2.1), Inches(1.9), Inches(1.8), Inches(4.5)

    res_headers = [("Month", RES_PINK), ("Actual Headcount", RES_BLUE), ("Target", RES_GREEN), ("Shortfall", RES_YELLOW), ("Remarks", RES_STEEL)]
    for c_idx, (h_title, h_bg) in enumerate(res_headers):
        format_cell_native(tbl_res.cell(0, c_idx), h_bg, h_title, font_size=14, bold=True, text_color=TEXT_BLACK if c_idx in [0, 3] else WHITE_TEXT)

    for r_idx, (y_val, m_val) in enumerate(rolling_5, start=1):
        _, m_days = calendar.monthrange(y_val, m_val)
        m_start, m_end = date(y_val, m_val, 1), date(y_val, m_val, m_days)
        actual_hc, new_hires, resignees = 0, [], []

        for stf in all_staff_hc:
            name_disp = stf["display_name"] or stf["formal_name"].split()[0]
            start_d = parse_iso_date(stf["joined_date"]) or parse_iso_date(stf["activation_date"])
            res_d = parse_iso_date(stf["resigned_date"])

            if (not start_d or start_d <= m_end) and not (res_d and res_d < m_start): actual_hc += 1
            if start_d and m_start <= start_d <= m_end: new_hires.append(name_disp)
            if res_d and m_start <= res_d <= m_end: resignees.append(name_disp)

        shortfall = target_headcount - actual_hc
        shortfall_str = str(shortfall) if shortfall > 0 else "-"
        nh_str = ", ".join(new_hires) if new_hires else "-"
        res_str = ", ".join(resignees) if resignees else "-"
        remarks_text = f"New Hire : {nh_str}\nResignee : {res_str}"

        format_cell_native(tbl_res.cell(r_idx, 0), RES_ROW_PINK, calendar.month_name[m_val], font_size=12, bold=True, text_color=TEXT_BLACK)
        format_cell_native(tbl_res.cell(r_idx, 1), RES_ROW_BLUE, str(actual_hc), font_size=12, bold=True, text_color=TEXT_BLACK)
        format_cell_native(tbl_res.cell(r_idx, 2), RES_ROW_GREEN, str(target_headcount), font_size=12, bold=True, text_color=TEXT_BLACK)
        format_cell_native(tbl_res.cell(r_idx, 3), RES_ROW_YELLOW, shortfall_str, font_size=12, bold=True, text_color=TEXT_BLACK)
        format_cell_native(tbl_res.cell(r_idx, 4), RES_ROW_STEEL, remarks_text, font_size=12, bold=False, text_color=TEXT_BLACK, align=PP_ALIGN.LEFT, margin_left=0.14)

    # SLIDE 5: DIVIDER 2
    slide5 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide5, prs, "Black.png", PURE_BLACK)
    t_box5 = slide5.shapes.add_textbox(Inches(1.0), Inches(2.6), Inches(11.33), Inches(2.3))
    tf5 = t_box5.text_frame
    p5_main = tf5.paragraphs[0]
    p5_main.text = "Daily Output & Productivity Index (P.I.)"
    p5_main.font.name, p5_main.font.bold, p5_main.font.size, p5_main.font.color.rgb = "Arial", True, Pt(48), WHITE_TEXT
    p5_main.alignment = PP_ALIGN.CENTER
    p5_sub = tf5.add_paragraph()
    p5_sub.text = "Detailed Summary by Station & AE"
    p5_sub.font.name, p5_sub.font.size, p5_sub.font.bold, p5_sub.font.color.rgb = "Arial", Pt(19), True, CORAL_TITLE
    p5_sub.alignment = PP_ALIGN.CENTER

    summary_slide_data = []

    # SLIDES 6+: STATIONS
    for st in stations:
        st_id, st_name = st["id"], st["name"]
        st_engineers = []
        for e in all_engineers:
            if not e["track_performance"]: continue
            u, act_d, res_d, u_count = e["username"], parse_iso_date(e["activation_date"]), parse_iso_date(e["resigned_date"]), user_cur_units_map.get(e["username"], 0)
            if not act_d or act_d > month_end: continue
            if res_d and res_d < month_start and u_count == 0: continue
            assigned_st_id, assigned_st_name, _, _, st_active = resolve_engineer_station(conn, e["id"], e["station_id"], year, month)
            if st_active and assigned_st_name == st_name: st_engineers.append(e)

        if not st_engineers: continue

        slide = prs.slides.add_slide(blank_layout)
        set_slide_background(slide, prs, "Grey Wave.png", BG_COLOR)

        title_box = slide.shapes.add_textbox(Inches(1.51), Inches(0.37), Inches(11.5), Inches(0.65))
        p_t = title_box.text_frame.paragraphs[0]
        p_t.text = f"OUTPUT TRACKING FOR {st_name.upper()} {month_name} {year}"
        p_t.font.name, p_t.font.bold, p_t.font.size, p_t.font.color.rgb = "Arial", True, Pt(24), CORAL_TITLE

        ae_display_names = [e["display_name"] or e["formal_name"].split()[0] for e in st_engineers]
        sub_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.95), Inches(12.5), Inches(0.35))
        p_sub = sub_box.text_frame.paragraphs[0]
        r_sh = p_sub.add_run(); r_sh.text = f"{st_name}: "; r_sh.font.bold, r_sh.font.size, r_sh.font.color.rgb = True, Pt(14), CORAL_TITLE
        r_ae = p_sub.add_run(); r_ae.text = ", ".join(ae_display_names); r_ae.font.size, r_ae.font.color.rgb = Pt(10), PINK_SUBTITLE

        st_cur_units, st_cur_days, st_prev_units, st_prev_days = 0, 0.0, 0, 0.0
        ae_contribution, ae_daily_matrix = {}, {}

        for e in st_engineers:
            disp_name, u, fn_clean, shift = e["display_name"] or e["formal_name"].split()[0], e["username"], e["formal_name"].strip().lower(), e["shift_hours"]
            e_cals = df_cal_cur[df_cal_cur["cal_by"] == u] if not df_cal_cur.empty else pd.DataFrame()
            u_count = len(e_cals)
            st_cur_units += u_count
            ae_contribution[disp_name] = u_count

            e_att = df_att_cur[df_att_cur["clean_name"] == fn_clean] if not df_att_cur.empty else pd.DataFrame()
            h_sum = e_att["capped_hours"].sum() if not e_att.empty else 0.0
            att_day_map = dict(zip(e_att["day"], e_att["raw_status"])) if not e_att.empty else {}
            st_cur_days += (h_sum / shift) if shift > 0 else 0.0

            day_counts = e_cals["cal_day"].value_counts().to_dict() if not e_cals.empty else {}
            raw_hours_list = [str(att_day_map.get(d, "-")) for d in range(1, num_days + 1)]
            raw_output_list = [day_counts.get(d, 0) for d in range(1, num_days + 1)]
            working_flags = [is_working_day_status(raw_hours_list[i]) for i in range(num_days)]

            adj_out_list, carry = list(raw_output_list), 0
            for i in range(num_days):
                if not working_flags[i]:
                    if adj_out_list[i] > 0: carry += adj_out_list[i]; adj_out_list[i] = 0
                else:
                    if carry > 0: adj_out_list[i] += carry; carry = 0
            if carry > 0:
                for i in range(num_days - 1, -1, -1):
                    if working_flags[i]: adj_out_list[i] += carry; break

            ae_daily_matrix[disp_name] = {"hours": raw_hours_list, "output": adj_out_list, "is_working": working_flags}
            prev_u, prev_h = prev_user_cal_map.get(u, 0), prev_user_hrs_map.get(fn_clean, 0.0)
            st_prev_units += prev_u
            st_prev_days += (prev_h / shift) if shift > 0 else 0.0

        st_pi = round(st_cur_units / st_cur_days, 1) if st_cur_days > 0 else 0.0
        st_prev_pi = round(st_prev_units / st_prev_days, 1) if st_prev_days > 0 else 0.0

        summary_slide_data.append({"name": st_name, "units": st_cur_units, "units_delta": format_delta(st_cur_units, st_prev_units, False), "pi": st_pi, "pi_delta": format_delta(st_pi, st_prev_pi, True), "prev_pi": st_prev_pi, "prev_units": st_prev_units})

        rows_cnt, cols_cnt = 2 + (len(st_engineers) * 2), 2 + num_days
        table_shape = slide.shapes.add_table(rows_cnt, cols_cnt, Inches(0.22), Inches(1.35), Inches(12.89), Inches(0.28 * 2 + 0.20 * (len(st_engineers) * 2)))
        tbl = table_shape.table
        set_native_table_grid_style(tbl)
        tbl.columns[0].width, tbl.columns[1].width = Inches(1.10), Inches(0.85)
        day_w = int((Inches(12.89) - Inches(1.95)) / num_days)
        for c in range(2, cols_cnt): tbl.columns[c].width = day_w

        c_day_master = tbl.cell(0, 0); c_day_master.merge(tbl.cell(0, 1))
        format_cell_native(c_day_master, TEAL_HEADER, "DAY", font_size=9, bold=True, text_color=TEXT_BLACK)
        for d in range(1, num_days + 1):
            cur_d = date(year, month, d)
            h_color = PH_HEADER if any(ae_daily_matrix[name]["hours"][d - 1] == "PH" for name in ae_contribution) else (ORANGE_HEADER if cur_d.weekday() >= 5 else TEAL_HEADER)
            format_cell_native(tbl.cell(0, d + 1), h_color, cur_d.strftime("%a")[:3].upper(), font_size=9, bold=True, text_color=TEXT_BLACK)

        c_date_master = tbl.cell(1, 0); c_date_master.merge(tbl.cell(1, 1))
        format_cell_native(c_date_master, TEAL_HEADER, "DATE", font_size=9, bold=True, text_color=TEXT_BLACK)
        for d in range(1, num_days + 1):
            h_color = PH_HEADER if any(ae_daily_matrix[name]["hours"][d - 1] == "PH" for name in ae_contribution) else (ORANGE_HEADER if date(year, month, d).weekday() >= 5 else TEAL_HEADER)
            format_cell_native(tbl.cell(1, d + 1), h_color, f"{d:02d}", font_size=9, bold=True, text_color=TEXT_BLACK)

        curr_row = 2
        for person_idx, e in enumerate(st_engineers):
            disp_name = e["display_name"] or e["formal_name"].split()[0]
            matrix, is_odd_person = ae_daily_matrix[disp_name], (person_idx % 2 == 0)
            w_bg = ODD_WEEKDAY if is_odd_person else EVEN_WEEKDAY
            wknd_bg = ODD_WEEKEND if is_odd_person else EVEN_WEEKEND
            ph_bg = ODD_PH if is_odd_person else EVEN_PH

            c_name_master = tbl.cell(curr_row, 0); c_name_master.merge(tbl.cell(curr_row + 1, 0))
            format_cell_native(c_name_master, w_bg, disp_name, font_size=8.5, bold=True, text_color=TEXT_BLACK)
            format_cell_native(tbl.cell(curr_row, 1), w_bg, "Working Hrs", font_size=8, bold=True, text_color=TEXT_BLACK)

            for d in range(1, num_days + 1):
                val = matrix["hours"][d - 1]
                cell_bg = ph_bg if val == "PH" else (wknd_bg if date(year, month, d).weekday() >= 5 else w_bg)
                disp_val = "" if (date(year, month, d).weekday() >= 5 and val in ["NW", "-", "0", ""]) else val
                t_color = CODE_SPECIAL if val in ["ADV", "MC", "AL", "PL", "CP", "UPL"] else TEXT_BLACK
                format_cell_native(tbl.cell(curr_row, d + 1), cell_bg, disp_val, font_size=8, bold=(val in ["ADV", "MC", "AL", "PL", "CP", "UPL"]), text_color=t_color)

            format_cell_native(tbl.cell(curr_row + 1, 1), w_bg, "Output", font_size=8, bold=True, text_color=TEXT_BLACK)
            for d in range(1, num_days + 1):
                out_val, h_val, is_work = matrix["output"][d - 1], matrix["hours"][d - 1], matrix["is_working"][d - 1]
                cell_bg = ph_bg if h_val == "PH" else (wknd_bg if date(year, month, d).weekday() >= 5 else w_bg)
                disp_val = ("" if out_val == 0 else str(out_val)) if not is_work else str(out_val)
                t_color = CODE_SPECIAL if (is_work and out_val == 0) else TEXT_BLACK
                format_cell_native(tbl.cell(curr_row + 1, d + 1), cell_bg, disp_val, font_size=8, bold=(is_work and out_val == 0), text_color=t_color)
            curr_row += 2

        # -------------------------------------------------------------
        # LEFT CHART: DAILY OUTPUT (STACKED COLUMN)
        # -------------------------------------------------------------
        c1_x, c1_y, c1_w, c1_h = Inches(0.22), Inches(3.61), Inches(6.80), Inches(3.41)
        card1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, c1_x, c1_y, c1_w, c1_h)
        card1.fill.solid(); card1.fill.fore_color.rgb, card1.line.color.rgb = CONTAINER_BG, CARD_BORDER

        chart_data_bar = CategoryChartData()
        chart_data_bar.categories = [f"{d:02d}" for d in range(1, num_days + 1)]
        for disp_name in ae_contribution: chart_data_bar.add_series(disp_name, ae_daily_matrix[disp_name]["output"])

        ch_shape1 = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, c1_x + Inches(0.12), c1_y + Inches(0.12), c1_w - Inches(0.24), c1_h - Inches(0.24), chart_data_bar)
        bar_chart = ch_shape1.chart
        bar_chart.has_title = True; bar_chart.chart_title.text_frame.text = f"DAILY OUTPUT OF {st_name.upper()} STATION"
        bar_chart.chart_title.text_frame.paragraphs[0].font.name, bar_chart.chart_title.text_frame.paragraphs[0].font.size, bar_chart.chart_title.text_frame.paragraphs[0].font.bold = "Aptos", Pt(16), True

        # RULE: Remove legend if station has only 1 AE!
        if len(st_engineers) <= 1:
            bar_chart.has_legend = False
        else:
            bar_chart.has_legend = True
            bar_chart.legend.position = XL_LEGEND_POSITION.TOP
            bar_chart.legend.include_in_layout = False
            bar_chart.legend.font.name = "Aptos"
            bar_chart.legend.font.size = Pt(8)

        # Explicit Aptos 8pt on Category and Value Axes
        cat_axis = bar_chart.category_axis
        cat_axis.tick_labels.font.name = "Aptos"
        cat_axis.tick_labels.font.size = Pt(8)
        cat_axis.tick_labels.offset = 100
        cat_axis.has_major_gridlines = False

        val_axis = bar_chart.value_axis
        val_axis.tick_labels.font.name = "Aptos"
        val_axis.tick_labels.font.size = Pt(8)
        val_axis.has_major_gridlines = True
        val_axis.major_gridlines.format.line.color.rgb = RGBColor(230, 230, 230)

        for s_idx, series in enumerate(bar_chart.series):
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = CHART_SERIES_COLORS[s_idx % len(CHART_SERIES_COLORS)]

        # -------------------------------------------------------------
        # RIGHT CHART: CONTRIBUTION PIE (STRICT MATCHING COLORS)
        # -------------------------------------------------------------
        c2_x, c2_y, c2_w, c2_h = Inches(7.17), Inches(3.61), Inches(4.78), Inches(2.81)
        card2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, c2_x, c2_y, c2_w, c2_h)
        card2.fill.solid(); card2.fill.fore_color.rgb, card2.line.color.rgb = CONTAINER_BG, CARD_BORDER

        chart_data_pie = CategoryChartData()
        chart_data_pie.categories = list(ae_contribution.keys())
        chart_data_pie.add_series("Contribution", [v if v > 0 else 0.001 for v in ae_contribution.values()])

        pie_chart = slide.shapes.add_chart(XL_CHART_TYPE.PIE, c2_x + Inches(0.1), c2_y + Inches(0.08), c2_w - Inches(0.2), c2_h - Inches(0.16), chart_data_pie).chart
        pie_chart.style = 1; pie_chart.has_title = True; pie_chart.chart_title.text_frame.text = "CONTRIBUTION"
        pie_chart.chart_title.text_frame.paragraphs[0].font.name, pie_chart.chart_title.text_frame.paragraphs[0].font.size, pie_chart.chart_title.text_frame.paragraphs[0].font.bold = "Aptos", Pt(16), True
        pie_chart.has_legend = False
        
        plots = pie_chart.plots[0]
        plots.has_data_labels = True
        plots.data_labels.position = XL_DATA_LABEL_POSITION.OUTSIDE_END
        plots.data_labels.show_category_name, plots.data_labels.show_percentage, plots.data_labels.show_value = True, True, False
        plots.data_labels.font.name, plots.data_labels.font.size = "Aptos", Pt(7)

        # STRICT COLORING FOR PIE SLICES (1st Yellow, 2nd Blue, 3rd Green...)
        try:
            for s_idx, point in enumerate(plots.series[0].points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = CHART_SERIES_COLORS[s_idx % len(CHART_SERIES_COLORS)]
        except Exception:
            pass

        # MoM Pill Box
        p_left, p_top, p_w, p_h = Inches(7.17), Inches(6.52), Inches(4.78), Inches(0.51)
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, p_left, p_top, p_w, p_h)
        pill.fill.solid(); pill.fill.fore_color.rgb, pill.line.color.rgb, pill.line.width = PILL_BG, PILL_BORDER, Pt(1.5)
        tf_p = pill.text_frame; tf_p.word_wrap, tf_p.vertical_anchor = True, MSO_ANCHOR.MIDDLE
        p_pi = tf_p.paragraphs[0]; p_pi.alignment = PP_ALIGN.CENTER

        r1 = p_pi.add_run(); r1.text = "Overall Productivity Index: "; r1.font.bold, r1.font.size, r1.font.color.rgb = True, Pt(10.5), TEXT_BLACK
        r2 = p_pi.add_run(); r2.text = f"{st_pi:.1f} "; r2.font.size, r2.font.color.rgb = Pt(12), TEXT_BLACK
        r3 = p_pi.add_run(); r3.text = f"{format_delta(st_pi, st_prev_pi, True)}  |  "; r3.font.size = Pt(10.5); r3.font.color.rgb = GREEN_DELTA if "+" in format_delta(st_pi, st_prev_pi, True) else (RED_DELTA if "-" in format_delta(st_pi, st_prev_pi, True) else TEXT_BLACK)
        r4 = p_pi.add_run(); r4.text = "Total Units: "; r4.font.bold, r4.font.size, r4.font.color.rgb = True, Pt(10.5), TEXT_BLACK
        r5 = p_pi.add_run(); r5.text = f"{st_cur_units} "; r5.font.size, r5.font.color.rgb = Pt(12), TEXT_BLACK
        r6 = p_pi.add_run(); r6.text = format_delta(st_cur_units, st_prev_units, False); r6.font.size = Pt(10.5); r6.font.color.rgb = GREEN_DELTA if "+" in format_delta(st_cur_units, st_prev_units, False) else (RED_DELTA if "-" in format_delta(st_cur_units, st_prev_units, False) else TEXT_BLACK)

    # DEPARTMENT SUMMARY SLIDE
    sum_slide = prs.slides.add_slide(blank_layout)
    set_slide_background(sum_slide, prs, "Grey Wave.png", BG_COLOR)

    num_st = len(summary_slide_data)
    if num_st == 0:
        empty_box = sum_slide.shapes.add_textbox(Inches(2.0), Inches(3.0), Inches(9.33), Inches(1.5))
        empty_box.text_frame.paragraphs[0].text = f"No calibration output recorded for {month_title} {year}."
    else:
        s_rows, s_cols = 4, 1 + num_st + 1
        s_tbl = sum_slide.shapes.add_table(s_rows, s_cols, Inches(0.4), Inches(0.35), Inches(12.5), Inches(2.2)).table
        set_native_table_grid_style(s_tbl)
        s_tbl.columns[0].width, s_tbl.columns[s_cols - 1].width = Inches(1.3), Inches(1.2)
        col_w = int((Inches(12.5) - Inches(2.5)) / max(num_st, 1))
        for c in range(1, s_cols - 1): s_tbl.columns[c].width = col_w

        format_cell_native(s_tbl.cell(0, 0), RUST_ORANGE, "STATION", font_size=12, bold=True, text_color=WHITE_TEXT)
        tot_units_all = sum(d["units"] for d in summary_slide_data)
        prev_tot_units = sum(d["prev_units"] for d in summary_slide_data)
        tot_pi_sum = round(sum(d["pi"] for d in summary_slide_data), 1)
        prev_tot_pi_sum = round(sum(d["prev_pi"] for d in summary_slide_data), 1)
        dept_avg_pi = round(tot_pi_sum / num_st, 1) if num_st > 0 else 0.0
        prev_dept_avg_pi = round(prev_tot_pi_sum / num_st, 1) if num_st > 0 else 0.0

        for i, d in enumerate(summary_slide_data):
            c_idx = i + 1
            format_cell_native(s_tbl.cell(0, c_idx), RUST_ORANGE, d["name"], font_size=12, bold=True, text_color=WHITE_TEXT)
            format_cell_native(s_tbl.cell(1, c_idx), RGBColor(255, 255, 255), str(d["units"]), font_size=11, bold=False, text_color=TEXT_BLACK)
            format_cell_native(s_tbl.cell(2, c_idx), RGBColor(255, 255, 255), f"{d['pi']:.1f}", font_size=11, bold=False, text_color=TEXT_BLACK)

        format_cell_native(s_tbl.cell(0, s_cols - 1), RUST_ORANGE, "TOTAL", font_size=12, bold=True, text_color=WHITE_TEXT)
        format_cell_native(s_tbl.cell(1, s_cols - 1), RGBColor(255, 255, 255), f"{tot_units_all} {format_delta(tot_units_all, prev_tot_units, False)}", font_size=11, bold=True, text_color=TEXT_BLACK)
        format_cell_native(s_tbl.cell(2, s_cols - 1), RGBColor(255, 255, 255), f"{tot_pi_sum:.1f} {format_delta(tot_pi_sum, prev_tot_pi_sum, True)}", font_size=11, bold=True, text_color=TEXT_BLACK)

        format_cell_native(s_tbl.cell(1, 0), RGBColor(255, 255, 255), "OUTPUT", font_size=11, bold=True, text_color=TEXT_BLACK)
        format_cell_native(s_tbl.cell(2, 0), RGBColor(255, 255, 255), "P.I.", font_size=11, bold=True, text_color=TEXT_BLACK)

        c_avg_left = s_tbl.cell(3, 0)
        if s_cols > 2: c_avg_left.merge(s_tbl.cell(3, s_cols - 2))
        format_cell_native(c_avg_left, CREAM_BEIGE, "AVERAGE P.I. (Total P.I. divided by # of Stations)", font_size=11, bold=True, text_color=TEXT_BLACK, align=PP_ALIGN.RIGHT)
        format_cell_native(s_tbl.cell(3, s_cols - 1), CREAM_BEIGE, f"{dept_avg_pi:.1f} {format_delta(dept_avg_pi, prev_dept_avg_pi, True)}", font_size=12, bold=True, text_color=TEXT_BLACK)

        # Output & P.I. Combo Chart with Transparent Overlay
        ch_x, ch_y, ch_w, ch_h = Inches(3.28), Inches(2.94), Inches(6.74), Inches(3.97)
        sum_card = sum_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, ch_x, ch_y, ch_w, ch_h)
        sum_card.fill.solid(); sum_card.fill.fore_color.rgb, sum_card.line.color.rgb = RGBColor(255, 255, 255), RGBColor(210, 210, 210)

        st_categories = [d["name"] for d in summary_slide_data]
        data_col = CategoryChartData()
        data_col.categories = st_categories
        data_col.add_series("OUTPUT", [d["units"] for d in summary_slide_data])

        base_chart = sum_slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, ch_x + Inches(0.12), ch_y + Inches(0.08), ch_w - Inches(0.24), ch_h - Inches(0.48), data_col).chart
        base_chart.style = 1; base_chart.has_title = True; base_chart.chart_title.text_frame.text = "OUTPUT & P.I."
        base_chart.chart_title.text_frame.paragraphs[0].font.name = "Aptos"
        base_chart.chart_title.text_frame.paragraphs[0].font.size = Pt(16)
        base_chart.chart_title.text_frame.paragraphs[0].font.bold = True
        base_chart.series[0].format.fill.solid(); base_chart.series[0].format.fill.fore_color.rgb = OUTPUT_PURPLE
        base_chart.has_legend = False

        base_chart.category_axis.tick_labels.font.name = "Aptos"
        base_chart.category_axis.tick_labels.font.size = Pt(8)
        base_chart.value_axis.tick_labels.font.name = "Aptos"
        base_chart.value_axis.tick_labels.font.size = Pt(8)

        data_line = CategoryChartData()
        data_line.categories = st_categories
        data_line.add_series("P.I.", [d["pi"] for d in summary_slide_data])
        data_line.add_series("Min Target", [2.75] * len(st_categories))

        line_chart = sum_slide.shapes.add_chart(XL_CHART_TYPE.LINE, ch_x + Inches(0.12), ch_y + Inches(0.08), ch_w - Inches(0.24), ch_h - Inches(0.48), data_line).chart
        line_chart.has_title = False; line_chart.has_legend = False

        # Make overlay transparent
        try:
            sp_pr = line_chart._element.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}spPr')
            if sp_pr is None: line_chart._element.insert(0, parse_xml('<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></c:spPr>'))
            else: sp_pr.append(parse_xml('<a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'))
            plot_area = line_chart._element.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}plotArea')
            if plot_area is not None:
                pa_sp_pr = plot_area.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}spPr')
                if pa_sp_pr is None: plot_area.insert(0, parse_xml('<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></c:spPr>'))
                else: pa_sp_pr.append(parse_xml('<a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'))
        except Exception: pass

        val_ax2 = line_chart.value_axis; val_ax2.has_major_gridlines = False
        val_ax2.tick_labels.font.name = "Aptos"
        val_ax2.tick_labels.font.size = Pt(8)
        try:
            val_ax2_elem = val_ax2._element
            axPos = val_ax2_elem.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}axPos')
            if axPos is not None: axPos.set('val', 'r')
            crosses = val_ax2_elem.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}crosses')
            if crosses is not None: crosses.set('val', 'max')
        except Exception: pass
        line_chart.category_axis.visible = False

        s_pi = line_chart.series[0]; s_pi.format.line.color.rgb, s_pi.format.line.width, s_pi.format.line.dash_style = GOLD_PI, Pt(1.5), MSO_LINE.DASH_DOT
        s_tgt = line_chart.series[1]; s_tgt.format.line.color.rgb, s_tgt.format.line.width, s_tgt.format.line.dash_style = BLACK_TARGET, Pt(1.5), MSO_LINE.DASH

        # Legend
        leg_box = sum_slide.shapes.add_textbox(ch_x, ch_y + ch_h - Inches(0.38), ch_w, Inches(0.35))
        p_leg = leg_box.text_frame.paragraphs[0]; p_leg.alignment = PP_ALIGN.CENTER
        r_l1 = p_leg.add_run(); r_l1.text = "■ OUTPUT       "; r_l1.font.bold, r_l1.font.size, r_l1.font.color.rgb = True, Pt(8), OUTPUT_PURPLE
        r_l2 = p_leg.add_run(); r_l2.text = "— · — P.I.       "; r_l2.font.bold, r_l2.font.size, r_l2.font.color.rgb = True, Pt(8), GOLD_PI
        r_l3 = p_leg.add_run(); r_l3.text = "- - - Min Target"; r_l3.font.bold, r_l3.font.size, r_l3.font.color.rgb = True, Pt(8), BLACK_TARGET

    # SLIDE N+2: SECTION DIVIDER 3
    slide_div_q = prs.slides.add_slide(blank_layout)
    set_slide_background(slide_div_q, prs, "Black.png", PURE_BLACK)
    p_dq = slide_div_q.shapes.add_textbox(Inches(1.0), Inches(3.0), Inches(11.33), Inches(1.5)).text_frame.paragraphs[0]
    p_dq.text = "Quality Index & KPIs"; p_dq.font.name, p_dq.font.bold, p_dq.font.size, p_dq.font.color.rgb = "Arial", True, Pt(48), WHITE_TEXT; p_dq.alignment = PP_ALIGN.CENTER

    # SLIDE N+3: EXECUTIVE KPI SCORECARD (SLIDE 14)
    add_kpi_scorecard_slide(prs, conn, year, month)

    # SLIDE N+4: QUALITY INDEX
    slide_q = prs.slides.add_slide(blank_layout)
    set_slide_background(slide_q, prs, "Grey Wave.png", BG_COLOR)

    p_qt = slide_q.shapes.add_textbox(Inches(1.51), Inches(0.37), Inches(11.5), Inches(0.65)).text_frame.paragraphs[0]
    p_qt.text = "QUALITY INDEX"; p_qt.font.name, p_qt.font.bold, p_qt.font.size, p_qt.font.color.rgb = "Arial", True, Pt(24), CORAL_TITLE

    rolling_months = []
    cy, cm = year, month
    for _ in range(12):
        rolling_months.append((cy, cm))
        cm -= 1
        if cm == 0: cm, cy = 12, cy - 1
    rolling_months.reverse()

    df_all_cals = pd.read_sql("SELECT cal_year, cal_month, COUNT(*) as units FROM calibrations GROUP BY cal_year, cal_month", conn)
    cal_lookup = dict(zip(zip(df_all_cals["cal_year"], df_all_cals["cal_month"]), df_all_cals["units"])) if not df_all_cals.empty else {}
    df_all_rejs = pd.read_sql("SELECT cal_by, reject_year, reject_month, failure_reason, COUNT(*) as cnt FROM rejection_logs GROUP BY cal_by, reject_year, reject_month, failure_reason", conn)
    if not df_all_rejs.empty:
        df_all_rejs["clean_cal_by"] = df_all_rejs["cal_by"].astype(str).str.strip().str.lower()
        df_all_rejs = df_all_rejs[(~df_all_rejs["clean_cal_by"].isin(muted_users)) & (df_all_rejs["clean_cal_by"].isin(roster_users))]

    month_labels, acc_rates, coc_counts, mr_counts, others_counts = [], [], [], [], []
    for y_num, m_num in rolling_months:
        tot_c = cal_lookup.get((y_num, m_num), 0)
        m_rejs = df_all_rejs[(df_all_rejs["reject_year"] == y_num) & (df_all_rejs["reject_month"] == m_num)] if not df_all_rejs.empty else pd.DataFrame()
        tot_err = m_rejs["cnt"].sum() if not m_rejs.empty else 0
        coc_cnt = m_rejs[m_rejs["failure_reason"].str.lower().str.contains("coc")]["cnt"].sum() if not m_rejs.empty else 0
        mr_cnt = m_rejs[m_rejs["failure_reason"].str.lower().str.contains("measurement|report")]["cnt"].sum() if not m_rejs.empty else 0
        oth_cnt = tot_err - (coc_cnt + mr_cnt)
        coc_counts.append(int(coc_cnt)); mr_counts.append(int(mr_cnt)); others_counts.append(int(oth_cnt))
        acc_rates.append(round(((tot_c - tot_err) / tot_c) * 100.0, 1) if tot_c > 0 else (100.0 if tot_err == 0 else 98.0))
        month_labels.append(f"{calendar.month_abbr[m_num]} '{str(y_num)[-2:]}")

    q_tbl = slide_q.shapes.add_table(5, 14, Inches(0.4), Inches(0.95), Inches(12.5), Inches(2.2)).table
    set_native_table_grid_style(q_tbl)
    q_tbl.columns[0].width, q_tbl.columns[1].width = Inches(0.5), Inches(1.8)
    q_col_w = int((Inches(12.5) - Inches(2.3)) / 12)
    for c in range(2, 14): q_tbl.columns[c].width = q_col_w

    c_m0 = q_tbl.cell(0, 0); c_m0.merge(q_tbl.cell(0, 1))
    format_cell_native(c_m0, BLUE_HEADER, "MONTH", font_size=10, bold=True, text_color=WHITE_TEXT)
    for i, lbl in enumerate(month_labels): format_cell_native(q_tbl.cell(0, i + 2), BLUE_HEADER, lbl, font_size=9, bold=True, text_color=WHITE_TEXT)

    c_ar0 = q_tbl.cell(1, 0); c_ar0.merge(q_tbl.cell(1, 1))
    format_cell_native(c_ar0, RGBColor(230, 240, 250), "Acceptance Rate (%)", font_size=10, bold=True, text_color=TEXT_BLACK)
    for i, r_val in enumerate(acc_rates): format_cell_native(q_tbl.cell(1, i + 2), RGBColor(230, 240, 250), f"{r_val:.1f}", font_size=9, bold=False, text_color=TEXT_BLACK)

    c_mis = q_tbl.cell(2, 0); c_mis.merge(q_tbl.cell(4, 0))
    format_cell_native(c_mis, BLUE_HEADER, "MISTAKES", font_size=9, bold=True, text_color=WHITE_TEXT)
    format_cell_native(q_tbl.cell(2, 1), RGBColor(255, 255, 255), "COC Issue", font_size=9, bold=True, text_color=TEXT_BLACK)
    format_cell_native(q_tbl.cell(3, 1), RGBColor(255, 255, 255), "Measurement Report", font_size=9, bold=True, text_color=TEXT_BLACK)
    format_cell_native(q_tbl.cell(4, 1), RGBColor(255, 255, 255), "Others", font_size=9, bold=True, text_color=TEXT_BLACK)

    for i in range(12):
        c_i = i + 2
        format_cell_native(q_tbl.cell(2, c_i), RGBColor(255, 255, 255), str(coc_counts[i]) if coc_counts[i] > 0 else "", font_size=9, text_color=TEXT_BLACK)
        format_cell_native(q_tbl.cell(3, c_i), RGBColor(255, 255, 255), str(mr_counts[i]) if mr_counts[i] > 0 else "", font_size=9, text_color=TEXT_BLACK)
        format_cell_native(q_tbl.cell(4, c_i), RGBColor(255, 255, 255), str(others_counts[i]) if others_counts[i] > 0 else "", font_size=9, text_color=TEXT_BLACK)

    # Line Chart
    cq1_x, cq1_y, cq1_w, cq1_h = Inches(0.4), Inches(3.4), Inches(6.1), Inches(3.4)
    card_q1 = slide_q.shapes.add_shape(MSO_SHAPE.RECTANGLE, cq1_x, cq1_y, cq1_w, cq1_h)
    card_q1.fill.solid(); card_q1.fill.fore_color.rgb, card_q1.line.color.rgb = CONTAINER_BG, CARD_BORDER

    chart_data_qline = CategoryChartData()
    chart_data_qline.categories = month_labels
    chart_data_qline.add_series("Actual", acc_rates)
    chart_data_qline.add_series("Target", [98.0] * 12)

    ch_q1 = slide_q.shapes.add_chart(XL_CHART_TYPE.LINE, cq1_x + Inches(0.12), cq1_y + Inches(0.08), cq1_w - Inches(0.24), cq1_h - Inches(0.16), chart_data_qline).chart
    ch_q1.has_title = True; ch_q1.chart_title.text_frame.text = "QUALITY INDEX"
    ch_q1.has_legend = True; ch_q1.legend.position, ch_q1.legend.include_in_layout = XL_LEGEND_POSITION.TOP, False

    s_act = ch_q1.series[0]; s_act.format.line.color.rgb, s_act.format.line.width = GREEN_ACTUAL, Pt(1.5)
    s_tgt_q = ch_q1.series[1]; s_tgt_q.format.line.color.rgb, s_tgt_q.format.line.width, s_tgt_q.format.line.dash_style = BLACK_TARGET, Pt(1.5), MSO_LINE.DASH

    # Column Chart
    cq2_x, cq2_y, cq2_w, cq2_h = Inches(6.8), Inches(3.4), Inches(6.1), Inches(3.4)
    card_q2 = slide_q.shapes.add_shape(MSO_SHAPE.RECTANGLE, cq2_x, cq2_y, cq2_w, cq2_h)
    card_q2.fill.solid(); card_q2.fill.fore_color.rgb, card_q2.line.color.rgb = CONTAINER_BG, CARD_BORDER

    chart_data_qbar = CategoryChartData()
    chart_data_qbar.categories = month_labels
    chart_data_qbar.add_series("COC Issue", coc_counts)
    chart_data_qbar.add_series("Measurement Report", mr_counts)
    chart_data_qbar.add_series("Others", others_counts)

    ch_q2 = slide_q.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, cq2_x + Inches(0.12), cq2_y + Inches(0.08), cq2_w - Inches(0.24), cq2_h - Inches(0.16), chart_data_qbar).chart
    ch_q2.has_title = True; ch_q2.chart_title.text_frame.text = "TYPE OF MISTAKE"
    ch_q2.has_legend = True; ch_q2.legend.position, ch_q2.legend.include_in_layout = XL_LEGEND_POSITION.TOP, False

    if len(ch_q2.series) >= 3:
        ch_q2.series[0].format.fill.solid(); ch_q2.series[0].format.fill.fore_color.rgb = RGBColor(214, 40, 57)
        ch_q2.series[1].format.fill.solid(); ch_q2.series[1].format.fill.fore_color.rgb = RGBColor(237, 125, 49)
        ch_q2.series[2].format.fill.solid(); ch_q2.series[2].format.fill.fore_color.rgb = RGBColor(112, 48, 160)

    p_fn = slide_q.shapes.add_textbox(Inches(0.4), Inches(6.9), Inches(5.0), Inches(0.4)).text_frame.paragraphs[0]
    p_fn.text, p_fn.font.italic, p_fn.font.color.rgb = "Detail breakdown will be the next slide.", True, RGBColor(255, 235, 156)

    # SLIDE N+5: DESCRIPTION OF OUTGOING ERRORS
    slide_err = prs.slides.add_slide(blank_layout)
    set_slide_background(slide_err, prs, "Grey Wave.png", BG_COLOR)

    p_errt = slide_err.shapes.add_textbox(Inches(1.51), Inches(0.37), Inches(11.5), Inches(0.65)).text_frame.paragraphs[0]
    p_errt.text = "DESCRIPTION OF OUTGOING ERRORS"; p_errt.font.name, p_errt.font.bold, p_errt.font.size, p_errt.font.color.rgb = "Arial", True, Pt(24), CORAL_TITLE

    df_cur_rejs = pd.read_sql("SELECT cal_by, model_id, serial_no, reject_date, failure_reason, remark FROM rejection_logs WHERE reject_year = ? AND reject_month = ? ORDER BY reject_date ASC", conn, params=[year, month])
    if not df_cur_rejs.empty:
        df_cur_rejs["clean_cal_by"] = df_cur_rejs["cal_by"].astype(str).str.strip().str.lower()
        df_cur_rejs = df_cur_rejs[(~df_cur_rejs["clean_cal_by"].isin(muted_users)) & (df_cur_rejs["clean_cal_by"].isin(roster_users))]

    err_rows = max(2, len(df_cur_rejs) + 1)
    err_tbl = slide_err.shapes.add_table(err_rows, 6, Inches(0.4), Inches(1.15), Inches(12.5), Inches(0.35 * err_rows)).table
    set_native_table_grid_style(err_tbl)
    err_tbl.columns[0].width, err_tbl.columns[1].width, err_tbl.columns[2].width, err_tbl.columns[3].width, err_tbl.columns[4].width, err_tbl.columns[5].width = Inches(1.8), Inches(1.5), Inches(1.8), Inches(1.6), Inches(1.8), Inches(4.0)

    err_headers = ["AE Resp.", "Model No", "Serial No", "Rejected Date", "Failure Reason", "Remarks by Inspector"]
    for c_i, h_text in enumerate(err_headers): format_cell_native(err_tbl.cell(0, c_i), BLUE_HEADER, h_text, font_size=10, bold=True, text_color=WHITE_TEXT)

    if not df_cur_rejs.empty:
        for r_i, row in df_cur_rejs.reset_index(drop=True).iterrows():
            curr_r = r_i + 1
            raw_u = str(row["cal_by"]).strip().lower()
            ae_display = username_to_name.get(raw_u, str(row["cal_by"]))
            format_cell_native(err_tbl.cell(curr_r, 0), RGBColor(255, 255, 255), ae_display, font_size=9, bold=True, text_color=TEXT_BLACK)
            format_cell_native(err_tbl.cell(curr_r, 1), RGBColor(255, 255, 255), str(row["model_id"]), font_size=9, text_color=TEXT_BLACK)
            format_cell_native(err_tbl.cell(curr_r, 2), RGBColor(255, 255, 255), str(row["serial_no"]), font_size=9, text_color=TEXT_BLACK)
            format_cell_native(err_tbl.cell(curr_r, 3), RGBColor(255, 255, 255), str(row["reject_date"]), font_size=9, text_color=TEXT_BLACK)
            format_cell_native(err_tbl.cell(curr_r, 4), RGBColor(255, 255, 255), str(row["failure_reason"]), font_size=9, text_color=TEXT_BLACK)
            format_cell_native(err_tbl.cell(curr_r, 5), RGBColor(255, 255, 255), str(row["remark"]), font_size=9, text_color=TEXT_BLACK, align=PP_ALIGN.LEFT, margin_left=0.14)
    else:
        c_no_err = err_tbl.cell(1, 0); c_no_err.merge(err_tbl.cell(1, 5))
        format_cell_native(c_no_err, RGBColor(255, 255, 255), "No outgoing inspection errors recorded for this period.", font_size=11, bold=True, text_color=GREEN_DELTA)

    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream


# =============================================================================
# GENERATE QUARTERLY PRESENTATION DECK (2 SLIDES)
# =============================================================================
def generate_quarterly_deck(conn, selected_quarter):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    periods_db = conn.execute("SELECT DISTINCT cal_year as yr, cal_month as mo FROM calibrations WHERE cal_year IS NOT NULL AND cal_month IS NOT NULL ORDER BY yr ASC, mo ASC").fetchall()
    fiscal_quarters_map = {}
    for r in periods_db:
        cal_y, m = int(r["yr"]), int(r["mo"])
        q_label, fiscal_yr, _ = get_keysight_fiscal_info(cal_y, m)
        q_key = f"{q_label} FY'{str(fiscal_yr)[-2:]}"
        if q_key not in fiscal_quarters_map: fiscal_quarters_map[q_key] = []
        fiscal_quarters_map[q_key].append((cal_y, m))

    if not fiscal_quarters_map:
        return io.BytesIO()

    q_key = selected_quarter if selected_quarter in fiscal_quarters_map else list(fiscal_quarters_map.keys())[-1]
    quarter_months = fiscal_quarters_map[q_key]
    month_names = [calendar.month_abbr[m].upper() for (_, m) in quarter_months]

    target_groups = [
        "CSG", 
        "EiSG+Scope", 
        "Legacy+Modular+PMPS(A)+MMA+Others", 
        "STD LAB"
    ]

    all_engineers = conn.execute("""
        SELECT e.id, e.formal_name, e.username, e.station_id, e.shift_hours, e.activation_date, e.resigned_date, s.name as station_name, s.group_name 
        FROM engineers e LEFT JOIN stations s ON e.station_id = s.id 
        WHERE (e.track_performance = 1 OR e.track_performance IS NULL)
    """).fetchall()

    muted_rows = conn.execute("SELECT username FROM ignored_technicians").fetchall()
    muted_users = set([m["username"].strip().lower() for m in muted_rows])

    df_cals_q = pd.DataFrame()
    for (cy, cm) in quarter_months:
        df_m = pd.read_sql("SELECT service_id, model_id, serial_no, cal_date, cal_day, cal_month, cal_year, cal_by, category, mos FROM calibrations WHERE cal_year = ? AND cal_month = ?", conn, params=[cy, cm])
        if not df_m.empty: df_cals_q = pd.concat([df_cals_q, df_m], ignore_index=True)

    if not df_cals_q.empty:
        df_cals_q["clean_by"] = df_cals_q["cal_by"].astype(str).str.strip().str.lower()
        df_cals_q = df_cals_q[~df_cals_q["clean_by"].isin(muted_users)]

    df_att_q = pd.DataFrame()
    for (cy, cm) in quarter_months:
        df_m = pd.read_sql("SELECT formal_name, capped_hours, year, month FROM attendance WHERE year = ? AND month = ?", conn, params=[cy, cm])
        if not df_m.empty: df_att_q = pd.concat([df_att_q, df_m], ignore_index=True)
    if not df_att_q.empty:
        df_att_q["clean_name"] = df_att_q["formal_name"].apply(normalize_str)

    group_stats = {}
    for grp in target_groups:
        st_rows = conn.execute("SELECT id, name FROM stations WHERE is_active = 1 AND (group_name = ? OR (group_name LIKE '%STD%' AND ? = 'STD LAB') OR (group_name LIKE '%LEGACY%' AND ? LIKE '%LEGACY%'))", (grp, grp, grp)).fetchall()
        st_ids = [s["id"] for s in st_rows]

        m_vols, m_days, m_pis, m_tats = [], [], [], []

        for (cy, cm) in quarter_months:
            _, num_d = calendar.monthrange(cy, cm)
            m_start, m_end = date(cy, cm, 1), date(cy, cm, num_d)
            aes_in_st = set()
            month_days = 0.0

            for e in all_engineers:
                act_d, res_d = parse_iso_date(e["activation_date"]), parse_iso_date(e["resigned_date"])
                if act_d and act_d > m_end: continue
                if res_d and res_d < m_start: continue
                
                st_id, _, _, _, st_active = resolve_engineer_station(conn, e["id"], e["station_id"], cy, cm)
                if not st_active: continue

                if st_id in st_ids:
                    u_clean = e["username"].strip().lower()
                    aes_in_st.add(u_clean)
                    fn_norm = normalize_str(e["formal_name"])
                    shift = e["shift_hours"] if e["shift_hours"] > 0 else 8.0
                    hrs = df_att_q[(df_att_q["clean_name"] == fn_norm) & (df_att_q["year"] == cy) & (df_att_q["month"] == cm)]["capped_hours"].sum() if not df_att_q.empty else 0.0
                    month_days += (hrs / shift)

            m_cals = df_cals_q[(df_cals_q["cal_year"] == cy) & (df_cals_q["cal_month"] == cm) & (df_cals_q["clean_by"].isin(aes_in_st))] if (not df_cals_q.empty and aes_in_st) else pd.DataFrame()
            u_cnt = len(m_cals)
            pi_val = round(u_cnt / month_days, 1) if month_days > 0 else 0.0

            w_num, a_w = 0.0, 0.0
            for cat in TAT_CATEGORIES:
                k_key, weight = cat["key"], cat["weight"]
                c_sub = m_cals[m_cals["category"].str.strip().str.upper() == k_key.upper()] if not m_cals.empty else pd.DataFrame()
                if len(c_sub) > 0:
                    pct = int(round(((c_sub["mos"] == 1).sum() / len(c_sub)) * 100.0))
                    w_num += pct * (weight / 100.0)
                    a_w += (weight / 100.0)
            tat_val = int(round(w_num / a_w)) if a_w > 0 else 100

            m_vols.append(u_cnt)
            m_days.append(month_days)
            m_pis.append(pi_val)
            m_tats.append(tat_val)

        tot_vol = sum(m_vols)
        tot_days = sum(m_days)
        quarter_pi = round(tot_vol / tot_days, 1) if tot_days > 0 else 0.0
        avg_mo_vol = int(round(tot_vol / 3.0))

        grp_all_cals = df_cals_q[df_cals_q["clean_by"].isin(aes_in_st)] if (not df_cals_q.empty and aes_in_st) else pd.DataFrame()
        q_w_num, q_a_w = 0.0, 0.0
        for cat in TAT_CATEGORIES:
            k_key, weight = cat["key"], cat["weight"]
            c_sub = grp_all_cals[grp_all_cals["category"].str.strip().str.upper() == k_key.upper()] if not grp_all_cals.empty else pd.DataFrame()
            if len(c_sub) > 0:
                pct = int(round(((c_sub["mos"] == 1).sum() / len(c_sub)) * 100.0))
                q_w_num += pct * (weight / 100.0)
                q_a_w += (weight / 100.0)
        quarter_tat = int(round(q_w_num / q_a_w)) if q_a_w > 0 else 100

        group_stats[grp] = {
            "m_vols": m_vols, "tot_vol": tot_vol, "avg_mo_vol": avg_mo_vol,
            "m_pis": m_pis, "quarter_pi": quarter_pi,
            "m_tats": m_tats, "quarter_tat": quarter_tat
        }

    # =========================================================================
    # SLIDE 1: PRODUCTIVITY INDEX OF <QUARTER>
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide1, prs, "Grey Wave.png", BG_COLOR)

    t_box1 = slide1.shapes.add_textbox(Inches(1.51), Inches(0.37), Inches(11.5), Inches(0.65))
    p1 = t_box1.text_frame.paragraphs[0]
    p1.text = f"PRODUCTIVITY INDEX {q_key}"
    p1.font.name, p1.font.bold, p1.font.size, p1.font.color.rgb = "Arial", True, Pt(26), CORAL_TITLE

    c1_x, c1_y, c1_w, c1_h = Inches(2.2), Inches(1.15), Inches(8.9), Inches(3.8)
    card1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, c1_x, c1_y, c1_w, c1_h)
    card1.fill.solid(); card1.fill.fore_color.rgb, card1.line.color.rgb = RGBColor(255, 255, 255), RGBColor(210, 210, 210)

    data_c1 = CategoryChartData()
    data_c1.categories = target_groups
    for m_i, m_name in enumerate(month_names):
        data_c1.add_series(m_name, [group_stats[g]["m_pis"][m_i] for g in target_groups])

    ch_shape1 = slide1.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, c1_x + Inches(0.12), c1_y + Inches(0.08), c1_w - Inches(0.24), c1_h - Inches(0.48), data_c1)
    ch1 = ch_shape1.chart
    ch1.has_title = True; ch1.chart_title.text_frame.text = f"PRODUCTIVITY INDEX OF {q_key}"
    ch1.chart_title.text_frame.paragraphs[0].font.name, ch1.chart_title.text_frame.paragraphs[0].font.size, ch1.chart_title.text_frame.paragraphs[0].font.bold = "Aptos", Pt(16), True

    blue_palette = [RGBColor(160, 200, 240), RGBColor(46, 117, 182), RGBColor(24, 43, 73)]
    for s_idx, s in enumerate(ch1.series):
        s.format.fill.solid(); s.format.fill.fore_color.rgb = blue_palette[s_idx % len(blue_palette)]
        s.has_data_labels = True
        s.data_labels.font.name = "Aptos"
        s.data_labels.font.size = Pt(7.5)

    ch1.has_legend = True
    ch1.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch1.legend.font.name, ch1.legend.font.size = "Aptos", Pt(8)

    tbl1 = slide1.shapes.add_table(6, 5, Inches(2.2), Inches(5.15), Inches(8.9), Inches(1.9)).table
    set_native_table_grid_style(tbl1)
    tbl1.columns[0].width = Inches(3.7)
    for c_i in range(1, 5): tbl1.columns[c_i].width = Inches(1.3)

    c_m1 = tbl1.cell(0, 0); c_m1.merge(tbl1.cell(0, 4))
    format_cell_native(c_m1, BLUE_HEADER, "PRODUCTIVITY INDEX (P.I.)", font_size=12, bold=True, text_color=WHITE_TEXT)

    format_cell_native(tbl1.cell(1, 0), BLUE_HEADER, "GROUP", font_size=11, bold=True, text_color=WHITE_TEXT)
    for m_i, m_name in enumerate(month_names):
        format_cell_native(tbl1.cell(1, m_i + 1), BLUE_HEADER, m_name, font_size=11, bold=True, text_color=WHITE_TEXT)
    format_cell_native(tbl1.cell(1, 4), BLUE_HEADER, "AVERAGE", font_size=11, bold=True, text_color=WHITE_TEXT)

    for r_i, grp in enumerate(target_groups, start=2):
        format_cell_native(tbl1.cell(r_i, 0), RGBColor(255, 255, 255), grp, font_size=10, bold=True, text_color=TEXT_BLACK, align=PP_ALIGN.CENTER)
        for m_i in range(3):
            format_cell_native(tbl1.cell(r_i, m_i + 1), RGBColor(255, 255, 255), f"{group_stats[grp]['m_pis'][m_i]:.1f}", font_size=11, text_color=TEXT_BLACK)
        format_cell_native(tbl1.cell(r_i, 4), RGBColor(190, 218, 245), f"{group_stats[grp]['quarter_pi']:.1f}", font_size=12, bold=True, text_color=TEXT_BLACK)

    # =========================================================================
    # SLIDE 2: OUTPUT & TAT PERFORMANCE OF <QUARTER>
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide2, prs, "Grey Wave.png", BG_COLOR)

    t_box2 = slide2.shapes.add_textbox(Inches(1.51), Inches(0.37), Inches(11.5), Inches(0.65))
    p2 = t_box2.text_frame.paragraphs[0]
    p2.text = f"OUTPUT & TAT PERFORMANCE OF {q_key}"
    p2.font.name, p2.font.bold, p2.font.size, p2.font.color.rgb = "Arial", True, Pt(26), CORAL_TITLE

    # LEFT HALF: OUTPUT (VOLUME)
    c_out_x, c_out_y, c_out_w, c_out_h = Inches(0.4), Inches(1.15), Inches(6.0), Inches(3.8)
    card_out = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, c_out_x, c_out_y, c_out_w, c_out_h)
    card_out.fill.solid(); card_out.fill.fore_color.rgb, card_out.line.color.rgb = RGBColor(255, 255, 255), RGBColor(210, 210, 210)

    data_out = CategoryChartData()
    data_out.categories = target_groups
    for m_i, m_name in enumerate(month_names):
        data_out.add_series(m_name, [group_stats[g]["m_vols"][m_i] for g in target_groups])

    ch_shape_out = slide2.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, c_out_x + Inches(0.1), c_out_y + Inches(0.08), c_out_w - Inches(0.2), c_out_h - Inches(0.45), data_out)
    ch_out = ch_shape_out.chart
    ch_out.has_title = True; ch_out.chart_title.text_frame.text = f"OUTPUT OF {q_key}"
    p_to = ch_out.chart_title.text_frame.paragraphs[0]
    p_to.font.name, p_to.font.bold, p_to.font.size, p_to.font.color.rgb = "Aptos", True, Pt(14), RGBColor(60, 60, 60)

    teal_palette = [RGBColor(122, 210, 210), RGBColor(46, 139, 139), RGBColor(21, 75, 75)]
    for s_idx, s in enumerate(ch_out.series):
        s.format.fill.solid(); s.format.fill.fore_color.rgb = teal_palette[s_idx % len(teal_palette)]
        s.has_data_labels = True
        s.data_labels.font.name = "Aptos"
        s.data_labels.font.size = Pt(7)

    ch_out.has_legend = True
    ch_out.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch_out.legend.font.name, ch_out.legend.font.size = "Aptos", Pt(7.5)

    tbl_o = slide2.shapes.add_table(6, 6, Inches(0.4), Inches(5.15), Inches(6.0), Inches(1.9)).table
    set_native_table_grid_style(tbl_o)
    tbl_o.columns[0].width = Inches(2.2)
    for c_i in range(1, 6): tbl_o.columns[c_i].width = Inches(0.76)

    c_mo = tbl_o.cell(0, 0); c_mo.merge(tbl_o.cell(0, 5))
    format_cell_native(c_mo, TEAL_HEADER, "OUTPUT", font_size=11, bold=True, text_color=WHITE_TEXT)

    format_cell_native(tbl_o.cell(1, 0), TEAL_HEADER, "GROUP", font_size=9, bold=True, text_color=WHITE_TEXT)
    for m_i, m_name in enumerate(month_names):
        format_cell_native(tbl_o.cell(1, m_i + 1), TEAL_HEADER, m_name, font_size=9, bold=True, text_color=WHITE_TEXT)
    format_cell_native(tbl_o.cell(1, 4), TEAL_HEADER, "TOTAL", font_size=9, bold=True, text_color=WHITE_TEXT)
    format_cell_native(tbl_o.cell(1, 5), TEAL_HEADER, "AVERAGE", font_size=9, bold=True, text_color=WHITE_TEXT)

    for r_i, grp in enumerate(target_groups, start=2):
        format_cell_native(tbl_o.cell(r_i, 0), RGBColor(255, 255, 255), grp, font_size=8, bold=True, text_color=TEXT_BLACK, align=PP_ALIGN.CENTER)
        for m_i in range(3):
            format_cell_native(tbl_o.cell(r_i, m_i + 1), RGBColor(255, 255, 255), str(group_stats[grp]["m_vols"][m_i]), font_size=9, text_color=TEXT_BLACK)
        format_cell_native(tbl_o.cell(r_i, 4), RGBColor(220, 240, 240), str(group_stats[grp]["tot_vol"]), font_size=9.5, bold=True, text_color=TEXT_BLACK)
        format_cell_native(tbl_o.cell(r_i, 5), RGBColor(200, 230, 230), str(group_stats[grp]["avg_mo_vol"]), font_size=9.5, bold=True, text_color=TEXT_BLACK)

    # RIGHT HALF: TURNAROUND TIME
    c_tat_x, c_tat_y, c_tat_w, c_tat_h = Inches(6.9), Inches(1.15), Inches(6.0), Inches(3.8)
    card_tat = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, c_tat_x, c_tat_y, c_tat_w, c_tat_h)
    card_tat.fill.solid(); card_tat.fill.fore_color.rgb, card_tat.line.color.rgb = CONTAINER_BG, CARD_BORDER

    data_tat = CategoryChartData()
    data_tat.categories = target_groups
    for m_i, m_name in enumerate(month_names):
        data_tat.add_series(m_name, [group_stats[g]["m_tats"][m_i] for g in target_groups])

    ch_tat = slide2.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, c_tat_x + Inches(0.1), c_tat_y + Inches(0.08), c_tat_w - Inches(0.2), c_tat_h - Inches(0.45), data_tat).chart
    ch_tat.has_title = True; ch_tat.chart_title.text_frame.text = f"TURNAROUND TIME OF {q_key}"
    p_tt = ch_tat.chart_title.text_frame.paragraphs[0]
    p_tt.font.name, p_tt.font.bold, p_tt.font.size, p_tt.font.color.rgb = "Aptos", True, Pt(14), RGBColor(60, 60, 60)

    green_palette = [RGBColor(152, 228, 176), RGBColor(0, 176, 80), RGBColor(0, 104, 55)]
    for s_idx, s in enumerate(ch_tat.series):
        s.format.fill.solid(); s.format.fill.fore_color.rgb = green_palette[s_idx % len(green_palette)]
        s.has_data_labels = True
        s.data_labels.font.name = "Aptos"
        s.data_labels.font.size = Pt(7)

    ch_tat.has_legend = True
    ch_tat.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch_tat.legend.font.name, ch_tat.legend.font.size = "Aptos", Pt(7.5)

    tbl_t = slide2.shapes.add_table(6, 5, Inches(6.9), Inches(5.15), Inches(6.0), Inches(1.9)).table
    set_native_table_grid_style(tbl_t)
    tbl_t.columns[0].width = Inches(2.2)
    for c_i in range(1, 5): tbl_t.columns[c_i].width = Inches(0.95)

    c_mt = tbl_t.cell(0, 0); c_mt.merge(tbl_t.cell(0, 4))
    format_cell_native(c_mt, RGBColor(0, 176, 80), "TURNAROUND TIME (TAT)", font_size=11, bold=True, text_color=WHITE_TEXT)

    format_cell_native(tbl_t.cell(1, 0), RGBColor(0, 176, 80), "GROUP", font_size=9, bold=True, text_color=WHITE_TEXT)
    for m_i, m_name in enumerate(month_names):
        format_cell_native(tbl_t.cell(1, m_i + 1), RGBColor(0, 176, 80), m_name, font_size=9, bold=True, text_color=WHITE_TEXT)
    format_cell_native(tbl_t.cell(1, 4), RGBColor(0, 176, 80), "AVERAGE", font_size=9, bold=True, text_color=WHITE_TEXT)

    for r_i, grp in enumerate(target_groups, start=2):
        format_cell_native(tbl_t.cell(r_i, 0), RGBColor(255, 255, 255), grp, font_size=8, bold=True, text_color=TEXT_BLACK, align=PP_ALIGN.CENTER)
        for m_i in range(3):
            format_cell_native(tbl_t.cell(r_i, m_i + 1), RGBColor(255, 255, 255), f"{group_stats[grp]['m_tats'][m_i]}%", font_size=9, text_color=TEXT_BLACK)
        format_cell_native(tbl_t.cell(r_i, 4), RGBColor(200, 235, 210), f"{group_stats[grp]['quarter_tat']}%", font_size=9.5, bold=True, text_color=TEXT_BLACK)

    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream