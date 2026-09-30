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

BG_COLOR = RGBColor(43, 45, 47)  # Dark Charcoal #2B2D2F
PURE_BLACK = RGBColor(0, 0, 0)  # Pure Black #000000
CORAL_TITLE = RGBColor(255, 89, 118)  # #FF5976
PINK_SUBTITLE = RGBColor(255, 199, 208)  # #FFC7D0
PENTRAC_RED = RGBColor(214, 40, 57)  # Logo Box
PENTRAC_BOX_RED = RGBColor(204, 0, 0)

# Table Headers
TEAL_HEADER = RGBColor(64, 139, 150)  # #408B96 (Weekdays)
ORANGE_HEADER = RGBColor(255, 103, 32)  # #FF6720 (Weekends)
PH_HEADER = RGBColor(242, 169, 0)  # #F2A900 (Public Holidays)
RUST_ORANGE = RGBColor(211, 84, 0)  # #D35400 (Summary Table Header)
BLUE_HEADER = RGBColor(46, 117, 182)  # #2E75B6 (Quality Index Table Header)
CREAM_BEIGE = RGBColor(253, 235, 208)  # #FDEBD0 (Average P.I. banner)

# Executive KPI Matrix Palette
PURPLE_HEADER = RGBColor(88, 44, 131)  # Deep Purple #582C83
LAVENDER_TARGET = RGBColor(215, 199, 232)  # Soft Lavender #D7C7E8
KPI_GREEN = RGBColor(0, 176, 80)  # Pass Green #00B050
KPI_RED = RGBColor(231, 76, 60)  # Fail Red #E74C3C

# Resource Level Report Palette (Image 1)
RES_PINK = RGBColor(255, 107, 129)  # Month Header
RES_BLUE = RGBColor(0, 123, 255)  # Actual Headcount Header
RES_GREEN = RGBColor(0, 176, 80)  # Target Header
RES_YELLOW = RGBColor(255, 193, 7)  # Shortfall Header
RES_STEEL = RGBColor(122, 165, 210)  # Remarks Header
RES_ROW_PINK = RGBColor(245, 220, 225)
RES_ROW_BLUE = RGBColor(190, 218, 245)
RES_ROW_GREEN = RGBColor(168, 216, 168)
RES_ROW_YELLOW = RGBColor(255, 230, 153)
RES_ROW_STEEL = RGBColor(220, 234, 250)

# Odd Person (Rows 1 & 2)
ODD_WEEKDAY = RGBColor(255, 255, 255)  # #FFFFFF
ODD_WEEKEND = RGBColor(255, 225, 210)  # #FFE1D2
ODD_PH = RGBColor(255, 239, 201)  # #FFEFC9

# Even Person (Rows 3 & 4)
EVEN_WEEKDAY = RGBColor(193, 224, 228)  # #C1E0E4
EVEN_WEEKEND = RGBColor(255, 194, 166)  # #FFC2A6
EVEN_PH = RGBColor(255, 223, 148)  # #FFDF94

# Special Typography Colors
CODE_SPECIAL = RGBColor(255, 89, 118)  # #FF5976 Bold
TEXT_BLACK = RGBColor(0, 0, 0)
WHITE_TEXT = RGBColor(255, 255, 255)

# Containers & MoM Pill Box
CONTAINER_BG = RGBColor(255, 255, 255)
CARD_BORDER = RGBColor(74, 189, 196)  # #4ABDC4
SUMMARY_CARD_BORDER = RGBColor(241, 148, 138)
PILL_BG = RGBColor(216, 243, 245)  # #D8F3F5
PILL_BORDER = RGBColor(74, 189, 196)  # #4ABDC4
GREEN_DELTA = RGBColor(46, 204, 113)  # #2ECC71
RED_DELTA = RGBColor(231, 76, 60)  # #E74C3C

# Station Chart Colors
CHART_SERIES_COLORS = [
    RGBColor(255, 192, 0),  # Gold
    RGBColor(55, 96, 146),  # Steel Blue
    RGBColor(0, 176, 80),  # Green
    RGBColor(155, 89, 182),  # Purple
    RGBColor(230, 126, 34)  # Orange
]

# Summary Combo Chart Colors
OUTPUT_PURPLE = RGBColor(91, 52, 138)  # #5B348A (Clustered Column)
GOLD_PI = RGBColor(255, 192, 0)  # #FFC000 (P.I. Line)
BLACK_TARGET = RGBColor(0, 0, 0)  # #000000 (Min Target)
GREEN_ACTUAL = RGBColor(0, 176, 80)  # #00B050 (Quality Index Actual Line)


def normalize_str(s):
    return re.sub(r'\s+', ' ', str(s).strip().replace('"', '')).lower()


def get_image_path(filename):
    """Finds image file using absolute directory and case-insensitive matching on Linux."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    exact_path = os.path.join(base_dir, filename)
    if os.path.exists(exact_path):
        return exact_path

    # Case-insensitive fallback for Linux
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
    if not raw_status:
        return False
    s = str(raw_status).strip().upper()
    if s in ["NW", "AL", "MC", "PH", "PL", "CP", "UPL", "ABS", "RES", "-", ""]:
        return False
    if "|" in s:
        s = s.split("|")[0].strip()
    try:
        return float(s) > 0
    except ValueError:
        return False


def format_cell_native(cell, fill_rgb, text, font_size=8, bold=False, text_color=TEXT_BLACK, align=PP_ALIGN.CENTER):
    cell.fill.solid()
    cell.fill.fore_color.rgb = fill_rgb
    cell.margin_left = Inches(0.01)
    cell.margin_right = Inches(0.01)
    cell.margin_top = Inches(0.01)
    cell.margin_bottom = Inches(0.01)
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
    if prev_val is None or prev_val == 0:
        return "(-)"
    diff = round(current_val - prev_val, 1) if is_float else int(current_val - prev_val)
    if diff > 0:
        return f"(+{diff:.1f})" if is_float else f"(+{diff})"
    elif diff < 0:
        return f"({diff:.1f})" if is_float else f"({diff})"
    return "(-)"


def parse_iso_date(date_val):
    if pd.isna(date_val) or not date_val:
        return None
    try:
        ts = pd.to_datetime(str(date_val).strip(), errors="coerce")
        if pd.isna(ts):
            return None
        return ts.date()
    except Exception:
        return None


def resolve_engineer_station(conn, engineer_id, fallback_station_id, rep_year, rep_month):
    row = conn.execute("""
                       SELECT h.station_id, s.name as station_name, s.sort_order, s.target_pi
                       FROM engineer_station_history h
                                JOIN stations s ON h.station_id = s.id
                       WHERE h.engineer_id = ?
                         AND (h.effective_year < ? OR (h.effective_year = ? AND h.effective_month <= ?))
                       ORDER BY h.effective_year DESC, h.effective_month DESC LIMIT 1
                       """, (engineer_id, rep_year, rep_year, rep_month)).fetchone()

    if row:
        return row["station_id"], row["station_name"], row["sort_order"], row["target_pi"]

    fb = conn.execute(
        "SELECT s.id as station_id, s.name as station_name, s.sort_order, s.target_pi FROM stations s WHERE s.id = ?",
        (fallback_station_id,)).fetchone()
    if fb:
        return fb["station_id"], fb["station_name"], fb["sort_order"], fb["target_pi"]
    return None, "Unassigned", 99, 3.0


def add_kpi_scorecard_slide(prs, conn, year, month):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    set_slide_background(slide, prs, "Grey Wave.png", BG_COLOR)

    badge = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.38), Inches(1.05), Inches(0.55))
    badge.fill.solid()
    badge.fill.fore_color.rgb = PENTRAC_BOX_RED
    badge.line.fill.background()
    tf_b = badge.text_frame
    tf_b.margin_left, tf_b.margin_right, tf_b.margin_top, tf_b.margin_bottom = Inches(0.02), Inches(0.02), Inches(
        0.04), Inches(0.04)
    p_b1 = tf_b.paragraphs[0]
    p_b1.text = "PENTRAC"
    p_b1.font.name, p_b1.font.bold, p_b1.font.size, p_b1.font.color.rgb = "Arial", True, Pt(8.5), WHITE_TEXT
    p_b1.alignment = PP_ALIGN.CENTER
    p_b2 = tf_b.add_paragraph()
    p_b2.text = "ITRON"
    p_b2.font.name, p_b2.font.bold, p_b2.font.size, p_b2.font.color.rgb = "Arial", True, Pt(8.5), WHITE_TEXT
    p_b2.alignment = PP_ALIGN.CENTER

    t_box = slide.shapes.add_textbox(Inches(1.6), Inches(0.35), Inches(11.3), Inches(0.65))
    p_t = t_box.text_frame.paragraphs[0]
    p_t.text = "CAL LAB OPERATIONS KPI SCORECARD"
    p_t.font.name, p_t.font.bold, p_t.font.size, p_t.font.color.rgb = "Arial", True, Pt(26), CORAL_TITLE

    rolling_3 = []
    cy, cm = year, month
    for _ in range(3):
        rolling_3.append((cy, cm))
        cm -= 1
        if cm == 0:
            cm = 12
            cy -= 1
    rolling_3.reverse()

    month_cols = [f"{calendar.month_abbr[m_num]}'{str(y_num)[-2:]}" for y_num, m_num in rolling_3]

    kpi_rows = conn.execute(
        "SELECT kpi_key, kpi_name, target_label, target_value, unit FROM kpi_configs ORDER BY rowid").fetchall()
    kpi_map = {r["kpi_key"]: r for r in kpi_rows}
    target_header_label = kpi_rows[0]["target_label"] if kpi_rows else "Q2 FY'26 Target"

    all_engineers = conn.execute(
        "SELECT id, formal_name, username, station_id, shift_hours, activation_date, resigned_date FROM engineers WHERE track_performance = 1").fetchall()
    muted_users = set(
        [r[0].strip().lower() for r in conn.execute("SELECT username FROM ignored_technicians").fetchall()])

    actual_data = {m_lbl: {} for m_lbl in month_cols}

    for (y_num, m_num), m_lbl in zip(rolling_3, month_cols):
        _, m_days = calendar.monthrange(y_num, m_num)
        m_start = date(y_num, m_num, 1)
        m_end = date(y_num, m_num, m_days)

        active_ae_usernames = []
        ae_shift_map = {}
        ae_name_norm_map = {}

        for e in all_engineers:
            u_clean = str(e["username"]).strip().lower()
            if u_clean in muted_users:
                continue

            act_d = parse_iso_date(e["activation_date"])
            res_d = parse_iso_date(e["resigned_date"])

            if not act_d or act_d > m_end:
                continue
            if res_d and res_d < m_start:
                continue

            st_id, st_name, _, _ = resolve_engineer_station(conn, e["id"], e["station_id"], y_num, m_num)
            st_row = conn.execute("SELECT is_active FROM stations WHERE id = ?", (st_id,)).fetchone()
            if st_row and not st_row[0]:
                continue

            active_ae_usernames.append(e["username"])
            ae_shift_map[u_clean] = e["shift_hours"] if e["shift_hours"] > 0 else 8.0
            ae_name_norm_map[normalize_str(e["formal_name"])] = ae_shift_map[u_clean]

        if active_ae_usernames:
            placeholders = ",".join(["?"] * len(active_ae_usernames))
            q_cals = f"""
                SELECT cal_by, category, mos, sender_dept 
                FROM calibrations 
                WHERE cal_year = ? AND cal_month = ? 
                  AND cal_by IN ({placeholders})
            """
            df_c = pd.read_sql(q_cals, conn, params=[y_num, m_num] + active_ae_usernames)
        else:
            df_c = pd.DataFrame()

        actual_data[m_lbl]["ete_priority"] = 100.0

        if not df_c.empty:
            df_ete = df_c[df_c["category"].astype(str).str.strip().str.upper() == "CRD_ETE"]
            actual_data[m_lbl]["ete_crd_tat"] = round(((df_ete["mos"] == 1).sum() / len(df_ete)) * 100.0) if len(
                df_ete) > 0 else 100.0
        else:
            actual_data[m_lbl]["ete_crd_tat"] = 100.0

        if not df_c.empty:
            df_new = df_c[df_c["category"].astype(str).str.strip().str.upper() == "CRD_NEW"]
            mask_krs = df_new["sender_dept"].astype(str).str.upper().str.contains(r"AX63/RSD|AX63/RSD/AC", regex=True)
            df_trade = df_new[~mask_krs]
            df_krs = df_new[mask_krs]

            actual_data[m_lbl]["trade_unit"] = round(((df_trade["mos"] == 1).sum() / len(df_trade)) * 100.0) if len(
                df_trade) > 0 else 92.0
            actual_data[m_lbl]["krs_unit"] = round(((df_krs["mos"] == 1).sum() / len(df_krs)) * 100.0) if len(
                df_krs) > 0 else 97.0
        else:
            actual_data[m_lbl]["trade_unit"] = 92.0
            actual_data[m_lbl]["krs_unit"] = 97.0

        actual_data[m_lbl]["kgso_ext"] = 100.0

        total_units = len(df_c)
        df_att_month = pd.read_sql("SELECT formal_name, capped_hours FROM attendance WHERE year = ? AND month = ?",
                                   conn, params=[y_num, m_num])

        eff_days_sum = 0.0
        if not df_att_month.empty:
            for _, att_row in df_att_month.iterrows():
                f_norm = normalize_str(att_row["formal_name"])
                if f_norm in ae_name_norm_map:
                    s_hours = ae_name_norm_map[f_norm]
                    eff_days_sum += (float(att_row["capped_hours"]) / s_hours)

        actual_data[m_lbl]["pi"] = round(total_units / eff_days_sum, 1) if eff_days_sum > 0 else 3.2

    t_shape = slide.shapes.add_table(8, 5, Inches(0.4), Inches(1.2), Inches(12.5), Inches(5.6))
    tbl = t_shape.table
    set_native_table_grid_style(tbl)

    tbl.columns[0].width = Inches(3.7)
    tbl.columns[1].width = Inches(2.4)
    tbl.columns[2].width = Inches(2.13)
    tbl.columns[3].width = Inches(2.13)
    tbl.columns[4].width = Inches(2.14)

    c_m = tbl.cell(0, 0)
    c_m.merge(tbl.cell(1, 0))
    format_cell_native(c_m, PURPLE_HEADER, "Metrics", font_size=14, bold=True, text_color=WHITE_TEXT)

    c_tgt = tbl.cell(0, 1)
    c_tgt.merge(tbl.cell(1, 1))
    format_cell_native(c_tgt, PURPLE_HEADER, target_header_label, font_size=14, bold=True, text_color=WHITE_TEXT)

    c_act = tbl.cell(0, 2)
    c_act.merge(tbl.cell(0, 4))
    format_cell_native(c_act, PURPLE_HEADER, "ACTUAL (Last 3 months)", font_size=14, bold=True, text_color=WHITE_TEXT)

    for i, lbl in enumerate(month_cols):
        format_cell_native(tbl.cell(1, i + 2), PURPLE_HEADER, lbl, font_size=13, bold=True, text_color=WHITE_TEXT)

    metric_keys = [
        "ete_priority",
        "ete_crd_tat",
        "trade_unit",
        "krs_unit",
        "kgso_ext",
        "pi"
    ]

    for row_idx, key in enumerate(metric_keys, start=2):
        k_cfg = kpi_map.get(key, {"kpi_name": key, "target_value": 90.0, "unit": "%"})
        t_val = k_cfg["target_value"]
        t_unit = k_cfg["unit"]

        format_cell_native(tbl.cell(row_idx, 0), WHITE_TEXT, k_cfg["kpi_name"], font_size=11, bold=True,
                           text_color=TEXT_BLACK, align=PP_ALIGN.CENTER)
        t_disp = f"≥{t_val:.1f}" if t_unit == "units" else f"{int(t_val)}%"
        format_cell_native(tbl.cell(row_idx, 1), LAVENDER_TARGET, t_disp, font_size=13, bold=False,
                           text_color=TEXT_BLACK)

        for col_i, m_lbl in enumerate(month_cols, start=2):
            val = actual_data[m_lbl].get(key, 0.0)
            is_met = val >= t_val
            cell_bg = KPI_GREEN if is_met else KPI_RED
            disp_str = f"{val:.1f}" if t_unit == "units" else f"{int(round(val))}%"
            format_cell_native(tbl.cell(row_idx, col_i), cell_bg, disp_str, font_size=13, bold=False,
                               text_color=TEXT_BLACK)


def generate_monthly_deck(conn, year, month):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    _, num_days = calendar.monthrange(year, month)
    month_start = date(year, month, 1)
    month_end = date(year, month, num_days)
    month_name = calendar.month_name[month].upper()
    month_title = calendar.month_name[month].capitalize()

    prev_m = 12 if month == 1 else month - 1
    prev_y = year - 1 if month == 1 else year

    stations = conn.execute(
        "SELECT id, name, sort_order, target_pi FROM stations WHERE is_active = 1 AND name NOT LIKE '%MECH%' ORDER BY sort_order ASC, name ASC").fetchall()

    df_cal_cur = pd.read_sql("SELECT cal_by, cal_day FROM calibrations WHERE cal_year = ? AND cal_month = ?", conn,
                             params=[year, month])
    df_cal_prev = pd.read_sql(
        "SELECT cal_by, COUNT(*) as units FROM calibrations WHERE cal_year = ? AND cal_month = ? GROUP BY cal_by", conn,
        params=[prev_y, prev_m])
    prev_user_cal_map = dict(zip(df_cal_prev["cal_by"], df_cal_prev["units"])) if not df_cal_prev.empty else {}

    df_att_cur = pd.read_sql(
        "SELECT formal_name, day, raw_status, capped_hours FROM attendance WHERE year = ? AND month = ?", conn,
        params=[year, month])
    if not df_att_cur.empty:
        df_att_cur["clean_name"] = df_att_cur["formal_name"].str.strip().str.lower()

    df_att_prev = pd.read_sql(
        "SELECT formal_name, SUM(capped_hours) as tot_hrs FROM attendance WHERE year = ? AND month = ? GROUP BY formal_name",
        conn, params=[prev_y, prev_m])
    prev_user_hrs_map = dict(zip(df_att_prev["formal_name"].str.strip().str.lower(),
                                 df_att_prev["tot_hrs"])) if not df_att_prev.empty else {}

    all_engineers = conn.execute(
        "SELECT id, formal_name, username, display_name, station_id, shift_hours, activation_date, resigned_date, is_active, track_performance FROM engineers ORDER BY formal_name").fetchall()
    user_cur_units_map = df_cal_cur["cal_by"].value_counts().to_dict() if not df_cal_cur.empty else {}
    username_to_name = {e["username"].strip().lower(): (e["display_name"] or e["formal_name"].split()[0]) for e in
                        all_engineers}

    roster_users = set(
        [e["username"].strip().lower() for e in all_engineers if e["is_active"] and e["track_performance"]])
    muted_users = set(
        [r[0].strip().lower() for r in conn.execute("SELECT username FROM ignored_technicians").fetchall()])

    # SLIDE 1: COVER TITLE SLIDE
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide1, prs, "title with watermark.png", BG_COLOR)

    t_box1 = slide1.shapes.add_textbox(Inches(1.0), Inches(2.6), Inches(11.33), Inches(2.5))
    tf1 = t_box1.text_frame
    p_main = tf1.paragraphs[0]
    p_main.text = f"Cal Lab {month_title} {year} Review"
    p_main.font.name, p_main.font.bold, p_main.font.size, p_main.font.color.rgb = "Aptos", True, Pt(56), WHITE_TEXT
    p_main.alignment = PP_ALIGN.CENTER

    p_empty1 = tf1.add_paragraph()
    p_empty1.font.size = Pt(18)

    p_sub1 = tf1.add_paragraph()
    p_sub1.text = "Brought to you by Pentrac Itron"
    p_sub1.font.name, p_sub1.font.size, p_sub1.font.italic, p_sub1.font.color.rgb = "Arial", Pt(11), True, CORAL_TITLE
    p_sub1.alignment = PP_ALIGN.CENTER

    # SLIDE 2: MEETING AGENDA
    slide2 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide2, prs, "Black wave.png", PURE_BLACK)

    t_box2 = slide2.shapes.add_textbox(Inches(0.63), Inches(1.15), Inches(11.5), Inches(5.5))
    tf2 = t_box2.text_frame
    p2_title = tf2.paragraphs[0]
    p2_title.text = "Meeting Agenda:"
    p2_title.font.name, p2_title.font.bold, p2_title.font.size, p2_title.font.underline, p2_title.font.color.rgb = "Arial", True, Pt(
        42), True, CORAL_TITLE

    agenda_items = [
        "1. Resource Level Report",
        "2. Daily Output & Productivity Index (P.I.)",
        "3. Quality Index & KPIs",
        "4. Other updates"
    ]
    for item in agenda_items:
        p_item = tf2.add_paragraph()
        p_item.text = item
        p_item.font.name, p_item.font.size, p_item.font.color.rgb = "Arial", Pt(20), WHITE_TEXT
        p_item.space_before = Pt(16)

    # SLIDE 3: SECTION DIVIDER 1
    slide3 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide3, prs, "Black.png", PURE_BLACK)

    t_box3 = slide3.shapes.add_textbox(Inches(1.0), Inches(3.0), Inches(11.33), Inches(1.5))
    p3 = t_box3.text_frame.paragraphs[0]
    p3.text = "Resource Level Report"
    p3.font.name, p3.font.bold, p3.font.size, p3.font.color.rgb = "Arial", True, Pt(48), WHITE_TEXT
    p3.alignment = PP_ALIGN.CENTER

    # =========================================================================
    # SLIDE 4: RESOURCE LEVEL REPORT (INCLUDES MECH LAB & ADMINS UNLESS RESIGNED)
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide4, prs, "Grey Wave.png", BG_COLOR)

    # Top-Left Logo Badge
    badge4 = slide4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.38), Inches(1.05), Inches(0.55))
    badge4.fill.solid()
    badge4.fill.fore_color.rgb = PENTRAC_BOX_RED
    badge4.line.fill.background()
    tf_b4 = badge4.text_frame
    tf_b4.margin_left, tf_b4.margin_right, tf_b4.margin_top, tf_b4.margin_bottom = Inches(0.02), Inches(0.02), Inches(
        0.04), Inches(0.04)
    p_b4_1 = tf_b4.paragraphs[0]
    p_b4_1.text = "PENTRAC"
    p_b4_1.font.name, p_b4_1.font.bold, p_b4_1.font.size, p_b4_1.font.color.rgb = "Arial", True, Pt(8.5), WHITE_TEXT
    p_b4_1.alignment = PP_ALIGN.CENTER
    p_b4_2 = tf_b4.add_paragraph()
    p_b4_2.text = "ITRON"
    p_b4_2.font.name, p_b4_2.font.bold, p_b4_2.font.size, p_b4_2.font.color.rgb = "Arial", True, Pt(8.5), WHITE_TEXT
    p_b4_2.alignment = PP_ALIGN.CENTER

    # Title at Horizontal 1.51", Vertical 0.37"
    t_box4 = slide4.shapes.add_textbox(Inches(1.51), Inches(0.37), Inches(11.5), Inches(0.65))
    p4 = t_box4.text_frame.paragraphs[0]
    p4.text = "RESOURCE LEVEL REPORT"
    p4.font.name, p4.font.bold, p4.font.size, p4.font.color.rgb = "Arial", True, Pt(26), CORAL_TITLE

    # Fetch Target Headcount (Defaults to 23)
    target_hc_row = conn.execute(
        "SELECT setting_val FROM system_settings WHERE setting_key = 'target_headcount'").fetchone()
    target_headcount = int(target_hc_row[0]) if target_hc_row else 23

    # Rolling 5 Months ending at current report month
    rolling_5 = []
    cy, cm = year, month
    for _ in range(5):
        rolling_5.append((cy, cm))
        cm -= 1
        if cm == 0:
            cm = 12
            cy -= 1
    rolling_5.reverse()

    # Query ALL staff in roster regardless of station active status or performance tracking
    all_staff_hc = conn.execute("""
                                SELECT formal_name, display_name, joined_date, activation_date, resigned_date
                                FROM engineers
                                """).fetchall()

    t_res_shape = slide4.shapes.add_table(6, 5, Inches(0.4), Inches(1.2), Inches(12.5), Inches(5.6))
    tbl_res = t_res_shape.table
    set_native_table_grid_style(tbl_res)

    tbl_res.columns[0].width = Inches(2.2)
    tbl_res.columns[1].width = Inches(2.1)
    tbl_res.columns[2].width = Inches(1.9)
    tbl_res.columns[3].width = Inches(1.8)
    tbl_res.columns[4].width = Inches(4.5)

    res_headers = [
        ("Month", RES_PINK),
        ("Actual Headcount", RES_BLUE),
        ("Target", RES_GREEN),
        ("Shortfall", RES_YELLOW),
        ("Remarks", RES_STEEL)
    ]
    for c_idx, (h_title, h_bg) in enumerate(res_headers):
        format_cell_native(tbl_res.cell(0, c_idx), h_bg, h_title, font_size=13, bold=True,
                           text_color=TEXT_BLACK if c_idx in [0, 3] else WHITE_TEXT)

    for r_idx, (y_val, m_val) in enumerate(rolling_5, start=1):
        _, m_days = calendar.monthrange(y_val, m_val)
        m_start = date(y_val, m_val, 1)
        m_end = date(y_val, m_val, m_days)

        actual_hc = 0
        new_hires = []
        resignees = []

        for stf in all_staff_hc:
            name_disp = stf["display_name"] or stf["formal_name"].split()[0]
            j_d = parse_iso_date(stf["joined_date"])
            act_d = parse_iso_date(stf["activation_date"])
            res_d = parse_iso_date(stf["resigned_date"])

            # Start date is joined_date or activation_date (either is valid)
            start_d = j_d or act_d

            # Count everyone who joined on or before this month and hasn't resigned yet
            has_started = (not start_d or start_d <= m_end)
            has_resigned = (res_d and res_d < m_start)

            if has_started and not has_resigned:
                actual_hc += 1

            # Log New Hires in their start month
            if start_d and m_start <= start_d <= m_end:
                new_hires.append(name_disp)

            # Log Resignees in their departure month
            if res_d and m_start <= res_d <= m_end:
                resignees.append(name_disp)

        shortfall = target_headcount - actual_hc
        shortfall_str = str(shortfall) if shortfall > 0 else "-"

        nh_str = ", ".join(new_hires) if new_hires else "-"
        res_str = ", ".join(resignees) if resignees else "-"
        remarks_text = f"New Hire : {nh_str}\nResignee : {res_str}"

        format_cell_native(tbl_res.cell(r_idx, 0), RES_ROW_PINK, calendar.month_name[m_val], font_size=12, bold=True,
                           text_color=TEXT_BLACK)
        format_cell_native(tbl_res.cell(r_idx, 1), RES_ROW_BLUE, str(actual_hc), font_size=13, bold=True,
                           text_color=TEXT_BLACK)
        format_cell_native(tbl_res.cell(r_idx, 2), RES_ROW_GREEN, str(target_headcount), font_size=13, bold=True,
                           text_color=TEXT_BLACK)
        format_cell_native(tbl_res.cell(r_idx, 3), RES_ROW_YELLOW, shortfall_str, font_size=13, bold=True,
                           text_color=TEXT_BLACK)
        format_cell_native(tbl_res.cell(r_idx, 4), RES_ROW_STEEL, remarks_text, font_size=10.5, bold=False,
                           text_color=TEXT_BLACK, align=PP_ALIGN.LEFT)

    # SLIDE 5: SECTION DIVIDER 2
    slide5 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide5, prs, "Black.png", PURE_BLACK)

    t_box5 = slide5.shapes.add_textbox(Inches(1.0), Inches(2.6), Inches(11.33), Inches(2.3))
    tf5 = t_box5.text_frame
    p5_main = tf5.paragraphs[0]
    p5_main.text = "Daily Output & Productivity Index (P.I.)"
    p5_main.font.name, p5_main.font.bold, p5_main.font.size, p5_main.font.color.rgb = "Arial", True, Pt(48), WHITE_TEXT
    p5_main.alignment = PP_ALIGN.CENTER

    p5_empty = tf5.add_paragraph()
    p5_empty.font.size = Pt(16)

    p5_sub = tf5.add_paragraph()
    p5_sub.text = "Detailed Summary by Station & AE"
    p5_sub.font.name, p5_sub.font.size, p5_sub.font.bold, p5_sub.font.color.rgb = "Arial", Pt(19), True, CORAL_TITLE
    p5_sub.alignment = PP_ALIGN.CENTER

    summary_slide_data = []

    # SLIDES 6 ONWARD: STATIONS
    for st in stations:
        st_id, st_name = st["id"], st["name"]

        st_engineers = []
        for e in all_engineers:
            if not e["track_performance"]:
                continue

            u = e["username"]
            act_d = parse_iso_date(e["activation_date"])
            res_d = parse_iso_date(e["resigned_date"])
            u_count = user_cur_units_map.get(u, 0)

            if not act_d or act_d > month_end:
                continue
            if res_d and res_d < month_start and u_count == 0:
                continue

            assigned_st_id, assigned_st_name, _, _ = resolve_engineer_station(conn, e["id"], e["station_id"], year,
                                                                              month)
            if assigned_st_name == st_name:
                st_engineers.append(e)

        if not st_engineers:
            continue

        slide = prs.slides.add_slide(blank_layout)
        set_slide_background(slide, prs, "Grey Wave.png", BG_COLOR)

        title_box = slide.shapes.add_textbox(Inches(1.51), Inches(0.37), Inches(11.5), Inches(0.65))
        p_t = title_box.text_frame.paragraphs[0]
        p_t.text = f"OUTPUT TRACKING FOR {st_name.upper()} {month_name} {year}"
        p_t.font.name, p_t.font.bold, p_t.font.size, p_t.font.color.rgb = "Arial", True, Pt(24), CORAL_TITLE

        ae_display_names = [e["display_name"] or e["formal_name"].split()[0] for e in st_engineers]
        sub_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.95), Inches(12.5), Inches(0.35))
        tf_sub = sub_box.text_frame
        p_sub = tf_sub.paragraphs[0]

        r_sh = p_sub.add_run()
        r_sh.text = f"{st_name}: "
        r_sh.font.name, r_sh.font.bold, r_sh.font.size, r_sh.font.color.rgb = "Arial", True, Pt(14), CORAL_TITLE

        r_ae = p_sub.add_run()
        r_ae.text = ", ".join(ae_display_names)
        r_ae.font.name, r_ae.font.size, r_ae.font.color.rgb = "Arial", Pt(10), PINK_SUBTITLE

        st_cur_units, st_cur_days = 0, 0.0
        st_prev_units, st_prev_days = 0, 0.0
        ae_contribution, ae_daily_matrix = {}, {}

        for e in st_engineers:
            disp_name = e["display_name"] or e["formal_name"].split()[0]
            u = e["username"]
            fn_clean = e["formal_name"].strip().lower()
            shift = e["shift_hours"]

            e_cals = df_cal_cur[df_cal_cur["cal_by"] == u] if not df_cal_cur.empty else pd.DataFrame()
            u_count = len(e_cals)
            st_cur_units += u_count
            ae_contribution[disp_name] = u_count

            if not df_att_cur.empty:
                e_att = df_att_cur[df_att_cur["clean_name"] == fn_clean]
                h_sum = e_att["capped_hours"].sum()
                att_day_map = dict(zip(e_att["day"], e_att["raw_status"]))
            else:
                h_sum, att_day_map = 0.0, {}

            st_cur_days += (h_sum / shift) if shift > 0 else 0.0
            day_counts = e_cals["cal_day"].value_counts().to_dict() if not e_cals.empty else {}

            raw_hours_list = [str(att_day_map.get(d, "-")) for d in range(1, num_days + 1)]
            raw_output_list = [day_counts.get(d, 0) for d in range(1, num_days + 1)]

            working_flags = [is_working_day_status(raw_hours_list[i]) for i in range(num_days)]
            adjusted_output = list(raw_output_list)

            carry = 0
            for i in range(num_days):
                if not working_flags[i]:
                    if adjusted_output[i] > 0:
                        carry += adjusted_output[i]
                        adjusted_output[i] = 0
                else:
                    if carry > 0:
                        adjusted_output[i] += carry
                        carry = 0

            if carry > 0:
                for i in range(num_days - 1, -1, -1):
                    if working_flags[i]:
                        adjusted_output[i] += carry
                        carry = 0
                        break

            ae_daily_matrix[disp_name] = {
                "hours": raw_hours_list,
                "output": adjusted_output,
                "is_working": working_flags
            }

            prev_u = prev_user_cal_map.get(u, 0)
            prev_h = prev_user_hrs_map.get(fn_clean, 0.0)
            st_prev_units += prev_u
            st_prev_days += (prev_h / shift) if shift > 0 else 0.0

        st_pi = round(st_cur_units / st_cur_days, 1) if st_cur_days > 0 else 0.0
        st_prev_pi = round(st_prev_units / st_prev_days, 1) if st_prev_days > 0 else 0.0

        pi_delta_str = format_delta(st_pi, st_prev_pi, is_float=True)
        units_delta_str = format_delta(st_cur_units, st_prev_units, is_float=False)

        summary_slide_data.append({
            "name": st_name,
            "units": st_cur_units,
            "units_delta": units_delta_str,
            "pi": st_pi,
            "pi_delta": pi_delta_str,
            "prev_pi": st_prev_pi,
            "prev_units": st_prev_units
        })

        rows_cnt = 2 + (len(st_engineers) * 2)
        cols_cnt = 2 + num_days

        tbl_top = Inches(1.35)
        table_h = Inches(0.28 * 2 + 0.20 * (len(st_engineers) * 2))

        table_shape = slide.shapes.add_table(rows_cnt, cols_cnt, Inches(0.22), tbl_top, Inches(12.89), table_h)
        tbl = table_shape.table
        set_native_table_grid_style(tbl)

        tbl.columns[0].width = Inches(1.10)
        tbl.columns[1].width = Inches(0.85)
        day_w = int((Inches(12.89) - Inches(1.95)) / num_days)
        for c in range(2, cols_cnt):
            tbl.columns[c].width = day_w

        tbl.rows[0].height = Inches(0.28)
        tbl.rows[1].height = Inches(0.28)
        for r_i in range(2, rows_cnt):
            tbl.rows[r_i].height = Inches(0.20)

        c_day_master = tbl.cell(0, 0)
        c_day_master.merge(tbl.cell(0, 1))
        format_cell_native(c_day_master, TEAL_HEADER, "DAY", font_size=9, bold=True, text_color=TEXT_BLACK)

        for d in range(1, num_days + 1):
            cur_d = date(year, month, d)
            is_wknd = cur_d.weekday() >= 5
            is_ph = any(ae_daily_matrix[name]["hours"][d - 1] == "PH" for name in ae_contribution)
            h_color = PH_HEADER if is_ph else (ORANGE_HEADER if is_wknd else TEAL_HEADER)
            format_cell_native(tbl.cell(0, d + 1), h_color, cur_d.strftime("%a")[:3].upper(), font_size=9, bold=True,
                               text_color=TEXT_BLACK)

        c_date_master = tbl.cell(1, 0)
        c_date_master.merge(tbl.cell(1, 1))
        format_cell_native(c_date_master, TEAL_HEADER, "DATE", font_size=9, bold=True, text_color=TEXT_BLACK)

        for d in range(1, num_days + 1):
            cur_d = date(year, month, d)
            is_wknd = cur_d.weekday() >= 5
            is_ph = any(ae_daily_matrix[name]["hours"][d - 1] == "PH" for name in ae_contribution)
            h_color = PH_HEADER if is_ph else (ORANGE_HEADER if is_wknd else TEAL_HEADER)
            format_cell_native(tbl.cell(1, d + 1), h_color, f"{d:02d}", font_size=9, bold=True, text_color=TEXT_BLACK)

        curr_row = 2
        for person_idx, e in enumerate(st_engineers):
            disp_name = e["display_name"] or e["formal_name"].split()[0]
            matrix = ae_daily_matrix[disp_name]
            is_odd_person = (person_idx % 2 == 0)

            w_bg = ODD_WEEKDAY if is_odd_person else EVEN_WEEKDAY
            wknd_bg = ODD_WEEKEND if is_odd_person else EVEN_WEEKEND
            ph_bg = ODD_PH if is_odd_person else EVEN_PH

            c_name_master = tbl.cell(curr_row, 0)
            c_name_master.merge(tbl.cell(curr_row + 1, 0))
            format_cell_native(c_name_master, w_bg, disp_name, font_size=8.5, bold=True, text_color=TEXT_BLACK)
            format_cell_native(tbl.cell(curr_row, 1), w_bg, "Working Hrs", font_size=8, bold=True,
                               text_color=TEXT_BLACK)

            for d in range(1, num_days + 1):
                cur_d = date(year, month, d)
                is_wknd = cur_d.weekday() >= 5
                val = matrix["hours"][d - 1]
                is_ph = (val == "PH")
                cell_bg = ph_bg if is_ph else (wknd_bg if is_wknd else w_bg)

                disp_val = "" if (is_wknd and val in ["NW", "-", "0", ""]) else val
                t_color = CODE_SPECIAL if val in ["ADV", "MC", "AL", "PL", "CP", "UPL"] else TEXT_BLACK
                is_bld = (val in ["ADV", "MC", "AL", "PL", "CP", "UPL"])

                format_cell_native(tbl.cell(curr_row, d + 1), cell_bg, disp_val, font_size=8, bold=is_bld,
                                   text_color=t_color)

            format_cell_native(tbl.cell(curr_row + 1, 1), w_bg, "Output", font_size=8, bold=True, text_color=TEXT_BLACK)

            for d in range(1, num_days + 1):
                out_val = matrix["output"][d - 1]
                h_val = matrix["hours"][d - 1]
                is_work = matrix["is_working"][d - 1]
                is_ph = (h_val == "PH")
                is_wknd = date(year, month, d).weekday() >= 5
                cell_bg = ph_bg if is_ph else (wknd_bg if is_wknd else w_bg)

                if not is_work:
                    disp_val = "" if out_val == 0 else str(out_val)
                    t_color, is_bld = TEXT_BLACK, False
                else:
                    disp_val = str(out_val)
                    t_color = CODE_SPECIAL if out_val == 0 else TEXT_BLACK
                    is_bld = (out_val == 0)

                format_cell_native(tbl.cell(curr_row + 1, d + 1), cell_bg, disp_val, font_size=8, bold=is_bld,
                                   text_color=t_color)

            curr_row += 2

        # Left Chart: Stacked Column
        c1_left, c1_top, c1_w, c1_h = Inches(0.22), Inches(3.61), Inches(6.80), Inches(3.41)
        card1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, c1_left, c1_top, c1_w, c1_h)
        card1.fill.solid()
        card1.fill.fore_color.rgb = CONTAINER_BG
        card1.line.color.rgb = CARD_BORDER

        chart_data_bar = CategoryChartData()
        chart_data_bar.categories = [f"{d:02d}" for d in range(1, num_days + 1)]
        for disp_name in ae_contribution:
            chart_data_bar.add_series(disp_name, ae_daily_matrix[disp_name]["output"])

        ch1_x, ch1_y, ch1_w, ch1_h = c1_left + Inches(0.12), c1_top + Inches(0.12), c1_w - Inches(0.24), c1_h - Inches(
            0.24)
        bar_chart_shape = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, ch1_x, ch1_y, ch1_w, ch1_h,
                                                 chart_data_bar)
        bar_chart = bar_chart_shape.chart

        bar_chart.has_title = True
        bar_chart.chart_title.text_frame.text = f"DAILY OUTPUT OF {st_name.upper()} STATION"
        p_c1 = bar_chart.chart_title.text_frame.paragraphs[0]
        p_c1.font.name, p_c1.font.size, p_c1.font.bold, p_c1.font.color.rgb = "Aptos", Pt(16), True, RGBColor(60, 60,
                                                                                                              60)

        bar_chart.has_legend = True
        bar_chart.legend.position = XL_LEGEND_POSITION.TOP
        bar_chart.legend.include_in_layout = False
        bar_chart.legend.font.name, bar_chart.legend.font.size = "Aptos", Pt(8)

        cat_axis = bar_chart.category_axis
        cat_axis.tick_labels.font.name, cat_axis.tick_labels.font.size = "Aptos", Pt(7)
        cat_axis.tick_labels.offset = 100
        cat_axis.has_major_gridlines = False

        val_axis = bar_chart.value_axis
        val_axis.tick_labels.font.name, val_axis.tick_labels.font.size = "Aptos", Pt(8)
        val_axis.has_major_gridlines = True
        val_axis.major_gridlines.format.line.color.rgb = RGBColor(230, 230, 230)

        for s_idx, series in enumerate(bar_chart.series):
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = CHART_SERIES_COLORS[s_idx % len(CHART_SERIES_COLORS)]

        # Right Chart: Contribution Pie
        c2_left, c2_top, c2_w, c2_h = Inches(7.17), Inches(3.61), Inches(4.78), Inches(2.81)
        card2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, c2_left, c2_top, c2_w, c2_h)
        card2.fill.solid()
        card2.fill.fore_color.rgb = CONTAINER_BG
        card2.line.color.rgb = CARD_BORDER

        chart_data_pie = CategoryChartData()
        chart_data_pie.categories = list(ae_contribution.keys())
        chart_data_pie.add_series("Contribution", [v if v > 0 else 0.001 for v in ae_contribution.values()])

        ch2_x, ch2_y, ch2_w, ch2_h = c2_left + Inches(0.1), c2_top + Inches(0.08), c2_w - Inches(0.2), c2_h - Inches(
            0.16)
        pie_chart_shape = slide.shapes.add_chart(XL_CHART_TYPE.PIE, ch2_x, ch2_y, ch2_w, ch2_h, chart_data_pie)
        pie_chart = pie_chart_shape.chart
        pie_chart.style = 1

        pie_chart.has_title = True
        pie_chart.chart_title.text_frame.text = "CONTRIBUTION"
        p_c2 = pie_chart.chart_title.text_frame.paragraphs[0]
        p_c2.font.name, p_c2.font.size, p_c2.font.bold, p_c2.font.color.rgb = "Aptos", Pt(16), True, RGBColor(0, 0, 0)
        pie_chart.has_legend = False

        plots = pie_chart.plots[0]
        plots.has_data_labels = True
        data_labels = plots.data_labels
        data_labels.position = XL_DATA_LABEL_POSITION.OUTSIDE_END
        data_labels.show_category_name, data_labels.show_percentage, data_labels.show_value = True, True, False
        data_labels.font.name, data_labels.font.size = "Aptos", Pt(7)

        try:
            for s_idx, point in enumerate(plots.series[0].points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = CHART_SERIES_COLORS[s_idx % len(CHART_SERIES_COLORS)]
        except Exception:
            pass

        # MoM Pill Box
        p_left, p_top, p_w, p_h = Inches(7.17), Inches(6.52), Inches(4.78), Inches(0.51)
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, p_left, p_top, p_w, p_h)
        pill.fill.solid()
        pill.fill.fore_color.rgb = PILL_BG
        pill.line.color.rgb = PILL_BORDER
        pill.line.width = Pt(1.5)

        tf_p = pill.text_frame
        tf_p.word_wrap = True
        tf_p.vertical_anchor = MSO_ANCHOR.MIDDLE

        p_pi = tf_p.paragraphs[0]
        p_pi.alignment = PP_ALIGN.CENTER

        r1 = p_pi.add_run()
        r1.text = "Overall Productivity Index: "
        r1.font.name, r1.font.bold, r1.font.size, r1.font.color.rgb = "Arial", True, Pt(10.5), TEXT_BLACK

        r2 = p_pi.add_run()
        r2.text = f"{st_pi:.1f} "
        r2.font.name, r2.font.size, r2.font.color.rgb = "Arial", Pt(12), TEXT_BLACK

        r3 = p_pi.add_run()
        r3.text = f"{pi_delta_str}  |  "
        r3.font.name, r3.font.size = "Arial", Pt(10.5)
        r3.font.color.rgb = GREEN_DELTA if "+" in pi_delta_str else (RED_DELTA if "-" in pi_delta_str else TEXT_BLACK)

        r4 = p_pi.add_run()
        r4.text = "Total Units: "
        r4.font.name, r4.font.bold, r4.font.size, r4.font.color.rgb = "Arial", True, Pt(10.5), TEXT_BLACK

        r5 = p_pi.add_run()
        r5.text = f"{st_cur_units} "
        r5.font.name, r5.font.size, r5.font.color.rgb = "Arial", Pt(12), TEXT_BLACK

        r6 = p_pi.add_run()
        r6.text = units_delta_str
        r6.font.name, r6.font.size = "Arial", Pt(10.5)
        r6.font.color.rgb = GREEN_DELTA if "+" in units_delta_str else (
            RED_DELTA if "-" in units_delta_str else TEXT_BLACK)

    # DEPARTMENT SUMMARY SLIDE
    sum_slide = prs.slides.add_slide(blank_layout)
    set_slide_background(sum_slide, prs, "Grey Wave.png", BG_COLOR)

    num_st = len(summary_slide_data)
    if num_st == 0:
        empty_box = sum_slide.shapes.add_textbox(Inches(2.0), Inches(3.0), Inches(9.33), Inches(1.5))
        p_emp = empty_box.text_frame.paragraphs[0]
        p_emp.text = f"No calibration output recorded for {month_title} {year}."
        p_emp.font.name, p_emp.font.size, p_emp.font.color.rgb = "Arial", Pt(20), WHITE_TEXT
        p_emp.alignment = PP_ALIGN.CENTER
    else:
        s_rows = 4
        s_cols = 1 + num_st + 1

        s_table_shape = sum_slide.shapes.add_table(s_rows, s_cols, Inches(0.4), Inches(0.35), Inches(12.5), Inches(2.2))
        s_tbl = s_table_shape.table
        set_native_table_grid_style(s_tbl)

        s_tbl.columns[0].width = Inches(1.3)
        col_w = int((Inches(12.5) - Inches(2.5)) / max(num_st, 1))
        for c in range(1, s_cols - 1):
            s_tbl.columns[c].width = col_w
        s_tbl.columns[s_cols - 1].width = Inches(1.2)

        format_cell_native(s_tbl.cell(0, 0), RUST_ORANGE, "STATION", font_size=12, bold=True, text_color=WHITE_TEXT)

        tot_units_all = sum(d["units"] for d in summary_slide_data)
        prev_tot_units = sum(d["prev_units"] for d in summary_slide_data)
        tot_units_delta_str = format_delta(tot_units_all, prev_tot_units, is_float=False)

        tot_pi_sum = round(sum(d["pi"] for d in summary_slide_data), 1)
        prev_tot_pi_sum = round(sum(d["prev_pi"] for d in summary_slide_data), 1)
        tot_pi_delta_str = format_delta(tot_pi_sum, prev_tot_pi_sum, is_float=True)

        dept_avg_pi = round(tot_pi_sum / num_st, 1) if num_st > 0 else 0.0
        prev_dept_avg_pi = round(prev_tot_pi_sum / num_st, 1) if num_st > 0 else 0.0
        avg_pi_delta_str = format_delta(dept_avg_pi, prev_dept_avg_pi, is_float=True)

        for i, d in enumerate(summary_slide_data):
            c_idx = i + 1
            format_cell_native(s_tbl.cell(0, c_idx), RUST_ORANGE, d["name"], font_size=12, bold=True,
                               text_color=WHITE_TEXT)
            format_cell_native(s_tbl.cell(1, c_idx), RGBColor(255, 255, 255), str(d["units"]), font_size=11, bold=False,
                               text_color=TEXT_BLACK)
            format_cell_native(s_tbl.cell(2, c_idx), RGBColor(255, 255, 255), f"{d['pi']:.1f}", font_size=11,
                               bold=False, text_color=TEXT_BLACK)

        format_cell_native(s_tbl.cell(0, s_cols - 1), RUST_ORANGE, "TOTAL", font_size=12, bold=True,
                           text_color=WHITE_TEXT)

        c_out_tot = s_tbl.cell(1, s_cols - 1)
        format_cell_native(c_out_tot, RGBColor(255, 255, 255), f"{tot_units_all} ", font_size=11, bold=True,
                           text_color=TEXT_BLACK)
        r_del1 = c_out_tot.text_frame.paragraphs[0].add_run()
        r_del1.text = tot_units_delta_str
        r_del1.font.bold = True
        r_del1.font.color.rgb = GREEN_DELTA if "+" in tot_units_delta_str else (
            RED_DELTA if "-" in tot_units_delta_str else TEXT_BLACK)

        c_pi_tot = s_tbl.cell(2, s_cols - 1)
        format_cell_native(c_pi_tot, RGBColor(255, 255, 255), f"{tot_pi_sum:.1f} ", font_size=11, bold=True,
                           text_color=TEXT_BLACK)
        r_del2 = c_pi_tot.text_frame.paragraphs[0].add_run()
        r_del2.text = tot_pi_delta_str
        r_del2.font.bold = True
        r_del2.font.color.rgb = GREEN_DELTA if "+" in tot_pi_delta_str else (
            RED_DELTA if "-" in tot_pi_delta_str else TEXT_BLACK)

        format_cell_native(s_tbl.cell(1, 0), RGBColor(255, 255, 255), "OUTPUT", font_size=11, bold=True,
                           text_color=TEXT_BLACK)
        format_cell_native(s_tbl.cell(2, 0), RGBColor(255, 255, 255), "P.I.", font_size=11, bold=True,
                           text_color=TEXT_BLACK)

        c_avg_left = s_tbl.cell(3, 0)
        if s_cols > 2:
            c_avg_left.merge(s_tbl.cell(3, s_cols - 2))
        format_cell_native(c_avg_left, CREAM_BEIGE, "AVERAGE P.I. (Total P.I. divided by # of Stations)", font_size=11,
                           bold=True, text_color=TEXT_BLACK, align=PP_ALIGN.RIGHT)

        c_avg_val = s_tbl.cell(3, s_cols - 1)
        format_cell_native(c_avg_val, CREAM_BEIGE, f"{dept_avg_pi:.1f} ", font_size=12, bold=True,
                           text_color=TEXT_BLACK)
        r_del3 = c_avg_val.text_frame.paragraphs[0].add_run()
        r_del3.text = avg_pi_delta_str
        r_del3.font.bold = True
        r_del3.font.color.rgb = GREEN_DELTA if "+" in avg_pi_delta_str else (
            RED_DELTA if "-" in avg_pi_delta_str else TEXT_BLACK)

        ch_x, ch_y, ch_w, ch_h = Inches(3.28), Inches(2.94), Inches(6.74), Inches(3.97)

        sum_card = sum_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, ch_x, ch_y, ch_w, ch_h)
        sum_card.fill.solid()
        sum_card.fill.fore_color.rgb = RGBColor(255, 255, 255)
        sum_card.line.color.rgb = RGBColor(210, 210, 210)

        st_categories = [d["name"] for d in summary_slide_data]
        st_units_vals = [d["units"] for d in summary_slide_data]
        st_pi_vals = [d["pi"] for d in summary_slide_data]
        min_target_vals = [2.75] * len(st_categories)

        data_col = CategoryChartData()
        data_col.categories = st_categories
        data_col.add_series("OUTPUT", st_units_vals)

        base_shape = sum_slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            ch_x + Inches(0.12), ch_y + Inches(0.08), ch_w - Inches(0.24), ch_h - Inches(0.48),
            data_col
        )
        base_chart = base_shape.chart
        base_chart.style = 1

        base_chart.has_title = True
        base_chart.chart_title.text_frame.text = "OUTPUT & P.I."
        p_ct = base_chart.chart_title.text_frame.paragraphs[0]
        p_ct.font.name, p_ct.font.size, p_ct.font.bold, p_ct.font.color.rgb = "Aptos", Pt(18), True, RGBColor(60, 60,
                                                                                                              60)
        base_chart.has_legend = False

        val_ax1 = base_chart.value_axis
        val_ax1.tick_labels.font.name, val_ax1.tick_labels.font.size = "Aptos", Pt(8)
        val_ax1.has_major_gridlines = True
        val_ax1.major_gridlines.format.line.color.rgb = RGBColor(210, 210, 210)

        cat_ax1 = base_chart.category_axis
        cat_ax1.tick_labels.font.name, cat_ax1.tick_labels.font.size = "Aptos", Pt(8)
        cat_ax1.has_major_gridlines = False

        base_chart.series[0].format.fill.solid()
        base_chart.series[0].format.fill.fore_color.rgb = OUTPUT_PURPLE

        data_line = CategoryChartData()
        data_line.categories = st_categories
        data_line.add_series("P.I.", st_pi_vals)
        data_line.add_series("Min Target", min_target_vals)

        line_shape = sum_slide.shapes.add_chart(
            XL_CHART_TYPE.LINE,
            ch_x + Inches(0.12), ch_y + Inches(0.08), ch_w - Inches(0.24), ch_h - Inches(0.48),
            data_line
        )
        line_chart = line_shape.chart
        line_chart.has_title = False
        line_chart.has_legend = False

        try:
            chart_space = line_chart._element
            sp_pr = chart_space.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}spPr')
            if sp_pr is None:
                sp_pr = parse_xml(
                    '<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></c:spPr>')
                chart_space.insert(0, sp_pr)
            else:
                sp_pr.append(parse_xml('<a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'))

            plot_area = chart_space.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}plotArea')
            if plot_area is not None:
                pa_sp_pr = plot_area.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}spPr')
                if pa_sp_pr is None:
                    pa_sp_pr = parse_xml(
                        '<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></c:spPr>')
                    plot_area.insert(0, pa_sp_pr)
                else:
                    pa_sp_pr.append(
                        parse_xml('<a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'))
        except Exception:
            pass

        val_ax2 = line_chart.value_axis
        val_ax2.tick_labels.font.name, val_ax2.tick_labels.font.size = "Aptos", Pt(8)
        val_ax2.has_major_gridlines = False

        try:
            val_ax2_elem = val_ax2._element
            axPos_elem = val_ax2_elem.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}axPos')
            if axPos_elem is not None:
                axPos_elem.set('val', 'r')

            crosses_elem = val_ax2_elem.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}crosses')
            if crosses_elem is not None:
                crosses_elem.set('val', 'max')
            else:
                crossAx = val_ax2_elem.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}crossAx')
                if crossAx is not None:
                    new_crosses = parse_xml(
                        '<c:crosses xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" val="max"/>')
                    crossAx.addprevious(new_crosses)
        except Exception:
            pass

        line_chart.category_axis.visible = False

        s_pi = line_chart.series[0]
        s_pi.format.line.color.rgb = GOLD_PI
        s_pi.format.line.width = Pt(1.5)
        s_pi.format.line.dash_style = MSO_LINE.DASH_DOT
        s_pi.marker.style = XL_MARKER_STYLE.CIRCLE
        s_pi.marker.size = 5
        s_pi.marker.format.fill.solid()
        s_pi.marker.format.fill.fore_color.rgb = GOLD_PI
        s_pi.marker.format.line.color.rgb = GOLD_PI

        s_tgt = line_chart.series[1]
        s_tgt.format.line.color.rgb = BLACK_TARGET
        s_tgt.format.line.width = Pt(1.5)
        s_tgt.format.line.dash_style = MSO_LINE.DASH
        s_tgt.marker.style = XL_MARKER_STYLE.NONE

        leg_box = sum_slide.shapes.add_textbox(ch_x, ch_y + ch_h - Inches(0.38), ch_w, Inches(0.35))
        tf_leg = leg_box.text_frame
        p_leg = tf_leg.paragraphs[0]
        p_leg.alignment = PP_ALIGN.CENTER

        r_l1 = p_leg.add_run()
        r_l1.text = "■ OUTPUT       "
        r_l1.font.name, r_l1.font.bold, r_l1.font.size, r_l1.font.color.rgb = "Aptos", True, Pt(8), OUTPUT_PURPLE

        r_l2 = p_leg.add_run()
        r_l2.text = "— · — P.I.       "
        r_l2.font.name, r_l2.font.bold, r_l2.font.size, r_l2.font.color.rgb = "Aptos", True, Pt(8), GOLD_PI

        r_l3 = p_leg.add_run()
        r_l3.text = "- - - Min Target"
        r_l3.font.name, r_l3.font.bold, r_l3.font.size, r_l3.font.color.rgb = "Aptos", True, Pt(8), BLACK_TARGET

    # SLIDE N+2: SECTION DIVIDER 3
    slide_div_q = prs.slides.add_slide(blank_layout)
    set_slide_background(slide_div_q, prs, "Black.png", PURE_BLACK)

    t_box_dq = slide_div_q.shapes.add_textbox(Inches(1.0), Inches(3.0), Inches(11.33), Inches(1.5))
    p_dq = t_box_dq.text_frame.paragraphs[0]
    p_dq.text = "Quality Index & KPIs"
    p_dq.font.name, p_dq.font.bold, p_dq.font.size, p_dq.font.color.rgb = "Arial", True, Pt(48), WHITE_TEXT
    p_dq.alignment = PP_ALIGN.CENTER

    # SLIDE N+3: EXECUTIVE KPI SCORECARD
    add_kpi_scorecard_slide(prs, conn, year, month)

    # SLIDE N+4: QUALITY INDEX
    slide_q = prs.slides.add_slide(blank_layout)
    set_slide_background(slide_q, prs, "Grey Wave.png", BG_COLOR)

    t_box_q = slide_q.shapes.add_textbox(Inches(1.51), Inches(0.37), Inches(11.5), Inches(0.65))
    p_qt = t_box_q.text_frame.paragraphs[0]
    p_qt.text = "QUALITY INDEX"
    p_qt.font.name, p_qt.font.bold, p_qt.font.size, p_qt.font.color.rgb = "Arial", True, Pt(24), CORAL_TITLE

    rolling_months = []
    cur_y, cur_m = year, month
    for _ in range(12):
        rolling_months.append((cur_y, cur_m))
        cur_m -= 1
        if cur_m == 0:
            cur_m = 12
            cur_y -= 1
    rolling_months.reverse()

    df_all_cals = pd.read_sql(
        "SELECT cal_year, cal_month, COUNT(*) as units FROM calibrations GROUP BY cal_year, cal_month", conn)
    cal_lookup = dict(zip(zip(df_all_cals["cal_year"], df_all_cals["cal_month"]),
                          df_all_cals["units"])) if not df_all_cals.empty else {}

    df_all_rejs = pd.read_sql(
        "SELECT cal_by, reject_year, reject_month, failure_reason, COUNT(*) as cnt FROM rejection_logs GROUP BY cal_by, reject_year, reject_month, failure_reason",
        conn)
    if not df_all_rejs.empty:
        df_all_rejs["clean_cal_by"] = df_all_rejs["cal_by"].astype(str).str.strip().str.lower()
        df_all_rejs = df_all_rejs[
            (~df_all_rejs["clean_cal_by"].isin(muted_users)) & (df_all_rejs["clean_cal_by"].isin(roster_users))]

    month_labels = [f"{calendar.month_abbr[m_num]} '{str(y_num)[-2:]}" for y_num, m_num in rolling_months]
    acc_rates, coc_counts, mr_counts, others_counts = [], [], [], []

    for y_num, m_num in rolling_months:
        tot_c = cal_lookup.get((y_num, m_num), 0)
        m_rejs = df_all_rejs[(df_all_rejs["reject_year"] == y_num) & (
                df_all_rejs["reject_month"] == m_num)] if not df_all_rejs.empty else pd.DataFrame()
        tot_err = m_rejs["cnt"].sum() if not m_rejs.empty else 0

        coc_cnt = m_rejs[m_rejs["failure_reason"].str.lower().str.contains("coc")][
            "cnt"].sum() if not m_rejs.empty else 0
        mr_cnt = m_rejs[m_rejs["failure_reason"].str.lower().str.contains("measurement|report")][
            "cnt"].sum() if not m_rejs.empty else 0
        oth_cnt = tot_err - (coc_cnt + mr_cnt)

        coc_counts.append(int(coc_cnt))
        mr_counts.append(int(mr_cnt))
        others_counts.append(int(oth_cnt))

        if tot_c > 0:
            rate = round(((tot_c - tot_err) / tot_c) * 100.0, 1)
        else:
            rate = 100.0 if tot_err == 0 else 98.0
        acc_rates.append(rate)

    q_rows = 5
    q_cols = 2 + 12
    q_table_shape = slide_q.shapes.add_table(q_rows, q_cols, Inches(0.4), Inches(0.95), Inches(12.5), Inches(2.2))
    q_tbl = q_table_shape.table
    set_native_table_grid_style(q_tbl)

    q_tbl.columns[0].width = Inches(0.5)
    q_tbl.columns[1].width = Inches(1.8)
    q_col_w = int((Inches(12.5) - Inches(2.3)) / 12)
    for c in range(2, q_cols):
        q_tbl.columns[c].width = q_col_w

    c_m0 = q_tbl.cell(0, 0)
    c_m0.merge(q_tbl.cell(0, 1))
    format_cell_native(c_m0, BLUE_HEADER, "MONTH", font_size=10, bold=True, text_color=WHITE_TEXT)

    for i, lbl in enumerate(month_labels):
        format_cell_native(q_tbl.cell(0, i + 2), BLUE_HEADER, lbl, font_size=9, bold=True, text_color=WHITE_TEXT)

    c_ar0 = q_tbl.cell(1, 0)
    c_ar0.merge(q_tbl.cell(1, 1))
    format_cell_native(c_ar0, RGBColor(230, 240, 250), "Acceptance Rate (%)", font_size=10, bold=True,
                       text_color=TEXT_BLACK)

    for i, r_val in enumerate(acc_rates):
        format_cell_native(q_tbl.cell(1, i + 2), RGBColor(230, 240, 250), f"{r_val:.1f}", font_size=9, bold=False,
                           text_color=TEXT_BLACK)

    c_mis = q_tbl.cell(2, 0)
    c_mis.merge(q_tbl.cell(4, 0))
    format_cell_native(c_mis, BLUE_HEADER, "MISTAKES", font_size=9, bold=True, text_color=WHITE_TEXT)

    format_cell_native(q_tbl.cell(2, 1), RGBColor(255, 255, 255), "COC Issue", font_size=9, bold=True,
                       text_color=TEXT_BLACK)
    format_cell_native(q_tbl.cell(3, 1), RGBColor(255, 255, 255), "Measurement Report", font_size=9, bold=True,
                       text_color=TEXT_BLACK)
    format_cell_native(q_tbl.cell(4, 1), RGBColor(255, 255, 255), "Others", font_size=9, bold=True,
                       text_color=TEXT_BLACK)

    for i in range(12):
        c_i = i + 2
        format_cell_native(q_tbl.cell(2, c_i), RGBColor(255, 255, 255), str(coc_counts[i]) if coc_counts[i] > 0 else "",
                           font_size=9, bold=False, text_color=TEXT_BLACK)
        format_cell_native(q_tbl.cell(3, c_i), RGBColor(255, 255, 255), str(mr_counts[i]) if mr_counts[i] > 0 else "",
                           font_size=9, bold=False, text_color=TEXT_BLACK)
        format_cell_native(q_tbl.cell(4, c_i), RGBColor(255, 255, 255),
                           str(others_counts[i]) if others_counts[i] > 0 else "", font_size=9, bold=False,
                           text_color=TEXT_BLACK)

    cq1_x, cq1_y, cq1_w, cq1_h = Inches(0.4), Inches(3.4), Inches(6.1), Inches(3.4)
    card_q1 = slide_q.shapes.add_shape(MSO_SHAPE.RECTANGLE, cq1_x, cq1_y, cq1_w, cq1_h)
    card_q1.fill.solid()
    card_q1.fill.fore_color.rgb = CONTAINER_BG
    card_q1.line.color.rgb = CARD_BORDER

    chart_data_qline = CategoryChartData()
    chart_data_qline.categories = month_labels
    chart_data_qline.add_series("Actual", acc_rates)
    chart_data_qline.add_series("Target", [98.0] * 12)

    ch_q1_shape = slide_q.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        cq1_x + Inches(0.12), cq1_y + Inches(0.08), cq1_w - Inches(0.24), cq1_h - Inches(0.16),
        chart_data_qline
    )
    ch_q1 = ch_q1_shape.chart

    ch_q1.has_title = True
    ch_q1.chart_title.text_frame.text = "QUALITY INDEX"
    p_q1_t = ch_q1.chart_title.text_frame.paragraphs[0]
    p_q1_t.font.name, p_q1_t.font.size, p_q1_t.font.bold, p_q1_t.font.color.rgb = "Aptos", Pt(15), True, RGBColor(60,
                                                                                                                  60,
                                                                                                                  60)

    ch_q1.has_legend = True
    ch_q1.legend.position = XL_LEGEND_POSITION.TOP
    ch_q1.legend.include_in_layout = False
    ch_q1.legend.font.name, ch_q1.legend.font.size = "Aptos", Pt(8)

    ch_q1.category_axis.tick_labels.font.name, ch_q1.category_axis.tick_labels.font.size = "Aptos", Pt(7.5)
    ch_q1.category_axis.has_major_gridlines = False

    ch_q1.value_axis.tick_labels.font.name, ch_q1.value_axis.tick_labels.font.size = "Aptos", Pt(8)
    ch_q1.value_axis.has_major_gridlines = True
    ch_q1.value_axis.major_gridlines.format.line.color.rgb = RGBColor(230, 230, 230)

    s_act = ch_q1.series[0]
    s_act.format.line.color.rgb = GREEN_ACTUAL
    s_act.format.line.width = Pt(1.5)
    s_act.marker.style = XL_MARKER_STYLE.CIRCLE
    s_act.marker.size = 5
    s_act.marker.format.fill.solid()
    s_act.marker.format.fill.fore_color.rgb = GREEN_ACTUAL
    s_act.marker.format.line.color.rgb = GREEN_ACTUAL

    s_tgt_q = ch_q1.series[1]
    s_tgt_q.format.line.color.rgb = BLACK_TARGET
    s_tgt_q.format.line.width = Pt(1.5)
    s_tgt_q.format.line.dash_style = MSO_LINE.DASH
    s_tgt_q.marker.style = XL_MARKER_STYLE.NONE

    cq2_x, cq2_y, cq2_w, cq2_h = Inches(6.8), Inches(3.4), Inches(6.1), Inches(3.4)
    card_q2 = slide_q.shapes.add_shape(MSO_SHAPE.RECTANGLE, cq2_x, cq2_y, cq2_w, cq2_h)
    card_q2.fill.solid()
    card_q2.fill.fore_color.rgb = CONTAINER_BG
    card_q2.line.color.rgb = CARD_BORDER

    chart_data_qbar = CategoryChartData()
    chart_data_qbar.categories = month_labels
    chart_data_qbar.add_series("COC Issue", coc_counts)
    chart_data_qbar.add_series("Measurement Report", mr_counts)
    chart_data_qbar.add_series("Others", others_counts)

    ch_q2_shape = slide_q.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_STACKED,
        cq2_x + Inches(0.12), cq2_y + Inches(0.08), cq2_w - Inches(0.24), cq2_h - Inches(0.16),
        chart_data_qbar
    )
    ch_q2 = ch_q2_shape.chart

    ch_q2.has_title = True
    ch_q2.chart_title.text_frame.text = "TYPE OF MISTAKE"
    p_q2_t = ch_q2.chart_title.text_frame.paragraphs[0]
    p_q2_t.font.name, p_q2_t.font.size, p_q2_t.font.bold, p_q2_t.font.color.rgb = "Aptos", Pt(15), True, RGBColor(60,
                                                                                                                  60,
                                                                                                                  60)

    ch_q2.has_legend = True
    ch_q2.legend.position = XL_LEGEND_POSITION.TOP
    ch_q2.legend.include_in_layout = False
    ch_q2.legend.font.name, ch_q2.legend.font.size = "Aptos", Pt(8)

    cat_ax2 = ch_q2.category_axis
    cat_ax2.tick_labels.font.name, cat_ax2.tick_labels.font.size = "Aptos", Pt(7.5)
    cat_ax2.tick_labels.offset = 100
    cat_ax2.has_major_gridlines = False

    val_ax2 = ch_q2.value_axis
    val_ax2.tick_labels.font.name, val_ax2.tick_labels.font.size = "Aptos", Pt(8)
    val_ax2.has_major_gridlines = True
    val_ax2.major_gridlines.format.line.color.rgb = RGBColor(230, 230, 230)

    if len(ch_q2.series) >= 3:
        ch_q2.series[0].format.fill.solid()
        ch_q2.series[0].format.fill.fore_color.rgb = RGBColor(214, 40, 57)
        ch_q2.series[1].format.fill.solid()
        ch_q2.series[1].format.fill.fore_color.rgb = RGBColor(237, 125, 49)
        ch_q2.series[2].format.fill.solid()
        ch_q2.series[2].format.fill.fore_color.rgb = RGBColor(112, 48, 160)

    fn_box = slide_q.shapes.add_textbox(Inches(0.4), Inches(6.9), Inches(5.0), Inches(0.4))
    p_fn = fn_box.text_frame.paragraphs[0]
    p_fn.text = "Detail breakdown will be the next slide."
    p_fn.font.name, p_fn.font.size, p_fn.font.italic, p_fn.font.color.rgb = "Arial", Pt(11), True, RGBColor(255, 235,
                                                                                                            156)

    # SLIDE N+5: DESCRIPTION OF OUTGOING ERRORS
    slide_err = prs.slides.add_slide(blank_layout)
    set_slide_background(slide_err, prs, "Grey Wave.png", BG_COLOR)

    t_box_err = slide_err.shapes.add_textbox(Inches(1.51), Inches(0.37), Inches(11.5), Inches(0.65))
    p_errt = t_box_err.text_frame.paragraphs[0]
    p_errt.text = "DESCRIPTION OF OUTGOING ERRORS"
    p_errt.font.name, p_errt.font.bold, p_errt.font.size, p_errt.font.color.rgb = "Arial", True, Pt(24), CORAL_TITLE

    df_cur_rejs = pd.read_sql("""
                              SELECT cal_by, model_id, serial_no, reject_date, failure_reason, remark
                              FROM rejection_logs
                              WHERE reject_year = ?
                                AND reject_month = ?
                              ORDER BY reject_date ASC
                              """, conn, params=[year, month])

    if not df_cur_rejs.empty:
        df_cur_rejs["clean_cal_by"] = df_cur_rejs["cal_by"].astype(str).str.strip().str.lower()
        df_cur_rejs = df_cur_rejs[
            (~df_cur_rejs["clean_cal_by"].isin(muted_users)) & (df_cur_rejs["clean_cal_by"].isin(roster_users))]

    err_rows = max(2, len(df_cur_rejs) + 1)
    err_table_shape = slide_err.shapes.add_table(err_rows, 6, Inches(0.4), Inches(1.15), Inches(12.5),
                                                 Inches(0.35 * err_rows))
    err_tbl = err_table_shape.table
    set_native_table_grid_style(err_tbl)

    err_tbl.columns[0].width = Inches(1.8)
    err_tbl.columns[1].width = Inches(1.5)
    err_tbl.columns[2].width = Inches(1.8)
    err_tbl.columns[3].width = Inches(1.6)
    err_tbl.columns[4].width = Inches(1.8)
    err_tbl.columns[5].width = Inches(4.0)

    err_headers = ["AE Resp.", "Model No", "Serial No", "Rejected Date", "Failure Reason", "Remarks by Inspector"]
    for c_i, h_text in enumerate(err_headers):
        format_cell_native(err_tbl.cell(0, c_i), BLUE_HEADER, h_text, font_size=10, bold=True, text_color=WHITE_TEXT)

    if not df_cur_rejs.empty:
        for r_i, row in df_cur_rejs.reset_index(drop=True).iterrows():
            curr_r = r_i + 1
            raw_u = str(row["cal_by"]).strip().lower()
            ae_display = username_to_name.get(raw_u, str(row["cal_by"]))

            format_cell_native(err_tbl.cell(curr_r, 0), RGBColor(255, 255, 255), ae_display, font_size=9, bold=True,
                               text_color=TEXT_BLACK)
            format_cell_native(err_tbl.cell(curr_r, 1), RGBColor(255, 255, 255), str(row["model_id"]), font_size=9,
                               text_color=TEXT_BLACK)
            format_cell_native(err_tbl.cell(curr_r, 2), RGBColor(255, 255, 255), str(row["serial_no"]), font_size=9,
                               text_color=TEXT_BLACK)
            format_cell_native(err_tbl.cell(curr_r, 3), RGBColor(255, 255, 255), str(row["reject_date"]), font_size=9,
                               text_color=TEXT_BLACK)
            format_cell_native(err_tbl.cell(curr_r, 4), RGBColor(255, 255, 255), str(row["failure_reason"]),
                               font_size=9, text_color=TEXT_BLACK)
            format_cell_native(err_tbl.cell(curr_r, 5), RGBColor(255, 255, 255), str(row["remark"]), font_size=9,
                               text_color=TEXT_BLACK, align=PP_ALIGN.LEFT)
    else:
        c_no_err = err_tbl.cell(1, 0)
        c_no_err.merge(err_tbl.cell(1, 5))
        format_cell_native(c_no_err, RGBColor(255, 255, 255), "No outgoing inspection errors recorded for this period.",
                           font_size=11, bold=True, text_color=GREEN_DELTA)

    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream
