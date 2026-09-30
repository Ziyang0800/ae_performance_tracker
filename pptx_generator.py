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
BG_COLOR = RGBColor(43, 45, 47)  # Dark Charcoal #2B2D2F
PURE_BLACK = RGBColor(0, 0, 0)  # Pure Black #000000
CORAL_TITLE = RGBColor(255, 89, 118)  # #FF5976
PINK_SUBTITLE = RGBColor(255, 199, 208)  # #FFC7D0
PENTRAC_RED = RGBColor(214, 40, 57)  # Accent Red

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

# Resource Level Report Palette (Slide 4)
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

# Special Typography Colors
CODE_SPECIAL = RGBColor(255, 89, 118)
TEXT_BLACK = RGBColor(0, 0, 0)
WHITE_TEXT = RGBColor(255, 255, 255)

# Containers & MoM Pill Box
CONTAINER_BG = RGBColor(255, 255, 255)
CARD_BORDER = RGBColor(74, 189, 196)
PILL_BG = RGBColor(216, 243, 245)
PILL_BORDER = RGBColor(74, 189, 196)
GREEN_DELTA = RGBColor(46, 204, 113)
RED_DELTA = RGBColor(231, 76, 60)

# Station Chart Colors
CHART_SERIES_COLORS = [
    RGBColor(255, 192, 0),  # Gold
    RGBColor(55, 96, 146),  # Steel Blue
    RGBColor(0, 176, 80),  # Green
    RGBColor(155, 89, 182),  # Purple
    RGBColor(230, 126, 34)  # Orange
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
    try:
        return float(s) > 0
    except ValueError:
        return False


def format_cell_native(cell, fill_rgb, text, font_size=8, bold=False, text_color=TEXT_BLACK, align=PP_ALIGN.CENTER,
                       margin_left=None):
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
    return f"(+{diff:.1f})" if (is_float and diff > 0) else (
        f"(+{diff})" if diff > 0 else (f"({diff:.1f})" if is_float else f"({diff})"))


def parse_iso_date(date_val):
    if pd.isna(date_val) or not date_val: return None
    try:
        ts = pd.to_datetime(str(date_val).strip(), errors="coerce")
        return None if pd.isna(ts) else ts.date()
    except Exception:
        return None


def resolve_engineer_station(conn, engineer_id, fallback_station_id, rep_year, rep_month):
    row = conn.execute("""
                       SELECT h.station_id, s.name as station_name, s.sort_order, s.target_pi, s.is_active
                       FROM engineer_station_history h
                                JOIN stations s ON h.station_id = s.id
                       WHERE h.engineer_id = ?
                         AND (h.effective_year < ? OR (h.effective_year = ? AND h.effective_month <= ?))
                       ORDER BY h.effective_year DESC, h.effective_month DESC LIMIT 1
                       """, (engineer_id, rep_year, rep_year, rep_month)).fetchone()

    if row:
        return row["station_id"], row["station_name"], row["sort_order"], row["target_pi"], row["is_active"]

    fb = conn.execute(
        "SELECT s.id as station_id, s.name as station_name, s.sort_order, s.target_pi, s.is_active FROM stations s WHERE s.id = ?",
        (fallback_station_id,)).fetchone()
    if fb:
        return fb["station_id"], fb["station_name"], fb["sort_order"], fb["target_pi"], fb["is_active"]
    return None, "Unassigned", 99, 3.0, 1


# =============================================================================
# GENERATE QUARTERLY PRESENTATION DECK (SLIDE 1 & SLIDE 2)
# =============================================================================
def generate_quarterly_deck(conn, selected_quarter):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    periods_db = conn.execute(
        "SELECT DISTINCT cal_year as yr, cal_month as mo FROM calibrations WHERE cal_year IS NOT NULL AND cal_month IS NOT NULL ORDER BY yr ASC, mo ASC").fetchall()
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
                                 SELECT e.id,
                                        e.formal_name,
                                        e.username,
                                        e.station_id,
                                        e.shift_hours,
                                        e.activation_date,
                                        e.resigned_date,
                                        s.name as station_name,
                                        s.group_name
                                 FROM engineers e
                                          LEFT JOIN stations s ON e.station_id = s.id
                                 WHERE (e.track_performance = 1 OR e.track_performance IS NULL)
                                 """).fetchall()

    muted_rows = conn.execute("SELECT username FROM ignored_technicians").fetchall()
    muted_users = set([m["username"].strip().lower() for m in muted_rows])

    df_cals_q = pd.DataFrame()
    for (cy, cm) in quarter_months:
        df_m = pd.read_sql(
            "SELECT service_id, model_id, serial_no, cal_date, cal_day, cal_month, cal_year, cal_by, category, mos FROM calibrations WHERE cal_year = ? AND cal_month = ?",
            conn, params=[cy, cm])
        if not df_m.empty: df_cals_q = pd.concat([df_cals_q, df_m], ignore_index=True)

    if not df_cals_q.empty:
        df_cals_q["clean_by"] = df_cals_q["cal_by"].astype(str).str.strip().str.lower()
        df_cals_q = df_cals_q[~df_cals_q["clean_by"].isin(muted_users)]

    df_att_q = pd.DataFrame()
    for (cy, cm) in quarter_months:
        df_m = pd.read_sql("SELECT formal_name, capped_hours, year, month FROM attendance WHERE year = ? AND month = ?",
                           conn, params=[cy, cm])
        if not df_m.empty: df_att_q = pd.concat([df_att_q, df_m], ignore_index=True)
    if not df_att_q.empty:
        df_att_q["clean_name"] = df_att_q["formal_name"].apply(normalize_str)

    group_stats = {}
    for grp in target_groups:
        st_rows = conn.execute(
            "SELECT id, name FROM stations WHERE is_active = 1 AND (group_name = ? OR (group_name LIKE '%STD%' AND ? = 'STD LAB') OR (group_name LIKE '%LEGACY%' AND ? LIKE '%LEGACY%'))",
            (grp, grp, grp)).fetchall()
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

                # 5-value unpack fix
                st_id, _, _, _, st_active = resolve_engineer_station(conn, e["id"], e["station_id"], cy, cm)
                if not st_active: continue

                if st_id in st_ids:
                    u_clean = e["username"].strip().lower()
                    aes_in_st.add(u_clean)
                    fn_norm = normalize_str(e["formal_name"])
                    shift = e["shift_hours"] if e["shift_hours"] > 0 else 8.0
                    hrs = df_att_q[
                        (df_att_q["clean_name"] == fn_norm) & (df_att_q["year"] == cy) & (df_att_q["month"] == cm)][
                        "capped_hours"].sum() if not df_att_q.empty else 0.0
                    month_days += (hrs / shift)

            m_cals = df_cals_q[(df_cals_q["cal_year"] == cy) & (df_cals_q["cal_month"] == cm) & (
                df_cals_q["clean_by"].isin(aes_in_st))] if (not df_cals_q.empty and aes_in_st) else pd.DataFrame()
            u_cnt = len(m_cals)
            pi_val = round(u_cnt / month_days, 1) if month_days > 0 else 0.0

            w_num, a_w = 0.0, 0.0
            for cat in TAT_CATEGORIES:
                k_key, weight = cat["key"], cat["weight"]
                c_sub = m_cals[
                    m_cals["category"].str.strip().str.upper() == k_key.upper()] if not m_cals.empty else pd.DataFrame()
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

        grp_all_cals = df_cals_q[df_cals_q["clean_by"].isin(aes_in_st)] if (
                    not df_cals_q.empty and aes_in_st) else pd.DataFrame()
        q_w_num, q_a_w = 0.0, 0.0
        for cat in TAT_CATEGORIES:
            k_key, weight = cat["key"], cat["weight"]
            c_sub = grp_all_cals[grp_all_cals[
                                     "category"].str.strip().str.upper() == k_key.upper()] if not grp_all_cals.empty else pd.DataFrame()
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
    card1.fill.solid()
    card1.fill.fore_color.rgb = RGBColor(255, 255, 255)
    card1.line.color.rgb = RGBColor(210, 210, 210)

    data_c1 = CategoryChartData()
    data_c1.categories = target_groups
    for m_i, m_name in enumerate(month_names):
        data_c1.add_series(m_name, [group_stats[g]["m_pis"][m_i] for g in target_groups])

    ch_shape1 = slide1.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, c1_x + Inches(0.12), c1_y + Inches(0.08),
                                        c1_w - Inches(0.24), c1_h - Inches(0.48), data_c1)
    ch1 = ch_shape1.chart
    ch1.has_title = True
    ch1.chart_title.text_frame.text = f"PRODUCTIVITY INDEX OF {q_key}"
    p_t1 = ch1.chart_title.text_frame.paragraphs[0]
    p_t1.font.name, p_t1.font.bold, p_t1.font.size, p_t1.font.color.rgb = "Aptos", True, Pt(16), RGBColor(60, 60, 60)

    blue_palette = [RGBColor(160, 200, 240), RGBColor(46, 117, 182), RGBColor(24, 43, 73)]
    for s_idx, s in enumerate(ch1.series):
        s.format.fill.solid()
        s.format.fill.fore_color.rgb = blue_palette[s_idx % len(blue_palette)]
        s.has_data_labels = True
        s.data_labels.font.name = "Aptos"
        s.data_labels.font.size = Pt(7.5)

    ch1.has_legend = True
    ch1.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch1.legend.font.name, ch1.legend.font.size = "Aptos", Pt(8)

    tbl_shape1 = slide1.shapes.add_table(6, 5, Inches(2.2), Inches(5.15), Inches(8.9), Inches(1.9))
    tbl1 = tbl_shape1.table
    set_native_table_grid_style(tbl1)
    tbl1.columns[0].width = Inches(3.7)
    tbl1.columns[1].width, tbl1.columns[2].width, tbl1.columns[3].width, tbl1.columns[4].width = Inches(1.3), Inches(
        1.3), Inches(1.3), Inches(1.3)

    c_m1 = tbl1.cell(0, 0)
    c_m1.merge(tbl1.cell(0, 4))
    format_cell_native(c_m1, BLUE_HEADER, "PRODUCTIVITY INDEX (P.I.)", font_size=12, bold=True, text_color=WHITE_TEXT)

    format_cell_native(tbl1.cell(1, 0), BLUE_HEADER, "GROUP", font_size=11, bold=True, text_color=WHITE_TEXT)
    for m_i, m_name in enumerate(month_names):
        format_cell_native(tbl1.cell(1, m_i + 1), BLUE_HEADER, m_name, font_size=11, bold=True, text_color=WHITE_TEXT)
    format_cell_native(tbl1.cell(1, 4), BLUE_HEADER, "AVERAGE", font_size=11, bold=True, text_color=WHITE_TEXT)

    for r_i, grp in enumerate(target_groups, start=2):
        format_cell_native(tbl1.cell(r_i, 0), RGBColor(255, 255, 255), grp, font_size=10, bold=True,
                           text_color=TEXT_BLACK, align=PP_ALIGN.CENTER)
        for m_i in range(3):
            format_cell_native(tbl1.cell(r_i, m_i + 1), RGBColor(255, 255, 255),
                               f"{group_stats[grp]['m_pis'][m_i]:.1f}", font_size=11, text_color=TEXT_BLACK)
        format_cell_native(tbl1.cell(r_i, 4), RGBColor(190, 218, 245), f"{group_stats[grp]['quarter_pi']:.1f}",
                           font_size=12, bold=True, text_color=TEXT_BLACK)

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
    card_out.fill.solid()
    card_out.fill.fore_color.rgb = RGBColor(255, 255, 255)
    card_out.line.color.rgb = RGBColor(210, 210, 210)

    data_out = CategoryChartData()
    data_out.categories = target_groups
    for m_i, m_name in enumerate(month_names):
        data_out.add_series(m_name, [group_stats[g]["m_vols"][m_i] for g in target_groups])

    ch_shape_out = slide2.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, c_out_x + Inches(0.1), c_out_y + Inches(0.08),
                                           c_out_w - Inches(0.2), c_out_h - Inches(0.45), data_out)
    ch_out = ch_shape_out.chart
    ch_out.has_title = True
    ch_out.chart_title.text_frame.text = f"OUTPUT OF {q_key}"
    p_to = ch_out.chart_title.text_frame.paragraphs[0]
    p_to.font.name, p_to.font.bold, p_to.font.size, p_to.font.color.rgb = "Aptos", True, Pt(14), RGBColor(60, 60, 60)

    teal_palette = [RGBColor(122, 210, 210), RGBColor(46, 139, 139), RGBColor(21, 75, 75)]
    for s_idx, s in enumerate(ch_out.series):
        s.format.fill.solid()
        s.format.fill.fore_color.rgb = teal_palette[s_idx % len(teal_palette)]
        s.has_data_labels = True
        s.data_labels.font.name = "Aptos"
        s.data_labels.font.size = Pt(7)

    ch_out.has_legend = True
    ch_out.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch_out.legend.font.name, ch_out.legend.font.size = "Aptos", Pt(7.5)

    tbl_out_s = slide2.shapes.add_table(6, 6, Inches(0.4), Inches(5.15), Inches(6.0), Inches(1.9))
    tbl_o = tbl_out_s.table
    set_native_table_grid_style(tbl_o)
    tbl_o.columns[0].width = Inches(2.2)
    for c_i in range(1, 6): tbl_o.columns[c_i].width = Inches(0.76)

    c_mo = tbl_o.cell(0, 0)
    c_mo.merge(tbl_o.cell(0, 5))
    format_cell_native(c_mo, TEAL_HEADER, "OUTPUT", font_size=11, bold=True, text_color=WHITE_TEXT)

    format_cell_native(tbl_o.cell(1, 0), TEAL_HEADER, "GROUP", font_size=9, bold=True, text_color=WHITE_TEXT)
    for m_i, m_name in enumerate(month_names):
        format_cell_native(tbl_o.cell(1, m_i + 1), TEAL_HEADER, m_name, font_size=9, bold=True, text_color=WHITE_TEXT)
    format_cell_native(tbl_o.cell(1, 4), TEAL_HEADER, "TOTAL", font_size=9, bold=True, text_color=WHITE_TEXT)
    format_cell_native(tbl_o.cell(1, 5), TEAL_HEADER, "AVERAGE", font_size=9, bold=True, text_color=WHITE_TEXT)

    for r_i, grp in enumerate(target_groups, start=2):
        format_cell_native(tbl_o.cell(r_i, 0), RGBColor(255, 255, 255), grp, font_size=8, bold=True,
                           text_color=TEXT_BLACK, align=PP_ALIGN.CENTER)
        for m_i in range(3):
            format_cell_native(tbl_o.cell(r_i, m_i + 1), RGBColor(255, 255, 255), str(group_stats[grp]["m_vols"][m_i]),
                               font_size=9, text_color=TEXT_BLACK)
        format_cell_native(tbl_o.cell(r_i, 4), RGBColor(220, 240, 240), str(group_stats[grp]["tot_vol"]), font_size=9.5,
                           bold=True, text_color=TEXT_BLACK)
        format_cell_native(tbl_o.cell(r_i, 5), RGBColor(200, 230, 230), str(group_stats[grp]["avg_mo_vol"]),
                           font_size=9.5, bold=True, text_color=TEXT_BLACK)

    # RIGHT HALF: TURNAROUND TIME
    c_tat_x, c_tat_y, c_tat_w, c_tat_h = Inches(6.9), Inches(1.15), Inches(6.0), Inches(3.8)
    card_tat = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, c_tat_x, c_tat_y, c_tat_w, c_tat_h)
    card_tat.fill.solid()
    card_tat.fill.fore_color.rgb = RGBColor(255, 255, 255)
    card_tat.line.color.rgb = RGBColor(210, 210, 210)

    data_tat = CategoryChartData()
    data_tat.categories = target_groups
    for m_i, m_name in enumerate(month_names):
        data_tat.add_series(m_name, [group_stats[g]["m_tats"][m_i] for g in target_groups])

    ch_shape_tat = slide2.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, c_tat_x + Inches(0.1),
                                           c_tat_y + Inches(0.08), c_tat_w - Inches(0.2), c_tat_h - Inches(0.45),
                                           data_tat)
    ch_tat = ch_shape_tat.chart
    ch_tat.has_title = True
    ch_tat.chart_title.text_frame.text = f"TURNAROUND TIME OF {q_key}"
    p_tt = ch_tat.chart_title.text_frame.paragraphs[0]
    p_tt.font.name, p_tt.font.bold, p_tt.font.size, p_tt.font.color.rgb = "Aptos", True, Pt(14), RGBColor(60, 60, 60)

    green_palette = [RGBColor(152, 228, 176), RGBColor(0, 176, 80), RGBColor(0, 104, 55)]
    for s_idx, s in enumerate(ch_tat.series):
        s.format.fill.solid()
        s.format.fill.fore_color.rgb = green_palette[s_idx % len(green_palette)]
        s.has_data_labels = True
        s.data_labels.font.name = "Aptos"
        s.data_labels.font.size = Pt(7)

    ch_tat.has_legend = True
    ch_tat.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch_tat.legend.font.name, ch_tat.legend.font.size = "Aptos", Pt(7.5)

    tbl_tat_s = slide2.shapes.add_table(6, 5, Inches(6.9), Inches(5.15), Inches(6.0), Inches(1.9))
    tbl_t = tbl_tat_s.table
    set_native_table_grid_style(tbl_t)
    tbl_t.columns[0].width = Inches(2.2)
    for c_i in range(1, 5): tbl_t.columns[c_i].width = Inches(0.95)

    c_mt = tbl_t.cell(0, 0)
    c_mt.merge(tbl_t.cell(0, 4))
    format_cell_native(c_mt, RGBColor(0, 176, 80), "TURNAROUND TIME (TAT)", font_size=11, bold=True,
                       text_color=WHITE_TEXT)

    format_cell_native(tbl_t.cell(1, 0), RGBColor(0, 176, 80), "GROUP", font_size=9, bold=True, text_color=WHITE_TEXT)
    for m_i, m_name in enumerate(month_names):
        format_cell_native(tbl_t.cell(1, m_i + 1), RGBColor(0, 176, 80), m_name, font_size=9, bold=True,
                           text_color=WHITE_TEXT)
    format_cell_native(tbl_t.cell(1, 4), RGBColor(0, 176, 80), "AVERAGE", font_size=9, bold=True, text_color=WHITE_TEXT)

    for r_i, grp in enumerate(target_groups, start=2):
        format_cell_native(tbl_t.cell(r_i, 0), RGBColor(255, 255, 255), grp, font_size=8, bold=True,
                           text_color=TEXT_BLACK, align=PP_ALIGN.CENTER)
        for m_i in range(3):
            format_cell_native(tbl_t.cell(r_i, m_i + 1), RGBColor(255, 255, 255), f"{group_stats[grp]['m_tats'][m_i]}%",
                               font_size=9, text_color=TEXT_BLACK)
        format_cell_native(tbl_t.cell(r_i, 4), RGBColor(200, 235, 210), f"{group_stats[grp]['quarter_tat']}%",
                           font_size=9.5, bold=True, text_color=TEXT_BLACK)

    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream