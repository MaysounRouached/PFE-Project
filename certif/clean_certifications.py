"""
Certification Data Pipeline
----------------------------
Turns raw partner-certification exports into a full set of insights:
  1. Removes universities (using an editable lookup table), flags unknowns
  2. Certificate count per partner per level
  3. Expiring certificates this year, urgency-sorted, with renewal suggestions
  4. Domain standardization
  5. Partner level (Basic/Advanced/Master) via an editable rules table
  6. Target vs. gap analysis (what's missing to reach a chosen level)
  7. New certificate detection (compares against last run's snapshot)

HOW TO USE:
    1. Put your raw Excel export in this folder (update RAW_FILE_PATH below).
    2. Fill in university_lookup.csv with real partner names/codes -> Yes/No.
    3. Fill in partner_targets.csv with each partner's target level PER
       DOMAIN (Basic/Advanced/Master) — one row per partner+domain they
       care about. Domain/partner combos left out are marked "Not set"
       in the Gap Analysis sheet instead of being assumed.
    4. Adjust partner_level_rules.csv if the level rule changes.
    5. Run: uv run clean_certifications.py
    6. Open certifications_report.xlsx — it has one sheet per insight.

Every run also saves a snapshot of today's certificate list, so the NEXT
run can detect which certificates are brand new since last time.
"""

import pandas as pd
from pathlib import Path
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties


HUAWEI_RED = "C8102E"
DARK_RED = "7A1414"
BAND_FILL_HEX = "FBE4E4"
GREEN_ACCENT = "C6E0B4"

HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color=HUAWEI_RED, end_color=HUAWEI_RED, fill_type="solid")
BODY_FONT = Font(name="Calibri", size=10)

Q1_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")  # light blue
Q2_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")  # light green
Q3_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")  # light yellow
Q4_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")  # light orange
QUARTER_FILLS = {"Q1": Q1_FILL, "Q2": Q2_FILL, "Q3": Q3_FILL, "Q4": Q4_FILL}


def style_workbook(path):
    """
    Applies the same sheet conventions used in the eSales report to every
    sheet here:
    - Row 1: an italic grey "Report Generated: <timestamp>" banner
    - Row 2: bold white header on the Huawei red brand fill
    - Frozen panes below the header, autofilter starting at the header
    - Auto-sized columns (wrapped, 2-line headers to keep the long
      bilingual column names readable) and a 75% zoom so more of the
      table fits on screen without scrolling
    - Urgency color-coding on the Expiring Certificates sheet
    """
    from openpyxl import load_workbook
    wb = load_workbook(path)

    # These two already carry their own "Report Generated" banner (the KPI
    # banner on Dashboard, the badge row on the domain matrix), so they're
    # styled separately (style_domain_matrix_sheet / add_dashboard) and
    # skipped here.
    SKIP_SHEETS = {"Partner Levels by Domain", "Dashboard"}

    for sheet_name in wb.sheetnames:
        if sheet_name in SKIP_SHEETS:
            continue

        ws = wb[sheet_name]
        if ws.max_row < 1:
            continue

        header_row = 2
        last_col = max(ws.max_column, 1)

        # Row 1: "Report Generated" timestamp, same as every eSales sheet
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
        gen_cell = ws.cell(row=1, column=1,
                            value=f"Report Generated: {TODAY.strftime('%Y-%m-%d %H:%M:%S')}")
        gen_cell.font = Font(italic=True, size=9, color="666666")
        ws.row_dimensions[1].height = 16

        # Header row styling (wrapped so long bilingual headers stay legible
        # at the tighter column widths below)
        for cell in ws[header_row]:
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[header_row].height = 30

        # Body font + freeze panes below the header + autofilter from the
        # header down (not including the Report Generated row)
        for row in ws.iter_rows(min_row=header_row + 1):
            for cell in row:
                cell.font = BODY_FONT
        ws.freeze_panes = f"A{header_row + 1}"
        last_col_letter = get_column_letter(last_col)
        ws.auto_filter.ref = f"A{header_row}:{last_col_letter}{ws.max_row}"

        # Auto-size columns based on content length, screen-fit like eSales
        for col_cells in ws.columns:
            length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
            col_letter = get_column_letter(col_cells[0].column)
            ws.column_dimensions[col_letter].width = min(max(length + 2, 8), 26)
        ws.sheet_view.zoomScale = 75

        # Bold + shade the TOTAL row, if present
        for row in ws.iter_rows(min_row=header_row + 1):
            if row[0].value == "TOTAL":
                for cell in row:
                    cell.font = Font(name="Calibri", size=10, bold=True)
                    cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

        # Quarter color-coding specific to the expiring certificates sheet
        if sheet_name == "Expiring Certificates":
            headers = [c.value for c in ws[header_row]]
            if "Quarter" in headers:
                q_col_idx = headers.index("Quarter") + 1
                for row in ws.iter_rows(min_row=header_row + 1):
                    q_cell = row[q_col_idx - 1]
                    fill = QUARTER_FILLS.get(q_cell.value)
                    if fill:
                        for cell in row:
                            cell.fill = fill

    wb.save(path)


def style_domain_matrix_sheet(path, badge_counts):
    """
    Fully styles the 'Partner Levels by Domain' sheet:
    - Badge row (Master/Advanced/Basic counts) at the top
    - Colored header + tinted columns per level (gold/silver/bronze)
    - Bright green highlight + bold on cells where a level was just
      reached this run (marked "X (NEW)" by build_partner_domain_matrix)
    - Grid borders, centered marks, frozen header + first 2 columns
    """
    from openpyxl import load_workbook
    wb = load_workbook(path)
    ws = wb["Partner Levels by Domain"]

    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ---- Badge row ----
    badge_font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    badge_fill = PatternFill(start_color=HUAWEI_RED, end_color=HUAWEI_RED, fill_type="solid")
    labels = [
        ("MASTER", badge_counts.get("Master", 0)),
        ("ADVANCED", badge_counts.get("Advanced", 0)),
        ("BASIC", badge_counts.get("Basic", 0)),
    ]
    col = 2
    for label, count in labels:
        cell = ws.cell(row=1, column=col, value=f"{label} : {count}")
        cell.font = badge_font
        cell.fill = badge_fill
        cell.alignment = Alignment(horizontal="center")
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 1)
        col += 3

    note = ws.cell(row=2, column=2,
                    value=f"Green highlight = level reached for the first time this run   ·   "
                          f"Report Generated: {TODAY.strftime('%Y-%m-%d %H:%M:%S')}")
    note.font = Font(italic=True, size=9, color="606060")

    # ---- Header row (data written with startrow=3, so header lands on row 4) ----
    header_row_idx = 4
    headers = [c.value for c in ws[header_row_idx]]

    LEVEL_HEADER_FILLS = {
        "Master": PatternFill(start_color="BF9000", end_color="BF9000", fill_type="solid"),
        "Advanced": PatternFill(start_color="595959", end_color="595959", fill_type="solid"),
        "Basic": PatternFill(start_color="833C00", end_color="833C00", fill_type="solid"),
    }
    LEVEL_COL_FILLS = {
        "Master": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
        "Advanced": PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"),
        "Basic": PatternFill(start_color="FBE5D6", end_color="FBE5D6", fill_type="solid"),
    }
    NEW_FILL = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    HIGHEST_LEVEL_FONT_COLORS = {"Master": "BF9000", "Advanced": "595959", "Basic": "833C00"}

    for col_idx, header in enumerate(headers, start=1):
        header_cell = ws.cell(row=header_row_idx, column=col_idx)
        header_cell.font = HEADER_FONT
        header_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        header_cell.border = border

        matched_level = next((lvl for lvl in ["Master", "Advanced", "Basic"]
                               if header and str(header).startswith(lvl + " -")), None)
        header_cell.fill = LEVEL_HEADER_FILLS.get(matched_level, HEADER_FILL)

        for data_row in ws.iter_rows(min_row=header_row_idx + 1, min_col=col_idx, max_col=col_idx):
            for cell in data_row:
                cell.border = border
                cell.alignment = Alignment(horizontal="center")
                if matched_level:
                    if cell.value and "NEW" in str(cell.value):
                        cell.value = "X"
                        cell.fill = NEW_FILL
                        cell.font = Font(bold=True, color="006100")
                    else:
                        cell.fill = LEVEL_COL_FILLS[matched_level]
                        cell.font = BODY_FONT
                elif header == "Highest Level":
                    cell.font = Font(bold=True, color=HIGHEST_LEVEL_FONT_COLORS.get(cell.value, "000000"))
                else:
                    cell.font = BODY_FONT

    # Freeze header row + first two columns (Partner, Code)
    ws.freeze_panes = ws.cell(row=header_row_idx + 1, column=3).coordinate
    ws.sheet_view.zoomScale = 80

    # Auto-size columns
    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        col_letter = get_column_letter(col_cells[0].column)
        ws.column_dimensions[col_letter].width = min(max(length + 2, 8), 30)

    wb.save(path)


def add_dashboard(path, partner_levels_df, expiring_df, clean_df, needs_review_df,
                   new_certs_df, leveled_up_details):
    """
    Adds a 'Dashboard' sheet at the front of the workbook:
    - Banner title
    - 5 KPI cards (partners, expiring, new certs, needs review, leveled up)
    - A "Recently Leveled Up" highlight table
    - 4 charts in clearly separated, labeled sections
    Helper data tables that feed the charts sit in columns N+, out of the
    way of the main view.
    """
    from openpyxl import load_workbook
    wb = load_workbook(path)
    if "Dashboard" in wb.sheetnames:
        del wb["Dashboard"]
    ws = wb.create_sheet("Dashboard", 0)

    NAVY = HUAWEI_RED  # kept as an alias so the rest of the function reads the same
    GOLD = "1E7B34"  # accent for "positive/new" callouts — same idea as eSales'
                      # green highlight for new opportunities and increases
    # The 5 KPI cards occupy columns 2-20 (col, +4 each). Helper tables that
    # feed the charts start well clear of that, so nothing collides and the
    # helper columns can be hidden as a block below.
    HELPER_COL_START = 23
    TITLE_FONT = Font(name="Calibri", size=20, bold=True, color="FFFFFF")
    SUBTITLE_FONT = Font(name="Calibri", size=10, italic=True, color="FBE4E4")
    SECTION_FONT = Font(name="Calibri", size=13, bold=True, color=NAVY)
    KPI_LABEL_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    KPI_VALUE_FONT = Font(name="Calibri", size=24, bold=True, color=NAVY)
    KPI_FILL = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    KPI_FILL_ACCENT = PatternFill(start_color=GREEN_ACCENT, end_color=GREEN_ACCENT, fill_type="solid")
    TABLE_HEADER_FONT = Font(bold=True)
    BANNER_FILL = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    THIN = Side(style="thin", color="D9D9D9")
    BOX_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    # ---- Title banner ----
    for row in range(1, 4):
        for col in range(1, 22):
            ws.cell(row=row, column=col).fill = BANNER_FILL
    ws["B2"] = "🏅 Certification Program Dashboard"
    ws["B2"].font = TITLE_FONT
    ws["B3"] = f"Generated {TODAY.strftime('%Y-%m-%d')}  ·  Auto-refreshed each pipeline run"
    ws["B3"].font = SUBTITLE_FONT

    # ---- KPI cards (5) ----
    total_partners = partner_levels_df[PARTNER_COL].nunique()
    total_expiring = len(expiring_df)
    total_new = len(new_certs_df)
    total_review = len(needs_review_df)
    total_leveled_up = len(leveled_up_details)

    kpis = [
        ("CERTIFIED PARTNERS", total_partners, KPI_FILL),
        ("EXPIRING THIS YEAR", total_expiring, KPI_FILL),
        ("NEW CERTIFICATES", total_new, KPI_FILL),
        ("NEEDS REVIEW", total_review, KPI_FILL),
        ("LEVELED UP THIS RUN", total_leveled_up, KPI_FILL_ACCENT),
    ]
    col = 2
    kpi_row_label, kpi_row_value = 5, 6
    for label, value, fill in kpis:
        label_cell = ws.cell(row=kpi_row_label, column=col, value=label)
        label_cell.font = KPI_LABEL_FONT
        label_cell.fill = fill
        label_cell.alignment = Alignment(horizontal="center")
        ws.merge_cells(start_row=kpi_row_label, start_column=col, end_row=kpi_row_label, end_column=col + 2)

        value_cell = ws.cell(row=kpi_row_value, column=col, value=value)
        value_cell.font = KPI_VALUE_FONT if fill is KPI_FILL else Font(name="Calibri", size=24, bold=True, color=GOLD)
        value_cell.alignment = Alignment(horizontal="center")
        value_cell.border = BOX_BORDER
        ws.merge_cells(start_row=kpi_row_value, start_column=col, end_row=kpi_row_value + 1, end_column=col + 2)
        col += 4

    # ---- Section: Recently Leveled Up ----
    section_row = 9
    ws.cell(row=section_row, column=2, value="🆕 Recently Leveled Up").font = SECTION_FONT

    table_row = section_row + 1
    if leveled_up_details.empty:
        ws.cell(row=table_row, column=2, value="No partners reached a new level this run.").font = Font(italic=True, color="808080")
    else:
        headers = ["Partner", "Domain", "New Level"]
        for i, h in enumerate(headers):
            c = ws.cell(row=table_row, column=2 + i, value=h)
            c.font = TABLE_HEADER_FONT
            c.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        display_rows = leveled_up_details.head(15)
        for i, (_, r) in enumerate(display_rows.iterrows(), start=1):
            ws.cell(row=table_row + i, column=2, value=r[PARTNER_COL])
            ws.cell(row=table_row + i, column=3, value=r["Domain"])
            level_cell = ws.cell(row=table_row + i, column=4, value=r["NewLevel"])
            level_cell.font = Font(bold=True, color="006100")
        if len(leveled_up_details) > 15:
            ws.cell(row=table_row + 16, column=2,
                    value=f"... and {len(leveled_up_details) - 15} more (see Partner Levels by Domain sheet)").font = Font(italic=True, size=9)

    # ---- Section: Certification Overview ----
    sec2_row = 28
    ws.cell(row=sec2_row, column=2, value="📊 Certification Overview").font = SECTION_FONT
    chart_row = sec2_row + 2

    level_order = ["Master", "Advanced", "Basic", "None"]
    level_counts = (
        partner_levels_df["PartnerLevel"].value_counts()
        .reindex(level_order).fillna(0).astype(int)
    )
    t1_row, t1_col = sec2_row, HELPER_COL_START  # off to the side, out of view
    ws.cell(row=t1_row, column=t1_col, value="Partner Level").font = TABLE_HEADER_FONT
    ws.cell(row=t1_row, column=t1_col + 1, value="Count").font = TABLE_HEADER_FONT
    for i, (level, count) in enumerate(level_counts.items(), start=1):
        ws.cell(row=t1_row + i, column=t1_col, value=level)
        ws.cell(row=t1_row + i, column=t1_col + 1, value=int(count))

    if level_counts.sum() > 0:
        pie = PieChart()
        pie.title = "Partner Level Distribution"
        data = Reference(ws, min_col=t1_col + 1, min_row=t1_row, max_row=t1_row + len(level_counts))
        cats = Reference(ws, min_col=t1_col, min_row=t1_row + 1, max_row=t1_row + len(level_counts))
        pie.add_data(data, titles_from_data=True)
        pie.set_categories(cats)
        pie.visible_cells_only = False  # keep plotting even though the source columns get hidden below
        # Same brand palette as the rest of the report: red for the top
        # tier, fading to grey for lower tiers/no level -- matches the
        # red+grey chart styling already used on the other three dashboards.
        pie_colors = [HUAWEI_RED, DARK_RED, "666666", "CCCCCC"]  # Master, Advanced, Basic, None
        pie.series[0].data_points = [
            DataPoint(idx=i, spPr=GraphicalProperties(solidFill=color))
            for i, color in enumerate(pie_colors[:len(level_counts)])
        ]
        pie.height, pie.width = 8, 12
        ws.add_chart(pie, f"B{chart_row}")

    if "Total Certificates" in partner_levels_df.columns:
        top_partners = partner_levels_df.nlargest(8, "Total Certificates")[[PARTNER_COL, "Total Certificates"]]
        t1b_row, t1b_col = sec2_row, HELPER_COL_START + 4
        ws.cell(row=t1b_row, column=t1b_col, value="Partner").font = TABLE_HEADER_FONT
        ws.cell(row=t1b_row, column=t1b_col + 1, value="Total Certs").font = TABLE_HEADER_FONT
        for i, (_, r) in enumerate(top_partners.iterrows(), start=1):
            ws.cell(row=t1b_row + i, column=t1b_col, value=r[PARTNER_COL])
            ws.cell(row=t1b_row + i, column=t1b_col + 1, value=int(r["Total Certificates"]))

        if len(top_partners) > 0:
            bar_top = BarChart()
            bar_top.title = "Top Partners by Total Certificates"
            bar_top.y_axis.title = "Certificates"
            data = Reference(ws, min_col=t1b_col + 1, min_row=t1b_row, max_row=t1b_row + len(top_partners))
            cats = Reference(ws, min_col=t1b_col, min_row=t1b_row + 1, max_row=t1b_row + len(top_partners))
            bar_top.add_data(data, titles_from_data=True)
            bar_top.set_categories(cats)
            bar_top.visible_cells_only = False  # keep plotting even though the source columns get hidden below
            bar_top.series[0].graphicalProperties.solidFill = HUAWEI_RED
            bar_top.height, bar_top.width = 8, 14
            ws.add_chart(bar_top, f"H{chart_row}")

    # ---- Section: Expiration Outlook ----
    sec3_row = 50
    ws.cell(row=sec3_row, column=2, value="⏰ Expiration Outlook").font = SECTION_FONT
    chart_row3 = sec3_row + 2

    t2_row, t2_col = sec3_row, HELPER_COL_START
    if not expiring_df.empty and EXPIRY_DATE_COL in expiring_df.columns:
        by_month = (
            expiring_df[EXPIRY_DATE_COL].dt.strftime("%Y-%m")
            .value_counts().sort_index()
        )
        ws.cell(row=t2_row, column=t2_col, value="Month").font = TABLE_HEADER_FONT
        ws.cell(row=t2_row, column=t2_col + 1, value="Expiring").font = TABLE_HEADER_FONT
        for i, (month, count) in enumerate(by_month.items(), start=1):
            ws.cell(row=t2_row + i, column=t2_col, value=month)
            ws.cell(row=t2_row + i, column=t2_col + 1, value=int(count))

        if len(by_month) > 0:
            bar = BarChart()
            bar.title = "Certificates Expiring by Month"
            bar.y_axis.title = "Count"
            data = Reference(ws, min_col=t2_col + 1, min_row=t2_row, max_row=t2_row + len(by_month))
            cats = Reference(ws, min_col=t2_col, min_row=t2_row + 1, max_row=t2_row + len(by_month))
            bar.add_data(data, titles_from_data=True)
            bar.set_categories(cats)
            bar.visible_cells_only = False  # keep plotting even though the source columns get hidden below
            bar.series[0].graphicalProperties.solidFill = HUAWEI_RED
            bar.height, bar.width = 8, 14
            ws.add_chart(bar, f"B{chart_row3}")

    # ---- Section: Domain Coverage ----
    sec4_row = 72
    ws.cell(row=sec4_row, column=2, value="🌐 Domain Coverage").font = SECTION_FONT
    chart_row4 = sec4_row + 2

    t3_row, t3_col = sec4_row, HELPER_COL_START
    if DOMAIN_COL in clean_df.columns:
        domain_counts = clean_df[DOMAIN_COL].value_counts().head(8)
        ws.cell(row=t3_row, column=t3_col, value="Domain").font = TABLE_HEADER_FONT
        ws.cell(row=t3_row, column=t3_col + 1, value="Certificates").font = TABLE_HEADER_FONT
        for i, (domain, count) in enumerate(domain_counts.items(), start=1):
            ws.cell(row=t3_row + i, column=t3_col, value=str(domain))
            ws.cell(row=t3_row + i, column=t3_col + 1, value=int(count))

        if len(domain_counts) > 0:
            bar2 = BarChart()
            bar2.title = "Certificates by Specialization Domain"
            bar2.y_axis.title = "Count"
            data = Reference(ws, min_col=t3_col + 1, min_row=t3_row, max_row=t3_row + len(domain_counts))
            cats = Reference(ws, min_col=t3_col, min_row=t3_row + 1, max_row=t3_row + len(domain_counts))
            bar2.add_data(data, titles_from_data=True)
            bar2.set_categories(cats)
            bar2.visible_cells_only = False  # keep plotting even though the source columns get hidden below
            bar2.series[0].graphicalProperties.solidFill = HUAWEI_RED
            bar2.height, bar2.width = 8, 14
            ws.add_chart(bar2, f"B{chart_row4}")

    # ---- Polish pass: bring this sheet in line with the other sheets ----

    # Auto-size the visible columns from actual content, the same way
    # style_workbook does it for every other sheet, instead of a flat guess.
    for col_idx in range(1, 12):  # A..K, the visible dashboard area
        col_letter = get_column_letter(col_idx)
        length = max(
            (len(str(c.value)) for c in ws[col_letter] if c.value is not None),
            default=0,
        )
        ws.column_dimensions[col_letter].width = min(max(length + 3, 12), 30)

    # The tables in columns W+ only exist to feed the charts — they're
    # scratch data, not something a reader needs to see. Hide them so the
    # sheet reads as clean and single-purpose as the rest of the report.
    # (Each chart above has visible_cells_only = False set, so hiding these columns
    # no longer empties the charts the way it did before.)
    for col_idx in range(HELPER_COL_START, HELPER_COL_START + 8):
        ws.column_dimensions[get_column_letter(col_idx)].hidden = True

    # Box the KPI cards and the "Recently Leveled Up" table with the same
    # thin grid border used on the "Partner Levels by Domain" sheet, so
    # they read as cards/tables instead of loose floating text.
    thin = Side(style="thin", color="B7B7B7")
    box_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    kpi_col = 2
    for _ in kpis:
        for r in range(kpi_row_label, kpi_row_value + 2):
            for c in range(kpi_col, kpi_col + 3):
                ws.cell(row=r, column=c).border = box_border
        kpi_col += 4

    if not leveled_up_details.empty:
        last_data_row = table_row + min(len(leveled_up_details), 15)
        for r in range(table_row, last_data_row + 1):
            for c in range(2, 5):
                ws.cell(row=r, column=c).border = box_border

    # Freeze the banner in place while scrolling, the same way every other
    # sheet freezes its header row.
    ws.freeze_panes = "A4"
    ws.sheet_view.zoomScale = 90

    wb.save(path)


def read_csv_safe(path, **kwargs):
    """
    Reads a CSV trying UTF-8 first, falling back to Windows-1252 (cp1252).
    Excel on Windows often saves CSVs as cp1252 by default, which breaks
    plain pd.read_csv() whenever the file has accented characters
    (e.g. École, Université).
    """
    try:
        return pd.read_csv(path, encoding="utf-8", **kwargs)
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="cp1252", **kwargs)


def get_requirement_count(row, cert_type_key):
    """
    Looks up a partner's certificate count for a rule column, supporting
    OR-alternatives written as "TypeA/TypeB" (e.g. "HCSA-Service/HCIA"
    means the requirement is met by holding EITHER certificate, not both).
    Returns the highest count among the alternatives.
    """
    alternatives = [c.strip() for c in str(cert_type_key).split("/")]
    counts = []
    for alt in alternatives:
        value = row.get(alt, 0)
        counts.append(0 if pd.isna(value) else value)
    return max(counts) if counts else 0

# ============ CONFIG — adjust to match your real column names ============

RAW_FILE_PATH = "raw_certifications.xlsx"
UNIVERSITY_LOOKUP_PATH = "university_lookup.csv"
LEVEL_RULES_PATH = "partner_level_rules.csv"
PARTNER_TARGETS_PATH = "partner_targets.csv"
SNAPSHOT_PATH = "certifications_snapshot_previous.csv"
PARTNER_LEVEL_SNAPSHOT_PATH = "partner_domain_levels_snapshot_previous.csv"
OUTPUT_PATH = "certifications_report.xlsx"

PARTNER_COL = "伙伴公司名称/Partner Companies"
PARTNER_CODE_COL = "伙伴编码/Partner Code"
PARTNER_TYPE_COL = "伙伴身份级别/Partner Type"
USER_ID_COL = "官网账号/UserID"
CERT_SERIAL_COL = "证书编号/Certificate Serial"
CERT_NAME_COL = "证书名称/Certificate Name"
CERT_LEVEL_COL = "证书级别/Certificate Level"          # e.g. HCSA-Sales, HCSP-Presales, ...
TECH_DIRECTION_COL = "技术方向/Technical Direction"
DOMAIN_COL = "归属经销商专业化方向/Belong To Reseller Specialization Domain"
GET_DATE_COL = "获得日期/Certificate Get Date"
EXPIRY_DATE_COL = "过期日期/Expiration Date"
STATUS_COL = "状态/Status"

KNOWN_DOMAINS = [
    "Data Communication", "Enterprise Services and Software",
    "Intelligent Collaboration", "Optical", "Storage"
]

# Domains that only count toward a partner's level when the raw field lists
# them ALONE (no comma). Previously Intelligent Collaboration, Enterprise
# Services and Software, and Optical were restricted this way; per request,
# ALL domains now count toward every domain they're listed under, whether
# alone or combined with others in a comma-separated field (e.g. a row
# with "Intelligent Collaboration,Storage" now counts toward BOTH domains).
# Left as an empty, editable list in case any domain needs this restriction
# again in the future.
RESTRICTED_ALONE_ONLY_DOMAINS = []

# Internal Huawei abbreviations used in partner capability matrices.
# CONFIRM "Intelligent Collaboration" -> "IC" is correct; not yet verified.
DOMAIN_ABBREVIATIONS = {
    "Data Communication": "IP",
    "Storage": "IT",
    "Enterprise Services and Software": "ESS",
    "Optical": "Optical",
    "Intelligent Collaboration": "IC",
}

# Assumption: level rows are INDEPENDENT (a partner qualifies for a level if
# they meet that row's own requirements, not the lower rows' requirements
# cumulatively). Change to True if levels should stack cumulatively.
LEVELS_ARE_CUMULATIVE = False

LEVEL_RANK = {"None": 0, "Basic": 1, "Advanced": 2, "Master": 3}

TODAY = datetime.today()
CURRENT_YEAR = TODAY.year


# ============ STEP 1: University filtering ============

def filter_universities(df):
    lookup = read_csv_safe(UNIVERSITY_LOOKUP_PATH)
    # The lookup CSV uses simple column names for easy manual editing;
    # map them to the real raw-file column names here.
    lookup = lookup.rename(columns={
        "PartnerCompany": PARTNER_COL,
        "PartnerCode": PARTNER_CODE_COL,
    })
    lookup[PARTNER_CODE_COL] = lookup[PARTNER_CODE_COL].astype(str).str.strip()
    df[PARTNER_CODE_COL] = df[PARTNER_CODE_COL].astype(str).str.strip()

    merged = df.merge(lookup[[PARTNER_CODE_COL, "IsUniversity"]],
                       on=PARTNER_CODE_COL, how="left")

    known_partners = merged[merged["IsUniversity"].notna()]
    unknown_partners = merged[merged["IsUniversity"].isna()]

    clean_df = known_partners[known_partners["IsUniversity"].str.strip().str.lower() == "no"]
    needs_review_df = unknown_partners[[PARTNER_COL, PARTNER_CODE_COL]].drop_duplicates()

    return clean_df.drop(columns=["IsUniversity"]), needs_review_df


# ============ STEP 2: Certificate count per partner per level ============

def certs_per_partner_level(df):
    pivoted = (
        df.groupby([PARTNER_COL, PARTNER_CODE_COL, CERT_LEVEL_COL])
        .size()
        .reset_index(name="Count")
        .pivot(index=[PARTNER_COL, PARTNER_CODE_COL], columns=CERT_LEVEL_COL, values="Count")
        .fillna(0)
        .astype(int)
        .reset_index()
    )
    level_cols = [c for c in pivoted.columns if c not in [PARTNER_COL, PARTNER_CODE_COL]]
    pivoted["Total Certificates"] = pivoted[level_cols].sum(axis=1)
    return pivoted


# ============ STEP 3: Expiring certificates + renewal suggestion ============

def expiring_this_year(df):
    df = df.copy()
    df[EXPIRY_DATE_COL] = pd.to_datetime(df[EXPIRY_DATE_COL], format="%Y%m%d", errors="coerce")
    expiring = df[df[EXPIRY_DATE_COL].dt.year == CURRENT_YEAR].copy()
    expiring["DaysUntilExpiry"] = (expiring[EXPIRY_DATE_COL] - TODAY).dt.days
    expiring["Quarter"] = "Q" + (((expiring[EXPIRY_DATE_COL].dt.month - 1) // 3) + 1).astype(str)
    expiring = expiring.sort_values("DaysUntilExpiry")

    suggestions = []
    for _, row in expiring.iterrows():
        candidate = suggest_best_candidate(df, row[PARTNER_COL], row[TECH_DIRECTION_COL],
                                            exclude_user=row.get(USER_ID_COL))
        suggestions.append(candidate)
    expiring["SuggestedRenewalCandidate"] = suggestions

    cols = [PARTNER_COL, USER_ID_COL, CERT_NAME_COL, CERT_LEVEL_COL,
            TECH_DIRECTION_COL, EXPIRY_DATE_COL, "Quarter", "DaysUntilExpiry",
            "SuggestedRenewalCandidate"]
    cols = [c for c in cols if c in expiring.columns]
    return expiring[cols]


def suggest_best_candidate(df, partner, tech_direction, exclude_user=None):
    """
    Heuristic: within the same partner, find the person with the most
    certificates in the same Technical Direction (excluding the person
    whose cert is expiring). More certs in that direction = more likely
    to pass a related/renewal exam. This rule is easy to swap out later
    if Huawei has a better signal (e.g. training completion data).
    """
    same_partner = df[df[PARTNER_COL] == partner]
    same_direction = same_partner[same_partner[TECH_DIRECTION_COL] == tech_direction]
    if exclude_user is not None:
        same_direction = same_direction[same_direction[USER_ID_COL] != exclude_user]

    if same_direction.empty:
        return "No other candidate found in this direction"

    counts = same_direction[USER_ID_COL].value_counts()
    return counts.idxmax()


# ============ STEP 4: Domain standardization ============

def standardize_domains(df):
    df = df.copy()
    df[DOMAIN_COL] = df[DOMAIN_COL].astype(str).str.strip()

    def all_parts_known(value):
        parts = [p.strip() for p in str(value).split(",")]
        return all(p in KNOWN_DOMAINS for p in parts)

    df["DomainRecognized"] = df[DOMAIN_COL].apply(all_parts_known)
    return df


def expand_partner_domain_rows(df):
    """
    Attributes each certificate row to the domain(s) it should count toward:
      - Every known domain listed in the raw field counts toward that
        domain's tally — whether it's the only domain listed or one of
        several comma-separated domains on that row. A row like
        "Intelligent Collaboration,Storage" now counts toward BOTH
        Intelligent Collaboration and Storage.
      - RESTRICTED_ALONE_ONLY_DOMAINS is kept as a hook (currently empty)
        in case a domain ever needs to go back to "only counts when
        listed alone."
    """
    records = []
    for _, row in df.iterrows():
        value = str(row[DOMAIN_COL]).strip()
        if value == "" or value.lower() == "nan":
            continue

        parts = [p.strip() for p in value.split(",")]
        is_multi = len(parts) > 1

        for part in parts:
            if part not in KNOWN_DOMAINS:
                continue
            if is_multi and part in RESTRICTED_ALONE_ONLY_DOMAINS:
                continue  # restricted domain only counts when alone
            record = row.to_dict()
            record["MatchedDomain"] = part
            records.append(record)

    return pd.DataFrame(records)


# ============ STEP 5: Partner level calculation ============

def calculate_partner_levels(cert_counts_df):
    rules = read_csv_safe(LEVEL_RULES_PATH).set_index("Level")
    level_order = ["Master", "Advanced", "Basic"]  # check highest first

    results = []
    for _, row in cert_counts_df.iterrows():
        assigned_level = "None"
        for level in level_order:
            requirements = rules.loc[level]
            meets_all = all(
                get_requirement_count(row, cert_type) >= min_required
                for cert_type, min_required in requirements.items()
                if min_required > 0
            )
            if meets_all:
                assigned_level = level
                break
        results.append(assigned_level)

    cert_counts_df = cert_counts_df.copy()
    cert_counts_df["PartnerLevel"] = results
    return cert_counts_df


# ============ STEP 5b: Certificate count + level, PER DOMAIN ============

def certs_per_partner_domain(df):
    expanded = expand_partner_domain_rows(df)
    if expanded.empty:
        return pd.DataFrame(columns=[PARTNER_COL, PARTNER_CODE_COL, "MatchedDomain", "Total Certificates"])

    pivoted = (
        expanded.groupby([PARTNER_COL, PARTNER_CODE_COL, "MatchedDomain", CERT_LEVEL_COL])
        .size()
        .reset_index(name="Count")
        .pivot(index=[PARTNER_COL, PARTNER_CODE_COL, "MatchedDomain"], columns=CERT_LEVEL_COL, values="Count")
        .fillna(0)
        .astype(int)
        .reset_index()
    )
    level_cols = [c for c in pivoted.columns if c not in [PARTNER_COL, PARTNER_CODE_COL, "MatchedDomain"]]
    pivoted["Total Certificates"] = pivoted[level_cols].sum(axis=1)
    return pivoted


def calculate_partner_domain_levels(cert_counts_domain_df):
    if cert_counts_domain_df.empty:
        cert_counts_domain_df = cert_counts_domain_df.copy()
        cert_counts_domain_df["PartnerLevel"] = []
        return cert_counts_domain_df

    rules = read_csv_safe(LEVEL_RULES_PATH).set_index("Level")
    level_order = ["Master", "Advanced", "Basic"]

    results = []
    for _, row in cert_counts_domain_df.iterrows():
        assigned_level = "None"
        for level in level_order:
            requirements = rules.loc[level]
            meets_all = all(
                get_requirement_count(row, cert_type) >= min_required
                for cert_type, min_required in requirements.items()
                if min_required > 0
            )
            if meets_all:
                assigned_level = level
                break
        results.append(assigned_level)

    cert_counts_domain_df = cert_counts_domain_df.copy()
    cert_counts_domain_df["PartnerLevel"] = results
    return cert_counts_domain_df


def get_leveled_up_keys(domain_levels_df):
    """
    Compares today's per-partner-per-domain levels against the previous
    run's snapshot and returns the set of (PartnerCode, Domain) pairs
    whose level increased since then (e.g. Advanced -> Master). Also
    saves today's levels as the new snapshot for the next run.
    On the very first run (no snapshot yet), nothing is flagged as
    "leveled up" — there's nothing to compare against.
    """
    snapshot_file = Path(PARTNER_LEVEL_SNAPSHOT_PATH)
    leveled_up = set()

    if snapshot_file.exists():
        previous_df = read_csv_safe(snapshot_file)
        prev_map = {
            (str(r[PARTNER_CODE_COL]).strip(), str(r["MatchedDomain"]).strip()): r["PartnerLevel"]
            for _, r in previous_df.iterrows()
        }
        for _, row in domain_levels_df.iterrows():
            key = (str(row[PARTNER_CODE_COL]).strip(), str(row["MatchedDomain"]).strip())
            prev_level = prev_map.get(key, "None")
            if LEVEL_RANK.get(row["PartnerLevel"], 0) > LEVEL_RANK.get(prev_level, 0):
                leveled_up.add(key)

    snapshot_cols = [PARTNER_COL, PARTNER_CODE_COL, "MatchedDomain", "PartnerLevel"]
    domain_levels_df[snapshot_cols].to_csv(snapshot_file, index=False)

    return leveled_up


def build_domain_level_matrix(domain_levels_df, clean_df):
    """
    Reshapes the long partner-domain-level table into a wide capability
    matrix: one row per partner, columns grouped by Level (Basic/Advanced/
    Master) with a sub-column per domain abbreviation, "X" marking which
    level a partner achieved in that domain. Matches Huawei's standard
    partner capability matrix layout.
    """
    achieved = domain_levels_df[domain_levels_df["PartnerLevel"] != "None"].copy()
    if achieved.empty:
        return pd.DataFrame()

    achieved["DomainAbbrev"] = achieved["MatchedDomain"].map(DOMAIN_ABBREVIATIONS).fillna(achieved["MatchedDomain"])

    if PARTNER_TYPE_COL in clean_df.columns:
        status_lookup = clean_df.groupby(PARTNER_CODE_COL)[PARTNER_TYPE_COL].first()
        achieved["Status"] = achieved[PARTNER_CODE_COL].map(status_lookup)
    else:
        achieved["Status"] = ""

    achieved["Mark"] = "X"

    matrix = achieved.pivot_table(
        index=[PARTNER_COL, PARTNER_CODE_COL, "Status"],
        columns=["PartnerLevel", "DomainAbbrev"],
        values="Mark",
        aggfunc="first",
        fill_value="",
    )

    level_order = ["Basic", "Advanced", "Master"]
    domain_order = ["IT", "IP", "Optical", "ESS", "IC"]
    ordered_cols = [
        (lvl, dom) for lvl in level_order for dom in domain_order
        if (lvl, dom) in matrix.columns
    ]
    matrix = matrix[ordered_cols]

    return matrix.reset_index()


# ============ STEP 6: Target vs gap analysis ============

# Abbreviations used in the partner x level x domain matrix view.
# "Intelligent Collaboration" wasn't in your example screenshot, so "IC"
# is my best-guess abbreviation for it — rename freely, it's just a label.
DOMAIN_ABBREVIATIONS = {
    "Data Communication": "IP",
    "Storage": "IT",
    "Enterprise Services and Software": "ESS",
    "Optical": "Optical",
    "Intelligent Collaboration": "IC",
}
MATRIX_LEVEL_ORDER = ["Basic", "Advanced", "Master"]
MATRIX_DOMAIN_ORDER = ["IT", "IP", "ESS", "Optical", "IC"]


def build_partner_domain_matrix(domain_levels_df, leveled_up_keys=None):
    """
    Reshapes the per-partner-per-domain levels into the wide matrix format:
    rows = partners (sorted best level first), columns grouped by
    Level -> Domain abbreviation, with 'X' where that partner achieved
    that level in that domain. Cells where the level was JUST reached
    this run are marked 'X (NEW)' for highlighting. Also adds a
    "Highest Level" column and returns badge counts.
    """
    leveled_up_keys = leveled_up_keys or set()
    achieved = domain_levels_df[domain_levels_df["PartnerLevel"] != "None"].copy()

    if achieved.empty:
        columns = [PARTNER_COL, PARTNER_CODE_COL] + [
            f"{level} - {domain}" for level in MATRIX_LEVEL_ORDER for domain in MATRIX_DOMAIN_ORDER
        ] + ["Highest Level"]
        return pd.DataFrame(columns=columns), {"Master": 0, "Advanced": 0, "Basic": 0}

    achieved["DomainAbbrev"] = achieved["MatchedDomain"].map(DOMAIN_ABBREVIATIONS).fillna(achieved["MatchedDomain"])
    achieved["ColumnKey"] = achieved["PartnerLevel"] + " - " + achieved["DomainAbbrev"]

    def mark_value(row):
        key = (str(row[PARTNER_CODE_COL]).strip(), str(row["MatchedDomain"]).strip())
        return "X (NEW)" if key in leveled_up_keys else "X"

    achieved["Mark"] = achieved.apply(mark_value, axis=1)

    wide = achieved.pivot_table(
        index=[PARTNER_COL, PARTNER_CODE_COL],
        columns="ColumnKey",
        values="Mark",
        aggfunc="first",
    ).reset_index()

    # Force a consistent, fully-populated column order/grid even if some
    # level+domain combos have zero partners (matches the fixed layout
    # of the reference screenshot).
    ordered_cols = [f"{level} - {domain}" for level in MATRIX_LEVEL_ORDER for domain in MATRIX_DOMAIN_ORDER]
    for col in ordered_cols:
        if col not in wide.columns:
            wide[col] = ""
    wide = wide[[PARTNER_COL, PARTNER_CODE_COL] + ordered_cols].fillna("")

    # Highest level achieved by each partner across all their domains
    highest_by_partner = achieved.groupby(PARTNER_CODE_COL)["PartnerLevel"].apply(
        lambda levels: max(levels, key=lambda lvl: LEVEL_RANK.get(lvl, 0))
    )
    wide["Highest Level"] = wide[PARTNER_CODE_COL].map(highest_by_partner)

    # Best partners first, then alphabetically within the same level
    wide["_sort_rank"] = wide["Highest Level"].map(LEVEL_RANK).fillna(0)
    wide = wide.sort_values(["_sort_rank", PARTNER_COL], ascending=[False, True]).drop(columns=["_sort_rank"])
    wide = wide.reset_index(drop=True)

    badge_counts = {
        level: int((achieved["PartnerLevel"] == level).sum())
        for level in ["Master", "Advanced", "Basic"]
    }
    return wide, badge_counts


def gap_analysis(domain_levels_df):
    """
    Reads each partner's target level PER DOMAIN from partner_targets.csv
    (matched by PartnerCode + Domain) and reports what's missing to reach
    it in that specific domain. A partner can have different targets for
    different domains (e.g. Master in Data Communication, Advanced in
    Storage). Domain/partner combos without a target are flagged "Not set"
    rather than assumed.
    """
    rules = read_csv_safe(LEVEL_RULES_PATH).set_index("Level")

    targets = read_csv_safe(PARTNER_TARGETS_PATH)
    targets = targets.rename(columns={
        "PartnerCompany": "PartnerName_fromTargets",
        "PartnerCode": PARTNER_CODE_COL,
    })
    targets[PARTNER_CODE_COL] = targets[PARTNER_CODE_COL].astype(str).str.strip()
    targets["Domain"] = targets["Domain"].astype(str).str.strip()

    cc = domain_levels_df.copy()
    cc[PARTNER_CODE_COL] = cc[PARTNER_CODE_COL].astype(str).str.strip()
    cc = cc.rename(columns={PARTNER_COL: "PartnerName_fromCerts", "MatchedDomain": "Domain"})

    merged = pd.merge(targets, cc, on=[PARTNER_CODE_COL, "Domain"], how="outer")
    merged[PARTNER_COL] = merged["PartnerName_fromTargets"].fillna(merged["PartnerName_fromCerts"])

    gap_rows = []
    for _, row in merged.iterrows():
        target_level = row.get("TargetLevel")
        partner_name = row.get(PARTNER_COL, "Unknown")
        domain = row.get("Domain", "Unknown")

        if pd.isna(target_level) or str(target_level).strip() == "":
            gap_rows.append({
                PARTNER_COL: partner_name,
                "Domain": domain,
                "TargetLevel": "Not set",
                "MissingCertificates": "No target set for this domain — add to partner_targets.csv",
            })
            continue

        target_level = str(target_level).strip()
        if target_level not in rules.index:
            gap_rows.append({
                PARTNER_COL: partner_name,
                "Domain": domain,
                "TargetLevel": target_level,
                "MissingCertificates": f"Unknown target level '{target_level}' — check partner_targets.csv",
            })
            continue

        requirements = rules.loc[target_level]
        missing = {}
        for cert_type, min_required in requirements.items():
            if min_required <= 0:
                continue
            current = get_requirement_count(row, cert_type)
            if current < min_required:
                missing[cert_type] = int(min_required - current)

        gap_rows.append({
            PARTNER_COL: partner_name,
            "Domain": domain,
            "TargetLevel": target_level,
            "MissingCertificates": ", ".join(
                f"{k} (need {v} more)" for k, v in missing.items()
            ) if missing else "Target already met",
        })
    return pd.DataFrame(gap_rows)


# ============ STEP 7: New certificate detection ============

def detect_new_certificates(current_df):
    snapshot_file = Path(SNAPSHOT_PATH)
    if not snapshot_file.exists():
        new_certs = current_df.copy()
        new_certs["Note"] = "First run — no previous snapshot to compare against"
    else:
        previous_df = read_csv_safe(snapshot_file)
        previous_serials = set(previous_df[CERT_SERIAL_COL].astype(str))
        current_df_copy = current_df.copy()
        current_df_copy[CERT_SERIAL_COL] = current_df_copy[CERT_SERIAL_COL].astype(str)
        new_certs = current_df_copy[~current_df_copy[CERT_SERIAL_COL].isin(previous_serials)]

    # Save today's snapshot for next time
    current_df.to_csv(snapshot_file, index=False)

    cols = [CERT_SERIAL_COL, PARTNER_COL, USER_ID_COL, CERT_NAME_COL,
            CERT_LEVEL_COL, GET_DATE_COL]
    cols = [c for c in cols if c in new_certs.columns]
    return new_certs[cols] if not new_certs.empty else new_certs


# ============ MAIN ============

def main():
    print(f"Loading raw file: {RAW_FILE_PATH}")
    raw_df = pd.read_excel(RAW_FILE_PATH)

    print("Filtering universities...")
    clean_df, needs_review_df = filter_universities(raw_df)
    print(f"  {len(clean_df)} rows kept, {len(needs_review_df)} unknown partners flagged for review")

    print("Standardizing domains...")
    clean_df = standardize_domains(clean_df)

    print("Counting certificates per partner/level...")
    cert_counts_df = certs_per_partner_level(clean_df)

    print("Finding expiring certificates...")
    expiring_df = expiring_this_year(clean_df)

    print("Calculating partner levels...")
    partner_levels_df = calculate_partner_levels(cert_counts_df)

    level_summary_df = (
        partner_levels_df["PartnerLevel"].value_counts()
        .reindex(["Master", "Advanced", "Basic", "None"]).fillna(0).astype(int)
        .reset_index()
    )
    level_summary_df.columns = ["Partner Level (Summary)", "Number of Partners"]

    # Export version: drop the single overall PartnerLevel column (per-domain
    # levels are the real detail now) and add a TOTAL row summing each
    # certificate type across all partners.
    partner_levels_export_df = partner_levels_df.drop(columns=["PartnerLevel"])
    numeric_cols = [c for c in partner_levels_export_df.columns if c not in [PARTNER_COL, PARTNER_CODE_COL]]
    total_row = {PARTNER_COL: "TOTAL", PARTNER_CODE_COL: ""}
    for c in numeric_cols:
        total_row[c] = partner_levels_export_df[c].sum()
    partner_levels_export_df = pd.concat(
        [partner_levels_export_df, pd.DataFrame([total_row])], ignore_index=True
    )

    print("Calculating partner levels per domain...")
    cert_counts_domain_df = certs_per_partner_domain(clean_df)
    domain_levels_df = calculate_partner_domain_levels(cert_counts_domain_df)

    print("Checking for level-ups since last run...")
    leveled_up_keys = get_leveled_up_keys(domain_levels_df)
    leveled_up_details = domain_levels_df[
        domain_levels_df.apply(
            lambda r: (str(r[PARTNER_CODE_COL]).strip(), str(r["MatchedDomain"]).strip()) in leveled_up_keys,
            axis=1,
        )
    ][[PARTNER_COL, "MatchedDomain", "PartnerLevel"]].rename(
        columns={"MatchedDomain": "Domain", "PartnerLevel": "NewLevel"}
    )
    print(f"  {len(leveled_up_details)} partner-domain level-ups detected")

    domain_matrix_df, domain_badge_counts = build_partner_domain_matrix(domain_levels_df, leveled_up_keys)

    print("Running gap analysis using each partner's per-domain target level...")
    gap_df = gap_analysis(domain_levels_df)

    print("Detecting new certificates since last run...")
    new_certs_df = detect_new_certificates(clean_df)

    print(f"Writing report to {OUTPUT_PATH}...")
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        # startrow=1 on the plain sheets leaves row 1 free for the
        # "Report Generated" banner that style_workbook adds afterward
        # (same layout as every eSales sheet). The domain matrix keeps its
        # own startrow=3 — it has its own badge banner instead.
        partner_levels_export_df.to_excel(writer, sheet_name="Partner Levels", index=False, startrow=1)
        level_summary_df.to_excel(writer, sheet_name="Partner Levels", index=False,
                                   startrow=1, startcol=len(partner_levels_export_df.columns) + 2)
        domain_matrix_df.to_excel(writer, sheet_name="Partner Levels by Domain",
                                   index=False, startrow=3)
        expiring_df.to_excel(writer, sheet_name="Expiring Certificates", index=False, startrow=1)
        gap_df.to_excel(writer, sheet_name="Gap Analysis", index=False, startrow=1)
        new_certs_df.to_excel(writer, sheet_name="New Certificates", index=False, startrow=1)
        needs_review_df.to_excel(writer, sheet_name="Needs Review (Universities)", index=False, startrow=1)
        clean_df.to_excel(writer, sheet_name="Clean Data (Full)", index=False, startrow=1)

    print("Applying formatting...")
    style_workbook(OUTPUT_PATH)
    style_domain_matrix_sheet(OUTPUT_PATH, domain_badge_counts)

    print("Building dashboard...")
    add_dashboard(OUTPUT_PATH, partner_levels_df, expiring_df, clean_df, needs_review_df,
                  new_certs_df, leveled_up_details)

    print("Done.")
    print(f"\n{len(needs_review_df)} partners need university classification review.")
    print(f"{len(expiring_df)} certificates expiring this year.")
    print(f"{len(new_certs_df)} new certificates since last run.")
    print(f"{len(leveled_up_details)} partner-domain level-ups this run.")


if __name__ == "__main__":
    main()