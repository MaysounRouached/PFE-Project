import pandas as pd
from pathlib import Path
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.table import Table, TableStyleInfo


def read_csv_safe(path, **kwargs):
    try:
        return pd.read_csv(path, encoding="utf-8", **kwargs)
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="cp1252", **kwargs)


def normalize_code(series):
    return (series.astype(str).str.strip()
                  .str.replace(r"\.0$", "", regex=True))


RAW_FILE_PATH = "raw_weekly_partner_data.xlsx"
TARGET_FILE_PATH = "partner_targets_revenue.csv"
HISTORICAL_FILE_PATH = "historical_revenue.csv"
SNAPSHOT_DIR = "weekly_snapshots"
OUTPUT_PATH = "WEEKLY_PARTNERS_PERFORMANCE_REPORT.xlsx"

PARTNER_NAME_COL = "Partner Name"
PARTNER_CODE_COL = "Partner Code"
YTD_REVENUE_COL = "YTD Revenue（USD）"
YTD_BOOKING_COL = "YTD Booking（USD）"
YTD_STORAGE_COL = "YTD STORAGE_Revenue（USD）"
YTD_NETWORK_COL = "YTD NETWORK_Revenue（USD）"

TODAY = datetime.today()
CURRENT_YEAR = TODAY.year
PRIOR_YEAR = CURRENT_YEAR - 1

HUAWEI_RED = "C8102E"
HUAWEI_DARK = "1A1A1A"
GREEN_UP = "1E7B34"
RED_DOWN = "C8102E"
AMBER = "E8A33D"
LIGHT_GREY = "F2F2F2"

CHANNEL_TARGET_SOURCE_COL = "TargetRevenue_YTD_MUSD"

# Partners included in the summary TOTAL row (and in the Dashboard's
# "Total YTD Revenue/Orders" KPI cards). Edit this list if the spelling
# in your Partner column differs slightly from what's below — matching
# is case-insensitive and ignores extra spaces, but the words must match.
TOTAL_ROW_PARTNERS = [
    "Partner Q5",
    "Partner T15",
    "Partner C10",
]

# ============ NC Achievement table config ============
RAW_DR_EXPORT_FILE = "raw_dr_export.xlsx"       # raw DR system export (Chinese/English headers)
HISTORICAL_PL_FILE = "historical_pl.xlsx"        # PL Booking/Revenue by partner for 2023-2025

# The NC Achievement table only ever includes these 7 partners — any other
# partner found in the DR export (even with real 2026 activity) is excluded
# from this specific table. Canonical spelling here is used for the output
# row regardless of how the name is spelled in the raw DR data (fixes
# "Partner T15" vs "Partner B14"-style duplicate rows too).
NC_ACHIEVEMENT_PARTNERS = [
    "Partner Q5",
    "Partner T15",
    "Partner C10",
    "Partner G22",
    "Partner L13",
    "Partner P30",
    "Partner I5",
]
# Known BG Code(s) for a given partner — a row matches that partner if its
# Signed Partner BG Code OR Tier-2 Partner BG Code is in this list, in
# ADDITION to (not instead of) the name-based match. Useful when the name
# columns are unreliable (typos, spelling variants) but the BG code is a
# stable unique identifier. Add more partners here as codes are confirmed.
NC_PARTNER_BG_CODES = {
    "Partner Q5": ["Partner L17"],
}
CHANNEL_REVENUE_REPORT_FILE = "../contracts/CHANNEL_REVENUE_REPORT.xlsx"  # Table 2's own output — used for the Channel/AM split

# ASSUMPTION (flagged, not yet confirmed against real data): a Signed
# Partner matching these keywords is treated as a pass-through distributor
# — the real partner credited is the Tier-2 Partner instead, same pattern
# as Table 2's TMI handling. Matching is done on BG Code where possible
# (set TMI_BG_CODES below once you confirm the real code) and falls back
# to name-keyword matching otherwise.
TMI_KEYWORDS_DR = ["Partner Q5", "Partner S27"]
TMI_CANONICAL_NAME_DR = "Partner Q5"
TMI_BG_CODES = []  # e.g. ["BG1000"] — exact BG code match, once confirmed; takes priority over name matching

DR_RAW_COLS = {
    "opp_no": "机会点编号/Opportunity No.",
    "signed_date": "签约日期/Signed Date",
    "signed_partner": "签约伙伴名称/Signed Partner Name",
    "signed_partner_bg": "签约伙伴BG编码/Signed Partner BG Code",
    "tier2_partner": "二级伙伴名称/Tier-2 Partner Name",
    "tier2_partner_bg": "二级伙伴BG编码/Tier-2 Partner BG Code",
    "customer_class": "客户层级/Customer Classification",
    "booking_usd": "订货金额(USD)/Booking Amount(USD)",
    "revenue_usd": "收入金额(USD)/Revenue Amount(USD)",
    "end_customer": "最终客户名称/End Customer Name",
}
SMA_CLASSIFICATION_VALUE = "Small&Medium Account"
NA_CLASSIFICATION_VALUE = "Core Account"

# Same exclusion as Table 2 — out of scope for these reports regardless of
# which raw file they show up in.
NC_EXCLUDED_ENTITIES = ["Partner R3", "Partner N6"]


def clean_text_dr(v):
    """Blanks out NaN/None properly — str(float('nan')) == 'nan', which was
    previously slipping through as a literal partner name."""
    if pd.isna(v):
        return ""
    return str(v).strip()


def _norm_name(s):
    return " ".join(clean_text_dr(s).casefold().split())

STATUS_COLORS = {
    "Diamond": "5B9BD5", "Distributor": "70AD47", "Gold": "FFC000",
    "Silver": "A6A6A6", "Authorized": "C00000",
}


def load_current(path):
    df = pd.read_excel(path)
    df[PARTNER_CODE_COL] = normalize_code(df[PARTNER_CODE_COL])
    df[PARTNER_NAME_COL] = df[PARTNER_NAME_COL].astype(str).str.strip()
    for col in [YTD_REVENUE_COL, YTD_BOOKING_COL, YTD_STORAGE_COL, YTD_NETWORK_COL]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df


def load_targets(path):
    targets = read_csv_safe(path)
    targets = targets.rename(columns={"PartnerCode": PARTNER_CODE_COL, "PartnerName": PARTNER_NAME_COL})
    targets[PARTNER_CODE_COL] = normalize_code(targets[PARTNER_CODE_COL])
    return targets


def load_historical(path):
    hist = read_csv_safe(path)
    hist = hist.rename(columns={"PartnerCode": PARTNER_CODE_COL, "PartnerName": PARTNER_NAME_COL})
    hist[PARTNER_CODE_COL] = normalize_code(hist[PARTNER_CODE_COL])
    rename_map = {}
    for year in (2023, 2024, 2025):
        rename_map[f"Order_{year}"] = f"TotalOrders_KUSD {year}"
        rename_map[f"Revenue_{year}"] = f"TotalRevenue_KUSD {year}"
    hist = hist.rename(columns=rename_map)
    return hist


def sum_duplicates(df, label):
    """When a partner appears in two or more rows (same Partner Code), sum
    their numeric columns into a single consolidated row instead of dropping
    the extras. Non-numeric columns (name, status, etc.) keep the first
    value seen."""
    dupes = df[df.duplicated(PARTNER_CODE_COL, keep=False)]
    if dupes.empty:
        return df

    codes = sorted(dupes[PARTNER_CODE_COL].unique())
    print(f"NOTE: {label} has partner(s) appearing in multiple rows — "
          f"summing their numeric columns into one row. Affected codes: {codes}")

    numeric_cols = [c for c in df.select_dtypes(include="number").columns if c != PARTNER_CODE_COL]
    agg = {c: "sum" for c in numeric_cols}
    for c in df.columns:
        if c not in agg and c != PARTNER_CODE_COL:
            agg[c] = "first"
    grouped = df.groupby(PARTNER_CODE_COL, as_index=False).agg(agg)
    return grouped[df.columns.tolist()]


def find_snapshot_near(days_ago, min_gap_days=20):
    folder = Path(SNAPSHOT_DIR)
    if not folder.exists():
        return None
    candidates = []
    for f in folder.glob("snapshot_*.csv"):
        date_str = f.stem.replace("snapshot_", "")
        try:
            d = datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            continue
        gap = (TODAY - d).days
        if gap >= min_gap_days:
            candidates.append((abs(gap - days_ago), f))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0])
    return read_csv_safe(candidates[0][1])


def save_mom_snapshot(df):
    """Saves today's Partner Code / YTD Booking / YTD Revenue to
    weekly_snapshots/ so find_snapshot_near() has something to compare
    against ~30 days from now. This was previously never actually being
    written — MoM Growth has likely shown 'Not enough history yet' on
    every single run until now, regardless of how long this script had
    been in use."""
    folder = Path(SNAPSHOT_DIR)
    folder.mkdir(exist_ok=True)
    snap = df[[PARTNER_CODE_COL, YTD_BOOKING_COL, YTD_REVENUE_COL]].copy()
    path = folder / f"snapshot_{TODAY.strftime('%Y-%m-%d')}.csv"
    snap.to_csv(path, index=False)
    return path


def build_report():
    current = sum_duplicates(load_current(RAW_FILE_PATH), "raw weekly export")
    targets = sum_duplicates(load_targets(TARGET_FILE_PATH), "partner_targets_revenue.csv")
    hist_wide = sum_duplicates(load_historical(HISTORICAL_FILE_PATH), "historical_revenue.csv")

    mom_snapshot = find_snapshot_near(30)
    if mom_snapshot is not None:
        mom_snapshot[PARTNER_CODE_COL] = normalize_code(mom_snapshot[PARTNER_CODE_COL])
        mom_lookup = mom_snapshot.set_index(PARTNER_CODE_COL)
    else:
        mom_lookup = None

    df = targets.merge(current, on=PARTNER_CODE_COL, how="left", suffixes=("", "_cur"))
    df = df.merge(hist_wide, on=PARTNER_CODE_COL, how="left", suffixes=("", "_hist"))

    missing_current = df[df[YTD_BOOKING_COL].isna()][PARTNER_CODE_COL].tolist()
    if missing_current:
        print(f"NOTE: {len(missing_current)} target-list partner(s) not found "
              f"in this week's raw export: {missing_current}")

    rows = []
    for _, r in df.iterrows():
        code = r[PARTNER_CODE_COL]

        ytd_booking = r[YTD_BOOKING_COL] if pd.notna(r[YTD_BOOKING_COL]) else 0
        ytd_revenue = r[YTD_REVENUE_COL] if pd.notna(r[YTD_REVENUE_COL]) else 0
        ytd_storage = r[YTD_STORAGE_COL] if pd.notna(r[YTD_STORAGE_COL]) else 0
        ytd_network = r[YTD_NETWORK_COL] if pd.notna(r[YTD_NETWORK_COL]) else 0

        orders_2026 = ytd_booking / 1000
        revenue_2026 = ytd_revenue / 1000
        storage_k = ytd_storage / 1000
        network_k = ytd_network / 1000
        other_revenue_k = revenue_2026 - storage_k - network_k

        target_booking = r.get("TargetBooking_YTD_MUSD")
        target_revenue = r.get("TargetRevenue_YTD_MUSD")
        channel_target_m = (r.get(CHANNEL_TARGET_SOURCE_COL)
                             if pd.notna(r.get(CHANNEL_TARGET_SOURCE_COL)) else "Not set")

        order_achievement = (round((ytd_booking / 1_000_000) / target_booking * 100, 1)
                              if pd.notna(target_booking) and target_booking else "Not set")
        revenue_achievement = (round((ytd_revenue / 1_000_000) / target_revenue * 100, 1)
                                if pd.notna(target_revenue) and target_revenue else "Not set")

        if mom_lookup is not None and code in mom_lookup.index:
            prev = mom_lookup.loc[code]
            mom_order = round((ytd_booking - prev[YTD_BOOKING_COL]) / 1000, 1)
            mom_revenue = round((ytd_revenue - prev[YTD_REVENUE_COL]) / 1000, 1)
        else:
            mom_order = "Not enough history yet"
            mom_revenue = "Not enough history yet"

        prior_orders_k = r.get(f"TotalOrders_KUSD {PRIOR_YEAR}")
        prior_revenue_k = r.get(f"TotalRevenue_KUSD {PRIOR_YEAR}")
        yoy_order = (round((orders_2026 - prior_orders_k) / prior_orders_k * 100, 1)
                     if pd.notna(prior_orders_k) and prior_orders_k else "No prior-year data")
        yoy_revenue = (round((revenue_2026 - prior_revenue_k) / prior_revenue_k * 100, 1)
                       if pd.notna(prior_revenue_k) and prior_revenue_k else "No prior-year data")

        rows.append({
            "Partner": r[PARTNER_NAME_COL],
            "Status": r.get("Status", "Not set"),
            "Channel Target (M$)": channel_target_m,
            "Tot Orders 2023 (K$)": r.get("TotalOrders_KUSD 2023"),
            "Tot Orders 2024 (K$)": r.get("TotalOrders_KUSD 2024"),
            "Total Orders 2025 (K$)": r.get("TotalOrders_KUSD 2025"),
            f"Total Orders {CURRENT_YEAR} (K$)": round(orders_2026),
            "Order Achievement %": order_achievement,
            "Order MoM Growth (K$)": mom_order,
            "Order YoY Growth %": yoy_order,
            "Tot Revenue 2023 (K$)": r.get("TotalRevenue_KUSD 2023"),
            "Tot Revenue 2024 (K$)": r.get("TotalRevenue_KUSD 2024"),
            "Tot Revenue 2025 (K$)": r.get("TotalRevenue_KUSD 2025"),
            "Storage Revenue (K$)": round(storage_k, 1),
            "Network Revenue (K$)": round(network_k, 1),
            "Other Revenue (K$)": round(other_revenue_k, 1),
            f"Total Revenues {CURRENT_YEAR} (K$)": round(revenue_2026),
            "Revenue Achievement %": revenue_achievement,
            "Revenue MoM Growth (K$)": mom_revenue,
            "Revenue YoY Growth %": yoy_revenue,
        })

    save_mom_snapshot(df)
    return pd.DataFrame(rows)


# ============ Summary TOTAL row (3 named partners only) ============

NUMERIC_TOTAL_COLS = [
    "Tot Orders 2023 (K$)", "Tot Orders 2024 (K$)", "Total Orders 2025 (K$)",
    f"Total Orders {CURRENT_YEAR} (K$)",
    "Tot Revenue 2023 (K$)", "Tot Revenue 2024 (K$)", "Tot Revenue 2025 (K$)",
    "Storage Revenue (K$)", "Network Revenue (K$)", "Other Revenue (K$)",
    f"Total Revenues {CURRENT_YEAR} (K$)",
]


def build_total_row(report_df):
    """Sums Orders/Revenue columns for only TOTAL_ROW_PARTNERS. Achievement %
    and Growth columns are left blank since summing percentages across
    partners isn't meaningful."""
    wanted = {_norm_name(p) for p in TOTAL_ROW_PARTNERS}
    subset = report_df[report_df["Partner"].apply(lambda p: _norm_name(p) in wanted)]

    found = set(subset["Partner"].apply(_norm_name))
    missing = wanted - found
    if missing:
        print(f"WARNING: TOTAL_ROW_PARTNERS not found in the report (check spelling "
              f"against your Partner column): {sorted(missing)}")

    total = {col: "" for col in report_df.columns}
    total["Partner"] = "TOTAL"
    total["Status"] = ""
    whole_number_cols = {f"Total Orders {CURRENT_YEAR} (K$)", f"Total Revenues {CURRENT_YEAR} (K$)"}
    for col in NUMERIC_TOTAL_COLS:
        if col in subset.columns:
            digits = 0 if col in whole_number_cols else 1
            total[col] = round(pd.to_numeric(subset[col], errors="coerce").sum(), digits)
    return total


# ============ Styling: main report sheet ============

def style_report(ws):
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color=HUAWEI_RED, end_color=HUAWEI_RED, fill_type="solid")
    body_font = Font(name="Calibri", size=10, color="7A1414")  # dark red instead of black
    thin = Side(style="thin", color="F0C2C2")
    band_fill = PatternFill(start_color="FBE4E4", end_color="FBE4E4", fill_type="solid")  # light red tint

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 45

    headers = [c.value for c in ws[1]]
    growth_cols = [i + 1 for i, h in enumerate(headers)
                   if h and ("Growth" in str(h) or "Achievement" in str(h))]

    last_row = ws.max_row
    for row in ws.iter_rows(min_row=2, max_row=last_row):
        for cell in row:
            cell.font = body_font
            cell.border = Border(bottom=thin)
        if row[0].row % 2 == 0:
            for cell in row:
                cell.fill = band_fill

    # Color scale on Achievement % columns (red -> amber -> green) — softer,
    # more refined tones than Excel's default heat-map colors, and matched
    # to the same red/green family used everywhere else in this workbook
    # instead of clashing with it.
    for idx, h in enumerate(headers, start=1):
        if h and "Achievement" in str(h):
            col_letter = get_column_letter(idx)
            rng = f"{col_letter}2:{col_letter}{last_row}"
            rule = ColorScaleRule(
                start_type="num", start_value=0, start_color="F4B6B6",
                mid_type="num", mid_value=100, mid_color="FCE4A6",
                end_type="num", end_value=150, end_color="B7DDB0",
            )
            ws.conditional_formatting.add(rng, rule)
        if h and "Growth" in str(h):
            col_letter = get_column_letter(idx)
            rng = f"{col_letter}2:{col_letter}{last_row}"
            rule = DataBarRule(start_type="min", end_type="max", color="C8102E")
            ws.conditional_formatting.add(rng, rule)

    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions

    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        col_letter = get_column_letter(col_cells[0].column)
        ws.column_dimensions[col_letter].width = min(max(length + 1, 7), 14)
    ws.column_dimensions["A"].width = 24  # Partner names — wide enough that none get cut off
    ws.sheet_view.zoomScale = 65  # whole table visible without horizontal scrolling

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# ============ Dashboard sheet ============

# Huawei-guided palette: red for the brand accent + "needs attention",
# charcoal/graphite for neutral totals, a restrained status scale for
# performance (green/amber/red), and cool grays for structure/status mix.
HUAWEI_GRAPHITE = "3C3C3C"
CARD_BG = "F4F4F4"
GRID_LINE = "E0E0E0"
STATUS_PALETTE = ["C8102E", "3C3C3C", "8C8C8C", "E8A33D", "8C1A1A", "70AD47"]


def to_num(series):
    return pd.to_numeric(series, errors="coerce")


def perf_color(value):
    if pd.isna(value):
        return "B0B0B0"
    if value >= 100:
        return GREEN_UP
    if value >= 70:
        return AMBER
    return RED_DOWN


def kpi_card(ws, col, row, label, value, accent):
    """A flat card: thin accent bar on top, dark label, large bold value."""
    span = 2
    accent_cell = ws.cell(row=row, column=col)
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
    accent_cell.fill = PatternFill(start_color=accent, end_color=accent, fill_type="solid")
    ws.row_dimensions[row].height = 5

    label_row = row + 1
    ws.merge_cells(start_row=label_row, start_column=col, end_row=label_row, end_column=col + span - 1)
    lc = ws.cell(row=label_row, column=col, value=label)
    lc.font = Font(name="Calibri", size=9, bold=True, color="808080")
    lc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[label_row].height = 32

    val_row = row + 2
    ws.merge_cells(start_row=val_row, start_column=col, end_row=val_row, end_column=col + span - 1)
    vc = ws.cell(row=val_row, column=col, value=value)
    vc.font = Font(name="Calibri", size=22, bold=True, color=accent)
    vc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[val_row].height = 34

    for r in (accent_cell.row, label_row, val_row):
        for c in range(col, col + span):
            cell = ws.cell(row=r, column=c)
            if r != accent_cell.row:
                cell.fill = PatternFill(start_color=CARD_BG, end_color=CARD_BG, fill_type="solid")
            cell.border = Border(
                left=Side(style="thin", color=GRID_LINE) if c == col else None,
                right=Side(style="thin", color=GRID_LINE) if c == col + span - 1 else None,
                top=Side(style="thin", color=GRID_LINE) if r == accent_cell.row else None,
                bottom=Side(style="thin", color=GRID_LINE) if r == val_row else None,
            )


def style_bar_chart(chart, title, y_title=None):
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import (RichTextProperties, Paragraph,
        ParagraphProperties, CharacterProperties)
    chart.title = title
    if y_title and hasattr(chart, "y_axis"):
        chart.y_axis.title = y_title
        # Give the (vertically rotated) axis title more breathing room from
        # the tick-number labels next to it, so it doesn't read as crowded
        # or overlapping — nudges the plot area right, freeing space on
        # the left for the title.
        from openpyxl.chart.layout import Layout, ManualLayout
        chart.plot_area.layout = Layout(manualLayout=ManualLayout(
            xMode="edge", yMode="edge", x=0.14, y=0.12, w=0.83, h=0.72))
    if hasattr(chart, "gapWidth"):
        chart.gapWidth = 40
    chart.legend = None
    if hasattr(chart, "x_axis"):
        small_font = CharacterProperties(sz=750)
        chart.x_axis.txPr = RichText(
            bodyPr=RichTextProperties(rot=-2700000, vert="horz"),
            p=[Paragraph(pPr=ParagraphProperties(defRPr=small_font), endParaRPr=small_font)]
        )
        # Explicitly force the category axis to actually render its tick
        # labels. Some Excel versions silently hide them (leaving a blank
        # gap under the bars, no partner names) for charts built from
        # hidden-column source data unless this is set explicitly —
        # leaving it unset apparently isn't a safe default.
        chart.x_axis.delete = False
        chart.x_axis.tickLblPos = "nextTo"


def set_axis_label_color(chart, color):
    """Color the category-axis (partner name) labels — used on the two
    achievement bar charts so partner names are readable and Huawei-red."""
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import (RichTextProperties, Paragraph,
        ParagraphProperties, CharacterProperties)
    from openpyxl.drawing.fill import ColorChoice
    red_font = CharacterProperties(sz=750, solidFill=color)
    chart.x_axis.txPr = RichText(
        bodyPr=RichTextProperties(rot=-2700000, vert="horz"),
        p=[Paragraph(pPr=ParagraphProperties(defRPr=red_font), endParaRPr=red_font)]
    )
    chart.x_axis.delete = False
    chart.x_axis.tickLblPos = "nextTo"


def add_partner_name_labels(series, color, position="outEnd"):
    """Attach the partner name directly to each bar as a data label. This is
    a belt-and-suspenders fix alongside the axis text color: some Excel
    versions/settings can suppress category-axis text for charts built from
    hidden columns, but a data label attached to the bar itself always shows."""
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import (RichTextProperties, Paragraph,
        ParagraphProperties, CharacterProperties)
    dl = DataLabelList()
    dl.showCatName = True
    dl.showVal = False
    dl.showSerName = False
    dl.showPercent = False
    dl.showLegendKey = False
    dl.dLblPos = position
    dl.txPr = RichText(
        bodyPr=RichTextProperties(rot=-2700000, vert="horz"),
        p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=700, solidFill=color)),
                      endParaRPr=CharacterProperties(sz=700, solidFill=color))]
    )
    series.dLbls = dl


def add_value_labels(series, color, position="outEnd"):
    """Attach the actual data value (achievement % or K$ amount) directly
    on each bar, horizontal (not rotated) so it's easy to read at a glance —
    used instead of per-bar partner names, which now show only via the
    x-axis category labels at the bottom (set_axis_label_color)."""
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import (RichTextProperties, Paragraph,
        ParagraphProperties, CharacterProperties)
    dl = DataLabelList()
    dl.showCatName = False
    dl.showVal = True
    dl.showSerName = False
    dl.showPercent = False
    dl.showLegendKey = False
    dl.dLblPos = position
    dl.txPr = RichText(
        bodyPr=RichTextProperties(rot=0, vert="horz"),
        p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=800, b=True, solidFill=color)),
                      endParaRPr=CharacterProperties(sz=800, b=True, solidFill=color))]
    )
    series.dLbls = dl


def colored_series(series, values, thresholds=True):
    """Apply per-bar Huawei performance colors (or a flat red) to a bar series."""
    from openpyxl.chart.marker import DataPoint
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.fill import PatternFillProperties
    points = []
    for i, v in enumerate(values):
        color = perf_color(v) if thresholds else HUAWEI_RED
        dp = DataPoint(idx=i)
        dp.graphicalProperties = GraphicalProperties(solidFill=color)
        points.append(dp)
    series.data_points = points


def build_dashboard(wb, report_df):
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False

    # ---- Header banner ----
    ws.merge_cells("B2:V2")
    title = ws["B2"]
    title.value = "WEEKLY_PARTNERS_PERFORMANCE_REPORT"
    title.font = Font(name="Calibri", size=20, bold=True, color="FFFFFF")
    title.fill = PatternFill(start_color=HUAWEI_RED, end_color=HUAWEI_RED, fill_type="solid")
    title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 38

    ws.merge_cells("B3:V3")
    sub = ws["B3"]
    sub.value = f"Generated {TODAY.strftime('%Y-%m-%d')}  ·  {len(report_df)} partners  ·  FY{CURRENT_YEAR}"
    sub.font = Font(name="Calibri", size=10, italic=True, color="8C8C8C")
    ws.row_dimensions[3].height = 16
    ws.row_dimensions[4].height = 8

    # ---- KPI metrics ----
    rev_ach = to_num(report_df["Revenue Achievement %"])
    ord_ach = to_num(report_df["Order Achievement %"])
    # Total YTD Revenue/Orders KPI cards only count TOTAL_ROW_PARTNERS,
    # matching the summary TOTAL row on the report sheet.
    _wanted = {_norm_name(p) for p in TOTAL_ROW_PARTNERS}
    _totals_subset = report_df[report_df["Partner"].apply(lambda p: _norm_name(p) in _wanted)]
    total_rev = _totals_subset[f"Total Revenues {CURRENT_YEAR} (K$)"].sum()
    total_ord = _totals_subset[f"Total Orders {CURRENT_YEAR} (K$)"].sum()
    avg_rev_ach = rev_ach.mean()
    avg_ord_ach = ord_ach.mean()
    on_track = int((ord_ach >= 100).sum())
    at_risk = int(((ord_ach >= 70) & (ord_ach < 100)).sum())
    behind = int((ord_ach < 70).sum())

    kpi_specs = [
        ("TOTAL YTD REVENUE (K$)", f"{total_rev:,.0f}", HUAWEI_GRAPHITE),
        ("TOTAL YTD ORDERS (K$)", f"{total_ord:,.0f}", HUAWEI_GRAPHITE),
        ("AVG REVENUE ACHIEVEMENT", f"{avg_rev_ach:,.0f}%", perf_color(avg_rev_ach)),
        ("AVG ORDER ACHIEVEMENT", f"{avg_ord_ach:,.0f}%", perf_color(avg_ord_ach)),
        ("ON TRACK (Orders) \u2265100%", str(on_track), GREEN_UP),
        ("APPROACHING TARGET (Orders) 70-99%", str(at_risk), AMBER),
        ("BEHIND (Orders) <70%", str(behind), RED_DOWN),
    ]
    kpi_row = 5
    col = 2
    for label, value, accent in kpi_specs:
        kpi_card(ws, col, kpi_row, label, value, accent)
        col += 3  # 2-col card + 1-col gap
    ws.row_dimensions[kpi_row + 3].height = 14

    # ---- data staging for charts, pushed to far-right hidden columns ----
    STAGE1_COL, STAGE2_COL, STAGE4_COL = 40, 43, 49
    STAGE5_COL, STAGE6_COL = 52, 55  # raw revenue/order value charts (new)
    data_start_row = 5

    sorted_by_rev_ach = report_df.assign(_v=rev_ach).sort_values("_v", ascending=False)
    ws.cell(row=data_start_row, column=STAGE1_COL, value="Partner")
    ws.cell(row=data_start_row, column=STAGE1_COL + 1, value="Revenue Achievement %")
    for i, (_, r) in enumerate(sorted_by_rev_ach.iterrows(), start=1):
        ws.cell(row=data_start_row + i, column=STAGE1_COL, value=r["Partner"])
        v = r["_v"]
        ws.cell(row=data_start_row + i, column=STAGE1_COL + 1, value=float(v) if pd.notna(v) else None)
    n = len(sorted_by_rev_ach)
    rev_vals = [r["_v"] for _, r in sorted_by_rev_ach.iterrows()]

    sorted_by_ord_ach = report_df.assign(_v=ord_ach).sort_values("_v", ascending=False)
    ws.cell(row=data_start_row, column=STAGE2_COL, value="Partner")
    ws.cell(row=data_start_row, column=STAGE2_COL + 1, value="Order Achievement %")
    for i, (_, r) in enumerate(sorted_by_ord_ach.iterrows(), start=1):
        ws.cell(row=data_start_row + i, column=STAGE2_COL, value=r["Partner"])
        v = r["_v"]
        ws.cell(row=data_start_row + i, column=STAGE2_COL + 1, value=float(v) if pd.notna(v) else None)
    n2 = len(sorted_by_ord_ach)
    ord_vals = [r["_v"] for _, r in sorted_by_ord_ach.iterrows()]

    _focus_wanted = {_norm_name(p) for p in TOTAL_ROW_PARTNERS}
    top5 = report_df.assign(_ord=report_df[f"Total Orders {CURRENT_YEAR} (K$)"])
    top5 = top5[top5["Partner"].apply(lambda p: _norm_name(p) in _focus_wanted)].sort_values("_ord", ascending=True)
    ws.cell(row=data_start_row, column=STAGE4_COL, value="Partner")
    ws.cell(row=data_start_row, column=STAGE4_COL + 1, value="Orders (K$)")
    for i, (_, r) in enumerate(top5.iterrows(), start=1):
        ws.cell(row=data_start_row + i, column=STAGE4_COL, value=r["Partner"])
        ws.cell(row=data_start_row + i, column=STAGE4_COL + 1, value=float(r["_ord"]))
    n_top5 = len(top5)

    # Raw Revenue (K$) by partner — actual dollar values, not achievement %
    sorted_by_rev_val = report_df.assign(_v=report_df[f"Total Revenues {CURRENT_YEAR} (K$)"]).sort_values("_v", ascending=False)
    ws.cell(row=data_start_row, column=STAGE5_COL, value="Partner")
    ws.cell(row=data_start_row, column=STAGE5_COL + 1, value="Revenue (K$)")
    for i, (_, r) in enumerate(sorted_by_rev_val.iterrows(), start=1):
        ws.cell(row=data_start_row + i, column=STAGE5_COL, value=r["Partner"])
        v = r["_v"]
        ws.cell(row=data_start_row + i, column=STAGE5_COL + 1, value=float(v) if pd.notna(v) else None)
    n5 = len(sorted_by_rev_val)

    # Raw Order (K$) by partner — actual dollar values, not achievement %
    sorted_by_ord_val = report_df.assign(_v=report_df[f"Total Orders {CURRENT_YEAR} (K$)"]).sort_values("_v", ascending=False)
    ws.cell(row=data_start_row, column=STAGE6_COL, value="Partner")
    ws.cell(row=data_start_row, column=STAGE6_COL + 1, value="Order (K$)")
    for i, (_, r) in enumerate(sorted_by_ord_val.iterrows(), start=1):
        ws.cell(row=data_start_row + i, column=STAGE6_COL, value=r["Partner"])
        v = r["_v"]
        ws.cell(row=data_start_row + i, column=STAGE6_COL + 1, value=float(v) if pd.notna(v) else None)
    n6 = len(sorted_by_ord_val)

    for c in [STAGE1_COL, STAGE1_COL + 1, STAGE2_COL, STAGE2_COL + 1,
              STAGE4_COL, STAGE4_COL + 1, STAGE5_COL, STAGE5_COL + 1, STAGE6_COL, STAGE6_COL + 1]:
        ws.column_dimensions[get_column_letter(c)].hidden = True

    # ---- Chart 1: Revenue Achievement % (colored red/amber/green per bar) ----
    bar = BarChart()
    bar.type = "col"
    style_bar_chart(bar, "Revenue Achievement % by Partner", "Achievement %")
    data = Reference(ws, min_col=STAGE1_COL + 1, min_row=data_start_row, max_row=data_start_row + n)
    cats = Reference(ws, min_col=STAGE1_COL, min_row=data_start_row + 1, max_row=data_start_row + n)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.visible_cells_only = False
    colored_series(bar.series[0], rev_vals, thresholds=False)  # flat Huawei red, no amber/green
    set_axis_label_color(bar, HUAWEI_RED)
    add_value_labels(bar.series[0], HUAWEI_DARK)
    bar.height = 10
    bar.width = 18
    ws.add_chart(bar, "B15")

    # ---- Chart 2: Order Achievement % (colored red/amber/green per bar) ----
    bar_o = BarChart()
    bar_o.type = "col"
    style_bar_chart(bar_o, "Order Achievement % by Partner", "Achievement %")
    data_o = Reference(ws, min_col=STAGE2_COL + 1, min_row=data_start_row, max_row=data_start_row + n2)
    cats_o = Reference(ws, min_col=STAGE2_COL, min_row=data_start_row + 1, max_row=data_start_row + n2)
    bar_o.add_data(data_o, titles_from_data=True)
    bar_o.set_categories(cats_o)
    bar_o.visible_cells_only = False
    colored_series(bar_o.series[0], ord_vals, thresholds=False)  # flat Huawei red, no amber/green
    set_axis_label_color(bar_o, HUAWEI_RED)
    add_value_labels(bar_o.series[0], HUAWEI_DARK)
    bar_o.height = 10
    bar_o.width = 18
    ws.add_chart(bar_o, "M15")

    # ---- Chart 3: Revenue (K$) by partner — raw values, not achievement % ----
    bar_rv = BarChart()
    bar_rv.type = "col"
    style_bar_chart(bar_rv, f"Revenue (K$) by Partner \u2014 {CURRENT_YEAR}", "K$")
    data_rv = Reference(ws, min_col=STAGE5_COL + 1, min_row=data_start_row, max_row=data_start_row + n5)
    cats_rv = Reference(ws, min_col=STAGE5_COL, min_row=data_start_row + 1, max_row=data_start_row + n5)
    bar_rv.add_data(data_rv, titles_from_data=True)
    bar_rv.set_categories(cats_rv)
    bar_rv.visible_cells_only = False
    colored_series(bar_rv.series[0], [1] * n5, thresholds=False)  # flat Huawei red
    set_axis_label_color(bar_rv, HUAWEI_RED)
    add_value_labels(bar_rv.series[0], HUAWEI_DARK)
    bar_rv.height = 10
    bar_rv.width = 18
    ws.add_chart(bar_rv, "B37")

    # ---- Chart 4: Order (K$) by partner — raw values, not achievement % ----
    bar_ov = BarChart()
    bar_ov.type = "col"
    style_bar_chart(bar_ov, f"Order (K$) by Partner \u2014 {CURRENT_YEAR}", "K$")
    data_ov = Reference(ws, min_col=STAGE6_COL + 1, min_row=data_start_row, max_row=data_start_row + n6)
    cats_ov = Reference(ws, min_col=STAGE6_COL, min_row=data_start_row + 1, max_row=data_start_row + n6)
    bar_ov.add_data(data_ov, titles_from_data=True)
    bar_ov.set_categories(cats_ov)
    bar_ov.visible_cells_only = False
    colored_series(bar_ov.series[0], [1] * n6, thresholds=False)  # flat Huawei red
    set_axis_label_color(bar_ov, HUAWEI_RED)
    add_value_labels(bar_ov.series[0], HUAWEI_DARK)
    bar_ov.height = 10
    bar_ov.width = 18
    ws.add_chart(bar_ov, "M37")

    ws.column_dimensions["A"].width = 2
    for letter in ["B","C","D","E","F","G","H","I","J","K","L","M",
                   "N","O","P","Q","R","S","T","U","V"]:
        ws.column_dimensions[letter].width = 9

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = "A1:X80"


def load_dr_raw(path):
    if not path or not Path(path).exists():
        print(f"NOTE: RAW_DR_EXPORT_FILE '{path}' not found — NC Achievement table will be skipped.")
        return None
    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_historical_pl(path):
    """Loads the historical PL sheet. Real column names confirmed:
    Partners, Status, Booking PL 2023/2024/2025 (K$), Revenue PL 2023/2024/2025 (K$).
    Normalizes to 'Partner' + the internal PL Booking/Revenue column names
    used elsewhere in this script."""
    if not path or not Path(path).exists():
        print(f"NOTE: HISTORICAL_PL_FILE '{path}' not found — PL 2023-2025 columns "
              f"will be blank, and the partner list will only include partners "
              f"with {CURRENT_YEAR} DR activity.")
        return None
    hist = pd.read_excel(path) if str(path).lower().endswith((".xlsx", ".xls")) else read_csv_safe(path)
    hist.columns = [str(c).strip() for c in hist.columns]
    print(f"HISTORICAL_PL_FILE columns as loaded (before rename): {hist.columns.tolist()}")
    rename_map = {
        "Partners": "Partner",
        "Booking PL 2023 (K$)": "PL Booking 2023 (K$)",
        "Booking PL 2024 (K$)": "PL Booking 2024 (K$)",
        "Booking PL 2025 (K$)": "PL Booking 2025 (K$)",
        "Revenue PL 2023 (K$)": "PL Revenue 2023 (K$)",
        "Revenue PL 2024 (K$)": "PL Revenue 2024 (K$)",
        "Revenue PL 2025 (K$)": "PL Revenue 2025 (K$)",
    }
    hist = hist.rename(columns={k: v for k, v in rename_map.items() if k in hist.columns})
    expected = set(rename_map.values())
    missing = expected - set(hist.columns)
    if missing:
        print(f"WARNING: after renaming, HISTORICAL_PL_FILE is still missing {sorted(missing)} — "
              f"those PL figures will silently show as 0 for every partner. This means the "
              f"original column header(s) in your file don't exactly match what this script "
              f"expects (extra space, different wording, etc.) — compare against the 'before "
              f"rename' list printed above.")
    return hist


def load_channel_lookup(path):
    """Loads Table 2's own output to get OPP NUMBER -> channel (Y/N),
    used to split SMA PL into Channel vs AM. Tries the sheet name a few
    ways since it's been referred to differently at different points —
    avoids silently returning nothing just because of a naming mismatch."""
    if not path or not Path(path).exists():
        print(f"NOTE: CHANNEL_REVENUE_REPORT_FILE '{path}' not found — SMA PL "
              f"will all be bucketed as AM (no Channel split).")
        return {}
    try:
        wb = pd.ExcelFile(path)
    except Exception as e:
        print(f"WARNING: couldn't open CHANNEL_REVENUE_REPORT_FILE '{path}': {e}")
        return {}

    sheet_name = None
    for candidate in ("CHANNEL_REVENUE_REPORT", "Channel Revenue Report"):
        if candidate in wb.sheet_names:
            sheet_name = candidate
            break
    if sheet_name is None:
        sheet_name = wb.sheet_names[0]
        print(f"NOTE: neither 'CHANNEL_REVENUE_REPORT' nor 'Channel Revenue Report' sheet "
              f"found in '{path}' — using the first sheet ('{sheet_name}') instead. "
              f"Actual sheets in that file: {wb.sheet_names}")

    df = pd.read_excel(path, sheet_name=sheet_name)
    df.columns = [str(c).strip() for c in df.columns]
    lookup = {}
    for _, r in df.iterrows():
        opp = normalize_code(pd.Series([r.get("OPP NUMBER", "")])).iloc[0]
        if opp:
            lookup[opp] = str(r.get("channel", "")).strip().upper()
    print(f"Channel lookup loaded from '{path}' sheet '{sheet_name}': {len(lookup)} opportunity(ies).")
    return lookup


def effective_partner(signed_name, signed_bg, tier2_name, tier2_bg):
    """Returns the partner name to credit — Tier-2 if Signed Partner is a
    pass-through distributor (TMI), else Signed Partner. BG code match
    takes priority over name-keyword match once TMI_BG_CODES is filled in.
    Returns "" (not the string "nan") when both are genuinely blank.
    When Signed Partner IS TMI (in any spelling/abbreviation) and there's
    no Tier-2 to hand credit to, the name is still standardized to the
    canonical "Partner Q5" — never left as a raw "Partner S27"
    that would show up as a separate, unmatched partner row."""
    signed_name = clean_text_dr(signed_name)
    tier2_name = clean_text_dr(tier2_name)
    signed_bg = clean_text_dr(signed_bg)

    is_tmi = False
    if TMI_BG_CODES and signed_bg in TMI_BG_CODES:
        is_tmi = True
    elif any(kw in signed_name.casefold() for kw in TMI_KEYWORDS_DR):
        is_tmi = True
    if is_tmi and tier2_name:
        return tier2_name
    if is_tmi:
        return TMI_CANONICAL_NAME_DR
    return signed_name


def build_nc_achievement(table1_report_df):
    dr = load_dr_raw(RAW_DR_EXPORT_FILE)
    if dr is None:
        return None

    hist_pl = load_historical_pl(HISTORICAL_PL_FILE)
    channel_lookup = load_channel_lookup(CHANNEL_REVENUE_REPORT_FILE)
    c = DR_RAW_COLS

    # Exclude out-of-scope entities regardless of which column they show up
    # in (same rule as Table 2's Partner R3 exclusion).
    before = len(dr)
    check_cols = [c["signed_partner"], c["tier2_partner"], c["end_customer"]]
    mask = pd.Series(False, index=dr.index)
    for col in check_cols:
        if col in dr.columns:
            mask |= dr[col].apply(lambda v: any(kw in clean_text_dr(v).casefold() for kw in NC_EXCLUDED_ENTITIES))
    if mask.any():
        print(f"Excluded {mask.sum()} row(s) matching NC_EXCLUDED_ENTITIES out of {before}.")
    dr = dr[~mask].copy()

    # A deal is credited to a whitelisted partner if that partner appears
    # in EITHER Signed Partner Name (K) OR Tier-2 Partner Name (O) — not
    # "pick one column, ignore the other". This means a single row CAN
    # contribute to two different partners' totals at once if K and O
    # name two different whitelisted partners (e.g. K="Partner T15",
    # O="Partner G22" both get this row's full value).
    # A row where K and O name the SAME partner is still only counted
    # once for that partner (it's one deal, not summed twice).
    dr["_k_norm"] = dr[c["signed_partner"]].apply(lambda v: _norm_name(clean_text_dr(v)))
    dr["_o_norm"] = dr[c["tier2_partner"]].apply(lambda v: _norm_name(clean_text_dr(v)))
    dr["_signed_bg"] = dr[c["signed_partner_bg"]].apply(clean_text_dr)
    dr["_tier2_bg"] = dr[c["tier2_partner_bg"]].apply(clean_text_dr)

    def partner_involved_mask(df, partner):
        """A row is 'involved' with this partner if the partner's name
        matches Signed Partner (K) or Tier-2 Partner (O), OR — for
        partners with a known BG code in NC_PARTNER_BG_CODES — if the
        Signed Partner BG Code or Tier-2 Partner BG Code matches. BG code
        is checked IN ADDITION to name, catching rows where the name
        column has a typo/spelling variant but the BG code is reliable."""
        p_norm = _norm_name(partner)
        involved = (df["_k_norm"] == p_norm) | (df["_o_norm"] == p_norm)
        bg_codes = {str(bc).strip() for bc in NC_PARTNER_BG_CODES.get(partner, [])}
        if bg_codes:
            involved = involved | df["_signed_bg"].isin(bg_codes) | df["_tier2_bg"].isin(bg_codes)
        return involved

    no_partner_count = ((dr["_k_norm"] == "") & (dr["_o_norm"] == "")).sum()
    if no_partner_count:
        print(f"NOTE: {no_partner_count} row(s) had no Signed Partner or Tier-2 Partner at "
              f"all — excluded (can't attribute to any partner).")
    dr = dr[(dr["_k_norm"] != "") | (dr["_o_norm"] != "")].copy()

    on_whitelist = pd.Series(False, index=dr.index)
    for _p in NC_ACHIEVEMENT_PARTNERS:
        on_whitelist |= partner_involved_mask(dr, _p)
    off_whitelist_count = (~on_whitelist).sum()
    if off_whitelist_count:
        off_names = sorted(set(dr.loc[~on_whitelist, c["signed_partner"]].apply(clean_text_dr)) |
                            set(dr.loc[~on_whitelist, c["tier2_partner"]].apply(clean_text_dr)))
        off_names = sorted(n for n in off_names if n)
        print(f"NOTE: {off_whitelist_count} row(s) out of {len(dr)} have neither Signed Partner "
              f"nor Tier-2 Partner matching NC_ACHIEVEMENT_PARTNERS (by name or known BG code) "
              f"— excluded from this table. Partner name(s) seen on those rows: {off_names}")
    dr = dr[on_whitelist].copy()

    dr["_signed_year"] = pd.to_datetime(dr[c["signed_date"]], errors="coerce").dt.year
    dr["_opp_key"] = normalize_code(dr[c["opp_no"]].astype(str))
    dr["_booking_k"] = pd.to_numeric(dr[c["booking_usd"]], errors="coerce").fillna(0) / 1000
    dr["_revenue_k"] = pd.to_numeric(dr[c["revenue_usd"]], errors="coerce").fillna(0) / 1000
    dr["_channel"] = dr["_opp_key"].map(channel_lookup).fillna("")
    # Robust classification match — trimmed + casefolded, so "core account",
    # "Core Account ", etc. all still match. A silent exact-match failure
    # here would make affected rows vanish from BOTH the SMA and NA
    # buckets without any warning, which is a very easy way for the
    # script's totals to quietly diverge from a manual Excel filter.
    dr["_class_norm"] = dr[c["customer_class"]].apply(lambda v: _norm_name(v))
    sma_norm = _norm_name(SMA_CLASSIFICATION_VALUE)
    na_norm = _norm_name(NA_CLASSIFICATION_VALUE)
    dr["_booking_k"] = pd.to_numeric(dr[c["booking_usd"]], errors="coerce").fillna(0) / 1000
    dr["_revenue_k"] = pd.to_numeric(dr[c["revenue_usd"]], errors="coerce").fillna(0) / 1000
    dr["_channel"] = dr["_opp_key"].map(channel_lookup).fillna("")

    # Uses ALL rows in the raw DR export, regardless of Signed Date — no
    # year restriction. The "PL 2026 (K$)" column label reflects the
    # sheet's naming convention, not an actual filter on signed year.
    dr_2026 = dr.copy()

    # Fixed partner list — always exactly NC_ACHIEVEMENT_PARTNERS, every row
    # appears whether or not it has 2026 DR activity or historical data.
    partners = list(NC_ACHIEVEMENT_PARTNERS)

    # Historical PL match check kept silent unless something's actually
    # missing — no need to print a full match report every run once this
    # is working correctly.
    if hist_pl is not None and "Partner" in hist_pl.columns:
        hist_keys = {_norm_name(p) for p in hist_pl["Partner"].dropna()}
        table_keys = {_norm_name(p) for p in partners}
        matched = table_keys & hist_keys
        if len(matched) < len(partners):
            unmatched_table = sorted(p for p in partners if _norm_name(p) not in hist_keys)
            print(f"NOTE: {len(unmatched_table)} partner(s) have no match in HISTORICAL_PL_FILE: "
                  f"{unmatched_table}")

    targets = read_csv_safe(TARGET_FILE_PATH) if Path(TARGET_FILE_PATH).exists() else pd.DataFrame()
    if not targets.empty:
        targets.columns = [str(cc).strip() for cc in targets.columns]

    t1_lookup = table1_report_df.set_index("Partner") if not table1_report_df.empty else pd.DataFrame()

    rows = []
    for partner in partners:
        involved = partner_involved_mask(dr_2026, partner)
        sub = dr_2026[involved]
        sma = sub[sub["_class_norm"] == sma_norm]
        na = sub[sub["_class_norm"] == na_norm]

        sma_channel = sma[sma["_channel"] == "Y"]
        sma_am = sma[sma["_channel"] != "Y"]

        booking_channel = sma_channel["_booking_k"].sum()
        booking_am = sma_am["_booking_k"].sum()
        booking_sma_total = booking_channel + booking_am
        booking_na = na["_booking_k"].sum()
        booking_pl_total = booking_sma_total + booking_na

        revenue_channel = sma_channel["_revenue_k"].sum()
        revenue_am = sma_am["_revenue_k"].sum()
        revenue_sma_total = revenue_channel + revenue_am
        revenue_na = na["_revenue_k"].sum()
        revenue_pl_total = revenue_sma_total + revenue_na

        # Status: prefer the historical PL sheet (has an entry for every
        # partner in this table by construction), fall back to Table 1's
        # targets file if blank/missing there.
        status = ""
        if hist_pl is not None and "Partner" in hist_pl.columns and "Status" in hist_pl.columns:
            hs_match = hist_pl[hist_pl["Partner"].apply(_norm_name) == _norm_name(partner)]
            if not hs_match.empty:
                status = hs_match.iloc[0].get("Status", "") or ""
        if not status and not targets.empty and "PartnerName" in targets.columns and "Status" in targets.columns:
            match = targets[targets["PartnerName"].apply(_norm_name) == _norm_name(partner)]
            if not match.empty:
                status = match.iloc[0]["Status"]

        pl_2023_b = pl_2024_b = pl_2025_b = pl_2023_r = pl_2024_r = pl_2025_r = 0
        if hist_pl is not None and "Partner" in hist_pl.columns:
            hmatch = hist_pl[hist_pl["Partner"].apply(_norm_name) == _norm_name(partner)]
            if not hmatch.empty:
                hr = hmatch.iloc[0]
                def _num(col):
                    v = pd.to_numeric(hr.get(col), errors="coerce")
                    return 0 if pd.isna(v) else v
                pl_2023_b = _num("PL Booking 2023 (K$)")
                pl_2024_b = _num("PL Booking 2024 (K$)")
                pl_2025_b = _num("PL Booking 2025 (K$)")
                pl_2023_r = _num("PL Revenue 2023 (K$)")
                pl_2024_r = _num("PL Revenue 2024 (K$)")
                pl_2025_r = _num("PL Revenue 2025 (K$)")

        tot_booking = tot_revenue = 0
        if not t1_lookup.empty:
            match_row = None
            for name in t1_lookup.index:
                if _norm_name(name) == _norm_name(partner):
                    match_row = t1_lookup.loc[name]
                    break
            if match_row is not None:
                tot_booking = pd.to_numeric(match_row.get(f"Total Orders {CURRENT_YEAR} (K$)", 0), errors="coerce") or 0
                tot_revenue = pd.to_numeric(match_row.get(f"Total Revenues {CURRENT_YEAR} (K$)", 0), errors="coerce") or 0

        hl_booking = max(tot_booking - booking_pl_total, 0) if tot_booking else 0
        hl_revenue = max(tot_revenue - revenue_pl_total, 0) if tot_revenue else 0
        pl_pct_booking = round(booking_pl_total / tot_booking * 100) if tot_booking else None
        pl_pct_revenue = round(revenue_pl_total / tot_revenue * 100) if tot_revenue else None

        rows.append({
            "Partner": partner, "Status": status,
            "PL Booking 2023 (K$)": round(pl_2023_b), "PL Booking 2024 (K$)": round(pl_2024_b),
            "PL Booking 2025 (K$)": round(pl_2025_b),
            "SMA PL Booking Channel (K$)": round(booking_channel), "SMA PL Booking AM (K$)": round(booking_am),
            "Total SMA PL Booking (K$)": round(booking_sma_total), "NA Booking (K$)": round(booking_na),
            "Total PL Booking (K$)": round(booking_pl_total), "PL % Booking": pl_pct_booking,
            "HL Booking (K$)": round(hl_booking), "Tot Booking (K$)": round(tot_booking),
            "PL Revenue 2023 (K$)": round(pl_2023_r), "PL Revenue 2024 (K$)": round(pl_2024_r),
            "PL Revenue 2025 (K$)": round(pl_2025_r),
            "SMA PL Revenue Channel (K$)": round(revenue_channel), "SMA PL Revenue AM (K$)": round(revenue_am),
            "Total SMA PL Revenue (K$)": round(revenue_sma_total), "NA Revenue (K$)": round(revenue_na),
            "Total PL Revenue (K$)": round(revenue_pl_total), "PL % Revenue": pl_pct_revenue,
            "HL Revenue (K$)": round(hl_revenue), "Tot Revenue (K$)": round(tot_revenue),
        })

    return pd.DataFrame(rows)


def write_nc_achievement_sheet(wb, nc_df):
    if nc_df is None or nc_df.empty:
        return
    ws = wb.create_sheet("NC Achievement")
    ws.sheet_view.showGridLines = False

    header_font = Font(bold=True, color="FFFFFF", size=9)
    banner_fill = PatternFill(start_color=HUAWEI_RED, end_color=HUAWEI_RED, fill_type="solid")
    section_fill = PatternFill(start_color="7A1414", end_color="7A1414", fill_type="solid")  # dark red, distinct from the bright banner
    col_header_fill = PatternFill(start_color="FBE4E4", end_color="FBE4E4", fill_type="solid")  # light red tint, matches Channel Target Report
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="E8A0A0")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)

    # Row 1: banner
    c1 = ws["A1"]
    c1.value = f"Data Update Time:  {TODAY.strftime('%Y-%m-%d %H:%M:%S')}"
    c1.font = Font(bold=True, size=9, color="666666")

    def block(start_col_idx):
        """Writes one Booking/Revenue block's headers starting at
        start_col_idx (1-based). Returns the column index after the block."""
        cols = [("PL 2023 (K$)", 1), ("PL 2024 (K$)", 1), ("PL 2025 (K$)", 1)]
        i = start_col_idx
        for label, span in cols:
            ws.cell(row=4, column=i, value=label)
            ws.merge_cells(start_row=4, start_column=i, end_row=6, end_column=i)
            i += 1
        pl2026_start = i
        ws.cell(row=4, column=i, value=f"PL {CURRENT_YEAR} (K$)")
        ws.merge_cells(start_row=4, start_column=pl2026_start, end_row=4, end_column=pl2026_start + 5)
        ws.cell(row=5, column=i, value="SMA PL")
        ws.merge_cells(start_row=5, start_column=i, end_row=5, end_column=i + 1)
        ws.cell(row=6, column=i, value="Channel")
        ws.cell(row=6, column=i + 1, value="AM")
        for j, label in enumerate(["Total SMA PL", "NA", "Total PL", "PL %"]):
            col = i + 2 + j
            ws.cell(row=5, column=col, value=label)
            ws.merge_cells(start_row=5, start_column=col, end_row=6, end_column=col)
        i += 6
        ws.cell(row=4, column=i, value=f"HL {CURRENT_YEAR} (K$)")
        ws.merge_cells(start_row=4, start_column=i, end_row=6, end_column=i)
        i += 1
        return i  # next available column, for the caller to write the Tot label into

    # Row 4/5/6: detailed headers
    ws.merge_cells("A4:A6"); ws["A4"].value = "Partner"
    ws.merge_cells("B4:B6"); ws["B4"].value = "Status"

    booking_start = 3
    tot_booking_col = block(booking_start)
    ws.cell(row=4, column=tot_booking_col, value="Tot Booking")
    ws.merge_cells(start_row=4, start_column=tot_booking_col, end_row=6, end_column=tot_booking_col)
    booking_end = tot_booking_col

    revenue_start = tot_booking_col + 1
    tot_revenue_col = block(revenue_start)
    ws.cell(row=4, column=tot_revenue_col, value="Tot Revenue")
    ws.merge_cells(start_row=4, start_column=tot_revenue_col, end_row=6, end_column=tot_revenue_col)
    revenue_end = tot_revenue_col
    last_col = tot_revenue_col

    # Row 2/3 banners — computed from the ACTUAL block boundaries above, not
    # hardcoded, so "PL 2023 (K$)"/"HL .../"Tot Revenue" always land under
    # the correct Booking/Revenue banner regardless of column count.
    ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=last_col)
    top_banner = ws.cell(row=2, column=3, value=f"NC achievement [Source: DR Report system]  \u00b7  Achievement {CURRENT_YEAR}")
    top_banner.font = Font(bold=True, size=12, color="FFFFFF")
    top_banner.fill = banner_fill
    top_banner.alignment = center

    ws.merge_cells(start_row=3, start_column=booking_start, end_row=3, end_column=booking_end)
    ws.cell(row=3, column=booking_start, value="Booking k USD")
    ws.merge_cells(start_row=3, start_column=revenue_start, end_row=3, end_column=revenue_end)
    ws.cell(row=3, column=revenue_start, value="Revenue k USD")
    for col in (booking_start, revenue_start):
        cell = ws.cell(row=3, column=col)
        cell.font = header_font
        cell.fill = section_fill
        cell.alignment = center

    for row in ws.iter_rows(min_row=1, max_row=6, max_col=last_col):
        for cell in row:
            if cell.value is not None:
                if cell.row >= 4:
                    cell.font = Font(bold=True, size=8, color="7A1414")
                    cell.fill = col_header_fill
                cell.alignment = center
                cell.border = border
    ws.row_dimensions[4].height = 26
    ws.row_dimensions[5].height = 16
    ws.row_dimensions[6].height = 16

    # Data rows
    data_cols = list(nc_df.columns)
    body_font = Font(size=9, color="7A1414")
    band_fill = PatternFill(start_color="FBE4E4", end_color="FBE4E4", fill_type="solid")
    r = 7
    for _, dr_row in nc_df.iterrows():
        for i, val in enumerate(dr_row, start=1):
            cell = ws.cell(row=r, column=i, value=val)
            cell.font = body_font
            cell.border = border
            cell.alignment = Alignment(horizontal="center")
            if r % 2 == 0:
                cell.fill = band_fill
        r += 1

    # TOTAL row — sums every $ column across all partners. PL % columns are
    # recomputed from the summed totals rather than summed themselves
    # (summing percentages isn't meaningful).
    total_row_num = r
    total_font = Font(bold=True, size=9, color="FFFFFF")
    total_fill = PatternFill(start_color=HUAWEI_RED, end_color=HUAWEI_RED, fill_type="solid")
    tot_pl_booking = nc_df["Total PL Booking (K$)"].sum()
    tot_booking_all = nc_df["Tot Booking (K$)"].sum()
    tot_pl_revenue = nc_df["Total PL Revenue (K$)"].sum()
    tot_revenue_all = nc_df["Tot Revenue (K$)"].sum()
    total_values = {"Partner": "TOTAL", "Status": ""}
    for col in data_cols:
        if col in ("Partner", "Status"):
            continue
        elif col == "PL % Booking":
            total_values[col] = round(tot_pl_booking / tot_booking_all * 100) if tot_booking_all else None
        elif col == "PL % Revenue":
            total_values[col] = round(tot_pl_revenue / tot_revenue_all * 100) if tot_revenue_all else None
        else:
            total_values[col] = round(nc_df[col].sum())
    for i, col in enumerate(data_cols, start=1):
        cell = ws.cell(row=total_row_num, column=i, value=total_values[col])
        cell.font = total_font
        cell.fill = total_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    for i in range(1, last_col + 1):
        ws.column_dimensions[get_column_letter(i)].width = 9
    ws.column_dimensions["A"].width = 26  # partner names — wide enough that none get cut off
    ws.column_dimensions["B"].width = 12
    ws.freeze_panes = "C7"
    ws.sheet_view.zoomScale = 70  # whole table visible without horizontal scrolling

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:{get_column_letter(last_col)}{total_row_num}"


COMPARISON_SNAPSHOT_DIR = "weekly_comparison_snapshots"

# Key metrics compared week-over-week for each table — kept to a focused
# set rather than every column, so the comparison sheet stays readable.
CHANNEL_TARGET_COMPARE_METRICS = [
    f"Total Orders {CURRENT_YEAR} (K$)", f"Total Revenues {CURRENT_YEAR} (K$)",
    "Order Achievement %", "Revenue Achievement %",
]
NC_ACHIEVEMENT_COMPARE_METRICS = [
    "Total PL Booking (K$)", "Total PL Revenue (K$)", "Tot Booking (K$)", "Tot Revenue (K$)",
]


def save_comparison_snapshots(report_df, nc_df):
    folder = Path(COMPARISON_SNAPSHOT_DIR)
    folder.mkdir(exist_ok=True)
    date_str = TODAY.strftime("%Y-%m-%d")
    report_df.to_csv(folder / f"channel_target_{date_str}.csv", index=False)
    if nc_df is not None and not nc_df.empty:
        nc_df.to_csv(folder / f"nc_achievement_{date_str}.csv", index=False)
    print(f"Saved this week's comparison snapshot(s) to {COMPARISON_SNAPSHOT_DIR}/ "
          f"for next week's Weekly Changes comparison.")


def load_previous_comparison_snapshot(prefix):
    """Finds the most recent snapshot strictly BEFORE today with this
    prefix. Returns (DataFrame, date_string) or (None, None) if there
    isn't one yet (e.g. first run)."""
    folder = Path(COMPARISON_SNAPSHOT_DIR)
    if not folder.exists():
        return None, None
    today_str = TODAY.strftime("%Y-%m-%d")
    candidates = sorted(
        f for f in folder.glob(f"{prefix}_*.csv")
        if f.stem.replace(f"{prefix}_", "") < today_str
    )
    if not candidates:
        return None, None
    latest = candidates[-1]
    date_str = latest.stem.replace(f"{prefix}_", "")
    return read_csv_safe(latest), date_str


def write_weekly_changes_sheet(wb, report_df, nc_df, prev_report_df, prev_report_date,
                                prev_nc_df, prev_nc_date):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    ws = wb.create_sheet("Weekly Changes")
    ws.sheet_view.showGridLines = False
    header_font = Font(bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill(start_color=HUAWEI_RED, end_color=HUAWEI_RED, fill_type="solid")
    increase_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    decrease_fill = PatternFill(start_color="F4B6B6", end_color="F4B6B6", fill_type="solid")
    increase_font = Font(bold=True, color="1E7B34", size=9)
    decrease_font = Font(bold=True, color="A61B1B", size=9)
    thin = Side(style="thin", color="DDDDDD")
    border = Border(top=thin, bottom=thin, left=thin, right=thin)

    def write_section(title, cur_df, prev_df, prev_date, metrics, start_row):
        r = start_row
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3 + len(metrics))
        c = ws.cell(row=r, column=1, value=title if prev_df is None else
                    f"{title}  (vs. {prev_date})")
        c.font = Font(bold=True, size=12, color="FFFFFF")
        c.fill = header_fill
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        r += 1

        if prev_df is None:
            ws.cell(row=r, column=1,
                     value="No prior snapshot yet — comparison will appear starting next run.").font = \
                Font(italic=True, color="808080", size=10)
            return r + 2

        headers = ["Partner"]
        for m in metrics:
            headers += [m, f"{m} (prev)", f"{m} \u0394"]
        for i, h in enumerate(headers, start=1):
            cell = ws.cell(row=r, column=i, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        header_row = r
        r += 1

        prev_indexed = prev_df.set_index(prev_df["Partner"].apply(_norm_name))
        for _, row in cur_df.iterrows():
            partner = row["Partner"]
            ws.cell(row=r, column=1, value=partner).border = border
            col = 2
            prev_row = prev_indexed.loc[_norm_name(partner)] if _norm_name(partner) in prev_indexed.index else None
            for m in metrics:
                cur_val = pd.to_numeric(row.get(m), errors="coerce")
                prev_val = pd.to_numeric(prev_row.get(m), errors="coerce") if prev_row is not None else None
                delta = (cur_val - prev_val) if (pd.notna(cur_val) and prev_val is not None and pd.notna(prev_val)) else None

                cell_cur = ws.cell(row=r, column=col, value=cur_val if pd.notna(cur_val) else None)
                cell_prev = ws.cell(row=r, column=col + 1, value=prev_val if prev_val is not None and pd.notna(prev_val) else None)
                cell_delta = ws.cell(row=r, column=col + 2, value=round(delta, 1) if delta is not None else None)
                for cc in (cell_cur, cell_prev, cell_delta):
                    cc.border = border
                    cc.alignment = Alignment(horizontal="center")
                if delta is not None and abs(delta) > 0.01:
                    fill = increase_fill if delta > 0 else decrease_fill
                    font = increase_font if delta > 0 else decrease_font
                    cell_delta.fill = fill
                    cell_delta.font = font
                col += 3
            r += 1

        for i in range(1, len(headers) + 1):
            ws.column_dimensions[get_column_letter(i)].width = 14
        ws.column_dimensions["A"].width = 24
        return r + 2

    next_row = write_section("Channel Target Report \u2014 Weekly Changes", report_df,
                              prev_report_df, prev_report_date, CHANNEL_TARGET_COMPARE_METRICS, 2)
    write_section("NC Achievement \u2014 Weekly Changes", nc_df, prev_nc_df, prev_nc_date,
                  NC_ACHIEVEMENT_COMPARE_METRICS, next_row)

    ws.sheet_view.zoomScale = 75
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def build_weekly_changes_dashboard(wb, report_df, nc_df, prev_report_df, prev_nc_df):
    """Appends a 'Weekly Changes' KPI section onto the bottom of the existing
    Dashboard sheet — same sheet, not a separate tab — placed safely below
    the main dashboard's own charts."""
    from openpyxl.styles import Font, PatternFill, Alignment

    ws = wb["Dashboard"]
    WEEKLY_SECTION_ROW = 62  # well clear of the main dashboard's charts, which end around row 57

    ws.merge_cells(start_row=WEEKLY_SECTION_ROW, start_column=2, end_row=WEEKLY_SECTION_ROW, end_column=16)
    title = ws.cell(row=WEEKLY_SECTION_ROW, column=2)
    title.value = "Weekly Changes"
    title.font = Font(size=18, bold=True, color="FFFFFF")
    title.fill = PatternFill(start_color=HUAWEI_RED, end_color=HUAWEI_RED, fill_type="solid")
    title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[WEEKLY_SECTION_ROW].height = 32

    kpi_row = WEEKLY_SECTION_ROW + 3

    if prev_report_df is None:
        ws.cell(row=kpi_row, column=2, value=(
            "No prior snapshot yet — this section will populate starting next run, "
            "once there's a previous week to compare against."
        )).font = Font(italic=True, color="808080", size=11)
        return

    def delta_sum(cur_df, prev_df, metric):
        prev_indexed = prev_df.set_index(prev_df["Partner"].apply(_norm_name))
        total = 0
        for _, row in cur_df.iterrows():
            key = _norm_name(row["Partner"])
            if key in prev_indexed.index:
                cur_val = pd.to_numeric(row.get(metric), errors="coerce")
                prev_val = pd.to_numeric(prev_indexed.loc[key].get(metric), errors="coerce")
                if pd.notna(cur_val) and pd.notna(prev_val):
                    total += (cur_val - prev_val)
        return total

    order_delta = delta_sum(report_df, prev_report_df, f"Total Orders {CURRENT_YEAR} (K$)")
    revenue_delta = delta_sum(report_df, prev_report_df, f"Total Revenues {CURRENT_YEAR} (K$)")
    nc_booking_delta = delta_sum(nc_df, prev_nc_df, "Total PL Booking (K$)") if (nc_df is not None and prev_nc_df is not None) else None
    nc_revenue_delta = delta_sum(nc_df, prev_nc_df, "Total PL Revenue (K$)") if (nc_df is not None and prev_nc_df is not None) else None

    specs = [
        ("ORDERS CHANGE (K$)", f"{order_delta:+,.0f}", GREEN_UP if order_delta >= 0 else RED_DOWN),
        ("REVENUE CHANGE (K$)", f"{revenue_delta:+,.0f}", GREEN_UP if revenue_delta >= 0 else RED_DOWN),
    ]
    if nc_booking_delta is not None:
        specs.append(("NC PL BOOKING CHANGE (K$)", f"{nc_booking_delta:+,.0f}",
                       GREEN_UP if nc_booking_delta >= 0 else RED_DOWN))
        specs.append(("NC PL REVENUE CHANGE (K$)", f"{nc_revenue_delta:+,.0f}",
                       GREEN_UP if nc_revenue_delta >= 0 else RED_DOWN))

    col = 2
    for label, value, accent in specs:
        kpi_card(ws, col, kpi_row, label, value, accent)
        col += 3

    ws.print_area = "A1:X90"  # extended to cover the new Weekly Changes section


def load_previous_manual_fields(path):
    """Reads Status and Channel Target (M$) from the script's own previous
    output (before it gets overwritten) — these are the two fields in this
    table most plausibly hand-corrected, so a manual edit here carries
    forward into the next run instead of being silently overwritten."""
    p = Path(path)
    if not p.exists():
        return {}
    try:
        prev = pd.read_excel(p, sheet_name="Channel Target Report")
    except Exception as e:
        print(f"NOTE: couldn't read previous output '{path}' for manual-value carry-forward: {e}")
        return {}
    prev.columns = [str(c).strip() for c in prev.columns]
    if "Partner" not in prev.columns:
        return {}
    result = {}
    for _, r in prev.iterrows():
        partner = str(r.get("Partner", "")).strip()
        if not partner or partner.upper() == "TOTAL":
            continue
        result[_norm_name(partner)] = {
            "Status": r.get("Status"),
            "Channel Target (M$)": r.get("Channel Target (M$)"),
        }
    return result


def apply_weekly_change_coloring(ws, cur_df, prev_df, metrics, data_start_row):
    """Tints each metric cell directly (not a separate delta column) green
    if it increased vs the previous snapshot, red if it decreased. Column
    positions come straight from the DataFrame's own column order, which
    is exactly the order these sheets were written in."""
    if prev_df is None or cur_df is None or cur_df.empty:
        return
    from openpyxl.styles import PatternFill, Font
    increase_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    decrease_fill = PatternFill(start_color="F4B6B6", end_color="F4B6B6", fill_type="solid")
    increase_font = Font(bold=True, color="1E7B34", size=ws.cell(row=data_start_row, column=1).font.size or 10)
    decrease_font = Font(bold=True, color="A61B1B", size=ws.cell(row=data_start_row, column=1).font.size or 10)

    prev_indexed = prev_df.set_index(prev_df["Partner"].apply(_norm_name))
    col_indices = {m: list(cur_df.columns).index(m) + 1 for m in metrics if m in cur_df.columns}

    for row_offset, (_, row) in enumerate(cur_df.iterrows()):
        r = data_start_row + row_offset
        key = _norm_name(row["Partner"])
        if key not in prev_indexed.index:
            continue
        prev_row = prev_indexed.loc[key]
        for m, col_idx in col_indices.items():
            cur_val = pd.to_numeric(row.get(m), errors="coerce")
            prev_val = pd.to_numeric(prev_row.get(m), errors="coerce")
            if pd.notna(cur_val) and pd.notna(prev_val):
                delta = cur_val - prev_val
                if abs(delta) > 0.01:
                    cell = ws.cell(row=r, column=col_idx)
                    cell.fill = increase_fill if delta > 0 else decrease_fill
                    cell.font = increase_font if delta > 0 else decrease_font


def main():
    print("Building Channel Target Report...")
    report_df = build_report()

    previous_manual = load_previous_manual_fields(OUTPUT_PATH)
    manual_carried_count = 0
    for idx, row in report_df.iterrows():
        prev_vals = previous_manual.get(_norm_name(row["Partner"]))
        if not prev_vals:
            continue
        for field in ("Status", "Channel Target (M$)"):
            prev_val = prev_vals.get(field)
            if prev_val is not None and pd.notna(prev_val) and str(prev_val).strip() and str(prev_val) != str(row.get(field)):
                report_df.at[idx, field] = prev_val
                manual_carried_count += 1
    if manual_carried_count:
        print(f"Carried forward {manual_carried_count} manually-edited Status/Channel Target "
              f"value(s) from the previous run's output.")

    prev_report_df, prev_report_date = load_previous_comparison_snapshot("channel_target")
    prev_nc_df, prev_nc_date = load_previous_comparison_snapshot("nc_achievement")

    print(f"Writing {OUTPUT_PATH}...")
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Channel Target Report"
    ws.append(list(report_df.columns))
    for _, r in report_df.iterrows():
        ws.append(list(r.values))

    total_row = build_total_row(report_df)
    ws.append([total_row[col] for col in report_df.columns])
    total_row_num = ws.max_row

    print("Applying report formatting...")
    style_report(ws)
    apply_weekly_change_coloring(ws, report_df, prev_report_df, CHANNEL_TARGET_COMPARE_METRICS,
                                  data_start_row=2)

    # Highlight the TOTAL row in Huawei red, overriding the banded styling
    # style_report just applied to every data row.
    for cell in ws[total_row_num]:
        cell.fill = PatternFill(start_color=HUAWEI_RED, end_color=HUAWEI_RED, fill_type="solid")
        cell.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    print("Building dashboard...")
    build_dashboard(wb, report_df)

    print("Building NC Achievement table...")
    nc_df = build_nc_achievement(report_df)
    write_nc_achievement_sheet(wb, nc_df)
    apply_weekly_change_coloring(wb["NC Achievement"], nc_df, prev_nc_df, NC_ACHIEVEMENT_COMPARE_METRICS,
                                  data_start_row=7)

    print("Building weekly comparison...")
    write_weekly_changes_sheet(wb, report_df, nc_df, prev_report_df, prev_report_date,
                                prev_nc_df, prev_nc_date)
    build_weekly_changes_dashboard(wb, report_df, nc_df, prev_report_df, prev_nc_df)

    wb.save(OUTPUT_PATH)
    save_comparison_snapshots(report_df, nc_df)

    print("Done.")
    print(f"{len(report_df)} partners in the report.")


if __name__ == "__main__":
    main()