import pandas as pd
from pathlib import Path
from datetime import datetime

# ============ CONFIG ============

CURRENT_FILE = "esales_current.xlsx"
PREVIOUS_FILE = "esales_previous.xlsx"
DR_FILE = "dr_export.xlsx"  # Deal Registration export — grey = opportunity has NO entry here
OUTPUT_PATH = "esales_report.xlsx"
PRESERVE_MANUAL_T1_T2 = True

TMI_KEYWORDS = ["Partner Q5", "Partner S27"]
TMI_CANONICAL_NAME = "Partner Q5"

# ============ Ongoing Pipe table config ============
TODAY = datetime.today()
CURRENT_YEAR = TODAY.year
CURRENT_QUARTER = (TODAY.month - 1) // 3 + 1  # 1-4

PIPE_PARTNERS = [
    "Partner Q5",
    "Partner T15",
    "Partner C10",
    "Partner G22",
    "Partner L13",
    "Partner P30",
    "Partner I5",
    "Partner M21",
    "Partner L1",
    "Partner H3",
]

PIPE_PARTNERS_MATCHED_ON_T1 = {"Partner Q5", "Partner T15", "Partner C10"}
PIPE_PARTNER_STATUS = {
    "Partner Q5": "Diamond",
    "Partner T15": "Distributor",
    "Partner C10": "Distributor",
    "Partner G22": "Gold",
    "Partner L13": "Gold",
    "Partner P30": "Silver",
    "Partner I5": "Authorized",
    "Partner M21": "Authorized",
    "Partner L1": "Authorized",
    "Partner H3": "Authorized",
}


def pipe_match_column(partner):
    return "T1 Partner" if partner in PIPE_PARTNERS_MATCHED_ON_T1 else "T2 Partner"

PIPE_SNAPSHOT_DIR = "esales_pipe_snapshots"
PIPE_SOURCE_PL_VALUE = "hdr"
PIPE_SOURCE_HL_VALUE = "cc4"

HUAWEI_RED = "C8102E"
GREY_FILL = "D9D9D9"
KNOWN_GREY_HEXES = ["E8E8E8"]
KNOWN_RED_HEXES = []
OFFICE_THEME = ["000000", "FFFFFF", "44546A", "E7E6E6", "4472C4", "ED7D31",
                "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"]


def apply_tint(hex_rgb, tint):
    """Excel's tint algorithm: darkens (tint<0) or lightens (tint>0) a base
    color. This is what turns theme color 'Background 1' + a tint into
    something like E8E8E8."""
    r, g, b = int(hex_rgb[0:2], 16), int(hex_rgb[2:4], 16), int(hex_rgb[4:6], 16)
    def adj(c):
        if tint < 0:
            return round(c * (1 + tint))
        return round(c * (1 - tint) + 255 * tint)
    return f"{adj(r):02X}{adj(g):02X}{adj(b):02X}"


def resolve_fill_hex(fgColor):
    """Best-effort: returns a 6-char hex string for a cell's fill color,
    whichever way Excel actually stored it (raw RGB, theme+tint, or an
    indexed palette color)."""
    if fgColor is None:
        return None
    if fgColor.type == "rgb" and fgColor.rgb and fgColor.rgb != "00000000":
        return fgColor.rgb[-6:]
    if fgColor.type == "theme":
        try:
            base = OFFICE_THEME[fgColor.theme]
        except (IndexError, TypeError):
            return None
        tint = fgColor.tint or 0
        return apply_tint(base, tint) if tint else base
    if fgColor.type == "indexed":
        from openpyxl.styles.colors import COLOR_INDEX
        try:
            idx_hex = COLOR_INDEX[fgColor.indexed]
            return idx_hex[-6:] if idx_hex else None
        except (IndexError, TypeError):
            return None
    return None


def read_any(path, **kwargs):
    path = str(path)
    if path.lower().endswith(".csv"):
        try:
            return pd.read_csv(path, encoding="utf-8", **kwargs)
        except UnicodeDecodeError:
            return pd.read_csv(path, encoding="cp1252", **kwargs)
    return pd.read_excel(path, **kwargs)


def clean_text(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def norm_key(v):
    return " ".join(clean_text(v).casefold().split())


def contains_tmi(text):
    t = text.casefold()
    return any(kw in t for kw in TMI_KEYWORDS)


def has_multiple_partners(text):
    return "," in text


# ============ Load previous week (with colors) ============

def classify_fill(hex_rgb):
    """Buckets a resolved 6-char hex color into 'red', 'grey', or None.
    Checks your confirmed exact colors first, then falls back to an RGB
    heuristic for anything not yet confirmed."""
    if not hex_rgb:
        return None
    hex_rgb = hex_rgb.upper()
    if hex_rgb in [h.upper() for h in KNOWN_GREY_HEXES]:
        return "grey"
    if hex_rgb in [h.upper() for h in KNOWN_RED_HEXES]:
        return "red"
    try:
        r, g, b = int(hex_rgb[0:2], 16), int(hex_rgb[2:4], 16), int(hex_rgb[4:6], 16)
    except ValueError:
        return None
    if r > 170 and g < 110 and b < 110:
        return "red"
    if abs(r - g) < 25 and abs(g - b) < 25 and abs(r - b) < 25 and 120 < r < 245:
        return "grey"
    return None


def load_previous_week(path):
    """Returns a dict keyed by normalized e-Opp No. -> {T1, T2, color}.
    Reads via openpyxl (not pandas) specifically to get cell fill colors,
    which pandas doesn't preserve."""
    import openpyxl
    p = Path(path)
    if not p.exists():
        print(f"NOTE: PREVIOUS_FILE '{path}' not found — this will be treated "
              f"as the first week: no color carry-forward, no comma-partner "
              f"fallback, and every opportunity will show as 'new'.")
        return None

    wb = openpyxl.load_workbook(p)
    ws = wb.active
    header_row_num = None
    for candidate_row in (1, 2, 3):
        candidate_headers = [c.value for c in ws[candidate_row]]
        if "e-Opp No." in candidate_headers:
            header_row_num = candidate_row
            headers = candidate_headers
            break
    if header_row_num is None:
        print(f"WARNING: PREVIOUS_FILE has no 'e-Opp No.' column in its first few rows — "
              f"can't read it for carry-forward.")
        return None
    try:
        opp_col = headers.index("e-Opp No.") + 1
        t1_col = headers.index("T1 Partner") + 1
        t2_col = headers.index("T2 Partner") + 1
    except ValueError as e:
        print(f"WARNING: PREVIOUS_FILE is missing an expected column: {e}")
        return None

    color_counts = {}
    result = {}
    for row in ws.iter_rows(min_row=header_row_num + 1):
        opp_no = clean_text(row[opp_col - 1].value)
        if not opp_no:
            continue
        color = None
        raw_color = None  # exact hex, kept even if it's not red/grey — lets
                           # an arbitrary manually-applied color carry forward
        for cell in row:
            fill = cell.fill
            resolved_hex = resolve_fill_hex(fill.fgColor) if fill else None
            if resolved_hex:
                color_counts[resolved_hex] = color_counts.get(resolved_hex, 0) + 1
                if raw_color is None:
                    raw_color = resolved_hex
            if color is None:
                c = classify_fill(resolved_hex)
                if c:
                    color = c
        result[norm_key(opp_no)] = {
            "T1": clean_text(row[t1_col - 1].value),
            "T2": clean_text(row[t2_col - 1].value),
            "color": color,
            "raw_color": raw_color,
        }

    print(f"Previous week loaded: {len(result)} opportunities.")
    print(f"  Distinct cell fill colors found (hex: count): {color_counts}")
    print(f"  (classified as red/grey using an RGB heuristic — check this "
          f"list against what you know the real red/grey fills to be)")
    return result


def load_current_output_colors(path):
    """Reads ONLY color data (not T1/T2, not used for new-opportunity
    detection) from the script's own most recent output. This is what
    lets a color you manually applied minutes ago survive an immediate
    re-run, even before you've promoted this week's file to
    PREVIOUS_FILE. Returns {} if the file doesn't exist yet."""
    import openpyxl
    p = Path(path)
    if not p.exists():
        return {}
    wb = openpyxl.load_workbook(p)
    ws = wb.active
    header_row_num = None
    headers = None
    for candidate_row in (1, 2, 3):
        candidate_headers = [c.value for c in ws[candidate_row]]
        if "e-Opp No." in candidate_headers:
            header_row_num = candidate_row
            headers = candidate_headers
            break
    if header_row_num is None:
        return {}
    opp_col = headers.index("e-Opp No.") + 1

    result = {}
    for row in ws.iter_rows(min_row=header_row_num + 1):
        opp_no = clean_text(row[opp_col - 1].value)
        if not opp_no:
            continue
        color = None
        raw_color = None
        for cell in row:
            fill = cell.fill
            resolved_hex = resolve_fill_hex(fill.fgColor) if fill else None
            if resolved_hex and raw_color is None:
                raw_color = resolved_hex
            if color is None:
                c = classify_fill(resolved_hex)
                if c:
                    color = c
        result[norm_key(opp_no)] = {"color": color, "raw_color": raw_color}
    return result


# ============ Load Deal Registration export (drives grey) ============

def load_dr_opportunities(path):
    """Returns a set of normalized Opportunity No. values from the DR
    export. An eSales opportunity NOT in this set gets marked grey — no
    matching Deal Registration exists for it. Returns None if not
    configured / file not found (grey is then simply never applied)."""
    if not path:
        return None
    p = Path(path)
    if not p.exists():
        print(f"NOTE: DR_FILE '{path}' not found — grey (no matching DR) won't be applied.")
        return None
    dr = read_any(p)
    dr.columns = [str(c).strip() for c in dr.columns]
    if "Opportunity No." not in dr.columns:
        print(f"WARNING: DR_FILE has no 'Opportunity No.' column — found: {dr.columns.tolist()}")
        return None
    opp_nos = {norm_key(v) for v in dr["Opportunity No."].apply(clean_text) if v}
    print(f"DR export loaded: {len(opp_nos)} unique opportunity number(s) with a Deal Registration.")
    return opp_nos


# ============ Load current week ============

def load_current_week(path):
    df = read_any(path)
    df.columns = [str(c).strip() for c in df.columns]
    text_cols = ["e-Opp No.", "Opportunity Name", "Country/Region", "Sales Stage",
                 "Named Account Type", "End Customer", "T1 Partner", "T2 Partner",
                 "Opportunity Source", "Opportunity Owner", "Tag", "Action"]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].apply(clean_text)
    return df


# ============ Build report ============

def classify_pl_hl_source(opportunity_source):
    """PL if Opportunity Source is 'HDR', HL if 'CC4' — case-insensitive,
    trimmed. Returns None for anything else (not counted as PL or HL)."""
    v = str(opportunity_source or "").strip().casefold()
    if v == PIPE_SOURCE_PL_VALUE:
        return "PL"
    if v == PIPE_SOURCE_HL_VALUE:
        return "HL"
    return None


def active_pipe_quarters():
    """Quarters shown on the Ongoing Pipe table — from the CURRENT quarter
    through Q4 of the current year. E.g. if today is in Q3, shows [3, 4].
    If today is in Q1, shows [1, 2, 3, 4]."""
    return list(range(CURRENT_QUARTER, 5))


def build_pipe_table(report_df):
    """Builds the Ongoing 2026 PIPE table — for each of the PIPE_PARTNERS,
    sums Estimated Order Amount(USD) into PL/HL buckets per quarter (based
    on Estimated Order Date), for each quarter from the current one through
    Q4. Matched by T1 Partner. Returns (DataFrame, quarters_list)."""
    df = report_df.copy()
    df["_parsed_date"] = df["Estimated Order Date"].apply(parse_order_date)
    df["_amount_k"] = pd.to_numeric(df["Estimated Order Amount(USD)"], errors="coerce").fillna(0) / 1000
    df["_pl_hl"] = df["Opportunity Source"].apply(classify_pl_hl_source)

    other_sources = sorted(df.loc[
        df["_pl_hl"].isna() & df["_parsed_date"].notna() &
        (df["_parsed_date"].apply(lambda d: d.year if d else None) == CURRENT_YEAR),
        "Opportunity Source"
    ].apply(lambda v: str(v).strip()).unique())
    other_sources = [s for s in other_sources if s]
    if other_sources:
        print(f"NOTE: {len(other_sources)} Opportunity Source value(s) in {CURRENT_YEAR} are "
              f"neither '{PIPE_SOURCE_PL_VALUE.upper()}' nor '{PIPE_SOURCE_HL_VALUE.upper()}' — "
              f"not counted as PL or HL in the pipe table: {other_sources}")

    quarters = active_pipe_quarters()
    rows = []
    for partner in PIPE_PARTNERS:
        match_col = pipe_match_column(partner)
        sub = df[df[match_col] == partner]
        row = {"Partner": partner, "Status": PIPE_PARTNER_STATUS.get(partner, "")}
        tot_year_pl = tot_year_hl = 0
        for q in quarters:
            if sub.empty:
                # Applying .apply() to an already-empty DataFrame's column
                # here silently collapses it to zero columns (a real pandas
                # quirk) — skip straight to zero instead of hitting that.
                pl_amt = hl_amt = 0
            else:
                q_sub = sub[sub["_parsed_date"].apply(
                    lambda d: d is not None and d.year == CURRENT_YEAR and (d.month - 1) // 3 + 1 == q)]
                pl_amt = q_sub.loc[q_sub["_pl_hl"] == "PL", "_amount_k"].sum()
                hl_amt = q_sub.loc[q_sub["_pl_hl"] == "HL", "_amount_k"].sum()
            row[f"Q{q} PL"] = round(pl_amt)
            row[f"Q{q} HL"] = round(hl_amt)
            row[f"Q{q} Tot"] = round(pl_amt + hl_amt)
            tot_year_pl += pl_amt
            tot_year_hl += hl_amt
        row["Tot PL"] = round(tot_year_pl)
        row["Tot HL"] = round(tot_year_hl)
        row["Tot Tot"] = round(tot_year_pl + tot_year_hl)
        rows.append(row)

    return pd.DataFrame(rows), quarters


def save_pipe_snapshot(pipe_df):
    folder = Path(PIPE_SNAPSHOT_DIR)
    folder.mkdir(exist_ok=True)
    path = folder / f"pipe_{TODAY.strftime('%Y-%m-%d')}.csv"
    pipe_df.to_csv(path, index=False)


def load_previous_pipe_snapshot():
    folder = Path(PIPE_SNAPSHOT_DIR)
    if not folder.exists():
        return None
    today_str = TODAY.strftime("%Y-%m-%d")
    candidates = sorted(f for f in folder.glob("pipe_*.csv") if f.stem.replace("pipe_", "") < today_str)
    if not candidates:
        return None
    try:
        return read_any(candidates[-1])
    except Exception:
        return pd.read_csv(candidates[-1])


def write_pipe_sheet(wb, pipe_df, quarters):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    ws = wb.create_sheet("Ongoing Pipe")
    ws.sheet_view.showGridLines = False
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color=HUAWEI_RED, end_color=HUAWEI_RED, fill_type="solid")
    sub_fill = PatternFill(start_color="7A1414", end_color="7A1414", fill_type="solid")
    col_fill = PatternFill(start_color="FBE4E4", end_color="FBE4E4", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="E8A0A0")
    thick = Side(style="medium", color=HUAWEI_RED)
    border = Border(top=thin, bottom=thin, left=thin, right=thin)

    ws.merge_cells("A1:B1")
    ws["A1"] = f"Report Generated: {TODAY.strftime('%Y-%m-%d %H:%M:%S')}"
    ws["A1"].font = Font(italic=True, size=9, color="666666")
    ws.row_dimensions[1].height = 18

    ws.merge_cells(start_row=2, start_column=1, end_row=4, end_column=1)
    ws.cell(row=2, column=1, value="Partner")
    ws.merge_cells(start_row=2, start_column=2, end_row=4, end_column=2)
    ws.cell(row=2, column=2, value="Status")

    n_groups = len(quarters) + 1  # + the "Tot {YEAR}" group
    last_col = 2 + n_groups * 3
    ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=last_col)
    ws.cell(row=2, column=3, value=f"Ongoing {CURRENT_YEAR} PIPE  (K$)")

    # Column index where each quarter-group (Q3, Q4, ..., Tot) starts —
    # used below to draw a thicker separator border between groups.
    group_start_cols = []
    col = 3
    for q in quarters:
        group_start_cols.append(col)
        ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col + 2)
        ws.cell(row=3, column=col, value=f"Q{q} {CURRENT_YEAR}")
        ws.cell(row=4, column=col, value="PL")
        ws.cell(row=4, column=col + 1, value="HL")
        ws.cell(row=4, column=col + 2, value=f"Tot {CURRENT_YEAR}")
        col += 3
    group_start_cols.append(col)
    ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col + 2)
    ws.cell(row=3, column=col, value=f"Tot {CURRENT_YEAR}")
    ws.cell(row=4, column=col, value="PL")
    ws.cell(row=4, column=col + 1, value="HL")
    ws.cell(row=4, column=col + 2, value=f"Tot {CURRENT_YEAR}")

    for row in ws.iter_rows(min_row=2, max_row=4, max_col=last_col):
        for cell in row:
            if cell.value is not None:
                cell.font = header_font if cell.row in (2, 3) else Font(bold=True, size=10, color="7A1414")
                cell.fill = header_fill if cell.row == 2 else (sub_fill if cell.row == 3 else col_fill)
                if cell.row == 4:
                    cell.font = Font(bold=True, size=10, color="7A1414")
                    cell.fill = col_fill
                cell.alignment = center
                cell.border = border
    ws.row_dimensions[2].height = 26
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 20

    data_cols = ["Partner", "Status"]
    for q in quarters:
        data_cols += [f"Q{q} PL", f"Q{q} HL", f"Q{q} Tot"]
    data_cols += ["Tot PL", "Tot HL", "Tot Tot"]

    r = 5
    band_fill = PatternFill(start_color="FBE4E4", end_color="FBE4E4", fill_type="solid")
    for _, dr in pipe_df.iterrows():
        for i, colname in enumerate(data_cols, start=1):
            cell = ws.cell(row=r, column=i, value=dr[colname])
            cell.font = Font(size=10, color="7A1414")
            cell.border = border
            cell.alignment = Alignment(horizontal="center")
            if i > 2 and isinstance(dr[colname], (int, float)):
                cell.number_format = "#,##0"
            if r % 2 == 0:
                cell.fill = band_fill
        r += 1

    total_row_num = r
    total_font = Font(bold=True, size=10, color="FFFFFF")
    total_fill = PatternFill(start_color=HUAWEI_RED, end_color=HUAWEI_RED, fill_type="solid")
    ws.cell(row=total_row_num, column=1, value="Total Tunisia")
    ws.cell(row=total_row_num, column=2, value="")
    for i, colname in enumerate(data_cols[2:], start=3):
        cell = ws.cell(row=total_row_num, column=i, value=round(pipe_df[colname].sum()))
        cell.number_format = "#,##0"
    for cell in ws[total_row_num]:
        cell.font = total_font
        cell.fill = total_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    # Thicker separator border between each quarter-group (Q3 | Q4 | Tot),
    # for visual clarity in a dense table like this.
    for gcol in group_start_cols:
        for row_num in range(2, total_row_num + 1):
            cell = ws.cell(row=row_num, column=gcol)
            existing = cell.border
            cell.border = Border(top=existing.top, bottom=existing.bottom,
                                  left=thick, right=existing.right)

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 12
    for i in range(3, last_col + 1):
        ws.column_dimensions[get_column_letter(i)].width = 10
    ws.freeze_panes = "C5"
    ws.sheet_view.zoomScale = 80
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:{get_column_letter(last_col)}{total_row_num}"

    return data_cols, total_row_num


def apply_pipe_change_coloring(ws, pipe_df, prev_pipe_df, data_cols, data_start_row=5):
    """Same green/red week-over-week coloring pattern as the other two
    reports — tints each metric cell directly based on change vs the most
    recent prior pipe snapshot."""
    if prev_pipe_df is None:
        return
    from openpyxl.styles import PatternFill, Font
    increase_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    decrease_fill = PatternFill(start_color="F4B6B6", end_color="F4B6B6", fill_type="solid")
    increase_font = Font(bold=True, color="1E7B34", size=10)
    decrease_font = Font(bold=True, color="A61B1B", size=10)

    prev_indexed = prev_pipe_df.set_index(prev_pipe_df["Partner"].apply(_norm_pipe_name))
    metric_cols = [c for c in data_cols if c not in ("Partner", "Status")]
    col_indices = {m: data_cols.index(m) + 1 for m in metric_cols}

    for row_offset, (_, row) in enumerate(pipe_df.iterrows()):
        r = data_start_row + row_offset
        key = _norm_pipe_name(row["Partner"])
        if key not in prev_indexed.index:
            continue
        prev_row = prev_indexed.loc[key]
        for m, col_idx in col_indices.items():
            if m not in prev_pipe_df.columns:
                continue
            cur_val = pd.to_numeric(row.get(m), errors="coerce")
            prev_val = pd.to_numeric(prev_row.get(m), errors="coerce")
            if pd.notna(cur_val) and pd.notna(prev_val):
                delta = cur_val - prev_val
                if abs(delta) > 0.01:
                    cell = ws.cell(row=r, column=col_idx)
                    cell.fill = increase_fill if delta > 0 else decrease_fill
                    cell.font = increase_font if delta > 0 else decrease_font


def _norm_pipe_name(s):
    return " ".join(str(s or "").strip().casefold().split())


def build_report():
    current = load_current_week(CURRENT_FILE)
    previous = load_previous_week(PREVIOUS_FILE)
    current_output_colors = load_current_output_colors(OUTPUT_PATH)
    dr_opportunities = load_dr_opportunities(DR_FILE)

    rows = []
    manual_review_count = 0
    dropped_count = 0
    fallback_used_count = 0
    fallback_unavailable_count = 0
    swapped_count = 0
    manual_t1_t2_overridden_count = 0
    t2_fallback_used_count = 0

    for _, r in current.iterrows():
        opp_no = r["e-Opp No."]
        t1 = r["T1 Partner"]
        t2 = r["T2 Partner"]

        if not t1 and not t2:
            dropped_count += 1
            continue

        manual_flag = ""

        def lookup_previous():
            entry = previous.get(norm_key(opp_no)) if previous else None
            if entry and (entry["T1"] or entry["T2"]):
                return entry["T1"], entry["T2"]
            return None, None

        # Comma-separated multi-partner fix: fall back to last week's clean values
        if has_multiple_partners(t1) or has_multiple_partners(t2):
            p1, p2 = lookup_previous()
            if p1 is not None or p2 is not None:
                t1, t2 = p1, p2
                fallback_used_count += 1
            else:
                manual_flag = "MANUAL REVIEW (multiple partners, no prior week data)"
                fallback_unavailable_count += 1

        # TMI normalization — TMI always ends up in T1, name standardized
        if contains_tmi(t2):
            t1 = TMI_CANONICAL_NAME
            t2 = ""
        elif contains_tmi(t1):
            t1 = TMI_CANONICAL_NAME

        # T1 must always be filled — if it arrived blank but T2 has a value,
        # that value belongs in T1 (a single partner given in the wrong
        # column), not left blank.
        if not t1:
            if t2:
                t1, t2 = t2, ""
                swapped_count += 1
            else:
                dropped_count += 1
                continue

        # T2 blank and T1 isn't TMI: this week's T1 might just be wrong/
        # incomplete data (not necessarily that T2 is genuinely missing) —
        # check last week's file for this same opportunity and trust that
        # instead of what this week's raw export shows.
        if not t2 and t1 != TMI_CANONICAL_NAME:
            p1, p2 = lookup_previous()
            if p1 is not None or p2 is not None:
                t1, t2 = p1, p2
                # re-check TMI in case last week's data has it
                if contains_tmi(t2):
                    t1, t2 = TMI_CANONICAL_NAME, ""
                elif contains_tmi(t1):
                    t1 = TMI_CANONICAL_NAME
                t2_fallback_used_count += 1
            else:
                manual_flag = manual_flag or "MANUAL REVIEW (T2 blank)"

        if manual_flag:
            manual_review_count += 1

        prev_entry = previous.get(norm_key(opp_no)) if previous else None

        # Manual T1/T2 correction carry-forward: whatever you last saved
        # for this opportunity wins outright, overriding this week's fresh
        # computation entirely (not just the specific fallback cases above).
        if PRESERVE_MANUAL_T1_T2 and prev_entry and (prev_entry["T1"] or prev_entry["T2"]):
            if prev_entry["T1"] != t1 or prev_entry["T2"] != t2:
                manual_t1_t2_overridden_count += 1
            t1, t2 = prev_entry["T1"], prev_entry["T2"]

        # Colors: prefer whatever's in the CURRENT output (freshest — could
        # be a same-week manual edit that hasn't been promoted to
        # PREVIOUS_FILE yet), falling back to PREVIOUS_FILE's color if this
        # opportunity isn't in the current output at all yet.
        current_entry = current_output_colors.get(norm_key(opp_no))
        if current_entry and (current_entry["color"] or current_entry["raw_color"]):
            color = current_entry["color"]
            raw_color = current_entry["raw_color"]
        else:
            color = prev_entry["color"] if prev_entry else None
            raw_color = prev_entry.get("raw_color") if prev_entry else None
        # New-opportunity / T1-T2 fallback logic always uses PREVIOUS_FILE
        # specifically — real week-over-week comparison, not "did I touch
        # this in the last five minutes".
        is_new = previous is not None and prev_entry is None
        grey = dr_opportunities is not None and norm_key(opp_no) not in dr_opportunities

        row = dict(r)
        row["T1 Partner"] = t1
        row["T2 Partner"] = t2
        row["_manual_flag"] = manual_flag
        row["_color"] = color
        row["_raw_color"] = raw_color
        row["_is_new"] = is_new
        row["_grey"] = grey
        rows.append(row)

    report_df = pd.DataFrame(rows)

    print(f"Dropped {dropped_count} opportunity(ies) with no partner at all.")
    if PRESERVE_MANUAL_T1_T2 and manual_t1_t2_overridden_count:
        print(f"Kept your manually-saved T1/T2 Partner for {manual_t1_t2_overridden_count} "
              f"opportunity(ies) instead of this week's freshly computed value.")
    if swapped_count:
        print(f"Moved T2 into T1 for {swapped_count} opportunity(ies) where T1 arrived blank "
              f"but T2 had a value.")
    if fallback_used_count:
        print(f"Fixed {fallback_used_count} comma-separated multi-partner cell(s) "
              f"using last week's data.")
    if t2_fallback_used_count:
        print(f"Corrected T1/T2 for {t2_fallback_used_count} opportunity(ies) where T2 was "
              f"blank (and T1 wasn't TMI) — used last week's data instead of trusting "
              f"this week's T1 as-is.")
    if fallback_unavailable_count:
        print(f"NOTE: {fallback_unavailable_count} opportunity(ies) have a comma-separated "
              f"partner problem with no prior week to fall back to — flagged for manual review.")
    if manual_review_count:
        print(f"NOTE: {manual_review_count} opportunity(ies) flagged for manual review overall.")

    return report_df


# ============ Overdue detection ============

def parse_order_date(v):
    """Estimated Order Date can show up as a real date, a YYYYMM string, an
    Excel serial number, or similar — normalize to an actual date for
    comparison. A YYYYMM value with no day component is treated as the 1st
    of that month."""
    if pd.isna(v) or v == "":
        return None
    if hasattr(v, "year") and hasattr(v, "month"):
        d = v
        return datetime(d.year, d.month, getattr(d, "day", 1)).date()
    # A raw Excel serial number (e.g. cell is numeric-formatted, not a real
    # Date cell) — pd.to_datetime() on a bare int/float interprets it as
    # nanoseconds-since-1970, giving a garbage 1970 date that's always
    # "in the past". Convert it properly from Excel's epoch instead.
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if 20000 < v < 80000:  # plausible range for a real-world date serial
            from datetime import timedelta
            return (datetime(1899, 12, 30) + timedelta(days=v)).date()
        return None

    raw = str(v).strip()
    stripped = raw.replace("-", "").replace("/", "")

    # YYYYMM, with or without separators (e.g. '2026-08', '202608')
    if stripped.isdigit() and len(stripped) == 6:
        return datetime(int(stripped[:4]), int(stripped[4:6]), 1).date()

    # A genuinely bare number with NO separators anywhere in the original
    # text — likely an Excel serial number written as a text string, as
    # opposed to a real YYYYMMDD date (which pd.to_datetime handles below).
    if raw.isdigit() and stripped == raw and len(raw) != 8:
        try:
            num = float(raw)
            if 20000 < num < 80000:
                from datetime import timedelta
                return (datetime(1899, 12, 30) + timedelta(days=num)).date()
        except ValueError:
            pass

    try:
        d = pd.to_datetime(v)
        return d.date()
    except Exception:
        return None


def is_overdue(v, today):
    """Overdue = the Estimated Order Date is strictly before today's actual
    date. Real day-of-month precision is available (format is YYYY/MM/DD),
    so this compares full dates, not just year+month."""
    parsed = parse_order_date(v)
    if parsed is None:
        return False
    return parsed < today.date()


# ============ Output ============

def style_and_write(report_df):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    def fit_sheet(ws, max_col_width=16):
        """Tightens columns and zooms out so the whole table fits on screen
        without horizontal scrolling, instead of Excel's wide defaults."""
        for col_cells in ws.columns:
            length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
            col_letter = get_column_letter(col_cells[0].column)
            ws.column_dimensions[col_letter].width = min(max(length + 1, 7), max_col_width)
        ws.sheet_view.zoomScale = 70
        ws.sheet_view.showGridLines = True

    today = datetime.today()
    display_cols = [c for c in report_df.columns if not c.startswith("_")]

    wb = Workbook()
    ws = wb.active
    ws.title = "eSales"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(display_cols))
    gen_cell = ws.cell(row=1, column=1, value=f"Report Generated: {today.strftime('%Y-%m-%d %H:%M:%S')}")
    gen_cell.font = Font(italic=True, size=9, color="666666")
    ws.row_dimensions[1].height = 16

    ws.append(display_cols)  # lands on row 2

    header_font = Font(bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill(start_color=HUAWEI_RED, end_color=HUAWEI_RED, fill_type="solid")
    for cell in ws[2]:
        cell.font = header_font
        cell.fill = header_fill
    ws.row_dimensions[2].height = 22

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    grey_fill = PatternFill(start_color=GREY_FILL, end_color=GREY_FILL, fill_type="solid")
    green_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")
    # Colors this script itself applies or reads as red/grey — a carried-forward
    # "arbitrary manual color" is skipped if it matches one of these, since
    # that case is already handled by the red/grey channels above (or is
    # this script's own header/banner color, not a real manual choice).
    SYSTEM_COLOR_HEXES = {HUAWEI_RED.upper(), GREY_FILL.upper(), "FFC7CE", "C6E0B4"} | \
                          {h.upper() for h in KNOWN_GREY_HEXES} | {h.upper() for h in KNOWN_RED_HEXES}

    am_notifications = []

    for _, r in report_df.iterrows():
        ws.append([r[c] for c in display_cols])
        row_idx = ws.max_row
        overdue = is_overdue(r.get("Estimated Order Date"), today)

        reasons = []
        raw_color = r.get("_raw_color")
        has_manual_color = isinstance(raw_color, str) and raw_color.upper() not in SYSTEM_COLOR_HEXES

        if r["_color"] == "red":
            # An explicit manual "flag for deletion" — the most urgent,
            # actionable signal, so it always wins even over an automatic
            # grey/green status.
            manual_fill = PatternFill(start_color=raw_color, end_color=raw_color, fill_type="solid") \
                if isinstance(raw_color, str) else red_fill
            for cell in ws[row_idx]:
                cell.fill = manual_fill
        elif r["_grey"]:
            for cell in ws[row_idx]:
                cell.fill = grey_fill
        elif r["_is_new"]:
            for cell in ws[row_idx]:
                cell.fill = green_fill
        elif has_manual_color:
            # Whatever other color was manually applied last time carries
            # forward exactly as-is.
            manual_fill = PatternFill(start_color=raw_color, end_color=raw_color, fill_type="solid")
            for cell in ws[row_idx]:
                cell.fill = manual_fill

        if r["_color"] == "red":
            reasons.append("Flagged for deletion")

        if overdue:
            reasons.append("Overdue (Estimated Order Date is before today)")

        if reasons:
            am_notifications.append({**{c: r[c] for c in display_cols},
                                      "Notification Reason": "; ".join(reasons)})

        if r["_manual_flag"]:
            ws.cell(row=row_idx, column=len(display_cols) + 1, value=r["_manual_flag"])

    ws.cell(row=2, column=len(display_cols) + 1, value="Manual Review Flag").font = header_font
    ws.cell(row=2, column=len(display_cols) + 1).fill = header_fill

    fit_sheet(ws)
    ws.freeze_panes = "A3"
    last_col_letter = get_column_letter(len(display_cols) + 1)
    ws.auto_filter.ref = f"A2:{last_col_letter}{ws.max_row}"

    def write_gen_date_row(sheet, num_cols):
        sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(num_cols, 1))
        c = sheet.cell(row=1, column=1, value=f"Report Generated: {today.strftime('%Y-%m-%d %H:%M:%S')}")
        c.font = Font(italic=True, size=9, color="666666")
        sheet.row_dimensions[1].height = 16

    # ---- AM Notifications sheet ----
    ws2 = wb.create_sheet("AM Notifications")
    if am_notifications:
        am_df = pd.DataFrame(am_notifications)
        am_cols = ["Opportunity Owner"] + [c for c in am_df.columns if c != "Opportunity Owner"]
        am_df = am_df[am_cols].sort_values("Opportunity Owner")
        write_gen_date_row(ws2, len(am_df.columns))
        ws2.append(list(am_df.columns))  # lands on row 2
        for cell in ws2[2]:
            cell.font = header_font
            cell.fill = header_fill
        for _, r in am_df.iterrows():
            ws2.append(list(r.values))
            if "Flagged for deletion" in str(r["Notification Reason"]):
                for cell in ws2[ws2.max_row]:
                    cell.fill = red_fill
        fit_sheet(ws2)
        ws2.freeze_panes = "A3"
        last_col_letter2 = get_column_letter(len(am_df.columns))
        ws2.auto_filter.ref = f"A2:{last_col_letter2}{ws2.max_row}"
    else:
        write_gen_date_row(ws2, 1)
        ws2.cell(row=2, column=1, value="No Account Manager notifications this week.")

    # ---- New Opportunities sheet ----
    ws3 = wb.create_sheet("New Opportunities")
    new_df = report_df[report_df["_is_new"]][display_cols]
    if not new_df.empty:
        write_gen_date_row(ws3, len(new_df.columns))
        ws3.append(list(new_df.columns))  # lands on row 2
        for cell in ws3[2]:
            cell.font = header_font
            cell.fill = header_fill
        for _, r in new_df.iterrows():
            ws3.append(list(r.values))
        fit_sheet(ws3)
        ws3.freeze_panes = "A3"
        last_col_letter3 = get_column_letter(len(new_df.columns))
        ws3.auto_filter.ref = f"A2:{last_col_letter3}{ws3.max_row}"
    else:
        write_gen_date_row(ws3, 1)
        ws3.cell(row=2, column=1, value="No new opportunities since last week (or no prior week to compare against).")

    print("Building Ongoing Pipe table...")
    pipe_df, quarters = build_pipe_table(report_df)
    prev_pipe_df = load_previous_pipe_snapshot()
    data_cols, pipe_total_row_num = write_pipe_sheet(wb, pipe_df, quarters)
    apply_pipe_change_coloring(wb["Ongoing Pipe"], pipe_df, prev_pipe_df, data_cols)

    wb.save(OUTPUT_PATH)
    save_pipe_snapshot(pipe_df)
    print(f"AM Notifications: {len(am_notifications)} row(s) needing attention.")
    print(f"New Opportunities: {len(new_df)} row(s).")


def main():
    print("Building eSales report...")
    report_df = build_report()
    if report_df.empty:
        print("Nothing to report.")
        return
    style_and_write(report_df)
    print("Done.")
    print(f"{len(report_df)} opportunities in the report.")


if __name__ == "__main__":
    main()