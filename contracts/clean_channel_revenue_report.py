import pandas as pd
from pathlib import Path


def read_any(path, **kwargs):
    path = str(path)
    if path.lower().endswith(".csv"):
        try:
            return pd.read_csv(path, encoding="utf-8", **kwargs)
        except UnicodeDecodeError:
            return pd.read_csv(path, encoding="cp1252", **kwargs)
    return pd.read_excel(path, **kwargs)


# ============ CONFIG ============

RAW_CONTRACT_FILE = "raw_channel_revenue.xlsx"
OPP_DR_FILE = "opportunity_dr_mapping.xlsx"
HISTORY_FILES = "history_data.xlsx"
REFERENCE_FILE = "reference_data.xlsx"
TIER1_PO_REFERENCE_FILE = "tier1_po_reference.xlsx"  # Contract / Tier 1 PO ref / T1 PO Amount Euro, matched by Contract
OUTPUT_PATH = "CHANNEL_REVENUE_REPORT.xlsx"
SNAPSHOT_DIR = "contract_snapshots"  # one CSV saved here every run, used for week-over-week comparison

CONTRACT_PREFIX = "0Y0"  # set to None to include every contract regardless of prefix

# The "TOTAL CONTRACTS" KPI and the Partner Filter sheet both scope to
# contracts starting with this specific prefix (narrower than the general
# CONTRACT_PREFIX filter above, which decides what's in the report at all).
CONTRACT_PREFIX_KPI = "0Y078826"

# Fixed dropdown list for the Dashboard's Partner Filter — only these
# partners are offered as filter choices, regardless of who else appears
# in the data.
CHANNEL_FILTER_PARTNERS = [
    "Partner L13",
    "Partner P30",
    "Partner G22",
    "Partner I5",
    "Partner M21",
    "Partner H3",
    "Partner T15",
    "Partner F22",
    "Partner E8",
    "Partner D12",
    "Partner L1",
    "Partner C10",
]

# Raw-data placeholder strings that mean "actually blank" (data quality
# artifacts seen in the export, e.g. SignedCustomer = "SOURCE IS NULL")
BLANK_PLACEHOLDERS = {"source is null", "null", "n/a", "na", "none", ""}

CARRIER_KEYWORDS = ["Partner X10", "Partner F4", "Partner E24"]
NC_TIER1_KEYWORDS = ["Partner S27", "Partner Q5"]
NC_TIER2_KEYWORDS = ["Partner L13", "Partner I5"]
NONNC_TIER1_KEYWORDS = ["Partner T15", "Partner B14", "computer distribution"]

# Rows involving any of these entities (as Tier 1, Tier 2, or End user) are
# dropped entirely before any other processing — out of scope for this report.
EXCLUDED_ENTITIES = ["Partner R3", "Partner N6"]

# "channel" = N when the End Customer is one of these — government
# ministries, public institutions, banks/insurers, utilities, transport
# authorities, and telecom operators are direct/public-sector accounts,
# not channel business. Everything else defaults to channel = Y.
# Add to this list as more non-channel end customers are identified.
NON_CHANNEL_KEYWORDS = [
    # Government ministries
    "ministry of defence", "ministere de la defense", "Partner Y11",
    "Partner X4", "ministere de l'interieur", "ministère de l'intérieur",
    "Partner V18", "ministere de la jeunesse", "ministère de la jeunesse",
    "ministry of higher education", "ministere de l'enseignement superieur",
    "Partner B26",
    "Partner C19", "ministere de la sante", "Partner P1",
    # Named public institutions
    "cck", "cnte", "fsi",
    # Utility
    "Partner L22",
    # Transport authorities
    "Partner E9", "att", "Partner J4", "enakel", "enaakel",
    # Telecom operators (as end customer, not as partner)
    "Partner F4", "Partner E24", "Partner X10", "Partner Z16",
    # Banks / insurers (generic — matches any name containing these words)
    "banque", "bank", "assurance", "insurance",
]


def contains_word(haystack, word):
    """Word-boundary match (case-insensitive) — avoids short codes like
    'ATT' or 'CCK' false-matching inside unrelated words."""
    import re
    return re.search(r"\b" + re.escape(word) + r"\b", haystack, re.IGNORECASE) is not None


def classify_channel(end_user):
    eu = str(end_user or "")
    if any(contains_word(eu, kw) for kw in NON_CHANNEL_KEYWORDS):
        return "N"
    return "Y"


def _norm_key(v):
    return " ".join(str(v or "").strip().casefold().split())


# Fields with no automated source at all (or only a partial one) — if you
# type a value into one of these by hand in the generated report, it's
# preserved on the next run instead of getting wiped back to blank.
MANUAL_CARRY_FORWARD_FIELDS = ["channel", "Partner AM", "Huawei OO", "Tier 1 PO ref", "T1 PO Amount Euro"]


def load_previous_manual_values(path):
    """Reads the 'Channel Revenue Report' sheet from the PREVIOUS run's
    output file (if it exists) and captures whatever's sitting in the
    manually-maintained fields, keyed by Contract. Used as a fallback: if
    this run's automated logic (raw data + reference table) comes up blank
    for one of these fields, whatever you'd previously typed in by hand
    gets carried forward instead of being erased."""
    p = Path(path)
    if not p.exists():
        return {}
    try:
        prev = pd.read_excel(p, sheet_name="Channel Revenue Report")
    except Exception as e:
        print(f"NOTE: couldn't read previous output '{path}' for manual-value "
              f"carry-forward: {e}")
        return {}
    prev.columns = [str(c).strip() for c in prev.columns]
    if "Contract" not in prev.columns:
        return {}

    result = {}
    for _, r in prev.iterrows():
        contract = clean_text(r.get("Contract", ""))
        if not contract:
            continue
        values = {}
        for field in MANUAL_CARRY_FORWARD_FIELDS:
            v = clean_text(r.get(field, ""))
            if v:
                values[field] = v
        if values:
            result[contract.casefold()] = values
    return result


def load_tier1_po_reference(path):
    """Loads your Tier 1 PO ref / T1 PO Amount Euro lookup sheet, matched by
    Contract. This is the authoritative source for these two fields — it
    takes priority over anything carried forward from a previous run's
    manual edits (see load_previous_manual_values), which only applies as
    a fallback for a contract this sheet doesn't cover yet."""
    p = Path(path)
    if not p.exists():
        return {}
    try:
        df = pd.read_excel(p)
    except Exception as e:
        print(f"NOTE: couldn't read TIER1_PO_REFERENCE_FILE '{path}': {e}")
        return {}
    df.columns = [str(c).strip() for c in df.columns]
    if "Contract" not in df.columns:
        print(f"WARNING: TIER1_PO_REFERENCE_FILE has no 'Contract' column — found: {df.columns.tolist()}")
        return {}

    result = {}
    filled_count = 0
    for _, r in df.iterrows():
        contract = clean_text(r.get("Contract", ""))
        if not contract:
            continue
        po_ref = clean_text(r.get("Tier 1 PO ref", ""))
        po_amount = clean_text(r.get("T1 PO Amount Euro", ""))
        if po_ref or po_amount:
            result[contract.casefold()] = {"Tier 1 PO ref": po_ref, "T1 PO Amount Euro": po_amount}
            filled_count += 1
    print(f"TIER1_PO_REFERENCE_FILE loaded: {len(df)} contract(s) listed, "
          f"{filled_count} with a PO ref and/or amount actually filled in.")
    return result


def load_reference_table(path):
    """Loads your manual Tier 2 Partner / End user / Huawei OO / channel /
    Partner AM lookup table. Returns two dicts:
      - specific: keyed by (Tier 2 Partner, End user) — used when a row in
        your table has BOTH filled in, for an exact partner+customer match.
      - by_enduser: keyed by End user alone — used for rows where your
        table left Tier 2 Partner blank, meaning "this applies to this end
        customer no matter which partner". Checked as a fallback when no
        specific match is found.
    Returns (None, None) if no path is configured or the file isn't found.
    """
    if not path:
        return None, None
    p = Path(path)
    if not p.exists():
        print(f"NOTE: REFERENCE_FILE '{path}' not found — Huawei OO / channel / "
              f"Partner AM will only come from the automated logic (raw data + DR match).")
        return None, None

    ref = read_any(p)
    ref.columns = [str(c).strip() for c in ref.columns]
    needed = ["Tier 2 Partner", "End user", "Huawei OO", "channel", "Partner AM"]
    missing = [c for c in needed if c not in ref.columns]
    if missing:
        print(f"WARNING: REFERENCE_FILE is missing column(s) {missing} — found: {ref.columns.tolist()}")
        for c in missing:
            ref[c] = ""
    for c in needed:
        ref[c] = ref[c].apply(clean_text)

    specific, by_enduser = {}, {}
    for _, r in ref.iterrows():
        eu_key = _norm_key(r["End user"])
        if not eu_key:
            continue
        payload = {"Huawei OO": r["Huawei OO"], "channel": r["channel"], "Partner AM": r["Partner AM"]}
        if r["Tier 2 Partner"]:
            specific[(_norm_key(r["Tier 2 Partner"]), eu_key)] = payload
        else:
            by_enduser[eu_key] = payload

    print(f"Reference table loaded: {len(specific)} partner+customer-specific row(s), "
          f"{len(by_enduser)} end-customer-only row(s).")
    return specific, by_enduser


def apply_reference_override(specific, by_enduser, tier2, end_user, huawei_oo, channel, partner_am):
    """Returns (huawei_oo, channel, partner_am, matched) — your reference
    table wins whenever it has an entry; otherwise the automated values
    passed in are kept unchanged."""
    if specific is None:
        return huawei_oo, channel, partner_am, False
    key_specific = (_norm_key(tier2), _norm_key(end_user))
    row = specific.get(key_specific) or by_enduser.get(_norm_key(end_user))
    if row is None:
        return huawei_oo, channel, partner_am, False
    return (
        row["Huawei OO"] or huawei_oo,
        row["channel"] or channel,
        row["Partner AM"] or partner_am,
        True,
    )


def clean_text(v):
    """Blank out known placeholder junk values; otherwise strip whitespace."""
    if pd.isna(v):
        return ""
    s = str(v).strip()
    return "" if s.casefold() in BLANK_PLACEHOLDERS else s


def contains_any(haystack, needles):
    h = haystack.casefold()
    return any(n in h for n in needles)


# ============ Load raw contract data ============

def load_raw_contracts(path):
    df = read_any(path)
    df.columns = [str(c).strip() for c in df.columns]

    text_cols = ["ContractNumber", "Oppowner", "Contract Name", "EndCustomerName",
                 "ProductName", "Product Group", "NAType", "Opp No.", "SignedCustomer",
                 "Tire2 Customer", "Account Responsible", "IsNA", "Pipeline Name", "Source"]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].apply(clean_text)

    for col in ["Booking", "Revenue"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["Period"] = df["Period"].astype(str).str.strip()

    if CONTRACT_PREFIX:
        before = len(df)
        df = df[df["ContractNumber"].str.startswith(CONTRACT_PREFIX)]
        print(f"Contract filter: kept {len(df)} of {before} raw rows "
              f"(ContractNumber starts with '{CONTRACT_PREFIX}')")

    if EXCLUDED_ENTITIES:
        before = len(df)
        check_cols = ["SignedCustomer", "Tire2 Customer", "EndCustomerName"]
        mask = pd.Series(False, index=df.index)
        for col in check_cols:
            if col in df.columns:
                mask |= df[col].apply(lambda v, col=col: contains_any(v, EXCLUDED_ENTITIES))
        excluded_contracts = sorted(df.loc[mask, "ContractNumber"].unique())
        df = df[~mask]
        if excluded_contracts:
            print(f"Excluded {before - len(df)} row(s) across {len(excluded_contracts)} "
                  f"contract(s) matching EXCLUDED_ENTITIES: {excluded_contracts}")

    return df


# ============ Load Opportunity/DR mapping ============

# Column-name candidates, in preference order (newer/richer export first,
# older simpler export as a fallback for backward compatibility)
OPP_NO_COL_CANDIDATES = ["Opportunity No.", "Opportunity No"]
DR_NO_COL_CANDIDATES = ["DR No.", "DR No"]
PARTNER_AM_SOURCE_COL = "Submitted By"

# The DR system's "Submitted By" field is a raw login/username, not a real
# name. This maps known logins to the actual person's name for the
# "Partner AM" column. Add new entries here as new logins show up — the
# script tells you which ones it couldn't map so you know what to add.
PARTNER_AM_NAME_MAP = {
    "Partner R35": "Partner L38",
    "Partner W5": "Partner Z33",
    "Partner Z7": "Partner V23",
    "Partner T19": "Partner R32",
    "Partner U32": "Partner G35",
    "Partner W18": "Partner G4",
    "Partner D3": "Partner L15",
    "Partner V10": "Partner L15",
    "Partner F16": "Partner Y14",
    "Partner Q12": "Partner O36",
    "Partner P39": "Partner B33",
    "Partner A21": "Partner B33",
    "Partner X29": "Partner C37",
    "Partner Y23": "Partner F23",
    "Partner R6": "Partner D20",
    "Partner X18": "Partner C18",
    "Partner V33": "Partner G38",
    "Partner O19": "Partner W20",
    "Partner I18": "Partner O24",
    "Partner J31": "Partner H11",
    "Partner V32": "Partner L15",
    "Partner R1": "Partner Q33",
    "Partner Q2": "Partner I21",
    "Partner T31": "Partner T2",
    "Partner I8": "Partner S14",
}


def clean_partner_am_name(raw_value):
    """Maps a raw 'Submitted By' login to a real name via
    PARTNER_AM_NAME_MAP. Unrecognized logins are left as-is (not blanked)
    so they're still visible in the report and easy to spot for adding to
    the map later."""
    key = str(raw_value or "").strip().casefold()
    return PARTNER_AM_NAME_MAP.get(key, raw_value)


def _first_present(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None


def load_opp_dr(path):
    df = read_any(path)
    df.columns = [str(c).strip() for c in df.columns]

    opp_col = _first_present(df, OPP_NO_COL_CANDIDATES)
    dr_col = _first_present(df, DR_NO_COL_CANDIDATES)
    if not opp_col or not dr_col:
        raise ValueError(
            f"Opportunity/DR file is missing an Opportunity/DR number column. "
            f"Looked for {OPP_NO_COL_CANDIDATES} and {DR_NO_COL_CANDIDATES}, "
            f"found columns: {df.columns.tolist()}"
        )

    df = df.rename(columns={opp_col: "Opportunity No", dr_col: "DR No"})
    df["Opportunity No"] = df["Opportunity No"].apply(clean_text)
    df["DR No"] = df["DR No"].apply(clean_text)
    if PARTNER_AM_SOURCE_COL in df.columns:
        df["Submitted By"] = df[PARTNER_AM_SOURCE_COL].apply(clean_text)
        unmapped_before = set(df.loc[df["Submitted By"] != "", "Submitted By"].unique())
        df["Submitted By"] = df["Submitted By"].apply(clean_partner_am_name)
        still_unmapped = sorted(v for v in unmapped_before if clean_partner_am_name(v) == v
                                 and v.casefold() not in PARTNER_AM_NAME_MAP)
        if still_unmapped:
            print(f"NOTE: {len(still_unmapped)} 'Submitted By' login(s) have no entry in "
                  f"PARTNER_AM_NAME_MAP, shown as-is in Partner AM: {still_unmapped}")
    else:
        df["Submitted By"] = ""

    df["_match_key"] = df["Opportunity No"].str.casefold()
    return df[["Opportunity No", "DR No", "Submitted By", "_match_key"]].drop_duplicates(
        subset="_match_key", keep="first")


# ============ Optional history for Recognized Revenue / Contract Date ============

def _period_sort_key(v):
    """Make YYYYMM and bare-YYYY values comparable. A bare year sorts as
    the start of that year (e.g. 'CODE-00752' -> 201900), since we don't know
    which month within it. Handles values read from Excel as floats
    (202210.0), ints, plain strings, and actual date/datetime objects —
    pandas silently turns numeric-looking Excel columns into floats, and
    "CODE-00730".isdigit() is False, which was previously causing every
    history date to fail to parse and lose to the current-week date."""
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    if hasattr(v, "strftime"):  # pandas Timestamp / datetime.date / datetime.datetime
        return int(v.strftime("%Y%m"))
    if isinstance(v, (int, float)):
        v = int(v)  # drops any trailing .0 cleanly, unlike string manipulation
        s = str(v)
    else:
        s = str(v).strip()
        if s.endswith(".0"):
            s = s[:-2]
    if not s or s.lower() == "nan":
        return None
    digits = s.replace("-", "")
    if digits.isdigit():
        if len(digits) == 6:
            return int(digits)
        if len(digits) == 4:
            return int(digits) * 100
    return None  # unparseable — ignored for comparison, but original value is kept for display


def pick_oldest(a, b):
    """Returns whichever of a/b is chronologically earliest (via
    _period_sort_key), applied uniformly to every contract — this is never
    a special case, it's the one rule used for every date comparison,
    whether both values come from history, both from the raw extract, or
    one of each. If one side is blank/unparseable, the other wins outright."""
    ka, kb = _period_sort_key(a), _period_sort_key(b)
    if ka is None and kb is None:
        return a if a else b
    if ka is None:
        return b
    if kb is None:
        return a
    return a if ka <= kb else b


def clean_period_display(v):
    """Normalize a date value for display — converts 202210.0 -> 'CODE-00730',
    a Timestamp -> 'CODE-00730', etc. Falls back to the original value if it
    doesn't look like one of the recognized period formats (e.g. bare years
    like 'CODE-00752' are kept as-is rather than forced into YYYYMM)."""
    if v is None or v == "":
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y%m")
    if isinstance(v, (int, float)):
        return str(int(v))
    return str(v).strip()


def load_history(paths):
    """Merges one or more prior-year history files. Each file only needs a
    'Contract' column plus 'Recognized Revenue Date' and 'Contract Date' —
    any other columns (Partner, Partner Type, etc.) are ignored, since
    those vary by year and aren't needed for this lookup. Returns a
    DataFrame indexed by Contract with the EARLIEST known date of each
    type across all supplied files, or None if no paths are configured."""
    if not paths:
        return None
    if isinstance(paths, (str, Path)):
        paths = [paths]

    frames = []
    for p in paths:
        p = Path(p)
        if not p.exists():
            print(f"NOTE: history file '{p}' not found — skipping.")
            continue
        h = read_any(p)
        h.columns = [str(c).strip() for c in h.columns]
        if "Contract" not in h.columns:
            print(f"WARNING: history file '{p}' has no 'Contract' column — skipping.")
            continue
        wanted_cols = ["Contract", "Recognized Revenue Date", "Contract Date",
                       "HW EBG order USD", "Hw EBG revenue USD"]
        keep = h[[c for c in wanted_cols if c in h.columns]].copy()
        for c in wanted_cols:
            if c not in keep.columns:
                keep[c] = None
        keep["Contract"] = keep["Contract"].apply(clean_text)
        keep["HW EBG order USD"] = pd.to_numeric(
            keep["HW EBG order USD"].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
        keep["Hw EBG revenue USD"] = pd.to_numeric(
            keep["Hw EBG revenue USD"].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
        frames.append(keep)

    if not frames:
        return None

    all_hist = pd.concat(frames, ignore_index=True)
    all_hist["_rev_key"] = all_hist["Recognized Revenue Date"].apply(_period_sort_key)
    all_hist["_ctr_key"] = all_hist["Contract Date"].apply(_period_sort_key)

    results = {}
    for contract, group in all_hist.groupby("Contract"):
        if not contract:
            continue
        rev_candidates = group.dropna(subset=["_rev_key"]).sort_values("_rev_key")
        ctr_candidates = group.dropna(subset=["_ctr_key"]).sort_values("_ctr_key")
        results[contract] = {
            "Recognized Revenue Date": rev_candidates.iloc[0]["Recognized Revenue Date"] if not rev_candidates.empty else None,
            "Contract Date": ctr_candidates.iloc[0]["Contract Date"] if not ctr_candidates.empty else None,
            # summed across every historical year this contract appears in —
            # this is prior-years' activity, added to the current extract's
            # total to get a lifetime cumulative figure.
            "Hist Order USD": group["HW EBG order USD"].sum(),
            "Hist Revenue USD": group["Hw EBG revenue USD"].sum(),
        }
    hist_df = pd.DataFrame.from_dict(results, orient="index")
    # Match case/whitespace-insensitively, same fix as the DR lookup — a
    # contract number differing only in case would otherwise silently miss.
    hist_df["_match_key"] = [c.casefold() for c in hist_df.index]
    print(f"History file(s) loaded: {len(hist_df)} unique contract(s) found. "
          f"Sample Contract values from the history file: {list(hist_df.index[:10])}")
    return hist_df


# ============ Partner Type classification ============

def classify_partner_type(tier1, tier2, end_user, project_desc):
    t1, t2 = tier1.casefold(), tier2.casefold()
    eu, pd_ = end_user.casefold(), project_desc.casefold()

    if contains_any(t1, CARRIER_KEYWORDS) or contains_any(t2, CARRIER_KEYWORDS):
        return "Carrier"
    if "demo" in pd_:
        return "DEMO"
    if "stock" in eu or "CODE-04950" in pd_:
        return "STOCK"
    if contains_any(t1, NC_TIER1_KEYWORDS) or contains_any(t2, NC_TIER2_KEYWORDS):
        return "NC"
    if contains_any(t1, NONNC_TIER1_KEYWORDS):
        return "Non-NC"
    if not tier1:
        return "MANUAL REVIEW (Tier 1 missing)"
    return "Direct"


def resolve_tier2(tier1, tier2_raw, partner_type):
    is_tmi = contains_any(tier1.casefold(), NC_TIER1_KEYWORDS)
    if is_tmi:
        return tier1
    if partner_type in ("Direct", "STOCK", "Carrier") and not tier2_raw:
        return tier1
    return tier2_raw


# ============ Per-contract consolidation ============

def first_positive_period(group, value_col):
    g = group.sort_values("Period")
    hit = g[g[value_col] > 0]
    if hit.empty:
        return None
    return hit.iloc[0]["Period"]


SNAPSHOT_COLS = ["Contract", "HW EBG order USD", "Hw EBG revenue USD"]


def save_snapshot(report_df):
    """Save today's Contract/Order/Revenue figures so next run can compare
    against them. One file per calendar day — running the script twice in
    the same day overwrites that day's snapshot rather than creating a
    second one."""
    from datetime import datetime
    folder = Path(SNAPSHOT_DIR)
    folder.mkdir(exist_ok=True)
    today_str = datetime.today().strftime("%Y-%m-%d")
    path = folder / f"snapshot_{today_str}.csv"
    report_df[SNAPSHOT_COLS].to_csv(path, index=False)
    return path


def load_previous_snapshot():
    """Finds the most recent snapshot from a PRIOR day (never today's own,
    even if this is a second run today) to compare against."""
    from datetime import datetime
    folder = Path(SNAPSHOT_DIR)
    if not folder.exists():
        return None, None
    today = datetime.today().date()
    candidates = []
    for f in folder.glob("snapshot_*.csv"):
        date_str = f.stem.replace("snapshot_", "")
        try:
            d = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            continue
        if d < today:
            candidates.append((d, f))
    if not candidates:
        return None, None
    candidates.sort(key=lambda x: x[0], reverse=True)
    prev_date, prev_path = candidates[0]
    return prev_date, read_any(prev_path)


def compare_to_previous(report_df):
    """Returns (new_contracts_df, changed_df, prev_date). Both DataFrames
    are empty (not None) if there's nothing to report, so callers don't
    need extra None-checks."""
    prev_date, prev_df = load_previous_snapshot()
    if prev_df is None:
        return report_df.iloc[0:0], pd.DataFrame(), None

    prev_df = prev_df.set_index("Contract")
    current_contracts = set(report_df["Contract"])
    prev_contracts = set(prev_df.index)

    new_contracts = report_df[report_df["Contract"].isin(current_contracts - prev_contracts)].copy()

    changed_rows = []
    for _, r in report_df[report_df["Contract"].isin(current_contracts & prev_contracts)].iterrows():
        prev = prev_df.loc[r["Contract"]]
        order_delta = round(r["HW EBG order USD"] - prev["HW EBG order USD"], 2)
        rev_delta = round(r["Hw EBG revenue USD"] - prev["Hw EBG revenue USD"], 2)
        if abs(order_delta) > 0.01 or abs(rev_delta) > 0.01:
            changed_rows.append({
                "Contract": r["Contract"],
                "Tier 1 Partner": r["Tier 1 Partner"],
                "Previous Order USD": prev["HW EBG order USD"],
                "Current Order USD": r["HW EBG order USD"],
                "Order Change": order_delta,
                "Previous Revenue USD": prev["Hw EBG revenue USD"],
                "Current Revenue USD": r["Hw EBG revenue USD"],
                "Revenue Change": rev_delta,
            })
    changed_df = pd.DataFrame(changed_rows)
    return new_contracts, changed_df, prev_date


def write_changes_sheet(wb, new_contracts, changed_df, prev_date):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    ws = wb.create_sheet("Weekly Changes")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color=HUAWEI_RED, end_color=HUAWEI_RED, fill_type="solid")
    thin = Side(style="thin", color="E8A0A0")
    full_border = Border(top=thin, bottom=thin, left=thin, right=thin)
    green_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    red_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")

    ws.cell(row=1, column=1, value=(
        f"Compared against snapshot from {prev_date}" if prev_date
        else "No prior snapshot found — this is the first run, nothing to compare yet. "
             "Run the script again next week to start seeing changes here."
    )).font = Font(italic=True, color="808080", size=10)

    row = 3
    ws.cell(row=row, column=1, value=f"New Contracts ({len(new_contracts)})").font = Font(bold=True, size=12, color=HUAWEI_RED)
    row += 1
    if not new_contracts.empty:
        cols = list(new_contracts.columns)  # full data — every column from the main report
        for j, h in enumerate(cols):
            c = ws.cell(row=row, column=1 + j, value=h)
            c.font = header_font
            c.fill = header_fill
            c.border = full_border
        for _, r in new_contracts[cols].iterrows():
            row += 1
            for j, col in enumerate(cols):
                cell = ws.cell(row=row, column=1 + j, value=r[col])
                cell.border = full_border
        row += 1
    else:
        ws.cell(row=row, column=1, value="No new contracts since the last snapshot.").font = Font(italic=True, color="808080")
        row += 1

    row += 2
    ws.cell(row=row, column=1, value=f"Revenue/Order Changes ({len(changed_df)})").font = Font(bold=True, size=12, color=HUAWEI_RED)
    row += 1
    if not changed_df.empty:
        cols = list(changed_df.columns)
        for j, h in enumerate(cols):
            c = ws.cell(row=row, column=1 + j, value=h)
            c.font = header_font
            c.fill = header_fill
            c.border = full_border
        for _, r in changed_df.iterrows():
            row += 1
            for j, col in enumerate(cols):
                cell = ws.cell(row=row, column=1 + j, value=r[col])
                cell.border = full_border
                if col in ("Order Change", "Revenue Change") and isinstance(r[col], (int, float)):
                    is_increase = r[col] >= 0
                    cell.font = Font(color="1E7B34" if is_increase else HUAWEI_RED, bold=True)
                    cell.fill = green_fill if is_increase else red_fill
    else:
        ws.cell(row=row, column=1, value="No revenue or order changes since the last snapshot.").font = Font(italic=True, color="808080")

    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        col_letter = get_column_letter(col_cells[0].column)
        ws.column_dimensions[col_letter].width = min(max(length + 2, 10), 30)


def build_report():
    raw = load_raw_contracts(RAW_CONTRACT_FILE)
    opp_dr = load_opp_dr(OPP_DR_FILE)
    history = load_history(HISTORY_FILES)
    ref_specific, ref_by_enduser = load_reference_table(REFERENCE_FILE)
    previous_manual = load_previous_manual_values(OUTPUT_PATH)
    tier1_po_lookup = load_tier1_po_reference(TIER1_PO_REFERENCE_FILE)
    manual_carried_count = 0

    if raw.empty:
        print("WARNING: no rows left after the ContractNumber filter — nothing to report.")
        return pd.DataFrame()

    rows = []
    manual_review_count = 0
    ref_matched_count = 0
    missing_oo_count = 0
    missing_am_count = 0
    unmatched_opp_numbers = []
    no_tier_count = 0
    hist_matched = hist_unmatched = 0
    unmatched_contracts = []

    for contract_no, group in raw.groupby("ContractNumber"):
        g_sorted = group.sort_values("Period")
        last = g_sorted.iloc[-1]  # most recent period's descriptive fields

        tier1 = last["SignedCustomer"]
        tier2_raw = last["Tire2 Customer"]
        end_user = last["EndCustomerName"]
        project_desc = last["Contract Name"]

        if not tier1 and not tier2_raw:
            no_tier_count += 1
            continue  # neither Tier 1 nor Tier 2 known — nothing to attribute this contract to, skip it

        partner_type = classify_partner_type(tier1, tier2_raw, end_user, project_desc)
        if partner_type.startswith("MANUAL REVIEW"):
            manual_review_count += 1
        tier2 = resolve_tier2(tier1, tier2_raw, partner_type)

        opp_no = last["Opp No."]
        opp_key = opp_no.casefold() if opp_no else ""
        dr_row = opp_dr[opp_dr["_match_key"] == opp_key] if opp_key else pd.DataFrame()
        dr_no = dr_row.iloc[0]["DR No"] if not dr_row.empty else ""
        partner_am = dr_row.iloc[0]["Submitted By"] if not dr_row.empty else ""
        hl_pl = "PL" if dr_no else "HL"
        if opp_no and dr_row.empty:
            unmatched_opp_numbers.append(opp_no)

        huawei_oo = last["Oppowner"]
        channel = classify_channel(end_user)
        huawei_oo, channel, partner_am, ref_matched = apply_reference_override(
            ref_specific, ref_by_enduser, tier2, end_user, huawei_oo, channel, partner_am)
        if ref_matched:
            ref_matched_count += 1

        # Tier 1 PO ref / Amount: the TIER1_PO_REFERENCE_FILE is the
        # authoritative source when it has a value for this contract.
        tier1_po_ref = ""
        t1_po_amount = ""
        po_entry = tier1_po_lookup.get(contract_no.casefold() if contract_no else "")
        if po_entry:
            tier1_po_ref = po_entry.get("Tier 1 PO ref", "")
            t1_po_amount = po_entry.get("T1 PO Amount Euro", "")

        # If automated logic + reference table still came up blank for any
        # manually-maintained field, fall back to whatever was typed in by
        # hand in the PREVIOUS run's output for this same contract, instead
        # of erasing it. For Tier 1 PO ref/Amount specifically, this only
        # applies when TIER1_PO_REFERENCE_FILE doesn't already cover the
        # contract — that sheet is the source of truth once it does.
        prev_values = previous_manual.get(contract_no.casefold() if contract_no else "", {})
        if prev_values:
            if not huawei_oo and "Huawei OO" in prev_values:
                huawei_oo = prev_values["Huawei OO"]
                manual_carried_count += 1
            if not partner_am and "Partner AM" in prev_values:
                partner_am = prev_values["Partner AM"]
                manual_carried_count += 1
            if "channel" in prev_values and not ref_matched:
                # only counts as a genuine manual override if it actually
                # differs from what auto-classification would produce —
                # otherwise it's just restating the same default, not a
                # real manual edit worth reporting.
                if prev_values["channel"] != classify_channel(end_user):
                    channel = prev_values["channel"]
                    manual_carried_count += 1
            if not tier1_po_ref and "Tier 1 PO ref" in prev_values:
                tier1_po_ref = prev_values["Tier 1 PO ref"]
                manual_carried_count += 1
            if not t1_po_amount and "T1 PO Amount Euro" in prev_values:
                t1_po_amount = prev_values["T1 PO Amount Euro"]
                manual_carried_count += 1

        if not huawei_oo:
            missing_oo_count += 1
        if not partner_am:
            missing_am_count += 1

        booking_total = round(group["Booking"].sum(), 2)
        revenue_total = round(group["Revenue"].sum(), 2)

        hist_order = hist_revenue = 0
        contract_date = first_positive_period(group, "Booking")
        rev_date = first_positive_period(group, "Revenue")
        contract_key = contract_no.casefold() if contract_no else ""
        hist_match = history[history["_match_key"] == contract_key] if (history is not None and contract_key) else None
        if hist_match is not None and not hist_match.empty:
            hist_matched += 1
            hist_row = hist_match.iloc[0]
            # Always take whichever is chronologically OLDER — this is the
            # one rule for every contract, not a history-wins default.
            contract_date = pick_oldest(contract_date, hist_row.get("Contract Date"))
            rev_date = pick_oldest(rev_date, hist_row.get("Recognized Revenue Date"))
            hist_order = hist_row.get("Hist Order USD") or 0
            hist_revenue = hist_row.get("Hist Revenue USD") or 0
        elif history is not None:
            hist_unmatched += 1
            unmatched_contracts.append(contract_no)

        rows.append({
            "Source": "EBG Tunisia",
            "Partner Type": partner_type,
            "Tier 1 Partner": tier1,
            "Tier 2 Partner": tier2,
            "End user": end_user,
            "Project Description": project_desc,
            "Contract": contract_no,
            "NA": last["IsNA"],
            "HL/PL": hl_pl,
            "Huawei OO": huawei_oo,
            "channel": channel,
            "Partner AM": partner_am,  # from your reference table, else the Opportunity/DR file's "Submitted By"
            "OPP NUMBER": opp_no,
            "DR": dr_no,
            "Recognized Revenue Date": clean_period_display(rev_date),
            "Contract Date": clean_period_display(contract_date),
            "Tier 1 PO ref": tier1_po_ref,          # blank unless manually entered previously
            "T1 PO Amount Euro": t1_po_amount,      # blank unless manually entered previously
            "HW EBG order USD": booking_total,          # 2026 activity only (official column)
            "Hw EBG revenue USD": revenue_total,        # 2026 activity only (official column)
            "Lifetime HW EBG order USD": round(booking_total + hist_order, 2),
            "Lifetime Hw EBG revenue USD": round(revenue_total + hist_revenue, 2),
        })


    report_df = pd.DataFrame(rows)
    report_df.insert(0, "Index", range(1, len(report_df) + 1))

    if manual_review_count:
        print(f"NOTE: {manual_review_count} contract(s) flagged 'MANUAL REVIEW' "
              f"(Tier 1 Partner missing) — Partner Type couldn't be determined automatically.")

    if no_tier_count:
        print(f"NOTE: {no_tier_count} contract(s) excluded entirely — had neither "
              f"Tier 1 nor Tier 2 partner, so there was nothing to attribute them to.")

    if ref_specific is not None:
        print(f"Reference table applied to {ref_matched_count} of {len(rows)} contract(s).")
    if manual_carried_count:
        print(f"Carried forward {manual_carried_count} manually-entered value(s) from the "
              f"previous run's output ({OUTPUT_PATH}) — channel/Huawei OO/Partner AM/Tier 1 "
              f"PO ref/T1 PO Amount, wherever nothing else could fill them in.")
    if missing_oo_count or missing_am_count:
        print(f"NOTE: after all sources (raw data, DR match, reference table), "
              f"{missing_oo_count} contract(s) still have no Huawei OO and "
              f"{missing_am_count} still have no Partner AM. Add those End Customers "
              f"to REFERENCE_FILE to fill the remaining gaps.")

    if history is not None:
        print(f"History matching: {hist_matched} of {hist_matched + hist_unmatched} contract(s) "
              f"found a match in HISTORY_FILES (dates/lifetime totals backfilled for those). "
              f"{hist_unmatched} had no history entry, so their dates come only from this raw extract.")
        if unmatched_contracts:
            print(f"  Sample unmatched Contract number(s) from this week's raw data "
                  f"(showing up to 10 of {len(unmatched_contracts)}): {unmatched_contracts[:10]}")
            print(f"  Compare these against the 'Sample Contract values from the history file' "
                  f"line printed above — if the formats look different (extra spaces, different "
                  f"prefix/suffix, etc.), that's the mismatch to fix in the history file.")
    else:
        print("NOTE: HISTORY_FILES is empty or none of the files were found — "
              "Recognized Revenue Date / Contract Date are based ONLY on this raw extract "
              "(any contract that first had activity before this file's data will show the wrong date).")

    total_with_opp = (report_df["OPP NUMBER"] != "").sum()
    matched = (report_df["DR"] != "").sum()
    print(f"DR matching: {matched} of {total_with_opp} contract(s) with an OPP NUMBER "
          f"found a matching DR (case/whitespace-insensitive match on Opportunity No).")
    if unmatched_opp_numbers:
        sample = unmatched_opp_numbers[:10]
        print(f"  Unmatched OPP NUMBER(s) not found in the Opportunity/DR file "
              f"(showing up to 10 of {len(unmatched_opp_numbers)}): {sample}")
        print("  This usually means: (a) that opportunity simply has no DR yet (expected — "
              "HL/PL will correctly show HL), (b) the Opportunity/DR file doesn't cover "
              "that period/region, or (c) a genuine formatting mismatch — check one of the "
              "unmatched numbers above against the Opportunity/DR file by eye.")

    return report_df


HUAWEI_RED = "C8102E"
HUAWEI_GRAPHITE = "3C3C3C"
CARD_BG = "F4F4F4"
GRID_LINE = "E0E0E0"
BAND_RED = "FBE4E4"
PALETTE = ["C8102E", "3C3C3C", "8C8C8C", "E8A33D", "8C1A1A", "70AD47"]


def style_report(ws, changed_df=None):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color=HUAWEI_RED, end_color=HUAWEI_RED, fill_type="solid")
    body_font = Font(name="Calibri", size=10, color="7A1414")
    thin = Side(style="thin", color="E8A0A0")
    full_border = Border(top=thin, bottom=thin, left=thin, right=thin)
    band_fill = PatternFill(start_color=BAND_RED, end_color=BAND_RED, fill_type="solid")
    increase_fill = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid")  # green
    decrease_fill = PatternFill(start_color="F4B6B6", end_color="F4B6B6", fill_type="solid")  # red
    increase_font = Font(name="Calibri", size=10, bold=True, color="1E7B34")
    decrease_font = Font(name="Calibri", size=10, bold=True, color="A61B1B")

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = full_border
    ws.row_dimensions[1].height = 30

    # Build a lookup of which contracts changed since last week and by how
    # much, so their Order/Revenue cells can be color-coded directly in
    # this main sheet — not just listed separately in "Weekly Changes".
    change_lookup = {}
    if changed_df is not None and not changed_df.empty:
        for _, r in changed_df.iterrows():
            change_lookup[r["Contract"]] = {
                "order_delta": r["Order Change"],
                "revenue_delta": r["Revenue Change"],
            }

    headers = [c.value for c in ws[1]]
    contract_col_idx = headers.index("Contract") + 1 if "Contract" in headers else None
    order_col_idx = headers.index("HW EBG order USD") + 1 if "HW EBG order USD" in headers else None
    revenue_col_idx = headers.index("Hw EBG revenue USD") + 1 if "Hw EBG revenue USD" in headers else None

    last_row = ws.max_row
    for row in ws.iter_rows(min_row=2, max_row=last_row):
        for cell in row:
            cell.font = body_font
            cell.border = full_border
        if row[0].row % 2 == 0:
            for cell in row:
                cell.fill = band_fill

        if change_lookup and contract_col_idx:
            contract = ws.cell(row=row[0].row, column=contract_col_idx).value
            change = change_lookup.get(contract)
            if change:
                if order_col_idx and abs(change["order_delta"]) > 0.01:
                    cell = ws.cell(row=row[0].row, column=order_col_idx)
                    cell.fill = increase_fill if change["order_delta"] > 0 else decrease_fill
                    cell.font = increase_font if change["order_delta"] > 0 else decrease_font
                if revenue_col_idx and abs(change["revenue_delta"]) > 0.01:
                    cell = ws.cell(row=row[0].row, column=revenue_col_idx)
                    cell.fill = increase_fill if change["revenue_delta"] > 0 else decrease_fill
                    cell.font = increase_font if change["revenue_delta"] > 0 else decrease_font

    ws.freeze_panes = "A2"  # only the header row is frozen — no column freeze
    ws.auto_filter.ref = ws.dimensions
    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        col_letter = get_column_letter(col_cells[0].column)
        ws.column_dimensions[col_letter].width = min(max(length + 1, 7), 15)
    ws.sheet_view.zoomScale = 65  # whole table visible without horizontal scrolling

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def kpi_card(ws, col, row, label, value, accent, sub_value=None):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
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
    vc.font = Font(name="Calibri", size=20, bold=True, color=accent)
    vc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[val_row].height = 34

    rows_used = [accent_cell.row, label_row, val_row]
    sub_row = None
    if sub_value:
        sub_row = row + 3
        ws.merge_cells(start_row=sub_row, start_column=col, end_row=sub_row, end_column=col + span - 1)
        sc = ws.cell(row=sub_row, column=col, value=sub_value)
        sc.font = Font(name="Calibri", size=9, bold=True, color=accent)
        sc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[sub_row].height = 20
        rows_used.append(sub_row)

    bottom_row = sub_row or val_row
    for r in rows_used:
        for c in range(col, col + span):
            cell = ws.cell(row=r, column=c)
            if r != accent_cell.row:
                cell.fill = PatternFill(start_color=CARD_BG, end_color=CARD_BG, fill_type="solid")
            cell.border = Border(
                left=Side(style="thin", color=GRID_LINE) if c == col else None,
                right=Side(style="thin", color=GRID_LINE) if c == col + span - 1 else None,
                top=Side(style="thin", color=GRID_LINE) if r == accent_cell.row else None,
                bottom=Side(style="thin", color=GRID_LINE) if r == bottom_row else None,
            )


def style_bar_chart(chart, title, y_title=None):
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import (RichTextProperties, Paragraph,
        ParagraphProperties, CharacterProperties)
    chart.title = title
    if y_title and hasattr(chart, "y_axis"):
        chart.y_axis.title = y_title
    if hasattr(chart, "gapWidth"):
        chart.gapWidth = 60  # more space between bars so labels don't crowd each other
    chart.legend = None
    if hasattr(chart, "y_axis"):
        chart.y_axis.majorGridlines = None  # remove horizontal gridlines for a cleaner look
        # Explicitly force the value axis to render its tick-number labels —
        # left unset, this can silently hide the Y-axis numbers in some
        # Excel versions (the same issue we already had to fix for the
        # x-axis category labels).
        chart.y_axis.delete = False
        chart.y_axis.tickLblPos = "nextTo"
    if hasattr(chart, "x_axis"):
        small_font = CharacterProperties(sz=800)
        chart.x_axis.txPr = RichText(
            bodyPr=RichTextProperties(rot=-1500000, vert="horz"),  # gentler ~25° angle, sits lower/further from the bars
            p=[Paragraph(pPr=ParagraphProperties(defRPr=small_font), endParaRPr=small_font)]
        )
        chart.x_axis.majorGridlines = None
        chart.x_axis.delete = False
        chart.x_axis.tickLblPos = "low"  # push labels below the axis line, away from the plot area


def flat_color_series(series, n, color):
    from openpyxl.chart.marker import DataPoint
    from openpyxl.chart.shapes import GraphicalProperties
    points = []
    for i in range(n):
        dp = DataPoint(idx=i)
        dp.graphicalProperties = GraphicalProperties(solidFill=color)
        points.append(dp)
    series.data_points = points


def palette_series(series, n):
    from openpyxl.chart.marker import DataPoint
    from openpyxl.chart.shapes import GraphicalProperties
    points = []
    for i in range(n):
        dp = DataPoint(idx=i)
        dp.graphicalProperties = GraphicalProperties(solidFill=PALETTE[i % len(PALETTE)])
        points.append(dp)
    series.data_points = points


def add_value_labels(series, color, position="outEnd", number_format=None):
    """Show the actual number directly on each bar so the value is visible
    without hovering."""
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import (RichTextProperties, Paragraph,
        ParagraphProperties, CharacterProperties)
    dl = DataLabelList()
    dl.showVal = True
    dl.showCatName = False
    dl.showSerName = False
    dl.showPercent = False
    dl.showLegendKey = False
    dl.dLblPos = position
    if number_format:
        dl.numFmt = number_format
    font = CharacterProperties(sz=800, b=True, solidFill=color)
    dl.txPr = RichText(
        bodyPr=RichTextProperties(),
        p=[Paragraph(pPr=ParagraphProperties(defRPr=font), endParaRPr=font)]
    )
    series.dLbls = dl


def add_distinct_helper_columns(report_df):
    """Adds two hidden helper columns used by the Dashboard's Partner
    Filter panel for accurate DISTINCT counts of Opportunities/DRs — a
    plain SUMIFS/COUNTIFS can't do a distinct count on its own, and
    array-formula tricks for this are fragile, so it's computed here in
    Python instead (reliable, testable) and just summed in Excel."""
    report_df = report_df.copy()
    report_df["_Opp First Occurrence"] = 0
    report_df["_DR First Occurrence"] = 0
    for partner, group in report_df.groupby("Tier 1 Partner"):
        seen_opp, seen_dr = set(), set()
        for idx, row in group.iterrows():
            opp = str(row.get("OPP NUMBER", "")).strip()
            if opp and opp not in seen_opp:
                report_df.loc[idx, "_Opp First Occurrence"] = 1
                seen_opp.add(opp)
            dr = str(row.get("DR", "")).strip()
            if dr and dr not in seen_dr:
                report_df.loc[idx, "_DR First Occurrence"] = 1
                seen_dr.add(dr)
    return report_df


def build_dashboard(wb, report_df, new_contracts=None, changed_df=None):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.chart import BarChart, PieChart, Reference
    from openpyxl.chart.label import DataLabelList
    from datetime import datetime

    if new_contracts is None:
        new_contracts = report_df.iloc[0:0]
    if changed_df is None:
        changed_df = pd.DataFrame()

    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False

    DASHBOARD_LAST_COL = 27  # 'AA' — matches the print area / chart extent below
    from openpyxl.utils import get_column_letter as _gcl
    last_col_letter = _gcl(DASHBOARD_LAST_COL)

    ws.merge_cells(f"B2:{last_col_letter}2")
    title = ws["B2"]
    title.value = "CHANNEL_REVENUE_REPORT"
    title.font = Font(name="Calibri", size=20, bold=True, color="FFFFFF")
    title.fill = PatternFill(start_color=HUAWEI_RED, end_color=HUAWEI_RED, fill_type="solid")
    title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 38

    ws.merge_cells(f"B3:{last_col_letter}3")
    sub = ws["B3"]
    sub.value = f"Generated {datetime.today().strftime('%Y-%m-%d')}  \u00b7  {len(report_df)} contracts"
    sub.font = Font(name="Calibri", size=10, italic=True, color="8C8C8C")

    # ---- KPIs: financial totals on one row, monitoring flags on the next ----
    total_contracts = report_df["Contract"].astype(str).str.startswith(CONTRACT_PREFIX_KPI).sum()
    rev_2026_k = report_df["Hw EBG revenue USD"].sum() / 1000
    ord_2026_k = report_df["HW EBG order USD"].sum() / 1000
    manual_count = report_df["Partner Type"].astype(str).str.startswith("MANUAL REVIEW").sum()
    backlog_mask = (report_df["HW EBG order USD"] > 0) & (report_df["Hw EBG revenue USD"] == 0)
    backlog_count = int(backlog_mask.sum())
    backlog_k = report_df.loc[backlog_mask, "HW EBG order USD"].sum() / 1000

    new_count = len(new_contracts)
    new_order_k = (new_contracts["HW EBG order USD"].sum() / 1000) if new_count else 0

    changed_count = len(changed_df)
    if changed_count and "Order Change" in changed_df.columns:
        order_change_k = changed_df["Order Change"].sum() / 1000
        revenue_change_k = changed_df["Revenue Change"].sum() / 1000
    else:
        order_change_k = revenue_change_k = 0

    financial_specs = [
        ("TOTAL CONTRACTS 2026", f"{total_contracts:,}", HUAWEI_GRAPHITE, None),
        ("2026 REVENUE (K$)", f"{rev_2026_k:,.0f}", HUAWEI_RED, None),
        ("2026 ORDER (K$)", f"{ord_2026_k:,.0f}", HUAWEI_RED, None),
    ]
    monitoring_specs = [
        ("NEW SINCE LAST WEEK", f"{new_count:,}",
         "70AD47" if new_count else HUAWEI_GRAPHITE,
         f"+{new_order_k:,.0f}K$ orders" if new_count else None),
        ("CHANGED SINCE LAST WEEK", f"{changed_count:,}",
         "E8A33D" if changed_count else HUAWEI_GRAPHITE,
         f"Ord {order_change_k:+,.0f}K$ / Rev {revenue_change_k:+,.0f}K$" if changed_count else None),
        ("UNBILLED BACKLOG (contracts)", f"{backlog_count:,}", "E8A33D" if backlog_count else "70AD47", None),
        ("UNBILLED BACKLOG (K$)", f"{backlog_k:,.0f}", "E8A33D" if backlog_k else "70AD47", None),
    ]

    FIN_ROW = 5
    MON_ROW = 10
    col = 2
    for label, value, accent, sub in financial_specs:
        kpi_card(ws, col, FIN_ROW, label, value, accent, sub_value=sub)
        col += 3
    col = 2
    for label, value, accent, sub in monitoring_specs:
        kpi_card(ws, col, MON_ROW, label, value, accent, sub_value=sub)
        col += 3

    # Two separate filter panels: Panel A covers the 3 entities that
    # genuinely sign as Tier 1 (Partner T15 / Partner C10 /
    # TMI); Panel B covers the 10 sub-partners, restricted by which of the
    # two distributors (Partner T15 / Partner C10) they
    # came through. Both start at the same row, in different columns, so
    # the dashboard doesn't get taller than a single panel's height.
    panel_a_end = write_distributor_filter_panel(ws, report_df, top_row=MON_ROW + 5, label_col=2)
    panel_b_end = write_partner_filter_panel(ws, report_df, top_row=MON_ROW + 5, label_col=9)
    chart_start_row = max(panel_a_end, panel_b_end) + 2

    # ---- staged chart data (hidden far-right columns) ----
    STAGE1, STAGE3, STAGE5, STAGE6, STAGE7 = 70, 76, 82, 85, 88
    STAGE8, STAGE9 = 92, 95
    row0 = 5

    # Exclude Manual Review from the Partner Type breakdown charts — it's not
    # a real classification, just a flag that needs human follow-up (still
    # counted in the KPI card above, just not plotted here).
    typed_df = report_df[~report_df["Partner Type"].astype(str).str.startswith("MANUAL REVIEW")]

    by_type_count = typed_df["Partner Type"].value_counts()
    ws.cell(row=row0, column=STAGE1, value="Partner Type")
    ws.cell(row=row0, column=STAGE1 + 1, value="Contracts")
    for i, (k, v) in enumerate(by_type_count.items(), start=1):
        ws.cell(row=row0 + i, column=STAGE1, value=k)
        ws.cell(row=row0 + i, column=STAGE1 + 1, value=int(v))
    n1 = len(by_type_count)

    by_type_ord = (typed_df.groupby("Partner Type")["HW EBG order USD"].sum() / 1000).sort_values(ascending=False)
    STAGE2 = STAGE1 + 3
    ws.cell(row=row0, column=STAGE2, value="Partner Type")
    ws.cell(row=row0, column=STAGE2 + 1, value="Orders (K$)")
    for i, (k, v) in enumerate(by_type_ord.items(), start=1):
        ws.cell(row=row0 + i, column=STAGE2, value=k)
        ws.cell(row=row0 + i, column=STAGE2 + 1, value=round(float(v), 1))
    n2 = len(by_type_ord)

    hl_pl_order = (report_df.groupby("HL/PL")["HW EBG order USD"].sum() / 1000)
    ws.cell(row=row0, column=STAGE3, value="HL/PL")
    ws.cell(row=row0, column=STAGE3 + 1, value="Order Amount (K$)")
    for i, (k, v) in enumerate(hl_pl_order.items(), start=1):
        ws.cell(row=row0 + i, column=STAGE3, value=k)
        ws.cell(row=row0 + i, column=STAGE3 + 1, value=round(float(v), 1))
    n3 = len(hl_pl_order)

    # NA vs non-NA split — by order amount, not contract count
    na_map = report_df["NA"].astype(str).str.strip().str.casefold().map(
        lambda v: "NA" if v in ("y", "yes") else ("Non-NA" if v in ("n", "no") else "Unspecified"))
    na_order = report_df.groupby(na_map)["HW EBG order USD"].sum() / 1000
    ws.cell(row=row0, column=STAGE5, value="NA Status")
    ws.cell(row=row0, column=STAGE5 + 1, value="Order Amount (K$)")
    for i, (k, v) in enumerate(na_order.items(), start=1):
        ws.cell(row=row0 + i, column=STAGE5, value=k)
        ws.cell(row=row0 + i, column=STAGE5 + 1, value=round(float(v), 1))
    n5 = len(na_order)

    # Contracts by Partner — COUNT, not revenue/orders
    partner_contract_counts = report_df["Tier 1 Partner"].value_counts().sort_values(ascending=False).head(10)
    ws.cell(row=row0, column=STAGE6, value="Partner")
    ws.cell(row=row0, column=STAGE6 + 1, value="Contracts")
    for i, (k, v) in enumerate(partner_contract_counts.items(), start=1):
        ws.cell(row=row0 + i, column=STAGE6, value=k)
        ws.cell(row=row0 + i, column=STAGE6 + 1, value=int(v))
    n6 = len(partner_contract_counts)

    # Revenue & Order trend across saved weekly snapshots
    trend_dates, trend_rev, trend_ord = [], [], []
    snap_folder = Path(SNAPSHOT_DIR)
    if snap_folder.exists():
        snap_files = sorted(snap_folder.glob("snapshot_*.csv"))
        for f in snap_files:
            date_str = f.stem.replace("snapshot_", "")
            try:
                snap_df = read_any(f)
                trend_dates.append(date_str)
                trend_rev.append(round(snap_df["Hw EBG revenue USD"].sum() / 1000, 1))
                trend_ord.append(round(snap_df["HW EBG order USD"].sum() / 1000, 1))
            except Exception:
                continue
    if trend_dates:
        print(f"Revenue & Order Trend chart data ({len(trend_dates)} snapshot(s)):")
        for d, r, o in zip(trend_dates, trend_rev, trend_ord):
            print(f"  {d}: Revenue={r}K$, Orders={o}K$")
    ws.cell(row=row0, column=STAGE7, value="Snapshot Date")
    ws.cell(row=row0, column=STAGE7 + 1, value="Revenue (K$)")
    ws.cell(row=row0, column=STAGE7 + 2, value="Orders (K$)")
    for i, (d, r, o) in enumerate(zip(trend_dates, trend_rev, trend_ord), start=1):
        ws.cell(row=row0 + i, column=STAGE7, value=d)
        ws.cell(row=row0 + i, column=STAGE7 + 1, value=r)
        ws.cell(row=row0 + i, column=STAGE7 + 2, value=o)
    n7 = len(trend_dates)

    # Distributor split — now THREE distributors (Computer Distribution
    # S.A, Partner T15, TMI), matched case-insensitively and covering
    # known spelling variants on Tier 1 Partner.
    dist_cf = report_df["Tier 1 Partner"].astype(str).str.casefold()
    is_cds = dist_cf.str.contains("computer distribution")
    is_smart = dist_cf.str.contains("Partner T15") | dist_cf.str.contains("Partner B14")
    is_tmi = dist_cf.str.contains("Partner S27") | dist_cf.str.contains("Partner Q5")
    distributor_order = pd.Series({
        "Partner C10": report_df.loc[is_cds, "HW EBG order USD"].sum() / 1000,
        "Partner T15": report_df.loc[is_smart, "HW EBG order USD"].sum() / 1000,
        "Partner S27": report_df.loc[is_tmi, "HW EBG order USD"].sum() / 1000,
    })
    ws.cell(row=row0, column=STAGE8, value="Distributor")
    ws.cell(row=row0, column=STAGE8 + 1, value="Orders (K$)")
    for i, (k, v) in enumerate(distributor_order.items(), start=1):
        ws.cell(row=row0 + i, column=STAGE8, value=k)
        ws.cell(row=row0 + i, column=STAGE8 + 1, value=round(float(v), 1))
    n8 = len(distributor_order)

    # Top Tier 2 Partners by Orders — genuine sub-partners only (rows where
    # Tier 2 Partner is just a fallback copy of Tier 1 Partner, set by
    # resolve_tier2() for Direct/STOCK/Carrier contracts, are excluded so
    # this chart reflects real Tier-2-level partners, not Tier 1 restated —
    # this is also why TMI never appears here: TMI's own contracts always
    # have Tier 2 Partner = Tier 1 Partner = TMI itself, by design in
    # resolve_tier2(), since TMI signs directly rather than through a
    # sub-partner. TMI's own activity shows in the Distributor chart above
    # instead, which is the correct place for it.
    tier2_only = report_df[(report_df["Tier 2 Partner"] != "") &
                            (report_df["Tier 2 Partner"] != report_df["Tier 1 Partner"])]
    top_tier2 = (tier2_only.groupby("Tier 2 Partner")["HW EBG order USD"].sum() / 1000
                 ).sort_values(ascending=False).head(10)
    ws.cell(row=row0, column=STAGE9, value="Tier 2 Partner")
    ws.cell(row=row0, column=STAGE9 + 1, value="Orders (K$)")
    for i, (k, v) in enumerate(top_tier2.items(), start=1):
        ws.cell(row=row0 + i, column=STAGE9, value=k)
        ws.cell(row=row0 + i, column=STAGE9 + 1, value=round(float(v), 1))
    n9 = len(top_tier2)

    for c in [STAGE1, STAGE1+1, STAGE2, STAGE2+1, STAGE3, STAGE3+1,
              STAGE5, STAGE5+1, STAGE6, STAGE6+1, STAGE7, STAGE7+1, STAGE7+2,
              STAGE8, STAGE8+1, STAGE9, STAGE9+1]:
        from openpyxl.utils import get_column_letter
        ws.column_dimensions[get_column_letter(c)].hidden = True

    # ---- Chart layout: 4 rows of paired charts, trend chart LAST (full width) ----
    row_a = chart_start_row
    row_b = row_a + 22
    row_c = row_b + 22
    row_d = row_c + 22
    row_trend = row_d + 22

    # ---- Chart A: Contracts by Partner Type ----
    a = BarChart()
    a.type = "col"
    style_bar_chart(a, "Contracts by Partner Type", "Contracts")
    data = Reference(ws, min_col=STAGE1+1, min_row=row0, max_row=row0+n1)
    cats = Reference(ws, min_col=STAGE1, min_row=row0+1, max_row=row0+n1)
    a.add_data(data, titles_from_data=True)
    a.set_categories(cats)
    a.visible_cells_only = False
    flat_color_series(a.series[0], n1, HUAWEI_RED)
    add_value_labels(a.series[0], HUAWEI_GRAPHITE)
    a.height = 10
    a.width = 18
    ws.add_chart(a, f"B{row_a}")

    # ---- Chart B: Orders by Partner Type ----
    b = BarChart()
    b.type = "col"
    style_bar_chart(b, "Orders by Partner Type (K$)", "K$")
    data = Reference(ws, min_col=STAGE2+1, min_row=row0, max_row=row0+n2)
    cats = Reference(ws, min_col=STAGE2, min_row=row0+1, max_row=row0+n2)
    b.add_data(data, titles_from_data=True)
    b.set_categories(cats)
    b.visible_cells_only = False
    flat_color_series(b.series[0], n2, HUAWEI_RED)
    add_value_labels(b.series[0], HUAWEI_GRAPHITE)
    b.height = 10
    b.width = 18
    ws.add_chart(b, f"M{row_a}")

    # ---- Chart C: HL vs PL split ----
    c = PieChart()
    style_bar_chart(c, "HL vs PL Split (Order Amount K$)")
    data = Reference(ws, min_col=STAGE3+1, min_row=row0, max_row=row0+n3)
    cats = Reference(ws, min_col=STAGE3, min_row=row0+1, max_row=row0+n3)
    c.add_data(data, titles_from_data=True)
    c.set_categories(cats)
    c.visible_cells_only = False
    palette_series(c.series[0], n3)
    c.dataLabels = DataLabelList()
    c.dataLabels.showCatName = True
    c.dataLabels.showPercent = False
    c.dataLabels.showVal = True
    c.dataLabels.showSerName = False
    c.dataLabels.showLegendKey = False
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties
    white_font = CharacterProperties(sz=900, b=True, solidFill="FFFFFF")
    c.dataLabels.txPr = RichText(
        p=[Paragraph(pPr=ParagraphProperties(defRPr=white_font), endParaRPr=white_font)]
    )
    c.legend = None
    c.height = 10
    c.width = 18
    ws.add_chart(c, f"B{row_b}")

    # ---- Chart E: NA vs non-NA split ----
    e = PieChart()
    style_bar_chart(e, "NA vs Non-NA Split (Order Amount K$)")
    data = Reference(ws, min_col=STAGE5+1, min_row=row0, max_row=row0+n5)
    cats = Reference(ws, min_col=STAGE5, min_row=row0+1, max_row=row0+n5)
    e.add_data(data, titles_from_data=True)
    e.set_categories(cats)
    e.visible_cells_only = False
    palette_series(e.series[0], n5)
    e.dataLabels = DataLabelList()
    e.dataLabels.showCatName = True
    e.dataLabels.showPercent = False
    e.dataLabels.showVal = True
    e.dataLabels.showSerName = False
    e.dataLabels.showLegendKey = False
    from openpyxl.chart.text import RichText as _RichText
    from openpyxl.drawing.text import Paragraph as _Paragraph, ParagraphProperties as _PP, CharacterProperties as _CP
    _white = _CP(sz=900, b=True, solidFill="FFFFFF")
    e.dataLabels.txPr = _RichText(p=[_Paragraph(pPr=_PP(defRPr=_white), endParaRPr=_white)])
    e.legend = None
    e.height = 10
    e.width = 18
    ws.add_chart(e, f"M{row_b}")

    # ---- Chart F: Contracts by Partner (count) ----
    f = BarChart()
    f.type = "col"
    style_bar_chart(f, "Contracts by Partner (Count)", "Contracts")
    data = Reference(ws, min_col=STAGE6+1, min_row=row0, max_row=row0+n6)
    cats = Reference(ws, min_col=STAGE6, min_row=row0+1, max_row=row0+n6)
    f.add_data(data, titles_from_data=True)
    f.set_categories(cats)
    f.visible_cells_only = False
    flat_color_series(f.series[0], n6, HUAWEI_RED)
    add_value_labels(f.series[0], HUAWEI_GRAPHITE)
    f.height = 10
    f.width = 18
    ws.add_chart(f, f"B{row_c}")

    # ---- Chart H: Distributor Orders (K$) — CDS / Partner T15 / TMI ----
    h = BarChart()
    h.type = "col"
    style_bar_chart(h, "Distributor Orders (K$)", "K$")
    data = Reference(ws, min_col=STAGE8+1, min_row=row0, max_row=row0+n8)
    cats = Reference(ws, min_col=STAGE8, min_row=row0+1, max_row=row0+n8)
    h.add_data(data, titles_from_data=True)
    h.set_categories(cats)
    h.visible_cells_only = False
    flat_color_series(h.series[0], n8, HUAWEI_RED)
    add_value_labels(h.series[0], HUAWEI_GRAPHITE)
    h.height = 10
    h.width = 18
    ws.add_chart(h, f"M{row_c}")

    # ---- Chart I: Top 10 Tier 2 Partners by Orders (K$) ----
    i_chart = BarChart()
    i_chart.type = "col"
    style_bar_chart(i_chart, "Top 10 Tier 2 Partners by Orders (K$)", "K$")
    data = Reference(ws, min_col=STAGE9+1, min_row=row0, max_row=row0+n9)
    cats = Reference(ws, min_col=STAGE9, min_row=row0+1, max_row=row0+n9)
    i_chart.add_data(data, titles_from_data=True)
    i_chart.set_categories(cats)
    i_chart.visible_cells_only = False
    flat_color_series(i_chart.series[0], n9, HUAWEI_RED)
    add_value_labels(i_chart.series[0], HUAWEI_GRAPHITE)
    i_chart.height = 11
    i_chart.width = 18
    ws.add_chart(i_chart, f"B{row_d}")

    # ---- Chart G: Revenue & Order trend — LAST chart, full width ----
    if n7 >= 2:
        from openpyxl.chart import LineChart
        g = LineChart()
        style_bar_chart(g, "Revenue & Order Trend (K$)", "K$")
        rev_ref = Reference(ws, min_col=STAGE7+1, min_row=row0, max_row=row0+n7)
        ord_ref = Reference(ws, min_col=STAGE7+2, min_row=row0, max_row=row0+n7)
        cats = Reference(ws, min_col=STAGE7, min_row=row0+1, max_row=row0+n7)
        g.add_data(rev_ref, titles_from_data=True)
        g.add_data(ord_ref, titles_from_data=True)
        g.set_categories(cats)
        g.visible_cells_only = False
        g.series[0].graphicalProperties.line.solidFill = HUAWEI_RED
        g.series[0].graphicalProperties.line.width = 25000
        g.series[0].smooth = False
        g.series[1].graphicalProperties.line.solidFill = HUAWEI_GRAPHITE
        g.series[1].graphicalProperties.line.width = 25000
        g.series[1].smooth = False
        from openpyxl.chart.legend import Legend
        from openpyxl.chart.layout import Layout, ManualLayout
        g.legend = Legend(legendPos="b")
        g.legend.layout = Layout(manualLayout=ManualLayout(
            xMode="edge", yMode="edge", x=0.35, y=0.93, h=0.07, w=0.3))
        g.height = 11
        g.width = 37
        ws.add_chart(g, f"B{row_trend}")
    else:
        ws.cell(row=row_trend, column=2, value=(
            "Revenue & Order Trend will appear here once at least 2 weekly snapshots exist "
            f"(currently {n7} saved in {SNAPSHOT_DIR}/). Run the script again next week to build history."
        )).font = Font(italic=True, color="808080", size=10)

    ws.column_dimensions["A"].width = 2
    for letter in ["B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA"]:
        ws.column_dimensions[letter].width = 9
    # Re-apply the filter panels' specific widths — the generic loop above
    # would otherwise reset them back to the default 9.
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 3
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["I"].width = 20
    ws.column_dimensions["J"].width = 3
    ws.column_dimensions["K"].width = 14
    ws.column_dimensions["L"].width = 14
    ws.column_dimensions["M"].width = 14

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:AB{row_trend + 25}"


DISTRIBUTOR_FILTER_OPTIONS = ["Partner T15", "Partner C10", "Partner S27"]
DISTRIBUTOR_MATCH_KEYWORDS = {
    "Partner T15": ["Partner T15", "Partner B14"],
    "Partner C10": ["computer distribution"],
    "Partner S27": ["Partner S27", "Partner Q5"],
}
# Filter B's 10 sub-partners — CHANNEL_FILTER_PARTNERS minus the two
# distributors now covered by Filter A instead.
SUB_PARTNER_FILTER_OPTIONS = [p for p in CHANNEL_FILTER_PARTNERS
                               if p not in ("Partner T15", "Partner C10")]
# Filter B's distributor-restriction only offers the two that sub-partners
# actually sit under — TMI signs directly, so it has no Tier-2 sub-partners.
SUB_PARTNER_DISTRIBUTOR_OPTIONS = ["Partner T15", "Partner C10"]


def _distributor_mask(report_df, distributor_name):
    cf = report_df["Tier 1 Partner"].astype(str).str.casefold()
    keywords = DISTRIBUTOR_MATCH_KEYWORDS[distributor_name]
    mask = pd.Series(False, index=report_df.index)
    for kw in keywords:
        mask |= cf.str.contains(kw, regex=False)
    return mask


def _metrics_row(sub, prefix_mask_for_row):
    contracts_2026 = int(prefix_mask_for_row.sum()) if len(sub) else 0
    hl = sub[sub["HL/PL"] == "HL"]
    pl = sub[sub["HL/PL"] == "PL"]
    na = sub[sub["NA"] == "Yes"]
    non_na = sub[sub["NA"] == "No"]
    channel = sub[sub["channel"] == "Y"]
    non_channel = sub[sub["channel"] == "N"]
    opps = sub.loc[sub["OPP NUMBER"].astype(str).str.strip() != "", "OPP NUMBER"].nunique()
    drs = sub.loc[sub["DR"].astype(str).str.strip() != "", "DR"].nunique()
    return {
        "Contracts 2026": contracts_2026,
        "Total Revenue (K$)": round(sub["Hw EBG revenue USD"].sum() / 1000, 1),
        "Total Order (K$)": round(sub["HW EBG order USD"].sum() / 1000, 1),
        "HL Contracts": len(hl),
        "HL Order (K$)": round(hl["HW EBG order USD"].sum() / 1000, 1),
        "HL Revenue (K$)": round(hl["Hw EBG revenue USD"].sum() / 1000, 1),
        "PL Contracts": len(pl),
        "PL Order (K$)": round(pl["HW EBG order USD"].sum() / 1000, 1),
        "PL Revenue (K$)": round(pl["Hw EBG revenue USD"].sum() / 1000, 1),
        "NA Contracts": len(na),
        "NA Order (K$)": round(na["HW EBG order USD"].sum() / 1000, 1),
        "NA Revenue (K$)": round(na["Hw EBG revenue USD"].sum() / 1000, 1),
        "Non-NA Contracts": len(non_na),
        "Non-NA Order (K$)": round(non_na["HW EBG order USD"].sum() / 1000, 1),
        "Non-NA Revenue (K$)": round(non_na["Hw EBG revenue USD"].sum() / 1000, 1),
        "Channel Contracts": len(channel),
        "Channel Order (K$)": round(channel["HW EBG order USD"].sum() / 1000, 1),
        "Channel Revenue (K$)": round(channel["Hw EBG revenue USD"].sum() / 1000, 1),
        "Non-Channel Contracts": len(non_channel),
        "Non-Channel Order (K$)": round(non_channel["HW EBG order USD"].sum() / 1000, 1),
        "Non-Channel Revenue (K$)": round(non_channel["Hw EBG revenue USD"].sum() / 1000, 1),
        "Opportunities": opps,
        "DRs": drs,
    }


def compute_distributor_filter_metrics(report_df):
    """Filter A's data: one row per distributor (Partner T15 / Partner C10 / TMI), standalone — matched purely by Tier 1 Partner,
    no sub-partner or second-distributor dimension needed since these ARE
    the top-level entities themselves."""
    prefix_mask = report_df["Contract"].astype(str).str.startswith(CONTRACT_PREFIX_KPI)
    rows = []
    for distributor in DISTRIBUTOR_FILTER_OPTIONS:
        mask = _distributor_mask(report_df, distributor)
        sub = report_df[mask]
        metrics = _metrics_row(sub, prefix_mask[mask])
        metrics["Distributor"] = distributor
        rows.append(metrics)
    return pd.DataFrame(rows)


def compute_partner_filter_metrics(report_df):
    """Filter B's data: one row per (sub-partner, distributor) pair — the
    10 sub-partners x the 2 distributors they can sit under. A contract
    counts toward a partner if they show up as EITHER Tier 1 OR Tier 2,
    AND Tier 1 Partner matches the selected distributor specifically —
    this is what lets you ask 'show me Partner L13's numbers, but
    only for deals that came through Partner C10'."""
    prefix_mask = report_df["Contract"].astype(str).str.startswith(CONTRACT_PREFIX_KPI)
    rows = []
    for partner in SUB_PARTNER_FILTER_OPTIONS:
        is_partner = (report_df["Tier 1 Partner"] == partner) | (report_df["Tier 2 Partner"] == partner)
        for distributor in SUB_PARTNER_DISTRIBUTOR_OPTIONS:
            dist_mask = _distributor_mask(report_df, distributor)
            mask = is_partner & dist_mask
            sub = report_df[mask]
            metrics = _metrics_row(sub, prefix_mask[mask])
            metrics["Partner"] = partner
            metrics["Distributor"] = distributor
            rows.append(metrics)
    return pd.DataFrame(rows)


def write_distributor_filter_panel(ws, report_df, top_row, label_col):
    """Filter A — single dropdown, 3 options (Partner T15 / Partner C10 / TMI). Returns the row number just below the panel."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.utils import get_column_letter

    header_font = Font(bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill(start_color=HUAWEI_GRAPHITE, end_color=HUAWEI_GRAPHITE, fill_type="solid")
    thin = Side(style="thin", color="DDDDDD")

    value_col = label_col + 2
    vcl = get_column_letter(value_col)
    sel = f"${vcl}${top_row}"

    ws.row_dimensions[top_row].height = 30
    ws.cell(row=top_row, column=label_col, value="Filter by Distributor:").font = Font(bold=True, size=13)

    ws.merge_cells(start_row=top_row, start_column=value_col, end_row=top_row, end_column=value_col + 2)
    sel_cell = ws.cell(row=top_row, column=value_col)
    for c in range(value_col, value_col + 3):
        cell = ws.cell(row=top_row, column=c)
        cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        cell.font = Font(bold=True, size=13)
        cell.border = Border(top=Side(style="medium", color=HUAWEI_RED), bottom=Side(style="medium", color=HUAWEI_RED),
                              left=Side(style="medium", color=HUAWEI_RED) if c == value_col else None,
                              right=Side(style="medium", color=HUAWEI_RED) if c == value_col + 2 else None)
    sel_cell.alignment = Alignment(horizontal="center", vertical="center")

    list_col = 110
    list_col_letter = get_column_letter(list_col)
    for i, p in enumerate(DISTRIBUTOR_FILTER_OPTIONS, start=1):
        ws.cell(row=i, column=list_col, value=p)
    ws.column_dimensions[list_col_letter].hidden = True

    dv = DataValidation(type="list", formula1=f"${list_col_letter}$1:${list_col_letter}${len(DISTRIBUTOR_FILTER_OPTIONS)}",
                         allow_blank=True, showDropDown=False)
    dv.error = "Choose a distributor from the dropdown list."
    dv.prompt = "Choose a distributor"
    ws.add_data_validation(dv)
    dv.add(sel_cell)
    sel_cell.value = DISTRIBUTOR_FILTER_OPTIONS[0]

    metrics_df = compute_distributor_filter_metrics(report_df)
    stage_col = 115
    stage_col_letter = get_column_letter(stage_col)
    metric_names = [c for c in metrics_df.columns if c != "Distributor"]
    ws.cell(row=1, column=stage_col, value="Distributor")
    for j, name in enumerate(metric_names, start=1):
        ws.cell(row=1, column=stage_col + j, value=name)
    for i, (_, mrow) in enumerate(metrics_df.iterrows(), start=2):
        ws.cell(row=i, column=stage_col, value=mrow["Distributor"])
        for j, name in enumerate(metric_names, start=1):
            ws.cell(row=i, column=stage_col + j, value=mrow[name])
    for c in range(stage_col, stage_col + len(metric_names) + 1):
        ws.column_dimensions[get_column_letter(c)].hidden = True

    key_range = f"${stage_col_letter}$2:${stage_col_letter}${len(metrics_df) + 1}"
    match_row = f"MATCH({sel},{key_range},0)"

    r = top_row + 2
    ws.cell(row=r, column=label_col, value="Metric").font = header_font
    ws.cell(row=r, column=label_col).fill = header_fill
    ws.cell(row=r, column=value_col, value="Value").font = header_font
    ws.cell(row=r, column=value_col).fill = header_fill
    for j, name in enumerate(metric_names, start=1):
        r += 1
        col_letter = get_column_letter(stage_col + j)
        formula = f'=IFERROR(INDEX(${col_letter}$2:${col_letter}${len(metrics_df) + 1},{match_row}),"")'
        lc = ws.cell(row=r, column=label_col, value=name)
        lc.border = Border(bottom=thin)
        vc = ws.cell(row=r, column=value_col, value=formula)
        vc.border = Border(bottom=thin)
        vc.font = Font(bold=True, color=HUAWEI_RED)
        vc.alignment = Alignment(horizontal="center")

    table_first_row = top_row + 2
    for row_num in range(table_first_row, r + 1):
        ws.row_dimensions[row_num].height = 16

    return r + 2


def write_partner_filter_panel(ws, report_df, top_row, label_col):
    """Filter B — two dropdowns: sub-partner name (10 options) AND which
    distributor (Partner T15 / Partner C10) their data is
    restricted to. Returns the row number just below the panel."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.utils import get_column_letter

    header_font = Font(bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill(start_color=HUAWEI_GRAPHITE, end_color=HUAWEI_GRAPHITE, fill_type="solid")
    thin = Side(style="thin", color="DDDDDD")

    value_col = label_col + 2
    vcl = get_column_letter(value_col)
    sel_partner = f"${vcl}${top_row}"
    sel_distributor = f"${vcl}${top_row + 1}"

    ws.row_dimensions[top_row].height = 30
    ws.row_dimensions[top_row + 1].height = 30
    ws.cell(row=top_row, column=label_col, value="Filter by Partner:").font = Font(bold=True, size=13)
    ws.cell(row=top_row + 1, column=label_col, value="...from Distributor:").font = Font(bold=True, size=13)

    def draw_dropdown_box(row_num):
        ws.merge_cells(start_row=row_num, start_column=value_col, end_row=row_num, end_column=value_col + 2)
        cell = ws.cell(row=row_num, column=value_col)
        for c in range(value_col, value_col + 3):
            cc = ws.cell(row=row_num, column=c)
            cc.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            cc.font = Font(bold=True, size=13)
            cc.border = Border(top=Side(style="medium", color=HUAWEI_RED), bottom=Side(style="medium", color=HUAWEI_RED),
                                left=Side(style="medium", color=HUAWEI_RED) if c == value_col else None,
                                right=Side(style="medium", color=HUAWEI_RED) if c == value_col + 2 else None)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        return cell

    partner_sel_cell = draw_dropdown_box(top_row)
    distributor_sel_cell = draw_dropdown_box(top_row + 1)

    list_col = 150
    list_col_letter = get_column_letter(list_col)
    for i, p in enumerate(SUB_PARTNER_FILTER_OPTIONS, start=1):
        ws.cell(row=i, column=list_col, value=p)
    ws.column_dimensions[list_col_letter].hidden = True

    dv = DataValidation(type="list", formula1=f"${list_col_letter}$1:${list_col_letter}${len(SUB_PARTNER_FILTER_OPTIONS)}",
                         allow_blank=True, showDropDown=False)
    dv.error = "Choose a partner from the dropdown list."
    dv.prompt = "Choose a partner"
    ws.add_data_validation(dv)
    dv.add(partner_sel_cell)
    partner_sel_cell.value = SUB_PARTNER_FILTER_OPTIONS[0]

    dist_list_col = 151
    dist_list_col_letter = get_column_letter(dist_list_col)
    for i, d in enumerate(SUB_PARTNER_DISTRIBUTOR_OPTIONS, start=1):
        ws.cell(row=i, column=dist_list_col, value=d)
    ws.column_dimensions[dist_list_col_letter].hidden = True

    dv2 = DataValidation(type="list", formula1=f"${dist_list_col_letter}$1:${dist_list_col_letter}${len(SUB_PARTNER_DISTRIBUTOR_OPTIONS)}",
                          allow_blank=True, showDropDown=False)
    dv2.error = "Choose a distributor."
    dv2.prompt = "Choose a distributor"
    ws.add_data_validation(dv2)
    dv2.add(distributor_sel_cell)
    distributor_sel_cell.value = SUB_PARTNER_DISTRIBUTOR_OPTIONS[0]

    metrics_df = compute_partner_filter_metrics(report_df)
    stage_col = 155
    stage_col_letter = get_column_letter(stage_col)
    metric_names = [c for c in metrics_df.columns if c not in ("Partner", "Distributor")]
    ws.cell(row=1, column=stage_col, value="Key")
    for j, name in enumerate(metric_names, start=1):
        ws.cell(row=1, column=stage_col + j, value=name)
    for i, (_, mrow) in enumerate(metrics_df.iterrows(), start=2):
        ws.cell(row=i, column=stage_col, value=f'{mrow["Partner"]}|{mrow["Distributor"]}')
        for j, name in enumerate(metric_names, start=1):
            ws.cell(row=i, column=stage_col + j, value=mrow[name])
    for c in range(stage_col, stage_col + len(metric_names) + 1):
        ws.column_dimensions[get_column_letter(c)].hidden = True

    key_range = f"${stage_col_letter}$2:${stage_col_letter}${len(metrics_df) + 1}"
    lookup_key = f'{sel_partner}&"|"&{sel_distributor}'
    match_row = f"MATCH({lookup_key},{key_range},0)"

    r = top_row + 3
    ws.cell(row=r, column=label_col, value="Metric").font = header_font
    ws.cell(row=r, column=label_col).fill = header_fill
    ws.cell(row=r, column=value_col, value="Value").font = header_font
    ws.cell(row=r, column=value_col).fill = header_fill
    for j, name in enumerate(metric_names, start=1):
        r += 1
        col_letter = get_column_letter(stage_col + j)
        formula = f'=IFERROR(INDEX(${col_letter}$2:${col_letter}${len(metrics_df) + 1},{match_row}),"")'
        lc = ws.cell(row=r, column=label_col, value=name)
        lc.border = Border(bottom=thin)
        vc = ws.cell(row=r, column=value_col, value=formula)
        vc.border = Border(bottom=thin)
        vc.font = Font(bold=True, color=HUAWEI_RED)
        vc.alignment = Alignment(horizontal="center")

    table_first_row = top_row + 3
    for row_num in range(table_first_row, r + 1):
        ws.row_dimensions[row_num].height = 16

    return r + 2


def main():
    print("Building Channel Revenue Report...")
    report_df = build_report()
    if report_df.empty:
        return

    report_df = add_distinct_helper_columns(report_df)

    print(f"Writing {OUTPUT_PATH}...")
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Channel Revenue Report"
    ws.append(list(report_df.columns))
    for _, r in report_df.iterrows():
        ws.append(list(r.values))

    print("Checking for changes since last run...")
    new_contracts, changed_df, prev_date = compare_to_previous(report_df)
    if prev_date is None:
        print("  No prior snapshot found — this is the first run. Nothing to compare yet; "
              "run the script again next week to start seeing week-over-week changes.")
    else:
        print(f"  Compared against snapshot from {prev_date}: "
              f"{len(new_contracts)} new contract(s), {len(changed_df)} contract(s) with revenue/order changes.")

    print("Applying report formatting...")
    style_report(ws, changed_df)
    # Hide the four distinct-count helper columns — they exist only for the
    # Dashboard's Partner Filter panels' formulas, not meant to be read directly.
    from openpyxl.utils import get_column_letter as _gcl2
    helper_start = len(report_df.columns) - 3  # 0-indexed position of "_Opp First Occurrence T1"
    for offset in (0, 1, 2, 3):
        ws.column_dimensions[_gcl2(helper_start + offset)].hidden = True

    write_changes_sheet(wb, new_contracts, changed_df, prev_date)

    # Save today's snapshot BEFORE building the dashboard — otherwise the
    # trend chart only shows snapshots from PRIOR runs and always excludes
    # today's own just-computed totals, making it look one run behind.
    snap_path = save_snapshot(report_df)
    print(f"Saved today's snapshot to {snap_path} for next week's comparison.")

    print("Building dashboard...")
    build_dashboard(wb, report_df, new_contracts, changed_df)

    wb.save(OUTPUT_PATH)

    print("Done.")
    print(f"{len(report_df)} contracts in the report.")


if __name__ == "__main__":
    main()