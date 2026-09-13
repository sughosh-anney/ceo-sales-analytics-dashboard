"""
refresh_ceo_dashboard.py

Regenerates ceo-dashboard.html from the CEO Sales Analytics workbook in this
folder, overwriting the HTML file in place. Run this any time you update the
source Excel file:

    py refresh_ceo_dashboard.py

Needs only openpyxl (no pandas), same as Dashboard/refresh_dashboard.py.

------------------------------------------------------------------------------
WHAT IT READS
------------------------------------------------------------------------------
Auto-detects the single .xlsx file in this folder (ignores Excel lock files
starting with "~$") -- the user replaces this file each quarter, so the exact
name is never hardcoded. Reads 8 sheets that already carry resolved, formatted
values (this workbook computes everything via formula/VLOOKUP into these
sheets, so data_only=True is enough -- no need to touch the 12 underlying
pivot sheets):

    CEO Executive Summary, Profitability_1, Vol vs Value & Price (+ _DC),
    Sales Dashboard, Top 10 & Tail (+ _Sub category), Growth Analysis

Cell coordinates below were verified against the workbook directly (a full
cell dump), not reconstructed from memory -- see CEO Dashboard/../ artifact
for the narrative structure map.

------------------------------------------------------------------------------
BADGES: recomputed from the live site's own legend rules, not copied
------------------------------------------------------------------------------
The source workbook's own Signal/Health/Trend emoji column is hand-set and
inconsistent (one clearly growing category was tagged "Degrow" in an early
cut), so none
of it is read here. Instead this script recomputes three badge systems from
scratch, using the exact threshold rules published in the incumbent
vendor-hosted dashboard's own legend panels (confirmed against every
row of its live data on 2026-07-22):
  - signal_from_growth(): Signal Indicator Legend (Exec Summary + Top 10 &
    Tail + Top 10 Sub-category "Status" columns) -- Strong Growing >=15%,
    Growing 5-15%, Slow Growing 0-5%, Degrowing <0%, NA ==0%.
  - growth_driver(): Growth Driver Legend (Vol vs Value tabs) -- sign-based
    on Volume effect (V) vs Price effect (P) vs Net value change, 7 states
    incl. "Price/Volume Cushioning Decline" (one effect masking a net loss).
  - trend_from_growth(): Trend Indicator Legend (Growth Analysis tab) --
    Accelerating/Decelerating from the 1M/2M/3M trajectory shape, else
    Steady Growth/Turnaround/Sustained Decline from Q1 vs 3M sign, else
    Volatile/Flat.
"""

import glob
import json
import os
import re
import sys
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent
HTML_PATH = BASE / "CEO_Dashboard_Standalone.html"
MC_FINANCIALS_PATH = BASE.parent / "Sales Automation" / "Master Consolidation" / "financials_data.json"

# Matches the live vendor dashboard's own 10-row entity order (2026-07-31):
# USA is split into its 3 constituent entities plus a dedicated Elimination
# row, rather than the single combined 'ADF USA' line -- see
# read_us_entity_pnl() in parse_financials.py for how those 4 rows are derived.
FS_ENTITY_ORDER = ['ADFL', 'ADFIL', 'TFIL', 'ADF Australia', 'ADF UK',
                   'ADF Holdings USA', 'ADF USA Ltd', 'Vibrant Foods NJ', 'Elimination', 'Total']


def build_fs_entity_pnl():
    """Entity-wise P&L for the CURRENT quarter (a real Q1 FY27, Apr-Jun), sourced
    from the board's own Financial Statement workbook via Master Consolidation's
    already-computed financials_data.json (read-only import, not re-parsed here --
    mirrors Dashboard/refresh_dashboard.py's load_mc_fin_data()). Replaces the
    Profitability_1 sheet's own 'Entity-wise P&L Snapshot' table, which is
    mislabeled 'Q1 FY27' but is actually only 2 months (Apr-May) of data.
    Same figures, same units (Rs Lakhs), same margin formula (EBITDA, PBT or PAT
    divided by Revenue) as that dashboard's Entity P&L -- current quarter table,
    so both agree exactly.

    IMPORTANT SCOPE CAVEAT (2026-07-30): this is TOTAL statutory revenue (all
    channels, all markets), not the export-only revenue the rest of this
    dashboard (Zone/Category/Brand signals, Sales Dashboard, Top 10 & Tail) is
    built around -- the two will NOT reconcile against each other, by design.

    Entity list (2026-07-31): matches the live vendor dashboard's 10 rows --
    USA is split into its 3 constituent entities (ADF Holdings USA, ADF USA
    Ltd, Vibrant Foods NJ) plus a dedicated Elimination row, rather than the
    single combined 'ADF USA' line used before. Both come from
    parse_financials.py's read_us_entity_pnl(): the 3 USA entities read
    straight off their own standalone sheets (ADFHL PNL / ADF USA PL / VIB USA
    PNL); Elimination combines the group-level intercompany elimination
    ('P & L Consol' sheet, already read for the other entities) with the
    intra-US-group elimination derived there (see that function's docstring
    for the reconciliation math and its one known small gap: prior-year
    EBITDA is a few lakhs off the vendor site's own figure, everything else
    ties out closely). 'ADF USA' (the old single combined line) is still in
    financials_data.json for any other consumer, just not in FS_ENTITY_ORDER
    above.

    Prior-year (Q1 FY26) figures come from 'pnl_entities_prior_quarter', added
    to parse_financials.py/financials_data.json on 2026-07-30 specifically for
    this table -- that block is missing 'ADF Australia' (didn't exist as an
    entity yet in Q1 FY26), which comes through as None here rather than a
    fabricated 0, and renders as a blank/zero-look figure in the table since
    the frontend has no separate "N/A" money format. (TFIL's PY EBITDA was
    also None at one point due to a blank cell upstream in the source
    workbook -- that's since been filled in and TFIL now has real PY EBITDA/
    PBT/PAT figures like every other entity; only ADF Australia's prior-year
    gap remains, and only because the entity is genuinely newer than FY26.)"""
    with open(MC_FINANCIALS_PATH, encoding='utf-8') as f:
        fin = json.load(f)
    entities = fin['pnl_entities_current_quarter']
    py_entities = fin.get('pnl_entities_prior_quarter', {})
    rows = []
    for name in FS_ENTITY_ORDER:
        s = entities.get(name)
        if s is None:
            continue
        rev = s.get('Revenue from operations')
        ebitda = s.get('EBITDA')
        pbt = s.get('Profit before Tax')
        pat = s.get('Profit for the year')
        py = py_entities.get(name) or {}
        rows.append({
            "name": name,
            "revenue": rev, "ebitda": ebitda,
            "ebitdaPct": (ebitda / rev) if rev else None,
            "pbt": pbt,
            "pbtPct": (pbt / rev) if rev else None,
            "pat": pat,
            "patPct": (pat / rev) if rev else None,
            "pyRevenue": py.get('Revenue from operations'),
            "pyEbitda": py.get('EBITDA'),
            "pyPbt": py.get('Profit before Tax'),
        })
    return rows


def build_gross_margin_kpi(category_margin_total, label):
    """Executive Summary's 'Gross Margin' tile.

    CORRECTED 2026-07-31: this was previously sourced from the board Financial
    Statement's Consolidated 'Gross profit margin' ratio (`build_fs_gross_margin_kpi()`,
    now removed), on the theory that the reference tile used a different,
    statutory calculation from the sales workbook's own product-level
    Price-minus-COGS margin -- reasonable at the time, since the two figures
    were far apart. That gap has since closed: a later cut of the workbook
    fixed the Category-wise Gross Margin Analysis table's own data, and its
    TOTAL row's marginPct/marginPctPy/marginDeltaPp now match the reference
    tile's VALUE, SUBTITLE, AND ROUNDING exactly -- confirmed by direct DOM
    comparison, not a coincidence of close-enough numbers. This proves the
    reference tile was never FS-sourced at all, it is the same product-level
    category-margin TOTAL row this dashboard already computes elsewhere.
    Reverted to that source; the FS ratio is not used here anymore.
    (The margin figures compared above were correct for the cut of the
    workbook current at the time; each later cut recomputes its own TOTAL
    row, which is expected behaviour, not a discrepancy.)

    `label` now comes from the caller's own sheet-derived text (A6, e.g.
    "Gross Margin 4M FY27") instead of being hardcoded to "...Q1 FY27" --
    that hardcode went stale the moment the reporting window widened past a
    single quarter (caught 2026-08 via a live browser check after the period
    changed to a 4-month YTD cut; only the VALUE needs the more-accurate
    category_margin_total source, the period wording should always track
    whatever the sheet itself currently says)."""
    return {
        "label": label,
        "value": category_margin_total["marginPct"],
        "sub": f"PY: {category_margin_total['marginPctPy']*100:.1f}%  Δ {category_margin_total['marginDeltaPp']*100:.1f} pp",
        "tone": "red", "fmt": "ratio",
    }


def find_source_xlsx():
    candidates = [p for p in glob.glob(str(BASE / "*.xlsx")) if not os.path.basename(p).startswith("~$")]
    if not candidates:
        raise FileNotFoundError(f"No .xlsx file found in {BASE}")
    if len(candidates) > 1:
        candidates.sort(key=os.path.getmtime, reverse=True)
        print(f"WARNING: multiple .xlsx files found, using most recently modified: {os.path.basename(candidates[0])}")
    return candidates[0]


SHEETS = {
    "exec": "\U0001F454 CEO Executive Summary",
    "profit": "\U0001F4B9 Profitability_1",
    "vol": "⚖️ Vol vs Value & Price",
    "vol_dc": "⚖️ Vol vs Value & Price _DC",
    "sales": "\U0001F4CA Sales Dashboard",
    "top": "\U0001F51D Top 10 & Tail",
    "top_sub": "\U0001F51D Top 10 & Tail_Sub category",
    "growth": "\U0001F4C8 Growth Analysis",
}


def v(ws, coord):
    val = ws[coord].value
    return val


def num(ws, coord):
    val = ws[coord].value
    return round(float(val), 6) if isinstance(val, (int, float)) else (val or 0)


def txt(ws, coord):
    val = ws[coord].value
    return str(val).strip() if val is not None else ""


def rows_block(ws, r1, r2, col_letters, keys):
    """Read rows r1..r2 (inclusive) across the given column letters into a
    list of dicts keyed by `keys` (same length/order as col_letters)."""
    out = []
    for r in range(r1, r2 + 1):
        row = {}
        for col, key in zip(col_letters, keys):
            cell = f"{col}{r}"
            val = ws[cell].value
            if isinstance(val, (int, float)):
                row[key] = round(float(val), 6)
            elif val is None:
                row[key] = None
            else:
                row[key] = str(val).strip()
        out.append(row)
    return out


def gr(cur, py):
    """None (renders "--") when there's no prior-year base to grow from,
    rather than a fabricated 0% or a divide-by-zero."""
    if py:
        return (cur - py) / py
    return None


# 2026-08-11, per explicit user instruction: relabel the source workbook's
# own zone names (ALL CAPS, as typed into the sheet) to the short form the
# user wants everywhere -- also matches Conso's own zone naming, which pulls
# in Standalone's own zone-wise split (see build_ceo_dashboard_conso.py's
# load_standalone_zone_rows()) and needs identical labels to merge cleanly.
ZONE_NAME_FIX = {
    "NORTH AMERICA": "North America", "UNITED KINGDOM": "UK",
    "WESTERN EUROPE": "Europe", "GULF COUNTRIES": "Middle East",
    "ASIA PACIFIC": "Asia Pacific", "INDIA": "India",
}


def fix_zone_names(rows):
    for r in rows:
        name = r.get("name")
        if name in ZONE_NAME_FIX:
            r["name"] = ZONE_NAME_FIX[name]
    return rows


# Growth-rate field -> (numerator_key, denominator_key) needed to recompute
# it correctly after merging two salesman rows into one -- growth rates are
# NOT additive ((a+c)/(b+d) != a/b + c/d), unlike money/share fields (cur,
# py, apr, ytdCur, sharePct, mixPct, ...), which sum safely since they're
# either absolute values or a share of a constant grand total.
_SALESMAN_GROWTH_FIELDS = {
    "grPct": ("cur", "py"), "grYoy": ("cur", "py"), "q1Gr": ("ytdCur", "ytdPy"),
    "aprGr": ("apr", "aprPy"), "mayGr": ("may", "mayPy"), "junGr": ("jun", "junPy"), "julGr": ("jul", "julPy"),
}
# gr1m..gr4m (Growth Analysis sheet's own trailing-month rates) and grQoq
# have no corresponding raw apr/may/jun/jul figures in that row shape to
# recompute a merged rate from -- left as None (renders "--") rather than
# keeping one entity's now-inapplicable rate or fabricating a blend.
_SALESMAN_UNRECOMPUTABLE_GROWTH_FIELDS = {"gr1m", "gr2m", "gr3m", "gr4m", "grQoq"}


def merge_salesmen(rows):
    """Two salesman rows that the business treats as one territory are merged
    into a single row, per explicit user instruction (2026-08-10) -- applied
    everywhere a Salesman block is built (Exec signals, Sales Dashboard,
    Top 10 & Tail, Growth Analysis). Must run BEFORE any downstream
    signal/status/trend computation and before rank_by_value() (for the
    Top 10 & Tail block) so those get correctly (re)computed off the
    merged totals, not copied from either original row.

    The pair to merge and the surviving row name are configuration, read
    from the environment as a comma-separated list (MERGE_SALESMAN_NAMES)
    with the first entry used as the merged row's name, so no individual's
    name is hard-coded in source."""
    _configured = [n.strip().upper() for n in
                   os.environ.get("MERGE_SALESMAN_NAMES", "").split(",") if n.strip()]
    NAMES = set(_configured)
    if len(NAMES) < 2:
        return rows  # merging not configured for this deployment
    matches = [r for r in rows if str(r.get("name", "")).strip().upper() in NAMES]
    if len(matches) < 2:
        return rows  # nothing to merge in this particular block/view

    skip_keys = {"name", "cumPct"} | set(_SALESMAN_GROWTH_FIELDS) | _SALESMAN_UNRECOMPUTABLE_GROWTH_FIELDS
    merged = {"name": _configured[0]}  # first configured entry is the surviving row name
    for key in set(k for m in matches for k in m) - skip_keys:
        vals = [m[key] for m in matches if isinstance(m.get(key), (int, float))]
        merged[key] = sum(vals) if vals else None

    for key, (num_key, den_key) in _SALESMAN_GROWTH_FIELDS.items():
        if any(key in m for m in matches):
            merged[key] = gr(merged.get(num_key) or 0, merged.get(den_key) or 0)
    for key in _SALESMAN_UNRECOMPUTABLE_GROWTH_FIELDS:
        if any(key in m for m in matches):
            merged[key] = None
    if "cumPct" in matches[0]:
        merged["cumPct"] = None  # recomputed by rank_by_value() right after this, where applicable

    out = [r for r in rows if str(r.get("name", "")).strip().upper() not in NAMES]
    out.append(merged)
    return out


def signal_from_growth(gr_pct):
    """Signal Indicator Legend, exactly as published on the live site:
    Strong Growing >=15%, Growing 5-15%, Slow Growing 0-5%, Degrowing <0%,
    NA ==0%."""
    gr = gr_pct or 0
    if gr == 0:
        return "NA"
    if gr < 0:
        return "Degrowing"
    if gr < 0.05:
        return "Slow Growing"
    if gr < 0.15:
        return "Growing"
    return "Strong Growing"


def trend_from_growth(q1_gr, gr2m, gr3m, gr4m):
    """Trend Indicator Legend, exactly as published on the live site's Growth
    & Degrowth Analysis tab. Trajectory shape (Accelerating/Decelerating) is
    checked before the aggregate Q1-vs-terminal-month rules, but ONLY
    qualifies if the terminal month's own sign agrees with the trajectory's
    direction -- a strictly-falling trajectory that is STILL POSITIVE by the
    terminal month (a brand whose growth fell month on month but stayed
    positive throughout) is NOT
    "Decelerating" on live, it falls through to "Steady Growth" (Q1 Gr% >= 5%
    AND terminal Gr% >= 0%) instead -- confirmed by finding a live row with
    this exact shape and comparing its label against another row with an
    outwardly-identical falling trajectory that DOES end negative (Ready To
    Eat, correctly "Decelerating") -- 2026-07-31. The original unconditional
    version of this function got this backwards for any monotonic trajectory
    that doesn't cross zero by the terminal month, verified against all 16
    rows of two full dimension tables before applying this fix, not just the
    one example.

    Window shifted 2026-08 from (1M, 2M, 3M) to (2M, 3M, 4M) now that the
    source workbook publishes a 4th trailing month -- gr4m is the new
    "terminal"/most-current rate, same role gr3m played before; the
    trajectory-shape check and all sign thresholds are otherwise unchanged."""
    q1_gr, gr2m, gr3m, gr4m = q1_gr or 0, gr2m or 0, gr3m or 0, gr4m or 0
    if gr2m < gr3m < gr4m and gr4m >= 0:
        return "Accelerating"
    if gr2m > gr3m > gr4m and gr4m < 0:
        return "Decelerating"
    if q1_gr >= 0.05 and gr4m >= 0:
        return "Steady Growth"
    if q1_gr < 0 and gr4m >= 0.05:
        return "Turnaround"
    if q1_gr < 0 and gr4m < 0:
        return "Sustained Decline"
    return "Volatile / Flat"


def growth_driver(val_from_vol, val_from_price, net_val_delta=None):
    """Growth Driver Legend, exactly as published on the live site's Vol vs
    Value tabs -- sign-based on Volume effect (V), Price effect (P), and Net
    value change, not a volume-share heuristic."""
    v = val_from_vol or 0
    p = val_from_price or 0
    net = net_val_delta if net_val_delta is not None else (v + p)
    if v == 0 and p == 0 and net == 0:
        return "NA"
    if v > 0 and p > 0:
        return "Both Driving"
    if v < 0 and p < 0:
        return "Both Declining"
    if v > 0 and p <= 0:
        if abs(v) > abs(p):
            return "Volume Led"
        if net < 0:
            return "Volume Cushioning Decline"
        return "Volume Led"
    if p > 0 and v <= 0:
        if abs(p) > abs(v):
            return "Price Led"
        if net < 0:
            return "Price Cushioning Decline"
        return "Price Led"
    return "NA"


# ==============================================================================
# 1. CEO Executive Summary
# ==============================================================================

# Live's exact KPI-tile icon per position (its own lucide-icon set, verified
# 2026-07-31 by reading each tile's icon class directly off the live DOM: ₹ =
# lucide-indian-rupee, the growth tiles = lucide-trending-up, No-of-FCL AND
# Gross-Margin both = lucide-layers (rendered here as the same 📦 stand-in for
# consistency), and the 3 bottom-row reference tiles all = lucide-chart-column).
# Applied by position, not by re-deriving from the sheet, since some of these
# labels come with no icon at all in the source cell and one (Realization
# Growth) came with the wrong one (💲) baked in.
KPI_ICONS = ['₹', '📈', '📦', '🌐', '📈', '📦', '📊', '📊', '📊']


def _apply_kpi_icon(label, icon):
    """Strip whatever leading icon/emoji (if any) the sheet cell came with and
    replace it with the verified-correct one for this tile position."""
    stripped = re.sub(r'^[^\w\n]+\s*', '', label)
    return f'{icon} {stripped}'


def build_exec(ws, export_ws=None, category_margin_total=None):
    """
    A handful of these 9 KPI tiles need more than a straight cell read -- the
    sheet's own pre-formatted text cells (A5/D5/H5/D6) round to whole numbers,
    truncate the label, or hold what live actually renders as the SUBTITLE
    rather than the LABEL. Verified cell-by-cell against the live site's own
    DOM (2026-07-31) after live's displayed KPI tiles were found to have
    regressed back to these raw-cell values -- a prior session's fix had only
    ever patched the generated HTML directly, not this function, so it didn't
    survive the next `py refresh_ceo_dashboard.py` run. Fixed at the source
    this time:
      - Revenue's PY sub (A5) is pre-rounded to a whole number by the sheet
        -- rebuilt from J7 (the prior-year KPI tile's own precise value, to
        one decimal) instead of trusting the display text.
      - Revenue Growth's label (D3) carries a trailing "Vs Q1" the live site
        drops ("Q1 Revenue Growth", not "...Growth Vs Q1"); its sub (D5) is a
        hand-typed annotation ("↑ Growing") live doesn't use at all -- live
        shows a fixed period-comparison caption instead.
      - No of FCL's value (G4) is a bare count; the reference view appends
        "FCL" and shows one decimal.
      - Export Value's PY sub (H5) is likewise pre-rounded by the sheet --
        the one extra decimal of precision only exists on the separate
        'Containers & Export value' sheet, not here.
      - Realization Growth's value (J4) needs a forced sign + the "₹.../kg"
        unit the reference view always shows, not a bare signed number.
      - "Q1 Growth vs Q4"'s label/sub are actually SWAPPED from what live
        shows: D6 ("🔢  Q1 FY27\\nGrowth vs Q4") is live's SUBTITLE, not its
        label -- live's actual label is the fixed "Q1 Growth vs Q4 FY26".
      - Q4 FY26 / Q1 FY26's values (G7/J7) render as plain grouped integers on
        live (plain grouped integers -- no ₹ symbol, no decimal), not money-
        formatted; both need a "Sales (₹L)" subtitle live shows and this
        sheet has no cell for at all.
    """
    gross_margin_kpi = {"label": txt(ws, "A6"), "value": num(ws, "A7"), "sub": txt(ws, "A8"), "tone": "red", "fmt": "pct"}
    if category_margin_total is not None:
        gross_margin_kpi = build_gross_margin_kpi(category_margin_total, txt(ws, "A6"))

    q1_fy26_revenue = num(ws, "J7")
    revenue_growth_label = re.sub(r'\s*[Vv]s\.?\s+Q\d\s*$', '', txt(ws, "D3"))
    fcl_value = num(ws, "G4")
    realization = num(ws, "J4")
    if export_ws is not None:
        export_py = num(export_ws, "D4")
        export_sub = f"vs ₹{export_py:,.1f} L PY"
    else:
        export_sub = txt(ws, "H5")

    kpis = [
        {"label": txt(ws, "A3"), "value": num(ws, "A4"), "sub": f"vs ₹{q1_fy26_revenue:,.1f} L PY", "tone": "navy", "fmt": "money"},
        {"label": revenue_growth_label, "value": num(ws, "D4"), "sub": "4M FY27 vs 4M FY26", "fmt": "pct"},
        {"label": txt(ws, "G3"), "value": f"{fcl_value:,.1f} FCL", "sub": txt(ws, "G5"), "tone": "green"},
        {"label": txt(ws, "H3"), "value": num(ws, "H4"), "sub": export_sub, "tone": "green", "fmt": "money"},
        {"label": txt(ws, "J3"), "value": f"{'+' if realization >= 0 else ''}₹{realization:.1f}/kg", "sub": txt(ws, "J5"), "fmt": "num1"},
        gross_margin_kpi,
        # D6/D7 no longer hold a growth% vs Q4 -- the sheet repurposed this tile into a
        # 3rd plain revenue reference figure, same shape as the G6/G7 and J6/J7 tiles
        # beside it (confirmed 2026-08: D7 is now an absolute ₹L revenue value, D6 a
        # bare period-name label, not a growth sub-caption). Mirror those two tiles
        # rather than rendering a (now nonexistent) growth percentage.
        {"label": txt(ws, "D6"), "value": f"{num(ws, 'D7'):,.0f}", "sub": "Sales (₹L)", "tone": "navy"},
        {"label": txt(ws, "G6"), "value": f"{num(ws, 'G7'):,.0f}", "sub": "Sales (₹L)", "tone": "teal"},
        {"label": txt(ws, "J6"), "value": f"{num(ws, 'J7'):,.0f}", "sub": "Sales (₹L)", "tone": "purple"},
    ]
    for kpi, icon in zip(kpis, KPI_ICONS):
        kpi["label"] = _apply_kpi_icon(kpi["label"], icon)
    category = rows_block(ws, 12, 19, "ABCD", ["name", "cur", "py", "grPct"])
    brand = rows_block(ws, 12, 19, "GHIJ", ["name", "cur", "py", "grPct"])
    zone = fix_zone_names(rows_block(ws, 23, 28, "ABCD", ["name", "cur", "py", "grPct"]))
    salesman = merge_salesmen(rows_block(ws, 23, 35, "GHIJ", ["name", "cur", "py", "grPct"]))
    for block in (category, brand, zone, salesman):
        for r in block:
            r["signal"] = signal_from_growth(r["grPct"])
    return {"kpis": kpis, "category": category, "brand": brand, "zone": zone, "salesman": salesman}


# ==============================================================================
# 2. Profitability_1
# ==============================================================================
# Profitability_1's own entity codes, cleaned up to board-ready display
# names (the sheet uses abbreviated internal codes, e.g. 'VIBRANT ' with a
# trailing space) -- ADFL/ADFIL/Elimination/CONSO TOTAL kept as-is, matching
# this dashboard's own long-standing naming for those.
PROFIT_ENTITY_NAME_MAP = {
    "TFL": "Telluric Foods", "Australia": "ADF Foods Australia",
    "ADFUKL": "ADF Foods UK", "ADFUSAHLD": "ADF Holdings USA",
    "ADFUSA": "ADF Foods USA", "VIBRANT": "Vibrant Foods NJ LLC",
    "CONSO TOTAL": "Total",
}


def build_profit(ws):
    keys = ["name", "revenue", "ebitda", "ebitdaPct", "pbt", "pbtPct", "pat", "patPct", "pyRevenue", "pyEbitda", "pyPbt"]
    entities = rows_block(ws, 5, 14, "ABCDEFGHIJK", keys)
    for e in entities:
        raw = str(e["name"]).strip()
        e["name"] = PROFIT_ENTITY_NAME_MAP.get(raw, raw)
    keys2 = ["name", "kgs", "cif", "pricePerKg", "cogsPerKg", "margin", "marginPct", "kgsPy", "cifPy", "pricePerKgPy", "marginPctPy", "marginDeltaPp"]
    category_margin = rows_block(ws, 19, 27, "ABCDEFGHIJKL", keys2)
    return {"entities": entities, "categoryMargin": category_margin}


# ==============================================================================
# 3. Vol vs Value & Price
# ==============================================================================
def build_vol(ws):
    kpis = [
        {"label": txt(ws, "A3"), "value": num(ws, "A4"), "sub": txt(ws, "A5"), "signColor": True, "fmt": "pct"},
        {"label": txt(ws, "D3"), "value": num(ws, "D4"), "sub": txt(ws, "D5"), "signColor": True, "fmt": "pct"},
        {"label": txt(ws, "G3"), "value": f"₹{num(ws, 'G4'):.1f}/kg", "sub": txt(ws, "G5"), "signColor": True, "fmt": "num1"},
        {"label": txt(ws, "J3"), "value": num(ws, "J4"), "sub": txt(ws, "J5") + " (4M)", "signColor": True, "fmt": "money"},
    ]
    keys = ["name", "volCur", "volPy", "volGrPct", "valCur", "valPy", "valGrPct", "priceCur", "pricePy",
            "priceDelta", "valFromVol", "valFromPrice", "netValDelta"]
    decomposition = rows_block(ws, 9, 17, "ABCDEFGHIJKLM", keys)
    for r in decomposition:
        r["growthDriver"] = growth_driver(r["valFromVol"], r["valFromPrice"], r["netValDelta"])
    keys2 = ["name", "priceCur", "pricePy", "priceDelta", "priceGrPct", "cogsCur", "cogsPy", "cogsDelta", "spread", "spreadPy", "spreadDelta"]
    cogs_spread = rows_block(ws, 21, 29, "ABCDEFGHIJK", keys2)
    return {"kpis": kpis, "decomposition": decomposition, "cogsSpread": cogs_spread}


# ==============================================================================
# 4. Vol vs Value & Price _DC (by invoice currency)
# ==============================================================================
def build_vol_dc(ws):
    kpis = [
        {"label": txt(ws, "A3"), "value": num(ws, "A4"), "sub": txt(ws, "A5"), "signColor": True, "fmt": "pct"},
        {"label": txt(ws, "D3"), "value": num(ws, "D4"), "sub": "4M FY27 vs 4M FY26", "signColor": True, "fmt": "pct"},
        {"label": txt(ws, "J3"), "value": f"₹{num(ws, 'J4'):.1f}/kg", "sub": txt(ws, "J5"), "signColor": True, "fmt": "num1"},
        {"label": txt(ws, "M3"), "value": num(ws, "M4"), "sub": txt(ws, "M5"), "signColor": True, "fmt": "money"},
    ]
    keys = ["name", "volCur", "volPy", "volGrPct", "docValCur", "docValPy", "valGrLocalPct",
            "valCurInr", "valPyInr", "valGrPct", "priceCur", "pricePy", "priceDelta",
            "valFromVol", "valFromPrice", "netValDelta"]
    cols = "ABCDEFGHIJKLMNOP"

    def block(r1, r2):
        rows = rows_block(ws, r1, r2, cols, keys)
        for r in rows:
            r["growthDriver"] = growth_driver(r["valFromVol"], r["valFromPrice"], r["netValDelta"])
        return rows

    return {
        "kpis": kpis,
        "currencies": {
            "USD": block(9, 17),
            "GBP": block(22, 30),
            "INR": block(35, 40),
            "ALL": block(45, 53),
        },
    }


# ==============================================================================
# 5. Sales Dashboard
# ==============================================================================
def build_sales(ws):
    # Workbook widened from a 3-month Q1 cut to a 4-month YTD cut (Apr-Jul) --
    # a 4th current-month/PY-month/growth% column was inserted after each of the
    # old 3-month groups, shifting everything after. Column O's header text is a
    # stale "Jun-27 Gr%" copy-paste leftover in the SOURCE sheet itself (verified
    # against columns E/J: the underlying data is unambiguously July, not June) --
    # keyed by position here, not by trusting that header string.
    keys = ["name", "apr", "may", "jun", "jul", "ytdCur", "aprPy", "mayPy", "junPy", "julPy", "ytdPy",
            "aprGr", "mayGr", "junGr", "julGr", "q1Gr", "mixPct", "q4Py"]
    cols = "ABCDEFGHIJKLMNOPQR"
    return {
        "category": rows_block(ws, 5, 13, cols, keys),
        "brand": rows_block(ws, 17, 25, cols, keys),
        "zone": fix_zone_names(rows_block(ws, 29, 35, cols, keys)),
        "salesman": merge_salesmen(rows_block(ws, 39, 52, cols, keys)),
    }


def rank_by_value(rows):
    """Sort a Top-10/Tail-style block by its own 'cur' value, descending, and
    recompute 'cumPct' as a fresh running sum in that corrected order.

    The source sheet's own row order is supposed to already be descending by
    value (that's the entire premise of a "Top" list, and cumPct is meant to
    accumulate as you read down the list) but was found to have several rows
    genuinely out of order in this cut (verified 2026-08 via independent
    exhaustive re-check: e.g. Chutney sat below the smaller Pickles row,
    Khansaama Brand below the smaller Soul Brand row, and a 5-row stretch of
    the Salesman block was jumbled). 'sharePct' is a static ratio (this row's
    value over the block's total) so it doesn't depend on row order and is
    left as-is; 'cumPct' does depend on order, so it's recalculated here
    rather than carried over from the sheet's own (mis-ordered) running sum --
    otherwise it would stop increasing monotonically top-to-bottom once the
    rows are put back in the right order."""
    rows = sorted(rows, key=lambda r: r["cur"] or 0, reverse=True)
    running = 0.0
    for r in rows:
        running += r["sharePct"] or 0
        r["cumPct"] = round(running, 6)
    return rows


# ==============================================================================
# 6. Top 10 & Tail
# ==============================================================================
def build_top(ws):
    # The Q4 FY26 reference and QoQ growth% columns were dropped from this sheet
    # entirely in the 4-month YTD cut -- no longer computed anywhere in the source,
    # not just relabeled. Removed here to match; the dashboard UI drops the
    # corresponding table columns too.
    keys = ["name", "cur", "py", "absDelta", "grYoy", "sharePct", "cumPct"]
    cols = "ABCDEFG"
    result = {
        "category": rank_by_value(rows_block(ws, 5, 12, cols, keys)),
        "brand": rank_by_value(rows_block(ws, 17, 24, cols, keys)),
        "zone": rank_by_value(fix_zone_names(rows_block(ws, 29, 34, cols, keys))),
        "salesman": rank_by_value(merge_salesmen(rows_block(ws, 39, 51, cols, keys))),
    }
    for rows in result.values():
        for r in rows:
            r["status"] = signal_from_growth(r["grYoy"])
    return result


# ==============================================================================
# 7. Top 10 & Tail_Sub category
# ==============================================================================
def build_top_sub(ws):
    # Same Q4/QoQ column removal as build_top() -- see comment there.
    keys = ["subcategory", "category", "cur", "py", "absDelta", "grYoy", "sharePct", "cumPct"]
    cols = "ABCDEFGH"
    result = {
        "top10": rank_by_value(rows_block(ws, 5, 14, cols, keys)),
        "tail10": rank_by_value(rows_block(ws, 19, 28, cols, keys)),
    }
    for rows in result.values():
        for r in rows:
            r["status"] = signal_from_growth(r["grYoy"])
    return result


# ==============================================================================
# 8. Growth Analysis
# ==============================================================================
def build_growth(ws):
    # Gained a new trailing "4M Gr%" column (I) alongside the existing 1M/2M/3M --
    # NOTE this is a DIFFERENT metric from column E despite sharing the identical
    # header text "4M Gr%": E is the overall YoY change for the whole 4-month cut,
    # I is the rolling/trailing 4-month growth rate (same family as 1M/2M/3M).
    # Disambiguated by column position, not header text.
    keys = ["name", "cur", "py", "absDelta", "grYoy", "gr1m", "gr2m", "gr3m", "gr4m"]
    cols = "ABCDEFGHI"
    result = {
        "category": rows_block(ws, 5, 12, cols, keys),
        "brand": rows_block(ws, 17, 24, cols, keys),
        "zone": fix_zone_names(rows_block(ws, 29, 34, cols, keys)),
        "salesman": merge_salesmen(rows_block(ws, 38, 50, cols, keys)),
    }
    for rows in result.values():
        for r in rows:
            r["trend"] = trend_from_growth(r["grYoy"], r["gr2m"], r["gr3m"], r["gr4m"])
    return result


# ==============================================================================
# Splice into HTML (same pattern as Dashboard/refresh_dashboard.py)
# ==============================================================================
def splice_const(html_content, const_name, value_obj):
    payload = json.dumps(value_obj, ensure_ascii=False, separators=(",", ":"))
    new_line = f"const {const_name} = {payload};\n"
    pattern = re.compile(r"const " + re.escape(const_name) + r" = \{.*?\};\n", re.S)
    if pattern.search(html_content):
        return pattern.sub(lambda m: new_line, html_content, count=1)
    marker = "/* CEO_DATA_INSERT_MARKER */"
    if marker not in html_content:
        raise RuntimeError(f"Could not find insertion marker for {const_name}")
    return html_content.replace(marker, new_line + marker, 1)


def main():
    sys.stdout.reconfigure(encoding="utf-8")

    if not HTML_PATH.exists():
        print(f"ERROR: dashboard HTML not found at {HTML_PATH}")
        sys.exit(1)

    source_path = find_source_xlsx()
    print(f"Reading {os.path.basename(source_path)} ...")
    wb = openpyxl.load_workbook(source_path, data_only=True)

    for key, name in SHEETS.items():
        if name not in wb.sheetnames:
            print(f"WARNING: sheet '{name}' not found in workbook -- skipping {key}.")

    # build_fs_entity_pnl() (board FS-sourced, Master Consolidation pipeline) is
    # NOT used here anymore -- as of 2026-08-09 it's still Q1 FY27 (Apr-Jun)
    # only, since the board FS is prepared quarterly and hasn't caught up to a
    # July cut. profit["entities"] below comes straight from build_profit()'s
    # own read of the Profitability_1 sheet, which is now genuinely 4M FY27
    # (Apr-Jul), correctly labeled, and independently verified 2026-08-09 to
    # tie out exactly to the audited-basis Consolidated MIS figures used
    # elsewhere (Revenue from op./EBITDA/PBT/PAT) -- the earlier "mislabeled
    # Q1, actually 2 months" data-quality problem build_fs_entity_pnl()'s
    # docstring describes does not appear to affect the current cut.
    profit = build_profit(wb[SHEETS["profit"]])
    profit["entitiesPeriodLabel"] = "4M FY27 (Apr-Jul 2026), from Profitability_1 sheet"

    export_ws = wb["Containers & Export value"] if "Containers & Export value" in wb.sheetnames else None
    category_margin_total = next((r for r in profit["categoryMargin"] if str(r.get("name", "")).upper() == "TOTAL"), None)
    data = {
        "sourceFile": os.path.basename(source_path),
        "exec": build_exec(wb[SHEETS["exec"]], export_ws, category_margin_total),
        "profit": profit,
        "vol": build_vol(wb[SHEETS["vol"]]),
        "volDc": build_vol_dc(wb[SHEETS["vol_dc"]]),
        "sales": build_sales(wb[SHEETS["sales"]]),
        "top": build_top(wb[SHEETS["top"]]),
        "topSub": build_top_sub(wb[SHEETS["top_sub"]]),
        "growth": build_growth(wb[SHEETS["growth"]]),
    }

    html = HTML_PATH.read_text(encoding="utf-8")
    html = splice_const(html, "CEO_DATA", data)
    HTML_PATH.write_text(html, encoding="utf-8")
    print(f"Done. Wrote {HTML_PATH}")


if __name__ == "__main__":
    main()
