"""
build_ceo_dashboard_conso.py

Regenerates the Consolidated view of the CEO dashboard -- the all-entity
mirror of the Standalone view -- from read-only source workbooks.

Two classes of source are used, for two different purposes:

  * A transaction-level extract, covering every group entity. Used for the
    Category, Brand and Zone breakdowns, which need row-level detail.

  * The management reporting pack. Used for the headline KPI tiles and the
    entity profit-and-loss table, because it is the reconciled figure.

Why both: the transaction extract sums entities before intercompany
elimination, so its grand total is a pre-elimination number. The management
pack carries the eliminated, reconciled total. The dashboard shows the
reconciled total in its headline tiles and uses the transaction extract only
where a dimensional breakdown is required.

Design notes:

  * Consolidated figures are derived from transaction rows and reconciled to
    the management report, rather than read from a pre-built pivot in the
    source. Pre-built aggregates in spreadsheets can carry a stale filter;
    deriving from rows and reconciling makes any divergence visible.

  * Aggregation uses the workbook's canonical grouping columns rather than
    its raw labels, because the raw label columns contain near-duplicate
    spellings that would split a single group across several rows.

  * Growth signal, growth driver and trend classifications are computed here
    rather than read from the source, so that one definition applies across
    every entity and period.

  * Entity-to-zone mapping treats a single-country subsidiary as belonging to
    its own country's zone, but not the parent entity, which sells across
    every zone and needs its real zone split read from the sales extract.

  * All source locations are read from environment configuration. Nothing in
    this file points at a specific machine, share or folder.

Output is a single self-contained HTML file with no runtime data dependency.
"""

import json
import os
import re
import sys
from pathlib import Path
from collections import defaultdict

import openpyxl

BASE = Path(__file__).resolve().parent
XLSM_PATH = BASE / "CEO Standalone and Conso Dashboard Data_V3.xlsm"
# Board MIS workbook -- absolute location supplied by the environment, never
# hard-coded (CONSO_MIS_WORKBOOK).
MIS_PATH = Path(os.environ.get("CONSO_MIS_WORKBOOK", ""))
OUT_HTML_PATH = BASE / "CEO_Dashboard_Conso.html"

MONTHS = ["Apr", "May", "Jun", "Jul"]

# Verified 2026-08 against the transaction extract’s own ENTITY column values exactly.
# 2026-08-11: this is still used as-is for the other 6 entities (each is a
# single-country subsidiary, so its own country IS its real zone), but NO
# LONGER for "ADF Foods Ltd (Standalone)" -- Standalone is predominantly an exporter
# whose sales actually span all 6 zones, not just India; its real zone-wise
# split is pulled in separately by load_standalone_zone_rows() below and
# merged in main(), so its ENTITY_ZONE entry here is unused (kept only so
# ENTITY_ZONE.get(entity, "Other") never has to special-case a missing key).
ENTITY_ZONE = {
    "ADF Foods Ltd (Standalone)": "India",
    "Telluric Foods": "India",
    "ADF Holdings USA": "North America",
    "ADF Foods USA": "North America",
    "Vibrant Foods NJ LLC": "North America",
    "ADF Foods UK": "UK",
    "ADF Foods Australia": "Asia Pacific",
}
# Filter-pane display order, grouped by zone.
ENTITY_DISPLAY_ORDER = [
    "ADF Foods Ltd (Standalone)", "Telluric Foods",
    "ADF Holdings USA", "ADF Foods USA", "Vibrant Foods NJ LLC",
    "ADF Foods UK", "ADF Foods Australia",
]

# Canonical taxonomy = the workbook's own FINANCIAL GROUP / FINANCIAL BRAND
# columns (verified 2026-08: 10 and 9 distinct values respectively), listed
# with Standalone's own 8 first for a familiar order, then the Group-only
# extras appended.
CATEGORY_ORDER = ["FROZEN FOODS", "READY TO EAT", "CHUTNEY", "PICKLES", "PASTE & SAUCES", "SPICES", "OTHERS", "TAMARIND", "GHEE", "TEA"]
BRAND_ORDER = ["ASHOKA BRAND", "OTHERS BRAND", "UNBRANDED", "CAMEL BRAND", "TRULY INDIAN", "AEROPLANE BRAND", "KHANSAAMA BRAND", "SOUL BRAND"]

# the transaction extract’s FINANCIAL BRAND column has both 'SOUL' and 'SOUL BRAND' as
# distinct raw values across different entities/rows -- confirmed 2026-08 by
# a reader to be the same brand, not two real ones (a data-entry
# inconsistency in the source workbook, not a genuine second brand).
# Normalized here at ingestion so they aggregate as one.
BRAND_NAME_FIX = {"SOUL": "SOUL BRAND"}
ZONE_ORDER = ["North America", "UK", "Middle East", "Europe", "Asia Pacific", "India"]

# Standalone's own workbook (same one refresh_ceo_dashboard.py reads) has a
# real, monthly-granular Zone-wise-sales breakdown across these same 6
# zones -- added 2026-08-11 so Conso's Zone table isn't just a 4-zone
# entity-location proxy that's silently missing Middle East/Europe
# entirely (no Conso entity is based there, but Standalone sells there).
# Same short-form labels as refresh_ceo_dashboard.py's ZONE_NAME_FIX, kept
# as a deliberate small duplication rather than importing that module.
STANDALONE_XLSX_PATH = BASE / os.environ.get("STANDALONE_SALES_WORKBOOK", "")
STANDALONE_ZONE_NAME_FIX = {
    "NORTH AMERICA": "North America", "UNITED KINGDOM": "UK",
    "WESTERN EUROPE": "Europe", "GULF COUNTRIES": "Middle East",
    "ASIA PACIFIC": "Asia Pacific", "INDIA": "India",
}

# MIS entity-column order in `Summary-ADFL Conso 2025` -- confirmed by exact
# numeric match against the transaction extract (e.g. ADFIL column = 'ADF Foods Australia',
# not a separate India entity as the name suggests).
MIS_ENTITY_TO_DISPLAY = {
    "ADFL": "ADF Foods Ltd (Standalone)",
    "ADFIL": "ADF Foods Australia",
    "TELLURIC": "Telluric Foods",
    "ADF UK": "ADF Foods UK",
    "ADFHL": "ADF Holdings USA",
    "ADF USA": "ADF Foods USA",
    "VIB NJ LLC": "Vibrant Foods NJ LLC",
}
MIS_ENTITY_COLS = {
    "Apr": list("CDEFGHIJ"), "May": list("KLMNOPQR"),
    "Jun": list("STUVWXYZ"), "Jul": "AA AB AC AD AE AF AG AH".split(),
}
MIS_ENTITY_ORDER = ["ADFL", "ADFIL", "TELLURIC", "ADF UK", "ADFHL", "ADF USA", "VIB NJ LLC", "Total"]

# ==============================================================================
# 12M (FY26 vs FY25) dataset -- added 2026-08-10, sourced ENTIRELY from a new
# audited Financial-Statement-basis workbook the operator provides (real FY25
# Group-wide data didn't exist anywhere before this). See module docstring
# for why this REPLACES the earlier extract-quarterly-based 12M approach
# rather than being layered alongside it (mixing pre-elimination quarterly
# sales with audited annual FS figures would produce a mismatched-basis
# growth% -- the exact class of bug already found and fixed once this
# session for the 4M KPI reference tiles).
# ==============================================================================
ANNUAL_XLSX_PATH = BASE / os.environ.get("ANNUAL_SALES_WORKBOOK", "")
FS_SEGMENT_PATH = BASE / os.environ.get("FS_SEGMENT_WORKBOOK", "")

ANNUAL_FY_CUR, ANNUAL_FY_PY = "FY 2025-26", "FY 2024-25"

# 'Region wise Sales Base' sheet's own entity spelling -> this script's
# canonical display names (verified 1:1, no ambiguity like the 4M MIS
# ADFIL/Australia mapping).
ANNUAL_ENTITY_FIX = {
    "Standalone": "ADF Foods Ltd (Standalone)",
    "ADF Holding USA": "ADF Holdings USA",
    "Vibrant LLC": "Vibrant Foods NJ LLC",
    "Telluric Foods Ltd": "Telluric Foods",
    "ADF Australia": "ADF Foods Australia",
    "ADF UK": "ADF Foods UK",
    # 'ADF Foods USA' already matches this script's own name.
}

# The sheet has case-inconsistent duplicates for a few values (TEA/Tea,
# SPICES/Spices, AUSTRALIA/Australia, INDIA/India) -- normalized to one
# canonical spelling per value so they aggregate as one, not two rows.
ANNUAL_CATEGORY_FIX = {
    "Frozen Breads": "FROZEN FOODS", "Frozen Snacks": "FROZEN FOODS", "Frozen Vegetables/Others": "FROZEN FOODS",
    "Ready to eat": "READY TO EAT", "Chutneys": "CHUTNEY", "Pickles": "PICKLES",
    "Spices": "SPICES", "SPICES": "SPICES", "Pastes& Sauce": "PASTE & SAUCES",
    "Tamarind": "TAMARIND", "Tea": "TEA", "TEA": "TEA", "Others": "OTHERS", "NPD": "OTHERS",
}
ANNUAL_CATEGORY_ORDER = ["FROZEN FOODS", "READY TO EAT", "CHUTNEY", "PICKLES", "PASTE & SAUCES", "SPICES", "TAMARIND", "TEA", "OTHERS"]

ANNUAL_BRAND_FIX = {
    "Ashoka": "ASHOKA BRAND", "Truly Indian": "TRULY INDIAN", "Camel": "CAMEL BRAND",
    "Aeroplane": "AEROPLANE BRAND", "Soul": "SOUL BRAND", "Khansama": "KHANSAAMA BRAND",
    "B2B brands": "OTHERS BRAND", "Others": "OTHERS BRAND",
}
ANNUAL_BRAND_ORDER = ["ASHOKA BRAND", "TRULY INDIAN", "CAMEL BRAND", "AEROPLANE BRAND", "SOUL BRAND", "KHANSAAMA BRAND", "OTHERS BRAND"]

# Region_SA is used, not the sheet's own 'ZONE' column (mostly blank/#N/A
# for most rows, confirmed 2026-08-10) -- Region_SA is customer-destination
# based (real export markets) and complete across all rows.
ANNUAL_REGION_FIX = {"AUSTRALIA": "Asia Pacific", "Australia": "Asia Pacific", "Others": "Asia Pacific",
                      "USA": "North America", "Canada": "North America",
                      "UK": "United Kingdom", "EU": "Western Europe", "GCC": "Gulf Countries",
                      "India": "India", "INDIA": "India"}
ANNUAL_REGION_ORDER = ["North America", "United Kingdom", "Western Europe", "Gulf Countries", "Asia Pacific", "India"]


# ==============================================================================
# Badge logic -- copied verbatim from refresh_ceo_dashboard.py (kept as a
# deliberate, small duplication rather than refactoring that already-verified
# file just for this one caller).
# ==============================================================================
def signal_from_growth(gr_pct):
    gr = gr_pct or 0
    if gr_pct is None:
        return "NA"
    if gr < 0:
        return "Degrowing"
    if gr < 0.05:
        return "Slow Growing"
    if gr < 0.15:
        return "Growing"
    return "Strong Growing"


def trend_from_growth(q1_gr, gr2, gr3, gr4):
    if gr4 is None:
        return "NA"
    q1_gr, gr2, gr3, gr4 = q1_gr or 0, gr2 or 0, gr3 or 0, gr4 or 0
    if gr2 < gr3 < gr4 and gr4 >= 0:
        return "Accelerating"
    if gr2 > gr3 > gr4 and gr4 < 0:
        return "Decelerating"
    if q1_gr >= 0.05 and gr4 >= 0:
        return "Steady Growth"
    if q1_gr < 0 and gr4 >= 0.05:
        return "Turnaround"
    if q1_gr < 0 and gr4 < 0:
        return "Sustained Decline"
    return "Volatile / Flat"


def gr(cur, py):
    """None when there's genuinely no prior-year base to grow from, rather
    than a fabricated 0% or a divide-by-zero -- renders as '-' client-side,
    not a misleading 0%."""
    if py:
        return (cur - py) / py
    return None


def rank_by_value(rows):
    """Sort descending by 'cur' and recompute 'cumPct' as a fresh running
    sum in that order -- see refresh_ceo_dashboard.py's identical helper and
    the sort-order bug it was written to fix; applied proactively here."""
    rows = sorted(rows, key=lambda r: r["cur"] or 0, reverse=True)
    running = 0.0
    for r in rows:
        running += r["sharePct"] or 0
        r["cumPct"] = round(running, 6)
    return rows


# ==============================================================================
# 1. the transaction extract -- raw transaction rows (cols A-M only; N onward is
#    unrelated scratch/documentation text baked into unused columns, not data)
# ==============================================================================
def load_mis_sales_by_entity_month():
    """Sales (row 15, `Summary-ADFL Conso 2025`) per entity per month for the
    current FY27 4M period -- the P&L's own control total, used by
    load_conso_data() to rescale the transaction extract’s raw entity-month figures. See
    that function's docstring for why."""
    wb = openpyxl.load_workbook(MIS_PATH, data_only=True)
    ws = wb["Summary-ADFL Conso 2025"]
    out = defaultdict(dict)
    for month, cols in MIS_ENTITY_COLS.items():
        for code, col in zip(MIS_ENTITY_ORDER, cols):
            if code not in MIS_ENTITY_TO_DISPLAY:
                continue
            v = ws[f"{col}15"].value
            if v is not None:
                out[MIS_ENTITY_TO_DISPLAY[code]][month] = out[MIS_ENTITY_TO_DISPLAY[code]].get(month, 0.0) + v
    return out


def load_conso_data():
    """2026-08-10: V3's raw entity-month totals for ADF Holdings USA / ADF
    Foods USA (Apr-Jun) no longer tie to the MIS-sourced Entity P&L the way
    V2's did -- a full per-entity, per-month reconciliation found material
    gaps for those two entities (July always matched exactly).
    Per explicit user instruction ("take as per the P&L only"), every FY27
    row here is rescaled by a per-(entity, month) factor = MIS Sales /
    the transaction extract’s own raw sum for that entity-month, so every downstream
    Category/Brand/Zone/Sub-category table ties exactly to the P&L again --
    the transaction extract still supplies the category/brand MIX (its actual strength),
    the P&L supplies the LEVEL. 'Consolidation Adjustment' -- a new
    pseudo-entity in V3 with no MIS/P&L counterpart -- is dropped entirely
    for the same reason (the P&L doesn't recognize it, so neither do we).
    PY (FY26) figures are left untouched -- no MIS entity-month history
    exists to rescale against; the transaction extract remains the sole PY source, same
    as before.
    """
    wb = openpyxl.load_workbook(XLSM_PATH, data_only=True, keep_vba=False)
    ws = wb["the transaction extract"]
    rows = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        entity, fy, qtr, month = row[0], row[1], row[2], row[3]
        if entity == "Consolidation Adjustment":
            continue
        sub_group = row[5].strip() if isinstance(row[5], str) else row[5]
        financial_group, financial_brand = row[6], row[7]
        financial_brand = BRAND_NAME_FIX.get(financial_brand, financial_brand)
        value = row[9]
        if value is None or fy not in ("FY25", "FY26", "FY27"):
            continue
        rows.append({
            "entity": entity, "fy": fy, "qtr": qtr, "month": month,
            "category": financial_group, "subcategory": sub_group, "brand": financial_brand,
            "zone": ENTITY_ZONE.get(entity, "Other"),
            "value": float(value),
            "customerType": row[11],
        })

    conso_sum = defaultdict(lambda: defaultdict(float))
    for r in rows:
        if r["fy"] == "FY27" and r["month"] in MONTHS:
            conso_sum[r["entity"]][r["month"]] += r["value"]

    mis_sales = load_mis_sales_by_entity_month()
    factor = defaultdict(lambda: defaultdict(lambda: 1.0))
    for entity, months in conso_sum.items():
        for month, total in months.items():
            mis_val = mis_sales.get(entity, {}).get(month)
            if mis_val is not None and total:
                factor[entity][month] = mis_val / total

    for r in rows:
        if r["fy"] == "FY27" and r["month"] in MONTHS:
            r["value"] *= factor[r["entity"]][r["month"]]

    return rows


def filter_entity(rows, entity):
    if entity == "ALL":
        return rows
    return [r for r in rows if r["entity"] == entity]


def compute_true_rescale_factors(rows, true_kpis):
    """2026-08-11, per explicit user instruction ('all the totals should
    match exactly as per the total sales'): external_only()'s RELATED PARTY
    exclusion ties CLOSE to the true, audited consolidated Sales figure
    (from the consolidated management sheet) but not exactly -- the transaction-
    level flag and the formal consolidation elimination entry are two
    independently-maintained numbers. This computes a second rescale
    factor, same technique as load_conso_data()'s MIS rescale: a
    multiplicative factor per period-bucket, computed from `rows`' own
    external-only sum vs the true figure, so the total ties exactly once
    applied (via apply_rescale_factors() below).

    MUST be computed from the FULL external-only dataset (all 7 entities'
    worth) -- calling this on a subset (e.g. just Standalone's zone rows)
    would divide the TRUE GROUP total by only a fraction of external sales,
    producing a wildly inflated factor. Compute once here, then apply the
    same factors to every row-set that needs to tie to the true total
    (Category/Brand's conso_rows_ext AND Standalone's separately-sourced
    zone rows alike).

    Only two true-figure data points exist in the MIS source for the
    current 4M window -- July 2026 alone, and the YTD-FY27 (Apr-Jul)
    cumulative -- no separate April/May/June split. So the rescale uses two
    buckets: July on its own (exact), and Apr+May+Jun combined as one
    (exact combined, not individually) -- same derivation already used for
    the 'Q1 FY27' reference KPI tile (YTD minus July). Same logic mirrored
    for the FY26 (PY) side using July 2025 / YTD-FY26.
    """
    sales_cur, sales_py, july_cur, july_py = true_kpis
    amj_cur_true = sales_cur - july_cur
    amj_py_true = sales_py - july_py

    def bucket_sum(fy, months):
        return sum(r["value"] for r in rows if r["fy"] == fy and r["month"] in months)

    ext_july_cur, ext_amj_cur = bucket_sum("FY27", ("Jul",)), bucket_sum("FY27", ("Apr", "May", "Jun"))
    ext_july_py, ext_amj_py = bucket_sum("FY26", ("Jul",)), bucket_sum("FY26", ("Apr", "May", "Jun"))

    return {
        ("FY27", "Jul"): (july_cur / ext_july_cur) if ext_july_cur else 1.0,
        ("FY27", "amj"): (amj_cur_true / ext_amj_cur) if ext_amj_cur else 1.0,
        ("FY26", "Jul"): (july_py / ext_july_py) if ext_july_py else 1.0,
        ("FY26", "amj"): (amj_py_true / ext_amj_py) if ext_amj_py else 1.0,
    }


def apply_rescale_factors(rows, factors):
    out = []
    for r in rows:
        r = dict(r)
        if r["fy"] in ("FY27", "FY26"):
            bucket = "Jul" if r["month"] == "Jul" else "amj"
            r["value"] *= factors[(r["fy"], bucket)]
        out.append(r)
    return out


def load_standalone_zone_rows(conso_rows_ext):
    """Standalone's real, monthly-granular Zone-wise-sales split (same
    workbook refresh_ceo_dashboard.py reads for the Standalone dashboard).

    2026-08-12, per explicit user instruction ("not even single decimal
    should be missed"): rescaled so each (fy, month) bucket's sum is EXACT
    to Standalone's own already-fully-corrected total in `conso_rows_ext`
    (the same external-only, true-rescaled figure Category/Brand use) --
    factor = target / raw_zone_sum, applied per zone. This replaces an
    earlier two-step approximation (an overall external-sales ratio, then
    a separate group-level true-rescale) that left a small residual: two
    independently-rounded paths to the same number don't land on bit-
    identical results. Scaling directly to the known-correct target
    guarantees Zone's grand total ties to Category/Brand's exactly, by
    construction -- sum(val * (target/raw)) = target, always.

    `conso_rows_ext` must already be the FULLY rescaled external-only rows
    (post external_only() + apply_rescale_factors()) -- passing the raw,
    unscaled conso_rows here would target the wrong (gross) number.
    """
    target = defaultdict(lambda: defaultdict(float))
    for r in conso_rows_ext:
        if r["entity"] == "ADF Foods Ltd (Standalone)":
            target[r["fy"]][r["month"]] += r["value"]

    wb = openpyxl.load_workbook(STANDALONE_XLSX_PATH, data_only=True)
    ws = wb["\U0001F4CA Sales Dashboard"]
    keys = ["name", "apr", "may", "jun", "jul", "ytdCur", "aprPy", "mayPy", "junPy", "julPy", "ytdPy",
            "aprGr", "mayGr", "junGr", "julGr", "q1Gr", "mixPct", "q4Py"]
    cols = "ABCDEFGHIJKLMNOPQR"
    raw = []  # (zone, fy, month, raw_value)
    for r in range(29, 36):
        row = {col: ws[f"{col}{r}"].value for col in cols}
        vals = dict(zip(keys, (row[c] for c in cols)))
        zone = STANDALONE_ZONE_NAME_FIX.get(str(vals["name"]).strip().upper()) if vals["name"] else None
        if not zone:
            continue
        for month, key in zip(MONTHS, ["apr", "may", "jun", "jul"]):
            cur_val = vals[key]
            if isinstance(cur_val, (int, float)):
                raw.append((zone, "FY27", month, float(cur_val)))
            py_val = vals[f"{key}Py"]
            if isinstance(py_val, (int, float)):
                raw.append((zone, "FY26", month, float(py_val)))

    raw_sum = defaultdict(lambda: defaultdict(float))
    for zone, fy, month, val in raw:
        raw_sum[fy][month] += val

    out = []
    for zone, fy, month, val in raw:
        r_sum = raw_sum[fy].get(month, 0.0)
        factor = (target[fy].get(month, 0.0) / r_sum) if r_sum else 0.0
        # qtr:None -- q4_fy26_by_dim() only cares about real qtr='Q4' rows
        # (the hidden Q4 FY26 reference column); no such data is sourced
        # here, so these rows just need the key to exist.
        out.append({"entity": "ADF Foods Ltd (Standalone)", "fy": fy, "qtr": None, "month": month,
                    "zone": zone, "value": val * factor, "customerType": "OTHER PARTY"})
    return out


def external_only(rows):
    """2026-08-11, per explicit user instruction: Category/Brand/Zone/
    Sub-category/Top/Growth tables should show the TRUE, elimination-
    adjusted figures, not the pre-elimination gross sum-of-entities basis --
    confirmed by an independent user-supplied reconciliation sheet
    ('Q1 FY27 PROPORTIONATELY ALIGNED TO MIS'), which ties almost exactly
    (within normal rounding) to excluding the transaction extract’s own 'RELATED PARTY'
    rows (CUSTOMER TYPE column) -- the actual intercompany flag, discovered
    2026-08-11. 'OTHER PARTY' = external/true; 'RELATED PARTY' = intercompany,
    eliminated in true consolidation. The Entity P&L table (build_entities)
    deliberately does NOT use this filter -- its 'TOTAL (Sum of Entities)'
    row is meant to stay the gross, pre-elimination figure (matching the MIS
    Summary-ADFL Conso 2025 sheet), with the single 'Eliminations' line
    (added earlier, per explicit user instruction to keep it as one lump
    Group-level row rather than a per-entity breakdown) bridging to
    'TOTAL (Consolidated)'.
    """
    return [r for r in rows if r["customerType"] != "RELATED PARTY"]


# ==============================================================================
# 4M (Apr-Jul FY27 vs FY26) tables
# ==============================================================================
def monthly_table(rows, dim_key, order, q4_lookup):
    """One row per distinct `dim_key` value: Apr/May/Jun/Jul FY27 + FY26,
    YTD totals, monthly + overall YoY growth, mix%, plus a genuine Q4 FY26
    reference pulled from `q4_lookup` (the transaction extract’s own QTR='Q4' rows for
    FY26 -- real data, not a placeholder, since the transaction extract carries full-year
    history). Appends a TOTAL row."""
    totals = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    for r in rows:
        totals[r[dim_key]][r["fy"]][r["month"]] += r["value"]

    dim_values = [v for v in order]
    for v in totals:
        if v and v not in dim_values:
            dim_values.append(v)

    out = []
    for v in dim_values:
        t = totals.get(v)
        if not t:
            continue
        fy27, fy26 = t.get("FY27", {}), t.get("FY26", {})
        apr, may, jun, jul = (fy27.get(m, 0.0) for m in MONTHS)
        aprPy, mayPy, junPy, julPy = (fy26.get(m, 0.0) for m in MONTHS)
        ytdCur, ytdPy = apr + may + jun + jul, aprPy + mayPy + junPy + julPy
        q4py_val = q4_lookup.get(v, 0.0)
        # 2026-08-11: previously dropped rows whenever Apr-Jul was all zero,
        # even if a real Q4 FY26 reference value existed for that name (a
        # category/brand that only sold in that quarter, not repeated in the
        # current comparison window) -- silently orphaning it from `out`
        # while the TOTAL row's q4Py (summed independently below from the
        # full q4_lookup) still counted it, so TOTAL didn't match the sum of
        # displayed rows. Now kept (as an all-zero-except-q4Py row) so the
        # two stay consistent.
        if ytdCur == 0 and ytdPy == 0 and q4py_val == 0:
            continue
        out.append({
            "name": v,
            "apr": apr, "may": may, "jun": jun, "jul": jul, "ytdCur": ytdCur,
            "aprPy": aprPy, "mayPy": mayPy, "junPy": junPy, "julPy": julPy, "ytdPy": ytdPy,
            "aprGr": gr(apr, aprPy), "mayGr": gr(may, mayPy), "junGr": gr(jun, junPy), "julGr": gr(jul, julPy),
            "q1Gr": gr(ytdCur, ytdPy),
            "q4Py": q4py_val,
        })

    total = {"name": "TOTAL", "q4Py": sum(q4_lookup.values())}
    for k in ("apr", "may", "jun", "jul", "ytdCur", "aprPy", "mayPy", "junPy", "julPy", "ytdPy"):
        total[k] = sum(r[k] for r in out)
    total["aprGr"], total["mayGr"], total["junGr"], total["julGr"] = (
        gr(total["apr"], total["aprPy"]), gr(total["may"], total["mayPy"]),
        gr(total["jun"], total["junPy"]), gr(total["jul"], total["julPy"]),
    )
    total["q1Gr"] = gr(total["ytdCur"], total["ytdPy"])

    grand = total["ytdCur"]
    for r in out:
        r["mixPct"] = (r["ytdCur"] / grand) if grand else 0
    total["mixPct"] = 1.0 if grand else 0

    out.append(total)
    return out


def q4_fy26_by_dim(rows, dim_key):
    totals = defaultdict(float)
    for r in rows:
        if r["fy"] == "FY26" and r["qtr"] == "Q4":
            totals[r[dim_key]] += r["value"]
    return totals


def top_and_growth_from_monthly(monthly_rows):
    non_total = [r for r in monthly_rows if r["name"] != "TOTAL"]
    grand = sum(r["ytdCur"] for r in non_total) or 1.0

    top_rows = []
    for r in non_total:
        share = r["ytdCur"] / grand
        top_rows.append({
            "name": r["name"], "cur": r["ytdCur"], "py": r["ytdPy"],
            "absDelta": r["ytdCur"] - r["ytdPy"], "grYoy": r["q1Gr"], "sharePct": share, "cumPct": 0,
        })
    top_rows = rank_by_value(top_rows)
    for r in top_rows:
        r["status"] = signal_from_growth(r["grYoy"])

    growth_rows = []
    for r in non_total:
        growth_rows.append({
            "name": r["name"], "cur": r["ytdCur"], "py": r["ytdPy"],
            "absDelta": r["ytdCur"] - r["ytdPy"], "grYoy": r["q1Gr"],
            "gr1m": r["aprGr"], "gr2m": r["mayGr"], "gr3m": r["junGr"], "gr4m": r["julGr"],
        })
    for r in growth_rows:
        r["trend"] = trend_from_growth(r["grYoy"], r["gr2m"], r["gr3m"], r["gr4m"])

    return top_rows, growth_rows


def build_top_sub(rows):
    totals = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    category_of = {}
    for r in rows:
        sub = r["subcategory"]
        if not sub:
            continue
        totals[sub][r["fy"]][r["month"]] += r["value"]
        category_of.setdefault(sub, r["category"])

    all_rows = []
    for sub, t in totals.items():
        fy27, fy26 = t.get("FY27", {}), t.get("FY26", {})
        cur = sum(fy27.get(m, 0.0) for m in MONTHS)
        py = sum(fy26.get(m, 0.0) for m in MONTHS)
        if cur == 0 and py == 0:
            continue
        all_rows.append({"subcategory": sub, "category": category_of.get(sub, ""), "cur": cur, "py": py})

    grand = sum(r["cur"] for r in all_rows) or 1.0
    for r in all_rows:
        r["absDelta"] = r["cur"] - r["py"]
        r["grYoy"] = gr(r["cur"], r["py"])
        r["sharePct"] = r["cur"] / grand

    all_rows.sort(key=lambda r: r["cur"], reverse=True)
    top10 = rank_by_value([dict(r) for r in all_rows[:10]])
    tail10 = rank_by_value([dict(r) for r in all_rows[-10:]])
    for r in top10 + tail10:
        r["status"] = signal_from_growth(r["grYoy"])
    return {"top10": top10, "tail10": tail10}


def build_dim_by_entity_4m(all_rows, dim_key, order):
    """{'ALL': [...], entity: [...]} monthly tables for one dimension,
    keyed for the entity filter pane."""
    out = {}
    for scope in ["ALL"] + ENTITY_DISPLAY_ORDER:
        scoped_rows = filter_entity(all_rows, scope)
        q4_lookup = q4_fy26_by_dim(scoped_rows, dim_key)
        out[scope] = monthly_table(scoped_rows, dim_key, order, q4_lookup)
    return out


# ==============================================================================
# 12M (FY26 vs FY25) tables -- ANNUAL ONLY, sourced from the audited
# annual sales workbook. No
# quarterly breakdown exists in this source for either year -- see module
# docstring for why this replaces rather than supplements the earlier
# the transaction extract-quarterly approach.
# ==============================================================================
def load_annual_transactions():
    wb = openpyxl.load_workbook(ANNUAL_XLSX_PATH, data_only=True)
    ws = wb["Region wise Sales Base"]
    rows = []
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
        entity, country, region_sa, zone, brand, cat_sa, subcat, fingroup, cif_abs, cif_cr, fy, remarks = row[:12]
        if entity is None or fy not in (ANNUAL_FY_CUR, ANNUAL_FY_PY) or cif_cr is None:
            continue
        rows.append({
            "entity": ANNUAL_ENTITY_FIX.get(entity, entity),
            "fy": "cur" if fy == ANNUAL_FY_CUR else "py",
            "category": ANNUAL_CATEGORY_FIX.get(cat_sa, cat_sa),
            "brand": ANNUAL_BRAND_FIX.get(brand, brand),
            "region": ANNUAL_REGION_FIX.get(region_sa, region_sa),
            "subcategory": str(subcat).strip().upper() if subcat else None,
            "value": float(cif_cr) * 100.0,  # CIF (Crores) -> Lakhs
        })
    return rows


def build_annual_table(rows, dim_key, order):
    """One row per distinct `dim_key` value: real audited cur (FY26) and py
    (FY25) on the SAME basis (both from this same workbook) -- growth% is
    genuine here, unlike the old the transaction extract-vs-nothing 12M attempt."""
    totals = defaultdict(lambda: defaultdict(float))
    for r in rows:
        totals[r[dim_key]][r["fy"]] += r["value"]

    dim_values = [v for v in order]
    for v in totals:
        if v and v not in dim_values:
            dim_values.append(v)

    out = []
    for v in dim_values:
        t = totals.get(v)
        if not t:
            continue
        cur, py = t.get("cur", 0.0), t.get("py", 0.0)
        if cur == 0 and py == 0:
            continue
        out.append({"name": v, "cur": cur, "py": py, "absDelta": cur - py, "grYoy": gr(cur, py)})

    total = {"name": "TOTAL"}
    total["cur"] = sum(r["cur"] for r in out)
    total["py"] = sum(r["py"] for r in out)
    total["absDelta"] = total["cur"] - total["py"]
    total["grYoy"] = gr(total["cur"], total["py"])

    grand = total["cur"]
    for r in out:
        r["mixPct"] = (r["cur"] / grand) if grand else 0
    total["mixPct"] = 1.0 if grand else 0

    out.append(total)
    return out


def rank_by_value_safe(rows):
    """Like rank_by_value, but tolerates a None sharePct (shouldn't happen
    here since sharePct is always cur/grand, but keeps the cumPct running
    sum well-defined either way)."""
    rows = sorted(rows, key=lambda r: r["cur"] or 0, reverse=True)
    running = 0.0
    for r in rows:
        running += r["sharePct"] or 0
        r["cumPct"] = round(running, 6)
    return rows


def top_and_growth_from_annual(annual_rows):
    non_total = [r for r in annual_rows if r["name"] != "TOTAL"]
    grand = sum(r["cur"] for r in non_total) or 1.0

    top_rows = []
    for r in non_total:
        top_rows.append({
            "name": r["name"], "cur": r["cur"], "py": r["py"] or None,
            "absDelta": r["absDelta"], "grYoy": r["grYoy"], "sharePct": r["cur"] / grand, "cumPct": 0,
        })
    top_rows = rank_by_value_safe(top_rows)
    for r in top_rows:
        r["status"] = signal_from_growth(r["grYoy"])

    growth_rows = []
    for r in non_total:
        growth_rows.append({"name": r["name"], "cur": r["cur"], "py": r["py"] or None, "absDelta": r["absDelta"], "grYoy": r["grYoy"]})
    for r in growth_rows:
        # No quarterly detail exists in this source -- trend collapses to a
        # YoY-only read (Steady Growth / Sustained Decline / Volatile Flat;
        # Accelerating/Decelerating/Turnaround need a real quarterly path,
        # which isn't available here, so those 3 states are unreachable by
        # construction rather than faked from a flat repeated value).
        r["trend"] = trend_from_growth(r["grYoy"], r["grYoy"], r["grYoy"], r["grYoy"])

    return top_rows, growth_rows


def build_dim_by_entity_annual(all_rows, dim_key, order):
    out = {}
    for scope in ["ALL"] + ENTITY_DISPLAY_ORDER:
        scoped = all_rows if scope == "ALL" else [r for r in all_rows if r["entity"] == scope]
        out[scope] = build_annual_table(scoped, dim_key, order)
    return out


def leader_rows_from_annual(annual_rows):
    """LEADER_COLS_12M shape (name/cur/py/gr12m/grQ4/signal) for the Exec
    Summary signal tables in 12M mode. grQ4 is None -- no quarterly
    breakdown in this source (see module docstring). `py` (2026-08-12, per
    explicit user instruction: show 12M FY25 beside 12M FY26 everywhere the
    latter appears) is both displayed as its own column and used by the
    page's withTotalRow()/buildTotalRow() to show a genuine Total 'FY25'
    figure instead of a misleading blank/zero."""
    out = []
    for r in annual_rows:
        if r["name"] == "TOTAL":
            continue
        out.append({"name": r["name"], "cur": r["cur"], "py": r["py"], "gr12m": r["grYoy"], "grQ4": None,
                    "signal": signal_from_growth(r["grYoy"])})
    return out


def build_top_sub_annual(rows):
    totals = defaultdict(lambda: defaultdict(float))
    category_of = {}
    for r in rows:
        sub = r["subcategory"]
        if not sub:
            continue
        totals[sub][r["fy"]] += r["value"]
        category_of.setdefault(sub, r["category"])

    all_rows = []
    for sub, t in totals.items():
        cur, py = t.get("cur", 0.0), t.get("py", 0.0)
        if cur == 0 and py == 0:
            continue
        all_rows.append({"subcategory": sub, "category": category_of.get(sub, ""), "cur": cur, "py": py or None})

    grand = sum(r["cur"] for r in all_rows) or 1.0
    for r in all_rows:
        r["absDelta"] = (r["cur"] - r["py"]) if r["py"] else None
        r["grYoy"] = gr(r["cur"], r["py"])
        r["sharePct"] = r["cur"] / grand

    all_rows.sort(key=lambda r: r["cur"], reverse=True)
    top10 = rank_by_value_safe([dict(r) for r in all_rows[:10]])
    tail10 = rank_by_value_safe([dict(r) for r in all_rows[-10:]])
    for r in top10 + tail10:
        r["status"] = signal_from_growth(r["grYoy"])
    return {"top10": top10, "tail10": tail10}


# ==============================================================================
# 2. MIS report -- Entity-wise P&L (Summary-ADFL Conso 2025) and the true,
#    elimination-adjusted Group KPIs (consolidated management sheet)
# ==============================================================================
def mis_entity_row(ws, row_num):
    totals = defaultdict(float)
    for month, cols in MIS_ENTITY_COLS.items():
        for code, col in zip(MIS_ENTITY_ORDER, cols):
            v = ws[f"{col}{row_num}"].value
            if v is not None:
                totals[code] += v
    return totals


def build_entities(conso_rows):
    wb = openpyxl.load_workbook(MIS_PATH, data_only=True)
    ws = wb["Summary-ADFL Conso 2025"]

    revenue = mis_entity_row(ws, 15)   # Sales -- ties exactly to the transaction extract, verified
    ebitda = mis_entity_row(ws, 49)    # EBIDTA
    pbt = mis_entity_row(ws, 51)       # PBT
    pat = mis_entity_row(ws, 55)       # PAT

    # PY (FY26) entity-wise EBITDA/PBT/PAT genuinely don't exist anywhere in
    # this MIS workbook (no prior-year entity-wise P&L sheet). PY Revenue
    # (Sales-basis, comparable to `revenue` above) IS available, from
    # the transaction extract’s own FY26 Apr-Jul entity totals.
    py_revenue_by_entity = defaultdict(float)
    for r in conso_rows:
        if r["fy"] == "FY26" and r["month"] in MONTHS:
            py_revenue_by_entity[r["entity"]] += r["value"]

    entities = []
    for code in MIS_ENTITY_ORDER:
        name = "TOTAL (Sum of Entities)" if code == "Total" else MIS_ENTITY_TO_DISPLAY[code]
        rev = revenue.get(code, 0.0)
        eb = ebitda.get(code, 0.0)
        pb = pbt.get(code, 0.0)
        pa = pat.get(code, 0.0)
        py_rev = (sum(py_revenue_by_entity.values()) if code == "Total"
                  else py_revenue_by_entity.get(MIS_ENTITY_TO_DISPLAY.get(code), 0.0))
        entities.append({
            "name": name, "revenue": rev, "ebitda": eb,
            "ebitdaPct": (eb / rev) if rev else None,
            "pbt": pb, "pbtPct": (pb / rev) if rev else None,
            "pat": pa, "patPct": (pa / rev) if rev else None,
            "pyRevenue": py_rev,
        })
    return entities


def compute_health_annual(revenue, ebitda, pbt, ebitda_pct, pbt_pct):
    """Python mirror of the client-side computeHealth() JS function --
    deliberate small duplication (same convention already used for
    signal_from_growth/trend_from_growth elsewhere in this file) rather
    than a cross-language shared source. Keep both in sync if the
    thresholds ever change."""
    if not revenue or revenue <= 0:
        return "N/A Consol"
    if ebitda <= 0 or pbt <= 0:
        return "Stressed"
    if ebitda_pct is not None and ebitda_pct >= 0.15 and pbt_pct is not None and pbt_pct >= 0.10 and pbt > 0:
        return "Healthy"
    if ebitda > 0 and pbt > 0 and ebitda_pct is not None and ebitda_pct >= 0.05:
        return "Moderate"
    return "Stressed"


def load_fs_entity_pnl():
    """Real, audited entity-wise Revenue/EBITDA/PBT/PAT for FY26 vs FY25 --
    added 2026-08-12 per explicit user instruction ('add EBITDA, PBT, PAT
    and its comparison ... as you are showing in Standalone'). An earlier
    investigation had concluded no such source existed; a follow-up deep
    dive (2026-08-12) found it does, in the SAME board Finalisation
    workbook already used for the group segment figures (FS_SEGMENT_PATH)
    -- each entity has its own statutory P&L sheet with FY26 and FY25
    columns. Exact cell references below were mapped and verified during
    that investigation (cross-checked against FORM AOC-I subsidiary
    statements and the Group P&L Consol sheet's own entity-column sanity
    totals).

    EBITDA is ALWAYS derived here as PBT + Depreciation&Amortisation +
    Finance Cost - Other Income (the permanent, non-negotiable house
    formula, verified: Impairment = 0 on every one of these sheets) --
    NEVER read from a sheet's own labelled 'EBITDA' row, which was found to
    be inconsistent between columns on the same sheet (it omits
    subtracting Other Income there, even though the FY26 column on the
    SAME sheet gets it right).

    Two source-data quirks handled per-entity:
      - ADF Holdings USA / ADF Foods USA / Vibrant Foods NJ LLC: their INR
        columns are labelled 'Rupees in Lakhs' but the cached values are
        actually absolute Rupees -- divided by 100,000 here.
      - ADF Foods Australia: acquired 9-Jul-2025 (confirmed via FORM
        AOC-I) -- genuinely did not exist in FY25, so its `py` is None
        (not a data gap, a real business fact), and even its FY26 column
        is a partial ~8.7-month period, not a full year.

    read_only=True: this 124-sheet, ~9MB workbook is dramatically slower to
    open in the default mode (which builds a full mutable object model,
    including styles, for every sheet) than in read_only/streaming mode --
    a first attempt without this timed out after 20+ minutes with no
    output. Direct cell access (ws[f"{col}{row}"]) still works fine in
    read_only mode, same as load_segment_disclosure_kpis() already does.
    """
    wb = openpyxl.load_workbook(FS_SEGMENT_PATH, data_only=True, read_only=True)

    def load_one(sheet, cur_col, py_col, revenue_row, other_income_row, finance_cost_row,
                 dep_amort_row, pbt_row, pat_row, unit_divisor=1.0, py_is_na=False):
        ws = wb[sheet]

        def val(row, col):
            v = ws[f"{col}{row}"].value
            return (v if isinstance(v, (int, float)) else 0.0) / unit_divisor

        def one_year(col):
            revenue = val(revenue_row, col)
            other_income = val(other_income_row, col)
            finance_cost = val(finance_cost_row, col)
            dep_amort = val(dep_amort_row, col)
            pbt = val(pbt_row, col)
            pat = val(pat_row, col)
            ebitda = pbt + dep_amort + finance_cost - other_income
            return {"revenue": revenue, "ebitda": ebitda, "pbt": pbt, "pat": pat}

        return one_year(cur_col), (None if py_is_na else one_year(py_col))

    specs = {
        "ADF Foods Ltd (Standalone)": ("P & L", "C", "E", 10, 11, 18, 19, 24, 33, 1.0, False),
        "Telluric Foods": ("TFIL CONSOL PNL", "E", "M", 8, 9, 16, 17, 22, 30, 1.0, False),
        "ADF Holdings USA": ("ADFHL PNL", "G", "I", 8, 9, 16, 17, 22, 30, 100000.0, False),
        "ADF Foods USA": ("ADF USA PL", "G", "I", 8, 9, 16, 17, 22, 30, 100000.0, False),
        "Vibrant Foods NJ LLC": ("VIB USA PNL", "G", "I", 8, 9, 16, 17, 22, 30, 100000.0, False),
        "ADF Foods UK": ("ADK UK PNL", "M", "P", 7, 8, 15, 16, 21, 30, 1.0, False),
        "ADF Foods Australia": ("ADF Australia PNL", "L", "O", 7, 8, 15, 16, 21, 30, 1.0, True),
    }
    out = {}
    for name, spec in specs.items():
        out[name] = load_one(*spec)

    # True, audited Group Consolidated (elimination-adjusted) -- same
    # 'P & L Consol' sheet, Total columns K (FY26) / AC (FY25). Verified:
    # revenue here matches the audited total revenue exactly
    # -- same figure, cross-checked two ways.
    out["__TRUE__"] = load_one("P & L Consol", "K", "AC", 10, 11, 18, 19, 24, 33, 1.0, False)
    return out


def build_entities_annual(fs_entity_pnl):
    """Entity-wise Revenue/EBITDA/PBT/PAT for 12M, FY26 vs FY25 -- added
    2026-08-12 (see load_fs_entity_pnl()'s docstring). Sourced entirely
    from each entity's own audited P&L (gross, pre-elimination, same as
    how the 4M Profitability tab's entity figures are gross/pre-
    elimination from the MIS Summary sheet) -- NOT the transactional
    annual sales workbook Category/Brand/Zone use, so
    this table's Revenue will not tie to Category/Brand/Zone below, by the
    same design already established and disclosed for the 4M tab.
    """
    entities = []
    for name in ENTITY_DISPLAY_ORDER:
        cur, py = fs_entity_pnl[name]
        entities.append({
            "name": name,
            "revenue": cur["revenue"], "pyRevenue": py["revenue"] if py else None,
            "revenueGr": gr(cur["revenue"], py["revenue"] if py else None),
            "ebitda": cur["ebitda"], "pyEbitda": py["ebitda"] if py else None,
            "ebitdaPct": (cur["ebitda"] / cur["revenue"]) if cur["revenue"] else None,
            "pbt": cur["pbt"], "pyPbt": py["pbt"] if py else None,
            "pbtPct": (cur["pbt"] / cur["revenue"]) if cur["revenue"] else None,
            "pat": cur["pat"], "pyPat": py["pat"] if py else None,
            "patPct": (cur["pat"] / cur["revenue"]) if cur["revenue"] else None,
        })
        e = entities[-1]
        e["health"] = compute_health_annual(e["revenue"], e["ebitda"], e["pbt"], e["ebitdaPct"], e["pbtPct"])

    def total_row(name, rows, key):
        cur = sum(r[key] for r in rows)
        py_vals = [r[f"py{key[0].upper()}{key[1:]}"] for r in rows]
        py = sum(v for v in py_vals if v is not None) if any(v is not None for v in py_vals) else None
        return cur, py

    sum_revenue, sum_py_revenue = total_row("sum", entities, "revenue")
    sum_ebitda, sum_py_ebitda = total_row("sum", entities, "ebitda")
    sum_pbt, sum_py_pbt = total_row("sum", entities, "pbt")
    sum_pat, sum_py_pat = total_row("sum", entities, "pat")
    sum_entities = {
        "name": "TOTAL (Sum of Entities)",
        "revenue": sum_revenue, "pyRevenue": sum_py_revenue, "revenueGr": gr(sum_revenue, sum_py_revenue),
        "ebitda": sum_ebitda, "pyEbitda": sum_py_ebitda, "ebitdaPct": (sum_ebitda / sum_revenue) if sum_revenue else None,
        "pbt": sum_pbt, "pyPbt": sum_py_pbt, "pbtPct": (sum_pbt / sum_revenue) if sum_revenue else None,
        "pat": sum_pat, "pyPat": sum_py_pat, "patPct": (sum_pat / sum_revenue) if sum_revenue else None,
    }
    sum_entities["health"] = compute_health_annual(sum_revenue, sum_ebitda, sum_pbt, sum_entities["ebitdaPct"], sum_entities["pbtPct"])
    entities.append(sum_entities)

    true_cur, true_py = fs_entity_pnl["__TRUE__"]
    elim = {
        "name": "Eliminations",
        "revenue": true_cur["revenue"] - sum_revenue, "pyRevenue": true_py["revenue"] - sum_py_revenue,
        "revenueGr": None,
        "ebitda": true_cur["ebitda"] - sum_ebitda, "pyEbitda": true_py["ebitda"] - sum_py_ebitda,
        "ebitdaPct": None,
        "pbt": true_cur["pbt"] - sum_pbt, "pyPbt": true_py["pbt"] - sum_py_pbt,
        "pbtPct": None,
        "pat": true_cur["pat"] - sum_pat, "pyPat": true_py["pat"] - sum_py_pat,
        "patPct": None,
        "health": "N/A Consol",
    }
    entities.append(elim)

    total_consol = {
        "name": "TOTAL (Consolidated)",
        "revenue": true_cur["revenue"], "pyRevenue": true_py["revenue"], "revenueGr": gr(true_cur["revenue"], true_py["revenue"]),
        "ebitda": true_cur["ebitda"], "pyEbitda": true_py["ebitda"], "ebitdaPct": (true_cur["ebitda"] / true_cur["revenue"]) if true_cur["revenue"] else None,
        "pbt": true_cur["pbt"], "pyPbt": true_py["pbt"], "pbtPct": (true_cur["pbt"] / true_cur["revenue"]) if true_cur["revenue"] else None,
        "pat": true_cur["pat"], "pyPat": true_py["pat"], "patPct": (true_cur["pat"] / true_cur["revenue"]) if true_cur["revenue"] else None,
    }
    total_consol["health"] = compute_health_annual(total_consol["revenue"], total_consol["ebitda"], total_consol["pbt"], total_consol["ebitdaPct"], total_consol["pbtPct"])
    entities.append(total_consol)
    return entities


def build_exec_kpis(sales_cur, sales_py, ebitda_cur, ebitda_py, pbt_cur, pbt_py, pat_cur, pat_py, margin_cur, margin_py,
                     q1_fy27_sales_true):
    """Mirrors Standalone's own 9-tile Exec Summary KPI row (see
    refresh_ceo_dashboard.py's build_exec()) wherever a Group-wide
    equivalent honestly exists. 3 of Standalone's 9 tiles are NOT included
    here, deliberately, not by oversight:
      - "No of FCL" (container-shipment count) -- sourced from a Standalone/
        India-specific export-logistics sheet with no Group-wide equivalent;
        the other 6 entities don't ship FCL containers from India.
      - "Export Value" -- meaningful for Standalone (an India-based
        manufacturer whose sales are predominantly export) but doesn't translate to
        a Group-wide concept, and the transaction extract has no export/domestic flag to
        compute one even if it did.
      - "Realization Growth" (₹/kg) -- blocked by the same volume-data gap
        already documented for the deferred Vol vs Value tabs: only
        Standalone tracks volume in true KG, the other 6 entities only have
        case/unit quantities in free-text item descriptions.
    Of Standalone's 3 trailing reference tiles (Q1 FY27 / Q4 FY26 / 4M FY26,
    all a bare "Sales (₹L)" figure), only 2 are included -- "Q4 FY26" is
    dropped, since no true elimination-adjusted figure for that specific
    sub-period exists anywhere in the MIS source (see caller for why). Both
    included tiles use the SAME true, elimination-adjusted basis as every
    other KPI here (see caller) -- not the the transaction extract pre-elimination sum
    used for the Category/Brand/Zone tables lower on this tab, which is
    disclosed separately.
    """
    rev_growth = gr(sales_cur, sales_py)
    kpis = [
        {"label": "₹ 4M FY27 Revenue", "value": sales_cur, "sub": f"vs ₹{sales_py:,.1f} L PY", "tone": "navy", "fmt": "money"},
        {"label": "📈 4M Revenue Growth", "value": rev_growth, "sub": "4M FY27 vs 4M FY26", "signColor": True, "fmt": "pct"},
        {"label": "📊 EBITDA (₹L)", "value": ebitda_cur, "sub": f"{ebitda_cur/sales_cur*100:.1f}% margin · vs ₹{ebitda_py:,.1f} L PY" if sales_cur else "", "tone": "teal", "fmt": "money"},
        {"label": "📊 PBT (₹L)", "value": pbt_cur, "sub": f"{pbt_cur/sales_cur*100:.1f}% margin · vs ₹{pbt_py:,.1f} L PY" if sales_cur else "", "tone": "purple", "fmt": "money"},
        {"label": "📊 PAT (₹L)", "value": pat_cur, "sub": f"{pat_cur/sales_cur*100:.1f}% margin · vs ₹{pat_py:,.1f} L PY" if sales_cur else "", "tone": "gold", "fmt": "money"},
        {"label": "💰 Gross Margin %", "value": margin_cur, "sub": f"PY: {margin_py*100:.1f}%  Δ {(margin_cur-margin_py)*100:+.1f} pp", "tone": "green", "fmt": "ratio"},
        {"label": "🌐 Group Entities", "value": "7", "sub": "India · USA (x3) · UK · Australia · Telluric", "tone": "green"},
        {"label": "📊 Q1 FY27", "value": f"{q1_fy27_sales_true:,.0f}", "sub": "Sales (₹L)", "tone": "navy"},
        {"label": "📊 4M FY26", "value": f"{sales_py:,.0f}", "sub": "Sales (₹L)", "tone": "purple"},
    ]
    return kpis


def add_eliminations_row(entities, true_kpis):
    """2026-08-10, per explicit user instruction: show 'Eliminations' as its
    own line in the Entity-wise P&L table, and make the table's grand total
    reflect the TRUE, elimination-adjusted Group figure -- not just the
    pre-elimination sum of entities. Previously the table only showed
    'TOTAL (Sum of Entities)' (pre-elim) while the Exec Summary KPI tiles
    showed the true consolidated figure elsewhere on the page, with an
    on-page note explaining the two wouldn't match; this closes that gap by
    laying out the full reconciliation on the Profitability tab itself:
        Entity 1..7  ->  TOTAL (Sum of Entities)  ->  Eliminations  ->  TOTAL (Consolidated)
    Eliminations = true_kpis - TOTAL (Sum of Entities), per line item
    (Revenue/EBITDA/PBT/PAT/PY Revenue) -- the only breakdown available,
    since the MIS source gives the elimination adjustment only at the
    Group level, not per entity. Percentages are left None (renders '--')
    for the Eliminations row -- a margin % of an elimination adjustment
    isn't a meaningful figure -- but ARE computed for the new true TOTAL
    row, against its own true revenue (matching the Exec Summary KPI tiles
    exactly).
    """
    sum_of_entities = next(e for e in entities if e["name"] == "TOTAL (Sum of Entities)")
    sales_cur, sales_py, ebitda_cur, pbt_cur, pat_cur = true_kpis

    elim_revenue = sales_cur - sum_of_entities["revenue"]
    elim_ebitda = ebitda_cur - sum_of_entities["ebitda"]
    elim_pbt = pbt_cur - sum_of_entities["pbt"]
    elim_pat = pat_cur - sum_of_entities["pat"]
    elim_py_revenue = sales_py - sum_of_entities["pyRevenue"]

    entities.append({
        "name": "Eliminations", "revenue": elim_revenue, "ebitda": elim_ebitda,
        "ebitdaPct": None, "pbt": elim_pbt, "pbtPct": None,
        "pat": elim_pat, "patPct": None, "pyRevenue": elim_py_revenue,
    })
    entities.append({
        "name": "TOTAL (Consolidated)", "revenue": sales_cur, "ebitda": ebitda_cur,
        "ebitdaPct": (ebitda_cur / sales_cur) if sales_cur else None,
        "pbt": pbt_cur, "pbtPct": (pbt_cur / sales_cur) if sales_cur else None,
        "pat": pat_cur, "patPct": (pat_cur / sales_cur) if sales_cur else None,
        "pyRevenue": sales_py,
    })


def build_exec_kpis_by_entity(entities, category_by_entity):
    """2026-08-11, per explicit user instruction: the entity-filtered Exec
    Summary Revenue tile now uses that entity's TRUE, external-only revenue
    (its Category table's own TOTAL row, same source as the Category/Brand/
    Zone tables shown right below it on the same filtered view) instead of
    its gross (pre-elimination) MIS figure -- the two could differ hugely
    for entities with heavy intercompany activity (for one US entity the
    gross figure was roughly double the external one), which looked like a bug when the
    KPI tile and the tables beneath it told very different stories for the
    same filtered entity.

    EBITDA/PBT/PAT stay on the gross/MIS basis -- no external-only P&L
    breakdown exists below Revenue (the transaction extract has no COGS/profit columns
    at all), so there is no alternative source for those three. This means
    their margin %s (computed against gross EBITDA/PBT/PAT over now-smaller
    external Revenue) will read higher than a true margin would -- disclosed
    on-page rather than silently shown as if it were apples-to-apples.
    """
    out = {}
    for e in entities:
        if e["name"].startswith("TOTAL"):
            continue
        cat_total = next((r for r in category_by_entity.get(e["name"], []) if r["name"] == "TOTAL"), None)
        rev = cat_total["ytdCur"] if cat_total else e["revenue"]
        py_rev = cat_total["ytdPy"] if cat_total else e["pyRevenue"]
        kpis = [
            {"label": "₹ 4M FY27 Revenue", "value": rev, "sub": f"vs ₹{py_rev:,.1f} L PY" if py_rev else "no FY26 baseline", "tone": "navy", "fmt": "money"},
            {"label": "📈 4M Revenue Growth", "value": gr(rev, py_rev), "sub": "4M FY27 vs 4M FY26", "signColor": True, "fmt": "pct"},
            {"label": "📊 EBITDA (₹L)", "value": e["ebitda"], "sub": f"{e['ebitdaPct']*100:.1f}% margin (of gross revenue)" if e["ebitdaPct"] is not None else "", "tone": "teal", "fmt": "money"},
            {"label": "📊 PBT (₹L)", "value": e["pbt"], "sub": f"{e['pbtPct']*100:.1f}% margin (of gross revenue)" if e["pbtPct"] is not None else "", "tone": "purple", "fmt": "money"},
            {"label": "📊 PAT (₹L)", "value": e["pat"], "sub": f"{e['patPct']*100:.1f}% margin (of gross revenue)" if e["patPct"] is not None else "", "tone": "gold", "fmt": "money"},
        ]
        out[e["name"]] = kpis
    return out


def load_segment_disclosure_kpis():
    """True, audited, elimination-adjusted Group Revenue and PBT for FY26 vs
    FY25 -- from the audited segment figures (the board financial-statement
    File_March 2026_v51.xlsx), the same figures published in the FY26
    annual results. Row/col references verified 2026-08-10 directly against
    the sheet: row 13 'Total Segment Revenue', row 24 'Total Profit Before
    Tax', col E = year ended 31-03-2026 (FY26), col F = year ended
    31-03-2025 (FY25)."""
    wb = openpyxl.load_workbook(FS_SEGMENT_PATH, data_only=True, read_only=True)
    ws = wb[os.environ.get("FS_SEGMENT_SHEET", "")]

    def val(row, col):
        v = ws.cell(row=row, column=col).value
        return v if isinstance(v, (int, float)) else 0.0

    return {
        "revenue_cur": val(13, 5), "revenue_py": val(13, 6),
        "pbt_cur": val(24, 5), "pbt_py": val(24, 6),
    }


def build_exec_kpis_annual(segment_kpis, entities_12m):
    """2026-08-12: EBITDA/PAT tiles added (previously only Revenue/PBT had a
    true Group-level figure available) -- pulled from entities_12m's own
    'TOTAL (Consolidated)' row (see build_entities_annual()), which is the
    SAME true, audited, elimination-adjusted figure as segment_kpis'
    Revenue/PBT (cross-verified: both sources' Revenue and PBT match
    exactly), just sourced from the fuller entity-wise P&L data pull that
    also yields EBITDA and PAT."""
    rev_cur, rev_py = segment_kpis["revenue_cur"], segment_kpis["revenue_py"]
    pbt_cur, pbt_py = segment_kpis["pbt_cur"], segment_kpis["pbt_py"]
    total_consol = next(e for e in entities_12m if e["name"] == "TOTAL (Consolidated)")
    ebitda_cur, ebitda_py = total_consol["ebitda"], total_consol["pyEbitda"]
    pat_cur, pat_py = total_consol["pat"], total_consol["pyPat"]
    kpis = [
        {"label": "₹ 12M FY26 Revenue", "value": rev_cur, "sub": f"vs ₹{rev_py:,.1f} L PY (FY25)", "tone": "navy", "fmt": "money"},
        {"label": "📈 12M Revenue Growth", "value": gr(rev_cur, rev_py), "sub": "12M FY26 vs 12M FY25", "signColor": True, "fmt": "pct"},
        {"label": "📊 EBITDA (₹L)", "value": ebitda_cur, "sub": f"{ebitda_cur/rev_cur*100:.1f}% margin · vs ₹{ebitda_py:,.1f} L PY" if rev_cur else "", "tone": "teal", "fmt": "money"},
        {"label": "📊 PBT (₹L)", "value": pbt_cur, "sub": f"{pbt_cur/rev_cur*100:.1f}% margin · vs ₹{pbt_py:,.1f} L PY" if rev_cur else "", "tone": "purple", "fmt": "money"},
        {"label": "📊 PAT (₹L)", "value": pat_cur, "sub": f"{pat_cur/rev_cur*100:.1f}% margin · vs ₹{pat_py:,.1f} L PY" if rev_cur else "", "tone": "gold", "fmt": "money"},
        {"label": "🌐 Group Entities", "value": "7", "sub": "India · USA (x3) · UK · Australia · Telluric", "tone": "green"},
        {"label": "📊 12M FY25 Revenue", "value": f"{rev_py:,.0f}", "sub": "Sales (₹L)", "tone": "purple"},
        {"label": "📊 12M FY25 PBT", "value": f"{pbt_py:,.0f}", "sub": "(₹L)", "tone": "teal"},
    ]
    return kpis


def build_exec_kpis_by_entity_annual(entities_12m):
    """2026-08-12: now shows EBITDA/PBT/PAT per entity too (previously
    Revenue-only), using the same real FS-sourced entity-wise P&L as the
    Profitability tab's entity table (see build_entities_annual())."""
    out = {}
    for e in entities_12m:
        if e["name"].startswith("TOTAL") or e["name"] == "Eliminations":
            continue
        kpis = [
            {"label": "₹ 12M FY26 Revenue", "value": e["revenue"],
             "sub": f"vs ₹{e['pyRevenue']:,.1f} L PY (FY25)" if e["pyRevenue"] else "no FY25 baseline",
             "tone": "navy", "fmt": "money"},
        ]
        if e["pyRevenue"]:
            kpis.append({"label": "📈 12M Revenue Growth", "value": e["revenueGr"], "sub": "12M FY26 vs 12M FY25", "signColor": True, "fmt": "pct"})
            kpis.append({"label": "📊 12M FY25 Revenue", "value": f"{e['pyRevenue']:,.0f}", "sub": "Sales (₹L)", "tone": "purple"})
        kpis.append({"label": "📊 EBITDA (₹L)", "value": e["ebitda"],
                      "sub": f"{e['ebitdaPct']*100:.1f}% margin" if e["ebitdaPct"] is not None else "", "tone": "teal", "fmt": "money"})
        kpis.append({"label": "📊 PBT (₹L)", "value": e["pbt"],
                      "sub": f"{e['pbtPct']*100:.1f}% margin" if e["pbtPct"] is not None else "", "tone": "purple", "fmt": "money"})
        kpis.append({"label": "📊 PAT (₹L)", "value": e["pat"],
                      "sub": f"{e['patPct']*100:.1f}% margin" if e["patPct"] is not None else "", "tone": "gold", "fmt": "money"})
        out[e["name"]] = kpis
    return out


# ==============================================================================
# Splice into HTML (same pattern as refresh_ceo_dashboard.py)
# ==============================================================================
def splice_const(html_content, const_name, value_obj):
    payload = json.dumps(value_obj, ensure_ascii=False, separators=(",", ":"))
    new_line = f"const {const_name} = {payload};\n"
    pattern = re.compile(r"const " + re.escape(const_name) + r" = \{.*?\};\n", re.S)
    if pattern.search(html_content):
        return pattern.sub(lambda m: new_line, html_content, count=1)
    raise RuntimeError(f"Could not find `const {const_name} = ...;` to replace")


def main():
    sys.stdout.reconfigure(encoding="utf-8")

    if not OUT_HTML_PATH.exists():
        print(f"ERROR: {OUT_HTML_PATH} not found -- run the HTML template setup first.")
        sys.exit(1)

    print("Reading the transaction extract ...")
    conso_rows = load_conso_data()
    print(f"  -> {len(conso_rows)} rows (FY25/FY26/FY27)")

    # 2026-08-11: Category/Brand/Zone/Sub-category/Top/Growth all switch to
    # the TRUE, elimination-adjusted basis (external-only) -- see
    # external_only()'s docstring. build_entities() below deliberately keeps
    # using the full, unfiltered `conso_rows` -- its "TOTAL (Sum of
    # Entities)" row is meant to stay pre-elimination.
    conso_rows_ext = external_only(conso_rows)
    print(f"  -> {len(conso_rows_ext)} rows after excluding RELATED PARTY (intercompany)")

    # Read the true, audited consolidated Sales figures early (needed to
    # rescale conso_rows_ext to tie exactly -- see rescale_external_to_true()).
    # Re-read again below via `val()` for the other P&L lines (EBITDA/PBT/PAT/
    # COGS) needed for the KPI tiles -- this workbook open is cheap and
    # keeping the two reads separate avoids restructuring the KPI section.
    wb_mis_early = openpyxl.load_workbook(MIS_PATH, data_only=True)
    ws_consol_early = wb_mis_early[os.environ.get("MIS_CONSOL_SHEET", "")]

    def val_early(row, col):
        v = ws_consol_early[f"{col}{row}"].value
        return v if isinstance(v, (int, float)) else 0.0

    true_sales_kpis = (val_early(4, "S"), val_early(4, "U"), val_early(4, "B"), val_early(4, "F"))
    # Computed ONCE from the full external-only dataset (see
    # compute_true_rescale_factors()'s docstring for why that matters), then
    # applied to every row-set that needs to tie to the true total.
    true_rescale_factors = compute_true_rescale_factors(conso_rows_ext, true_sales_kpis)
    conso_rows_ext = apply_rescale_factors(conso_rows_ext, true_rescale_factors)

    # 2026-08-11: Zone uses a hybrid source -- the other 6 entities are each
    # a single-country subsidiary, so their existing entity-location proxy
    # (conso_rows_ext, ENTITY_ZONE) IS their real zone; only Standalone (a
    # predominantly an exporter) needs its real, multi-zone split pulled in separately
    # from its own workbook. load_standalone_zone_rows() rescales that split
    # to tie EXACTLY to conso_rows_ext's own (already fully-rescaled)
    # Standalone total -- see that function's docstring.
    print("Reading Standalone's own Zone-wise-sales split ...")
    zone_source_rows = (
        [r for r in conso_rows_ext if r["entity"] != "ADF Foods Ltd (Standalone)"]
        + load_standalone_zone_rows(conso_rows_ext)
    )

    # ---------------- 4M dataset ----------------
    category_by_entity = build_dim_by_entity_4m(conso_rows_ext, "category", CATEGORY_ORDER)
    brand_by_entity = build_dim_by_entity_4m(conso_rows_ext, "brand", BRAND_ORDER)
    zone_by_entity = build_dim_by_entity_4m(zone_source_rows, "zone", ZONE_ORDER)

    top_category_by_entity, growth_category_by_entity = {}, {}
    top_brand_by_entity, growth_brand_by_entity = {}, {}
    top_zone_by_entity, growth_zone_by_entity = {}, {}
    topSub_by_entity = {}
    exec_category_by_entity, exec_brand_by_entity, exec_zone_by_entity = {}, {}, {}
    for scope in ["ALL"] + ENTITY_DISPLAY_ORDER:
        top_category_by_entity[scope], growth_category_by_entity[scope] = top_and_growth_from_monthly(category_by_entity[scope])
        top_brand_by_entity[scope], growth_brand_by_entity[scope] = top_and_growth_from_monthly(brand_by_entity[scope])
        top_zone_by_entity[scope], growth_zone_by_entity[scope] = top_and_growth_from_monthly(zone_by_entity[scope])
        topSub_by_entity[scope] = build_top_sub(filter_entity(conso_rows_ext, scope))
        exec_category_by_entity[scope] = [{"name": r["name"], "cur": r["ytdCur"], "py": r["ytdPy"], "grPct": r["q1Gr"], "signal": signal_from_growth(r["q1Gr"])} for r in category_by_entity[scope] if r["name"] != "TOTAL"]
        exec_brand_by_entity[scope] = [{"name": r["name"], "cur": r["ytdCur"], "py": r["ytdPy"], "grPct": r["q1Gr"], "signal": signal_from_growth(r["q1Gr"])} for r in brand_by_entity[scope] if r["name"] != "TOTAL"]
        exec_zone_by_entity[scope] = [{"name": r["name"], "cur": r["ytdCur"], "py": r["ytdPy"], "grPct": r["q1Gr"], "signal": signal_from_growth(r["q1Gr"])} for r in zone_by_entity[scope] if r["name"] != "TOTAL"]

    print("Reading MIS report (4M entity P&L + Group KPIs) ...")
    entities = build_entities(conso_rows)
    kpis_by_entity = build_exec_kpis_by_entity(entities, category_by_entity)

    wb_mis = openpyxl.load_workbook(MIS_PATH, data_only=True)
    ws_consol = wb_mis[os.environ.get("MIS_CONSOL_SHEET", "")]

    def val(row, col):
        v = ws_consol[f"{col}{row}"].value
        return v if isinstance(v, (int, float)) else 0.0

    sales_cur, sales_py = val(4, "S"), val(4, "U")
    ebitda_cur, ebitda_py = val(15, "S"), val(15, "U")
    pbt_cur, pbt_py = val(23, "S"), val(23, "U")
    pat_cur, pat_py = val(26, "S"), val(26, "U")
    cogs_cur, cogs_py = val(7, "S"), val(7, "U")
    margin_cur = (sales_cur - cogs_cur) / sales_cur if sales_cur else 0.0
    margin_py = (sales_py - cogs_py) / sales_py if sales_py else 0.0

    add_eliminations_row(entities, (sales_cur, sales_py, ebitda_cur, pbt_cur, pat_cur))

    # Reference tiles -- MUST use the SAME true, elimination-adjusted basis
    # as the KPIs above (row 4 col S/U), not the transaction extract’s pre-elimination
    # sum -- verified 2026-08-10: showing a pre-elimination "4M FY26"
    # reference tile right next to the Revenue tile's true-eliminated
    # "vs PY" figure looked like an outright wrong number (they're both
    # nominally "4M FY26" but materially apart), even though each was
    # individually correct on its own, undisclosed basis.
    # - 4M FY26 = sales_py itself (col U), already computed above.
    # - Q1 FY27 (Apr-Jun) = YTD FY27 (4M, col S) minus July 2026 alone
    #   (col B) -- derived by subtracting two independently-verified,
    #   unambiguously-labeled cells, NOT by trusting this sheet's own
    #   "Q2 FY26-27" column label, which turned out to just equal the
    #   July-2026 figure again (i.e. a quarter-to-date stub, not a real
    #   quarter total) rather than a genuine Apr-Jun total.
    # - Q4 FY26 (Jan-Mar 2026) has NO reliable true-eliminated source in
    #   this sheet at all (it only carries rolling current-year data) --
    #   dropped rather than guessed at.
    july_2026_sales = val(4, "B")
    q1_fy27_sales_true = sales_cur - july_2026_sales

    kpis_all = build_exec_kpis(sales_cur, sales_py, ebitda_cur, ebitda_py, pbt_cur, pbt_py, pat_cur, pat_py, margin_cur, margin_py,
                                q1_fy27_sales_true)

    data = {
        "sourceFile": f"{XLSM_PATH.name} + {MIS_PATH.name}",
        "entities": ENTITY_DISPLAY_ORDER,
        "exec": {
            "kpisAll": kpis_all, "kpisByEntity": kpis_by_entity,
            "category": exec_category_by_entity, "brand": exec_brand_by_entity, "zone": exec_zone_by_entity,
        },
        "profit": {"entities": entities},
        "sales": {"category": category_by_entity, "brand": brand_by_entity, "zone": zone_by_entity},
        "top": {"category": top_category_by_entity, "brand": top_brand_by_entity, "zone": top_zone_by_entity},
        "topSub": topSub_by_entity,
        "growth": {"category": growth_category_by_entity, "brand": growth_brand_by_entity, "zone": growth_zone_by_entity},
    }

    # ---------------- 12M dataset (annual only, audited FS-basis) ----------------
    print("Reading annual FY26/FY25 workbook (Category/Brand/Region/Entity) ...")
    annual_rows = load_annual_transactions()
    print(f"  -> {len(annual_rows)} rows (FY26/FY25)")

    category_by_entity_12 = build_dim_by_entity_annual(annual_rows, "category", ANNUAL_CATEGORY_ORDER)
    brand_by_entity_12 = build_dim_by_entity_annual(annual_rows, "brand", ANNUAL_BRAND_ORDER)
    zone_by_entity_12 = build_dim_by_entity_annual(annual_rows, "region", ANNUAL_REGION_ORDER)

    top_category_12, growth_category_12 = {}, {}
    top_brand_12, growth_brand_12 = {}, {}
    top_zone_12, growth_zone_12 = {}, {}
    topSub_12 = {}
    exec_category_12, exec_brand_12, exec_zone_12 = {}, {}, {}
    for scope in ["ALL"] + ENTITY_DISPLAY_ORDER:
        top_category_12[scope], growth_category_12[scope] = top_and_growth_from_annual(category_by_entity_12[scope])
        top_brand_12[scope], growth_brand_12[scope] = top_and_growth_from_annual(brand_by_entity_12[scope])
        top_zone_12[scope], growth_zone_12[scope] = top_and_growth_from_annual(zone_by_entity_12[scope])
        scoped_rows = annual_rows if scope == "ALL" else [r for r in annual_rows if r["entity"] == scope]
        topSub_12[scope] = build_top_sub_annual(scoped_rows)
        exec_category_12[scope] = leader_rows_from_annual(category_by_entity_12[scope])
        exec_brand_12[scope] = leader_rows_from_annual(brand_by_entity_12[scope])
        exec_zone_12[scope] = leader_rows_from_annual(zone_by_entity_12[scope])

    print("Reading audited segment figures (group revenue and PBT) ...")
    segment_kpis = load_segment_disclosure_kpis()
    print("Reading entity-wise audited P&L (Revenue/EBITDA/PBT/PAT, FY26 vs FY25) ...")
    fs_entity_pnl = load_fs_entity_pnl()
    entities_12m = build_entities_annual(fs_entity_pnl)
    kpis_all_12m = build_exec_kpis_annual(segment_kpis, entities_12m)
    kpis_by_entity_12m = build_exec_kpis_by_entity_annual(entities_12m)

    data_12m = {
        "entities": ENTITY_DISPLAY_ORDER,
        "exec": {
            "kpisAll": kpis_all_12m, "kpisByEntity": kpis_by_entity_12m,
            "category": exec_category_12, "brand": exec_brand_12, "zone": exec_zone_12,
        },
        "profit": {"entities": entities_12m},
        "sales": {"category": category_by_entity_12, "brand": brand_by_entity_12, "zone": zone_by_entity_12},
        "top": {"category": top_category_12, "brand": top_brand_12, "zone": top_zone_12},
        "topSub": topSub_12,
        "growth": {"category": growth_category_12, "brand": growth_brand_12, "zone": growth_zone_12},
    }

    html = OUT_HTML_PATH.read_text(encoding="utf-8")
    html = splice_const(html, "CEO_DATA", data)
    html = splice_const(html, "CEO_DATA_12M", data_12m)
    OUT_HTML_PATH.write_text(html, encoding="utf-8")
    print(f"Done. Wrote {OUT_HTML_PATH}")


if __name__ == "__main__":
    main()
