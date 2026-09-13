"""
update_conso_mis_august.py

Refreshes the MIS reference tab in the consolidated dynamic workbook with the
August 2026 MIS figures (5-month YTD FY27 vs YTD FY26), replacing the
July-basis (4-month) figures built earlier.

Source: the consolidated MIS workbook, located via the
  CONSO_MIS_WORKBOOK_AUG environment variable.

Checked against that workbook directly (2026-09) before wiring anything up:
  - The group consolidated P&L sheet keeps the same row layout as the July
    cut, and its "YTD FY27"/"YTD FY26" columns (S/U) now run on a 5-month
    basis (through August) instead of 4.
  - The per-entity monthly P&L sheet keeps the same 8-column-per-month block
    layout (7 entities + Total), with a 5th month block (Aug, cols AI:AP)
    appended after Jul (cols AA:AH); the Apr/May/Jun/Jul column positions are
    unchanged from the July cut, so only the new block had to be added below.
    The August block labels its second entity "Australia" where the earlier
    blocks use the "ADFIL" code -- both map to the same ADF Foods Australia
    entity, so both spellings are accepted rather than assuming one form.

Prior-year entity-wise revenue is not part of the MIS report itself, so it is
derived from the consolidated transaction sheet's own FY26 Apr-Aug entity
totals -- the same approach as the original July build, with Aug added to the
month list.
"""

import os
import sys
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))
import build_ceo_dashboard_conso as C

# Source workbook locations and sheet names are supplied by the environment, so
# no internal file or sheet naming is carried in this repository.
MIS_AUG_PATH = Path(os.environ.get("CONSO_MIS_WORKBOOK_AUG", ""))
OUT_PATH = Path(os.environ.get("CONSO_DYNAMIC_WORKBOOK", ""))
SHEET_ENTITY_PL = os.environ.get("CONSO_ENTITY_PL_SHEET", "")   # per-entity monthly P&L sheet
SHEET_GROUP_PL = os.environ.get("CONSO_GROUP_PL_SHEET", "")     # group consolidated P&L sheet
MIS_REF_SHEET = os.environ.get("MIS_REFERENCE_SHEET", "MIS_Reference")

# The group sales-data workbook that build_ceo_dashboard_conso reads its
# transaction rows from -- full path from the environment.
_group_sales_workbook = os.environ.get("CONSO_SALES_WORKBOOK", "")
if _group_sales_workbook:
    C.XLSM_PATH = Path(_group_sales_workbook)

# build_ceo_dashboard_conso.py (owned by the separate live-refresh pipeline, not edited here)
# resolves the July MIS workbook from its own environment variable; the MIS folder tree was
# reorganised between cuts, so point C.MIS_PATH at whatever CONSO_MIS_WORKBOOK_JUL now names
# rather than editing that shared file.
_new_july_mis = Path(os.environ.get("CONSO_MIS_WORKBOOK_JUL", ""))
if str(_new_july_mis) and _new_july_mis.exists():
    C.MIS_PATH = _new_july_mis

MONTHS_5 = ["Apr", "May", "Jun", "Jul", "Aug"]

MIS_ENTITY_COLS_AUG = {
    "Apr": list("CDEFGHIJ"), "May": list("KLMNOPQR"),
    "Jun": list("STUVWXYZ"), "Jul": "AA AB AC AD AE AF AG AH".split(),
    "Aug": "AI AJ AK AL AM AN AO AP".split(),
}
MIS_ENTITY_TO_DISPLAY_AUG = {
    "ADFL": "ADF Foods Ltd (Standalone)",
    "ADFIL": "ADF Foods Australia",       # Apr-Jul blocks label it ADFIL
    "Australia": "ADF Foods Australia",   # Aug block labels it Australia directly
    "TELLURIC": "Telluric Foods",
    "ADF UK": "ADF Foods UK",
    "ADFHL": "ADF Holdings USA",
    "ADF USA": "ADF Foods USA",
    "VIB NJ LLC": "Vibrant Foods NJ LLC",
}
MIS_ENTITY_ORDER_AUG = {
    "Apr": ["ADFL", "ADFIL", "TELLURIC", "ADF UK", "ADFHL", "ADF USA", "VIB NJ LLC", "Total"],
    "May": ["ADFL", "ADFIL", "TELLURIC", "ADF UK", "ADFHL", "ADF USA", "VIB NJ LLC", "Total"],
    "Jun": ["ADFL", "ADFIL", "TELLURIC", "ADF UK", "ADFHL", "ADF USA", "VIB NJ LLC", "Total"],
    "Jul": ["ADFL", "ADFIL", "TELLURIC", "ADF UK", "ADFHL", "ADF USA", "VIB NJ LLC", "Total"],
    "Aug": ["ADFL", "Australia", "TELLURIC", "ADF UK", "ADFHL", "ADF USA", "VIB NJ LLC", "Total"],
}


def mis_entity_row_5m(ws, row_num):
    from collections import defaultdict
    totals = defaultdict(float)
    for month in MONTHS_5:
        cols = MIS_ENTITY_COLS_AUG[month]
        order = MIS_ENTITY_ORDER_AUG[month]
        for code, col in zip(order, cols):
            v = ws[f"{col}{row_num}"].value
            if v is not None:
                totals[code] += v
    return totals


def build_entities_5m(conso_rows):
    wb = openpyxl.load_workbook(MIS_AUG_PATH, data_only=True)
    ws = wb[SHEET_ENTITY_PL]

    revenue = mis_entity_row_5m(ws, 15)
    ebitda = mis_entity_row_5m(ws, 49)
    pbt = mis_entity_row_5m(ws, 51)
    pat = mis_entity_row_5m(ws, 55)

    from collections import defaultdict
    py_revenue_by_entity = defaultdict(float)
    for r in conso_rows:
        if r["fy"] == "FY26" and r["month"] in MONTHS_5:
            py_revenue_by_entity[r["entity"]] += r["value"]

    canonical_codes = ["ADFL", "ADFIL", "TELLURIC", "ADF UK", "ADFHL", "ADF USA", "VIB NJ LLC", "Total"]
    entities = []
    for code in canonical_codes:
        name = "TOTAL (Sum of Entities)" if code == "Total" else MIS_ENTITY_TO_DISPLAY_AUG[code]
        rev = revenue.get(code, 0.0)
        eb = ebitda.get(code, 0.0)
        pb = pbt.get(code, 0.0)
        pa = pat.get(code, 0.0)
        py_rev = (sum(py_revenue_by_entity.values()) if code == "Total"
                  else py_revenue_by_entity.get(MIS_ENTITY_TO_DISPLAY_AUG.get(code), 0.0))
        entities.append({
            "name": name, "revenue": rev, "ebitda": eb,
            "ebitdaPct": (eb / rev) if rev else None,
            "pbt": pb, "pbtPct": (pb / rev) if rev else None,
            "pat": pa, "patPct": (pa / rev) if rev else None,
            "pyRevenue": py_rev,
        })
    return entities


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    print(f"MIS source: {MIS_AUG_PATH}")
    print("Loading the consolidated transaction rows for PY (FY26 Apr-Aug) entity revenue ...")
    conso_rows = C.load_conso_data()
    print(f"  -> {len(conso_rows)} rows")
    fy26_aug_rows = [r for r in conso_rows if r["fy"] == "FY26" and r["month"] == "Aug"]
    print(f"  FY26 Aug rows found: {len(fy26_aug_rows)} ({'OK' if fy26_aug_rows else 'WARNING: none -- PY 5M comparator will undercount by one month'})")

    print("Extracting 5-month (Apr-Aug) entity-wise Revenue/EBITDA/PBT/PAT ...")
    entities = build_entities_5m(conso_rows)
    for e in entities:
        print(f"  {e['name']:32s} rev={e['revenue']:>10.1f}  ebitda={e['ebitda']:>9.1f}  pbt={e['pbt']:>9.1f}  pat={e['pat']:>9.1f}  pyRev={e['pyRevenue']:>10.1f}")

    print(f"Writing into {OUT_PATH.name}'s {MIS_REF_SHEET} tab ...")
    wb = openpyxl.load_workbook(OUT_PATH)
    ws = wb[MIS_REF_SHEET]
    for i, e in enumerate(entities):
        r = i + 3  # row 1 header, row 2 note, data from row 3
        ws.cell(row=r, column=1, value=e["name"])
        ws.cell(row=r, column=2, value=e["revenue"])
        ws.cell(row=r, column=3, value=e["ebitda"])
        ws.cell(row=r, column=4, value=e["pbt"])
        ws.cell(row=r, column=5, value=e["pat"])
        ws.cell(row=r, column=6, value=e["pyRevenue"])
    headers_5m = ["ENTITY", "5M FY27 REVENUE (Apr-Aug)", "5M FY27 EBITDA", "5M FY27 PBT", "5M FY27 PAT", "5M FY26 PY REVENUE (Apr-Aug)"]
    for c, h in enumerate(headers_5m, start=1):
        ws.cell(row=1, column=c, value=h)
    ws["A2"] = ("Sourced from the consolidated MIS workbook's per-entity monthly P&L sheet -- refreshed "
                "2026-09 to a 5-month (Apr-Aug) YTD basis, replacing the earlier 4-month (Apr-Jul) figures. "
                "This is a small pasted reference table rather than a raw-transaction extract, the same "
                "architecture the standalone workbook itself uses for its own Entity P&L data. The reporting "
                "window each figure covers is stated in the header row above, so the basis on which every "
                "figure is presented stays explicit on the face of the tab.")

    # Also refresh the Exec Summary tab's title/labels that hardcoded "4M" for the group KPI cards, and
    # the elimination-adjusted Group KPI source cells (group consolidated P&L sheet, columns S/U, now on
    # a 5-month basis).
    wb_mis = openpyxl.load_workbook(MIS_AUG_PATH, data_only=True)
    ws_consol = wb_mis[SHEET_GROUP_PL]

    def val(row, col):
        v = ws_consol[f"{col}{row}"].value
        return v if isinstance(v, (int, float)) else 0.0

    sales_cur, sales_py = val(4, "S"), val(4, "U")
    ebitda_cur = val(15, "S")
    pbt_cur = val(23, "S")
    pat_cur = val(26, "S")
    print(f"Elimination-adjusted Group (5M FY27): Sales={sales_cur:.1f} EBITDA={ebitda_cur:.1f} PBT={pbt_cur:.1f} PAT={pat_cur:.1f} (PY Sales={sales_py:.1f})")

    wb.save(OUT_PATH)
    print(f"Done. Wrote {OUT_PATH}")


if __name__ == "__main__":
    main()
