"""
update_standalone_entity_pl.py (v2)

Wires ADFL's (Standalone) monthly Revenue/EBITDA/PBT/PAT into
' ENTITY-WISE P&L-'!B4/D4/F4/H4 -- previously static numbers labelled
"4M FY27", now live formulas driven by the same Settings!$B$6 Reporting
Window selector used everywhere else in this workbook.

v2 SWITCHES SOURCE from a separate standalone MIS file's own total sheet
(v1's source) to the per-entity monthly P&L sheet inside the consolidated MIS
workbook (located via the CONSO_MIS_WORKBOOK_AUG environment variable) -- the
SAME sheet already used for the other 6 entities in the consolidated dynamic
workbook. Two reasons:

  1. CONSISTENCY: that sheet states EBITDA on the board convention
     (PBT + Depreciation + Finance Cost), whereas this project's own EBITDA
     definition also subtracts Other Income/Forex, which can move the figure
     materially in a month carrying a one-off forex gain. Both are legitimate
     definitions; the decision by design is to present all 7 entities on the
     board convention here, labelled as such, so the same entity never shows
     two differently defined EBITDA figures depending on which workbook is
     open. Where this project computes EBITDA itself, it is built from its
     components rather than read from a summary row, so it stays consistent
     everywhere it appears.
  2. COMPLETENESS: this sheet also carries a genuine PAT line directly, which
     v1's source did not, so PAT is now live instead of static.

Revenue is taken as the Sales line only, by design -- deliberately not
combined with Other Operating Income (export incentives), matching how the
consolidated dynamic workbook's own MIS reference tab already sources Sales
for all 7 entities, so the two agree row for row. (v1 of this script combined
the two into "Revenue from Operations"; corrected here.)

Coverage: only Apr-Aug 2026 (the months this MIS cut carries, matching the
other 6 entities in the consolidated workbook) -- Sep-Mar stay 0 until later
months' MIS cuts are supplied and this tab is re-run, the same
manually-updated-reference-table architecture as the MIS reference tab in the
consolidated workbook (not raw-transaction-driven).
"""

import os
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

BASE = Path(__file__).resolve().parent
# Source workbook locations and sheet names are supplied by the environment, so
# no internal file or sheet naming is carried in this repository.
MIS_CONSO_PATH = Path(os.environ.get("CONSO_MIS_WORKBOOK_AUG", ""))
OUT_PATH = Path(os.environ.get("STANDALONE_DYNAMIC_WORKBOOK", ""))
SHEET_ENTITY_PL = os.environ.get("CONSO_ENTITY_PL_SHEET", "")   # per-entity monthly P&L sheet
MIS_REF_SHEET = os.environ.get("MIS_REFERENCE_SHEET", "MIS_Reference")
MIS_REF_SHEET_STANDALONE = MIS_REF_SHEET + "_Standalone"

MONTH_ORDER = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
# ADFL's own column per month block in the per-entity monthly P&L sheet -- the same layout used for
# the other 6 entities (update_conso_mis_august.py's MIS_ENTITY_COLS_AUG); ADFL is always the first
# entity column of each block.
ADFL_COL_BY_MONTH = {"Apr": "C", "May": "K", "Jun": "S", "Jul": "AA", "Aug": "AI"}

NAVY, WHITE, BODY_FONT = "FF1B3A6B", "FFFFFFFF", "FF4A4A4A"


def fill(c):
    return PatternFill("solid", fgColor=c)


def extract_monthly_pl():
    wb = openpyxl.load_workbook(MIS_CONSO_PATH, data_only=True)
    ws = wb[SHEET_ENTITY_PL]

    def val(col, row):
        v = ws[f"{col}{row}"].value
        return v if isinstance(v, (int, float)) else 0.0

    out = []
    for m in MONTH_ORDER:
        if m not in ADFL_COL_BY_MONTH:
            out.append({"month": m, "revenue": 0.0, "ebitda": 0.0, "pbt": 0.0, "pat": 0.0})
            continue
        col = ADFL_COL_BY_MONTH[m]
        # Sales line only -- by design, deliberately not combined with Other Operating Income, so this
        # matches how the consolidated dynamic workbook's own MIS reference tab sources Sales for all
        # 7 entities and the two agree row for row. (An earlier version of this script added the Other
        # Operating Income row to report "Revenue from Operations"; reverted here.)
        revenue = val(col, 15)
        ebitda = val(col, 49)
        pbt = val(col, 51)
        pat = val(col, 55)
        out.append({"month": m, "revenue": revenue, "ebitda": ebitda, "pbt": pbt, "pat": pat})
    wb.close()
    return out


def write_raw_tab(wb, monthly):
    name = MIS_REF_SHEET_STANDALONE
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    cols = ["MONTH", "SALES", "EBITDA (board convention)", "PBT", "PAT"]
    ws.append(cols)
    for c in range(1, len(cols) + 1):
        ws.cell(row=1, column=c).font = Font(bold=True, color=WHITE)
        ws.cell(row=1, column=c).fill = fill(NAVY)
    for r in monthly:
        ws.append([r["month"], r["revenue"], r["ebitda"], r["pbt"], r["pat"]])
    note = ("Source: the per-entity monthly P&L sheet inside the consolidated MIS workbook -- the same "
            "sheet used for the other 6 entities in the consolidated dynamic workbook -- reading ADFL's own "
            "column in each month block. Revenue is the Sales line only, by design, and deliberately not "
            "combined with Other Operating Income, matching how the consolidated dynamic workbook's own MIS "
            "reference tab sources Sales for all 7 entities, so the two agree row for row. EBITDA is stated "
            "on the board convention (PBT + Depreciation + Finance Cost, no Other Income adjustment) -- the "
            "same basis used for all 7 entities in the consolidated workbook, and labelled as such here; "
            "this project's own EBITDA definition also subtracts Other Income, so the basis is stated "
            "wherever the figure appears. PAT comes straight from the source's own tax-provisioned line, "
            "which v1's source did not carry. Only Apr-Aug 2026 exist in the MIS cut supplied so far -- "
            "Sep-Mar are 0 until later cuts are provided and this tab is re-run (manually-updated reference "
            "table, not raw-transaction-driven).")
    ws["A16"] = note
    ws.merge_cells("A16:E22")
    ws["A16"].font = Font(italic=True, size=9, color=BODY_FONT)
    ws["A16"].alignment = Alignment(wrap_text=True, vertical="top")
    for c, w in zip("ABCDE", [10, 18, 16, 12, 12]):
        ws.column_dimensions[c].width = w
    print(f"  Wrote {name}: {len(monthly)} months")


def rewire_entity_pl(wb):
    ws = wb[" ENTITY-WISE P&L-"]

    def choose_formula(metric_col_letter):
        running = []
        cum = []
        for i in range(12):
            running.append(f"{MIS_REF_SHEET_STANDALONE}!{metric_col_letter}{2+i}")
            cum.append("+".join(running))
        return "=CHOOSE(Settings!$B$7," + ",".join(cum) + ")"

    ws["B4"] = choose_formula("B")  # Revenue
    ws["D4"] = choose_formula("C")  # EBITDA
    ws["F4"] = choose_formula("D")  # PBT
    ws["H4"] = choose_formula("E")  # PAT -- now genuinely available
    for addr in ("B4", "D4", "F4", "H4"):
        ws[addr].number_format = "#,##0.0"

    ws["A1"] = '=Settings!$B$6&" "&TEXT(Settings!$B$3,"mmm-yy")&" onward -- CURRENT (ADFL row only, live; other entities static)"'

    note_row_num = 16
    ws.merge_cells(f"A{note_row_num}:S{note_row_num+3}")
    n = ws.cell(row=note_row_num, column=1, value=
                "ADFL (Standalone) row's Revenue/EBITDA/PBT/PAT are now live from the standalone MIS "
                "reference tab (sourced from the same board MIS workbook used for the other 6 entities), "
                "following the Settings!$B$6 Reporting Window selector. EBITDA here is stated on the board "
                "convention (PBT + Dep + Finance Cost, no Other Income adjustment) -- the same basis as the "
                "other 6 entity rows and the consolidated workbook, not this project's own EBITDA "
                "definition, which also subtracts Other Income. The FY26 comparator (cols L/N/P/R) remains "
                "the last statically pasted figure until a prior-year monthly source is added. The other 8 "
                "entity rows (ADFIL/Australia, TELLURIC, ADF UK, ADFHL, ADF USA, VIB NJ LLC, Elimination) "
                "are unchanged.")
    n.font = Font(size=9, italic=True, color=BODY_FONT)
    n.alignment = Alignment(wrap_text=True, vertical="top")
    print("  Rewired ' ENTITY-WISE P&L-' B4/D4/F4/H4 (ADFL Revenue/EBITDA/PBT/PAT) to Window Total formulas")


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    print(f"Extracting ADFL monthly P&L from {MIS_CONSO_PATH.name} (per-entity monthly P&L sheet) ...")
    monthly = extract_monthly_pl()
    for r in monthly:
        print(f"  {r['month']}: Revenue={r['revenue']:.1f} EBITDA={r['ebitda']:.1f} PBT={r['pbt']:.1f} PAT={r['pat']:.1f}")

    print(f"Opening {OUT_PATH.name} ...")
    wb = openpyxl.load_workbook(OUT_PATH)
    write_raw_tab(wb, monthly)
    rewire_entity_pl(wb)
    wb.save(OUT_PATH)
    print(f"Done. Wrote {OUT_PATH}")


if __name__ == "__main__":
    main()
