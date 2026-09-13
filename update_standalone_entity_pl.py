"""
update_standalone_entity_pl.py (v2)

Wires ADFL's (Standalone) monthly Revenue/EBITDA/PBT/PAT into
' ENTITY-WISE P&L-'!B4/D4/F4/H4 -- previously static numbers labeled
"4M FY27", now live formulas driven by the same Settings!$B$6 Reporting
Window selector used everywhere else in this workbook.

v2 SWITCHES SOURCE from the standalone "Standalone MIS_August 2026-v1.xlsx"
file's 'Total' sheet (v1's source) to 'Summary-ADFL Conso 2025' inside the
Conso MIS workbook (located via the CONSO_MIS_WORKBOOK_AUG environment
variable) -- the SAME sheet already used for the other 6
entities in the Conso Dynamic workbook. Two reasons, both raised by the
user 2026-09:

  1. CONSISTENCY: traced the actual formula behind this sheet's own
     'EBIDTA' row (C49 = C21-C45+C44+C41 = PBT + Depreciation + Finance
     Cost) -- it does NOT subtract Other Income/Forex the way this
     project's permanent EBITDA formula does (confirmed materially
     different for one entity-month, where the board convention and a
     properly Other-Income-adjusted figure diverged by a large margin,
     mostly a one-off forex gain). User's explicit decision: keep the board's own EBITDA
     convention as-is for all 7 entities in Conso, rather than rebuild
     each entity's Other Income adjustment. Standalone's ADFL row now
     matches that same convention/source for consistency -- same entity,
     shouldn't show two different EBITDA numbers depending which
     workbook you're looking at.
  2. COMPLETENESS: this sheet also gives genuine PAT (row 55) directly --
     the old 'Total' sheet source had no tax-provision line at all and
     PAT had to stay static/undisclosed-as-unavailable. Fixed here.

Revenue = Sales (row 15) + Other Operating Income (row 16) -- verified
this equals 'Revenue from Operations' exactly (row 17's own SUM(15:16)
formula in the source, and cross-checked against 'ADFL MIS '!B6 "Revenue
from op." for August: identical both ways). Using Sales (row 15) alone
understates revenue by leaving out real operating income (export
incentives) -- the mistake in v1 of this script, caught by the user.

Coverage: only Apr-Aug 2026 (the months this MIS cut actually has,
matching the other 6 entities in Conso) -- Sep-Mar stay 0 until later
months' Conso MIS files are supplied and this tab is re-run, same
manually-updated-reference-table architecture as RawData_MIS in the
Conso workbook (not raw-transaction-driven).
"""

import os
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

BASE = Path(__file__).resolve().parent
# Source workbook location comes from the environment -- nothing
# machine-specific is hard-coded here.
MIS_CONSO_PATH = Path(os.environ.get("CONSO_MIS_WORKBOOK_AUG", ""))
OUT_PATH = BASE / "CEO_Sales_Analytics_YTD JULY-26_StandaloneDynamic.xlsx"

MONTH_ORDER = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"]
# ADFL's own column per month block in 'Summary-ADFL Conso 2025' -- same layout used for the other 6
# entities (update_conso_mis_august.py's MIS_ENTITY_COLS_AUG), ADFL is always the first entity column.
ADFL_COL_BY_MONTH = {"Apr": "C", "May": "K", "Jun": "S", "Jul": "AA", "Aug": "AI"}

NAVY, WHITE, BODY_FONT = "FF1B3A6B", "FFFFFFFF", "FF4A4A4A"


def fill(c):
    return PatternFill("solid", fgColor=c)


def extract_monthly_pl():
    wb = openpyxl.load_workbook(MIS_CONSO_PATH, data_only=True)
    ws = wb["Summary-ADFL Conso 2025"]

    def val(col, row):
        v = ws[f"{col}{row}"].value
        return v if isinstance(v, (int, float)) else 0.0

    out = []
    for m in MONTH_ORDER:
        if m not in ADFL_COL_BY_MONTH:
            out.append({"month": m, "revenue": 0.0, "ebitda": 0.0, "pbt": 0.0, "pat": 0.0})
            continue
        col = ADFL_COL_BY_MONTH[m]
        # "Sales" only (row 15) -- NOT combined with Other Operating Income (row 16) -- per explicit
        # user instruction 2026-09 (confirmed against a screenshot of the month's Sales row, matching
        # exactly). An earlier version of this script added row 16 to get "Revenue from Operations";
        # reverted here per that instruction.
        revenue = val(col, 15)
        ebitda = val(col, 49)
        pbt = val(col, 51)
        pat = val(col, 55)
        out.append({"month": m, "revenue": revenue, "ebitda": ebitda, "pbt": pbt, "pat": pat})
    wb.close()
    return out


def write_raw_tab(wb, monthly):
    name = "RawData_MIS_Standalone"
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
    note = ("Source: Summary-ADFL Conso 2025 sheet inside the Conso MIS file (ADF_Conso MIS_August 26-... -- "
            "the SAME sheet used for the other 6 entities in the Conso Dynamic workbook), ADFL's own column per "
            "month block. Sales = row 15 only, per explicit user instruction 2026-09 (confirmed against a "
            "screenshot of the month's Sales row, matching this row exactly) -- NOT combined with Other "
            "Operating Income (row 16), consistent with how Conso Dynamic's own RawData_MIS already sources "
            "Sales for all 7 entities (verified same row, same convention). EBITDA = PBT + Depreciation + Finance Cost, per THIS SOURCE's "
            "own formula (does NOT subtract Other Income/Forex) -- deliberately matches the board-reported "
            "convention used for all 7 entities in Conso Dynamic (user's explicit decision 2026-09), NOT this "
            "project's own permanent EBITDA formula (which would subtract Other Income) -- the two would "
            "otherwise show two different EBITDA numbers for the same entity depending which workbook you open. "
            "PAT is now genuinely available (row 55) -- the earlier version of this tab sourced from a "
            "different file with no tax-provision line and had to leave PAT static/undisclosed. Only Apr-Aug "
            "2026 exist in the MIS cut supplied so far -- Sep-Mar are 0 until later months' Conso MIS files are "
            "provided and this tab is re-run (manually-updated reference table, not raw-transaction-driven).")
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
            running.append(f"RawData_MIS_Standalone!{metric_col_letter}{2+i}")
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
                "ADFL (Standalone) row's Revenue/EBITDA/PBT/PAT are now live from RawData_MIS_Standalone (sourced "
                "from the same board MIS file used for the other 6 entities), following the Settings!$B$6 "
                "Reporting Window selector. EBITDA here uses the board's own convention (PBT+Dep+Finance Cost, "
                "no Other Income adjustment) -- matches the other 6 entity rows and Conso Dynamic, NOT this "
                "project's permanent EBITDA formula. The FY26 comparator (cols L/N/P/R) remains the last "
                "statically pasted figure -- no prior-year-monthly source has been supplied yet. The other 8 "
                "entity rows (ADFIL/Australia, TELLURIC, ADF UK, ADFHL, ADF USA, VIB NJ LLC, Elimination) are "
                "unchanged.")
    n.font = Font(size=9, italic=True, color=BODY_FONT)
    n.alignment = Alignment(wrap_text=True, vertical="top")
    print("  Rewired ' ENTITY-WISE P&L-' B4/D4/F4/H4 (ADFL Revenue/EBITDA/PBT/PAT) to Window Total formulas")


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    print(f"Extracting ADFL monthly P&L from {MIS_CONSO_PATH.name} ('Summary-ADFL Conso 2025') ...")
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
