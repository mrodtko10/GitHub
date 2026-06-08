"""
update_dashboard.py
-------------------
Double-click (or run) this script any time you update Weekly_Commitments_Review.xlsx.
It re-reads the Excel data and regenerates budget_dashboard.html with fresh numbers.

Requirements: openpyxl  (pip install openpyxl)
"""

import json
import os
import re
import sys

try:
    import openpyxl
except ImportError:
    input("ERROR: openpyxl is not installed.\nRun:  pip install openpyxl\n\nPress Enter to close.")
    sys.exit(1)

# ---------------------------------------------------------------------------
# Paths – both files must live in the same folder as this script
# ---------------------------------------------------------------------------
HERE        = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH  = os.path.join(HERE, "Weekly_Commitments_Review.xlsx")
HTML_PATH   = os.path.join(HERE, "budget_dashboard.html")

# ---------------------------------------------------------------------------
# Column indices inside the Category Table tab (0-based)
# ---------------------------------------------------------------------------
COL_COST_CODE = 0
COL_ITEM      = 1

# CVR EAC  (SU=O, OC=P, MA=Q)
COL_SU_EAC, COL_OC_EAC, COL_MA_EAC = 14, 15, 16

# Accrued  (SU=U, OC=V, MA=W)
COL_SU_ACC, COL_OC_ACC, COL_MA_ACC = 20, 21, 22

# Actuals  (SU=X, OC=Y, MA=Z)
COL_SU_ACT, COL_OC_ACT, COL_MA_ACT = 23, 24, 25

# Committed (SU=AA, OC=AB, MA=AC)
COL_SU_COM, COL_OC_COM, COL_MA_COM = 26, 27, 28

# CTD      (SU=AG, OC=AH, MA=AI)
COL_SU_CTD, COL_OC_CTD, COL_MA_CTD = 32, 33, 34

# Rem EAC  (SU=AM, OC=AN, MA=AO)
COL_SU_REM, COL_OC_REM, COL_MA_REM = 38, 39, 40

# Over/Under (SU=AP, OC=AQ, MA=AR)
COL_SU_OU,  COL_OC_OU,  COL_MA_OU  = 41, 42, 43


def _num(cell_value):
    """Return float or 0 for blank/text cells."""
    try:
        return float(cell_value) if cell_value not in (None, "", "-") else 0.0
    except (TypeError, ValueError):
        return 0.0


def extract_data(ws):
    """
    Read Category Table worksheet.
    Row 0 = totals/CVR Month  (index 1 in openpyxl)
    Row 1 = SU/OC/MA sub-headers
    Row 2 = full column headers
    Row 3+ = data  → openpyxl row 4+
    """
    records = []
    all_rows = list(ws.iter_rows(values_only=True))

    # Pull CVR Month from row 0 (first row, somewhere around col 0 or visible header)
    cvr_month = None
    header_row = all_rows[0] if all_rows else []
    for cell in header_row:
        if isinstance(cell, str) and cell.strip():
            # e.g. "May, 2026" or "CVR Month: May, 2026"
            text = cell.strip().replace("CVR Month:", "").replace("CVR Month :", "").strip()
            if text:
                cvr_month = text
                break

    for row in all_rows[3:]:  # skip header rows 0-2
        cost_code = row[COL_COST_CODE] if len(row) > COL_COST_CODE else None
        item      = row[COL_ITEM]      if len(row) > COL_ITEM      else None

        # Skip blank / total rows
        if not cost_code or str(cost_code).strip() in ("", "Cost Code", "Total"):
            continue

        # Determine over/under status (majority vote across categories)
        ou_vals = [_num(row[c]) if len(row) > c else 0 for c in (COL_SU_OU, COL_OC_OU, COL_MA_OU)]
        ou_sum  = sum(ou_vals)
        ou_status = "over" if ou_sum < 0 else "under"

        def g(col):
            return _num(row[col]) if len(row) > col else 0.0

        records.append({
            "cost_code": str(cost_code).strip(),
            "item":      str(item).strip() if item else "",
            "su_eac": g(COL_SU_EAC), "oc_eac": g(COL_OC_EAC), "ma_eac": g(COL_MA_EAC),
            "su_acc": g(COL_SU_ACC), "oc_acc": g(COL_OC_ACC), "ma_acc": g(COL_MA_ACC),
            "su_act": g(COL_SU_ACT), "oc_act": g(COL_OC_ACT), "ma_act": g(COL_MA_ACT),
            "su_com": g(COL_SU_COM), "oc_com": g(COL_OC_COM), "ma_com": g(COL_MA_COM),
            "su_ctd": g(COL_SU_CTD), "oc_ctd": g(COL_OC_CTD), "ma_ctd": g(COL_MA_CTD),
            "su_rem": g(COL_SU_REM), "oc_rem": g(COL_OC_REM), "ma_rem": g(COL_MA_REM),
            "ou_status": ou_status,
        })

    return records, cvr_month


def update_html(records, cvr_month):
    with open(HTML_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    # Replace RAW_DATA
    new_raw = "const RAW_DATA = " + json.dumps(records, separators=(",", ":")) + ";"
    html = re.sub(r"const RAW_DATA\s*=\s*\[.*?\];", new_raw, html, flags=re.DOTALL)

    # Replace CVR Month in <title> and badge span
    if cvr_month:
        html = re.sub(
            r"(Budget Dashboard\s*[–-]\s*)([^<\"]+)",
            lambda m: m.group(1) + cvr_month,
            html,
        )
        html = re.sub(
            r"(CVR Month:\s*)([^<\"]+)",
            lambda m: m.group(1) + cvr_month,
            html,
        )

    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)


def main():
    if not os.path.exists(EXCEL_PATH):
        input(f"ERROR: Could not find Excel file at:\n  {EXCEL_PATH}\n\nPress Enter to close.")
        sys.exit(1)

    if not os.path.exists(HTML_PATH):
        input(f"ERROR: Could not find HTML dashboard at:\n  {HTML_PATH}\n\nPress Enter to close.")
        sys.exit(1)

    print("Reading Excel workbook...")
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

    # Find the Category Table sheet (case-insensitive)
    sheet_name = None
    for name in wb.sheetnames:
        if "category" in name.lower() and "table" in name.lower():
            sheet_name = name
            break
    if not sheet_name:
        # Fallback: try first sheet that isn't Dashboard or PO
        for name in wb.sheetnames:
            if "dashboard" not in name.lower() and "po" not in name.lower():
                sheet_name = name
                break
    if not sheet_name:
        sheet_name = wb.sheetnames[0]

    print(f"  Using sheet: {sheet_name}")
    ws = wb[sheet_name]

    records, cvr_month = extract_data(ws)
    print(f"  Extracted {len(records)} cost code records.")
    if cvr_month:
        print(f"  CVR Month: {cvr_month}")

    print("Updating dashboard HTML...")
    update_html(records, cvr_month)
    print(f"  Saved: {HTML_PATH}")

    input("\nDone! Dashboard updated successfully.\nPress Enter to close.")


if __name__ == "__main__":
    main()
