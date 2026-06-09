"""
Extract updated data from the new xlsm file and update budget_dashboard.html
"""
import json
import re
import openpyxl

XLSM = "/root/.claude/uploads/acf59d80-0730-52d2-ad54-91944376be86/69824955-244002_Weekly_Commitments_Review.xlsm"
HTML  = "/home/user/GitHub/budget_dashboard.html"

def N(v):
    try: return float(v) if v not in (None, "") else 0.0
    except: return 0.0

wb = openpyxl.load_workbook(XLSM, data_only=True)

# ── Find Category Table sheet ─────────────────────────────────────────────────
ws_cat = next((s for s in wb.sheetnames if "category" in s.lower() and "table" in s.lower()), None)
assert ws_cat, "Category Table sheet not found"
ws = wb[ws_cat]

# ── CVR Month from row 1 ──────────────────────────────────────────────────────
cvr_month = ""
for c in range(1, 9):
    v = str(ws.cell(1, c).value or "").strip()
    if v and v != "0":
        v = v.replace("CVR Month:", "").replace("CVR Month :", "").strip()
        cvr_month = v
        break

print(f"CVR Month: {cvr_month}")

# ── Column mapping (1-based) ──────────────────────────────────────────────────
# A=1 Cost Code, B=2 Item
# O=15 SU EAC, P=16 OC EAC, Q=17 MA EAC
# R=18 SU PO Log, S=19 OC PO Log, T=20 MA PO Log  (NEW)
# U=21 SU Accrued, V=22 OC Accrued, W=23 MA Accrued
# X=24 SU Actuals, Y=25 OC Actuals, Z=26 MA Actuals
# AA=27 SU Committed, AB=28 OC Committed, AC=29 MA Committed
# AG=33 SU CTD, AH=34 OC CTD, AI=35 MA CTD
# AJ=36 SU Man Adj, AK=37 OC Man Adj, AL=38 MA Man Adj  (NEW)
# AM=39 SU Rem, AN=40 OC Rem, AO=41 MA Rem
# AP=42 SU OU, AQ=43 OC OU, AR=44 MA OU

# Determine last row
last_row = ws.max_row
records = []
skip_vals = {"", "cost code", "total", "0"}

for i in range(4, last_row + 1):
    cc   = str(ws.cell(i, 1).value or "").strip()
    item = str(ws.cell(i, 2).value or "").strip()
    if cc.lower() in skip_vals:
        continue

    su_eac = N(ws.cell(i, 15).value)
    oc_eac = N(ws.cell(i, 16).value)
    ma_eac = N(ws.cell(i, 17).value)
    su_po  = N(ws.cell(i, 18).value)
    oc_po  = N(ws.cell(i, 19).value)
    ma_po  = N(ws.cell(i, 20).value)
    su_acc = N(ws.cell(i, 21).value)
    oc_acc = N(ws.cell(i, 22).value)
    ma_acc = N(ws.cell(i, 23).value)
    su_act = N(ws.cell(i, 24).value)
    oc_act = N(ws.cell(i, 25).value)
    ma_act = N(ws.cell(i, 26).value)
    su_com = N(ws.cell(i, 27).value)
    oc_com = N(ws.cell(i, 28).value)
    ma_com = N(ws.cell(i, 29).value)
    su_ctd = N(ws.cell(i, 33).value)
    oc_ctd = N(ws.cell(i, 34).value)
    ma_ctd = N(ws.cell(i, 35).value)
    su_adj = N(ws.cell(i, 36).value)
    oc_adj = N(ws.cell(i, 37).value)
    ma_adj = N(ws.cell(i, 38).value)
    su_rem = N(ws.cell(i, 39).value)
    oc_rem = N(ws.cell(i, 40).value)
    ma_rem = N(ws.cell(i, 41).value)

    # OU status: derive from rem (same logic as VBA)
    total_rem = su_rem + oc_rem + ma_rem
    total_eac = su_eac + oc_eac + ma_eac
    if total_eac == 0 and total_rem == 0:
        ou = "N/A"
    elif total_rem < 0:
        ou = "Over"
    else:
        ou = "Under"

    records.append({
        "cost_code": cc,
        "item": item,
        "su_eac": su_eac, "oc_eac": oc_eac, "ma_eac": ma_eac,
        "su_po":  su_po,  "oc_po":  oc_po,  "ma_po":  ma_po,
        "su_acc": su_acc, "oc_acc": oc_acc, "ma_acc": ma_acc,
        "su_act": su_act, "oc_act": oc_act, "ma_act": ma_act,
        "su_com": su_com, "oc_com": oc_com, "ma_com": ma_com,
        "su_ctd": su_ctd, "oc_ctd": oc_ctd, "ma_ctd": ma_ctd,
        "su_adj": su_adj, "oc_adj": oc_adj, "ma_adj": ma_adj,
        "su_rem": su_rem, "oc_rem": oc_rem, "ma_rem": ma_rem,
        "ou_status": ou,
    })

print(f"Category Table: {len(records)} records")

# ── PO Log sheet ──────────────────────────────────────────────────────────────
po_log_entries = []
ws_po = next((s for s in wb.sheetnames if "po" in s.lower() and "log" in s.lower()), None)
if ws_po:
    wpo = wb[ws_po]
    # Row 2 = headers, row 3+ = data; scan at least 200 rows, skip fully empty rows
    scan_to = max(wpo.max_row, 202)
    for i in range(3, scan_to + 1):
        po_num    = str(wpo.cell(i, 2).value or "").strip()
        cost_code = str(wpo.cell(i, 3).value or "").strip()
        amount_v  = wpo.cell(i, 5).value
        # skip row if both key identifiers are blank
        if not po_num and not cost_code:
            continue
        date_val = wpo.cell(i, 1).value
        category = str(wpo.cell(i, 4).value or "").strip()
        amount   = N(amount_v)
        desc     = str(wpo.cell(i, 6).value or "").strip()
        date_str = str(date_val)[:10] if date_val else ""
        po_log_entries.append({
            "date": date_str,
            "po_number": po_num,
            "cost_code": cost_code,
            "category": category,
            "amount": amount,
            "description": desc,
        })
    print(f"PO Log: {len(po_log_entries)} entries")
else:
    print("PO Log sheet not found")

# ── Write JSON files for inspection ──────────────────────────────────────────
with open("/tmp/raw_data.json", "w") as f:
    json.dump(records, f, indent=2)
with open("/tmp/po_log.json", "w") as f:
    json.dump(po_log_entries, f, indent=2)

print("Written to /tmp/raw_data.json and /tmp/po_log.json")
print(f"Sample ou_status values: {set(r['ou_status'] for r in records)}")
