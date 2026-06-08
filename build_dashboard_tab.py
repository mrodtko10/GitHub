"""
build_dashboard_tab.py
Builds a styled Excel Dashboard tab mirroring the HTML dashboard.
"""

import re, json, os
from collections import defaultdict

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.axis import ChartLines

HERE       = os.path.dirname(os.path.abspath(__file__))
HTML_PATH  = os.path.join(HERE, "budget_dashboard.html")
EXCEL_IN   = os.path.join(HERE, "Weekly_Commitments_Review.xlsx")
EXCEL_OUT  = os.path.join(HERE, "Weekly_Commitments_Review_Dashboard.xlsx")

# ── Colors ───────────────────────────────────────────────────────────────────
BG      = "0F1117"
SURFACE = "1A1D27"
SRF2    = "22263A"
BORD    = "2E3350"
ACCENT  = "4F8EF7"
PURPLE  = "7C5CFC"
GREEN   = "22C55E"
RED     = "EF4444"
YELLOW  = "F59E0B"
TEAL    = "14B8A6"
WHITE   = "E2E8F0"
SUB     = "94A3B8"
ORANGE  = "F97316"

def fill(h): return PatternFill("solid", fgColor=h)
def font(h=WHITE, bold=False, sz=10, name="Segoe UI"):
    return Font(color=h, bold=bold, size=sz, name=name)
def aln(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def put(ws, r, c, val=None, bg=None, fg=WHITE, bold=False, sz=10,
        h="center", v="center", nfmt=None, wrap=False):
    cell = ws.cell(row=r, column=c)
    if val is not None: cell.value = val
    if bg:              cell.fill  = fill(bg)
    cell.font      = font(fg, bold, sz)
    cell.alignment = aln(h, v, wrap)
    if nfmt: cell.number_format = nfmt
    return cell

def mput(ws, r1, c1, r2, c2, val=None, bg=None, fg=WHITE,
         bold=False, sz=10, h="center", v="center", nfmt=None, wrap=False):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    cell = put(ws, r1, c1, val, bg, fg, bold, sz, h, v, nfmt, wrap)
    if bg:
        for row in range(r1, r2+1):
            for col in range(c1, c2+1):
                ws.cell(row=row, column=col).fill = fill(bg)
    return cell

# ── Load data ─────────────────────────────────────────────────────────────────
with open(HTML_PATH) as f: html = f.read()
records   = json.loads(re.search(r"const RAW_DATA = (\[.*?\]);", html, re.DOTALL).group(1))
cvr_month = re.search(r"CVR Month: ([^<\"]+)", html).group(1).strip()

total_eac = sum(r["su_eac"]+r["oc_eac"]+r["ma_eac"] for r in records)
total_ctd = sum(r["su_ctd"]+r["oc_ctd"]+r["ma_ctd"] for r in records)
total_rem = sum(r["su_rem"]+r["oc_rem"]+r["ma_rem"] for r in records)
total_act = sum(r["su_act"]+r["oc_act"]+r["ma_act"] for r in records)
total_acc = sum(r["su_acc"]+r["oc_acc"]+r["ma_acc"] for r in records)
total_com = sum(r["su_com"]+r["oc_com"]+r["ma_com"] for r in records)
over_cnt  = sum(1 for r in records if r["ou_status"]=="over")
under_cnt = sum(1 for r in records if r["ou_status"]=="under")

su_eac = sum(r["su_eac"] for r in records)
oc_eac = sum(r["oc_eac"] for r in records)
ma_eac = sum(r["ma_eac"] for r in records)
su_rem = sum(r["su_rem"] for r in records)
oc_rem = sum(r["oc_rem"] for r in records)
ma_rem = sum(r["ma_rem"] for r in records)
su_ctd = sum(r["su_ctd"] for r in records)
oc_ctd = sum(r["oc_ctd"] for r in records)
ma_ctd = sum(r["ma_ctd"] for r in records)

by_cc = defaultdict(lambda: dict(eac=0,ctd=0,rem=0,act=0,item=""))
for r in records:
    cc = r["cost_code"]
    by_cc[cc]["eac"] += r["su_eac"]+r["oc_eac"]+r["ma_eac"]
    by_cc[cc]["ctd"] += r["su_ctd"]+r["oc_ctd"]+r["ma_ctd"]
    by_cc[cc]["rem"] += r["su_rem"]+r["oc_rem"]+r["ma_rem"]
    by_cc[cc]["act"] += r["su_act"]+r["oc_act"]+r["ma_act"]
    by_cc[cc]["item"] = r["item"]

top12 = sorted(by_cc.items(), key=lambda x: x[1]["eac"], reverse=True)[:12]
top10r = sorted(by_cc.items(), key=lambda x: x[1]["rem"], reverse=True)[:10]
all_s  = sorted(by_cc.items(), key=lambda x: x[1]["eac"], reverse=True)

# ── Workbook ───────────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(EXCEL_IN)
if "Dashboard" in wb.sheetnames: del wb["Dashboard"]
ws = wb.create_sheet("Dashboard", 0)
ws.sheet_view.showGridLines = False
ws.sheet_properties.tabColor = ACCENT

# ── Column widths ─────────────────────────────────────────────────────────────
# Layout: A(margin) | B-J (left half, 9 cols) | K(gap) | L-T (right half, 9 cols) | U(margin)
# Each data col = 11 units ≈ 1.85cm  →  9 cols ≈ 16.6cm  →  use chart width 15cm
CW = {1:1.2, 11:1.2, 21:1.2}  # A, K(gap), U(margin)
for c in range(2,11):  CW[c] = 11   # B-J left half
for c in range(12,21): CW[c] = 11   # L-T right half
for c, w in CW.items():
    ws.column_dimensions[get_column_letter(c)].width = w

# Fill entire canvas dark
for r in range(1, 100):
    for c in range(1, 22):
        ws.cell(r, c).fill = fill(BG)

# ── HEADER (rows 1-3) ─────────────────────────────────────────────────────────
ws.row_dimensions[1].height = 5
ws.row_dimensions[2].height = 34
ws.row_dimensions[3].height = 5
for r in (1,2,3):
    for c in range(1,22):
        ws.cell(r,c).fill = fill(SURFACE)
    ws.cell(r,1).fill = fill(ACCENT)           # accent left stripe

mput(ws,2,2,2,13, "Budget & Commitments  Dashboard",
     SURFACE, WHITE, bold=True, sz=18, h="left")
mput(ws,2,14,2,20, f"CVR Month:  {cvr_month}",
     ACCENT, WHITE, bold=True, sz=11, h="right")

# spacer
ws.row_dimensions[4].height = 10

# ── KPI CARDS (rows 5-12) ─────────────────────────────────────────────────────
# 8 cards × 2 cols each = 16 cols, starting at col 2
KPIS = [
    ("Total EAC",     total_eac, ACCENT,  "$#,##0"),
    ("Cost to Date",  total_ctd, TEAL,    "$#,##0"),
    ("Remaining",     total_rem, GREEN if total_rem>=0 else RED, "$#,##0"),
    ("Actuals",       total_act, YELLOW,  "$#,##0"),
    ("Accrued",       total_acc, PURPLE,  "$#,##0"),
    ("Committed",     total_com, ORANGE,  "$#,##0"),
    ("Over Budget",   over_cnt,  RED,     "0"),
    ("Under Budget",  under_cnt, GREEN,   "0"),
]
ws.row_dimensions[5].height  = 5   # accent top bar
ws.row_dimensions[6].height  = 18  # label
ws.row_dimensions[7].height  = 4
ws.row_dimensions[8].height  = 26  # value
ws.row_dimensions[9].height  = 4
ws.row_dimensions[10].height = 5   # bottom bar
ws.row_dimensions[11].height = 10  # spacer below cards

for i,(label,val,clr,nfmt) in enumerate(KPIS):
    c1 = 2 + i*2    # start col (B=2,D=4,F=6,H=8,J=10,L=12,N=14,P=16)
    c2 = c1 + 1
    for r in range(5,11):
        for c in range(c1,c2+1):
            ws.cell(r,c).fill = fill(SURFACE)
    # accent bar
    for c in range(c1,c2+1): ws.cell(5,c).fill = fill(clr)
    # label
    mput(ws,6,c1,6,c2, label, SURFACE, SUB, sz=9)
    # value
    cell = mput(ws,8,c1,8,c2, val, SURFACE, clr, bold=True, sz=14, nfmt=nfmt)

# ── CHART DATA (hidden sheet) ─────────────────────────────────────────────────
if "_CD" in wb.sheetnames: del wb["_CD"]
wc = wb.create_sheet("_CD")
wc.sheet_state = "hidden"

# Chart 1 data: EAC vs CTD top 12  (cols A-C)
wc["A1"]="Cost Code"; wc["B1"]="EAC"; wc["C1"]="CTD"
for i,(cc,d) in enumerate(top12,2):
    wc.cell(i,1).value=cc; wc.cell(i,2).value=round(d["eac"]); wc.cell(i,3).value=round(d["ctd"])
ch1_end = len(top12)+1

# Chart 2 data: Category EAC vs Remaining (cols E-G)
wc["E1"]="Category"; wc["F1"]="EAC"; wc["G1"]="Remaining"
for i,(cat,e,r2) in enumerate([("Subcontract",su_eac,su_rem),
                                 ("Other Costs",oc_eac,oc_rem),
                                 ("Materials",  ma_eac,ma_rem)],2):
    wc.cell(i,5).value=cat; wc.cell(i,6).value=round(e); wc.cell(i,7).value=round(r2)

# Chart 3 data: Top 10 remaining (cols I-J)
wc["I1"]="Cost Code"; wc["J1"]="Remaining"
for i,(cc,d) in enumerate(top10r,2):
    wc.cell(i,9).value=cc; wc.cell(i,10).value=round(d["rem"])
ch3_end = len(top10r)+1

# Chart 4 data: Actuals/Accrued/Committed (cols L-M)
wc["L1"]="Type"; wc["M1"]="Amount"
wc["L2"]="Actuals";   wc["M2"]=round(total_act)
wc["L3"]="Accrued";   wc["M3"]=round(total_acc)
wc["L4"]="Committed"; wc["M4"]=round(total_com)

# ── Helper: build & style a bar chart ─────────────────────────────────────────
def make_bar(title, labels_ref, data_refs, colors, horiz=False,
             x_title=None, y_title=None):
    ch = BarChart()
    ch.type     = "bar" if horiz else "col"
    ch.grouping = "clustered"
    ch.title    = title
    ch.style    = 10
    ch.width    = 15
    ch.height   = 11.5
    ch.add_data(data_refs[0], titles_from_data=True)
    if len(data_refs) > 1:
        ch.add_data(data_refs[1], titles_from_data=True)
    ch.set_categories(labels_ref)
    for idx,clr in enumerate(colors):
        s = ch.series[idx]
        s.graphicalProperties.solidFill = clr
        s.graphicalProperties.line.solidFill = clr
        s.dLbls = DataLabelList()
        s.dLbls.showVal = True
        s.dLbls.showLegendKey = False
        s.dLbls.showCatName   = False
        s.dLbls.showSerName   = False
        s.dLbls.numFmt = "$#,##0"
    # Axis titles
    if horiz:
        # For horizontal bar: catAx = vertical (items), valAx = horizontal ($)
        if x_title:
            ch.y_axis.title = x_title   # category axis (items)
        if y_title:
            ch.x_axis.title = y_title   # value axis ($)
        ch.x_axis.numFmt = "$#,##0"
    else:
        if x_title:
            ch.x_axis.title = x_title
        if y_title:
            ch.y_axis.title = y_title
        ch.y_axis.numFmt = "$#,##0"
    # Axis number format on value axis
    ch.y_axis.delete = False
    ch.x_axis.delete = False
    return ch

def make_pie(title, labels_ref, data_ref):
    ch = PieChart()
    ch.title  = title
    ch.style  = 10
    ch.width  = 15
    ch.height = 11.5
    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(labels_ref)
    s = ch.series[0]
    s.dLbls = DataLabelList()
    s.dLbls.showVal     = True
    s.dLbls.showPercent = True
    s.dLbls.showCatName = True
    s.dLbls.showLegendKey = False
    return ch

# ── CHART ROW 1 (anchored at row 12) ─────────────────────────────────────────
# Set rows 12-31 to fixed height so charts have a clean 20-row zone
for r in range(12, 32):
    ws.row_dimensions[r].height = 18

chart1 = make_bar(
    "CVR EAC vs Cost to Date",
    Reference(wc, min_col=1, min_row=2, max_row=ch1_end),
    [Reference(wc, min_col=2, min_row=1, max_row=ch1_end),
     Reference(wc, min_col=3, min_row=1, max_row=ch1_end)],
    [ACCENT, TEAL],
    x_title="Cost Code", y_title="Amount ($)"
)
ws.add_chart(chart1, "B12")   # left half

chart2 = make_bar(
    "Remaining Budget by Category",
    Reference(wc, min_col=5, min_row=2, max_row=4),
    [Reference(wc, min_col=6, min_row=1, max_row=4),
     Reference(wc, min_col=7, min_row=1, max_row=4)],
    [PURPLE, GREEN],
    x_title="Category", y_title="Amount ($)"
)
ws.add_chart(chart2, "L12")   # right half (col 12 = L)

# spacer row between chart rows
ws.row_dimensions[32].height = 12

# ── CHART ROW 2 (anchored at row 33) ─────────────────────────────────────────
for r in range(33, 53):
    ws.row_dimensions[r].height = 18

chart3 = make_bar(
    "Top 10 Items – Remaining Budget",
    Reference(wc, min_col=9,  min_row=2, max_row=ch3_end),
    [Reference(wc, min_col=10, min_row=1, max_row=ch3_end)],
    [ACCENT],
    horiz=True,
    x_title="Cost Code", y_title="Remaining ($)"
)
ws.add_chart(chart3, "B33")

chart4 = make_pie(
    "Actuals · Accrued · Committed",
    Reference(wc, min_col=12, min_row=2, max_row=4),
    Reference(wc, min_col=13, min_row=1, max_row=4)
)
ws.add_chart(chart4, "L33")

# spacer
ws.row_dimensions[53].height = 14

# ── DATA TABLE (row 54+) ──────────────────────────────────────────────────────
ws.row_dimensions[54].height = 20
mput(ws,54,2,54,20, "  Cost Code Summary",
     SRF2, WHITE, bold=True, sz=11, h="left")

HDR_ROW = 55
ws.row_dimensions[HDR_ROW].height = 20
HDR = [(2,"Cost Code",8,"center"),(3,"Item",16,"left"),
       (8,"EAC",9,"right"),(10,"CTD",9,"right"),
       (12,"Remaining",9,"right"),(14,"Actuals",9,"right"),
       (16,"O/U",8,"center")]
for col,label,_,ha in HDR:
    c = ws.cell(HDR_ROW, col)
    c.value = label; c.fill=fill(SRF2)
    c.font=font(WHITE,bold=True,sz=9); c.alignment=aln(ha)

for i,(cc,d) in enumerate(all_s):
    row = HDR_ROW + 1 + i
    ws.row_dimensions[row].height = 16
    rb = SURFACE if i%2==0 else BG
    ou = d["rem"] < 0
    ou_lbl = "▲ Over" if ou else "✓ Under"
    ou_clr = RED if ou else GREEN

    for c in range(2,18): ws.cell(row,c).fill = fill(rb)

    put(ws,row,2, cc,      rb, WHITE,  sz=9, h="center")
    put(ws,row,3, d["item"],rb,SUB,    sz=9, h="left", wrap=False)
    put(ws,row,8, d["eac"],rb,WHITE,   sz=9, h="right",  nfmt="$#,##0")
    put(ws,row,10,d["ctd"],rb,TEAL,    sz=9, h="right",  nfmt="$#,##0")
    put(ws,row,12,d["rem"],rb,ou_clr,  sz=9, h="right",  nfmt="$#,##0")
    put(ws,row,14,d["act"],rb,YELLOW,  sz=9, h="right",  nfmt="$#,##0")
    put(ws,row,16,ou_lbl,  rb,ou_clr,  sz=9, bold=True, h="center")

    # Span item col across 4 cols (3-6)
    try: ws.merge_cells(start_row=row,start_column=3,end_row=row,end_column=7)
    except: pass

# ── Freeze below KPI section ──────────────────────────────────────────────────
ws.freeze_panes = "B12"

# ── Save ──────────────────────────────────────────────────────────────────────
print("Saving...")
wb.save(EXCEL_OUT)
print(f"Saved: {EXCEL_OUT}")
