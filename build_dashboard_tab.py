"""
build_dashboard_tab.py
Builds an interactive Excel Dashboard tab that mirrors the HTML dashboard.
Reads data from the existing HTML file's RAW_DATA (already computed).
Saves as Weekly_Commitments_Review.xlsm
"""

import re, json, os, copy
from collections import defaultdict

import openpyxl
from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side,
                               GradientFill)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.layout import Layout
from openpyxl.chart.marker import Marker
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.worksheet.datavalidation import DataValidation

HERE       = os.path.dirname(os.path.abspath(__file__))
HTML_PATH  = os.path.join(HERE, "budget_dashboard.html")
EXCEL_IN   = os.path.join(HERE, "Weekly_Commitments_Review.xlsx")
EXCEL_OUT  = os.path.join(HERE, "Weekly_Commitments_Review_Dashboard.xlsx")

# ── Colors ──────────────────────────────────────────────────────────────────
BG       = "0F1117"
SURFACE  = "1A1D27"
SURFACE2 = "22263A"
BORDER_C = "2E3350"
ACCENT   = "4F8EF7"
PURPLE   = "7C5CFC"
GREEN    = "22C55E"
RED      = "EF4444"
YELLOW   = "F59E0B"
TEAL     = "14B8A6"
WHITE    = "E2E8F0"
SUBTEXT  = "94A3B8"
ORANGE   = "F97316"

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(hex_color=WHITE, bold=False, size=10, name="Segoe UI"):
    return Font(color=hex_color, bold=bold, size=size, name=name)

def align(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def border(hex_color=BORDER_C, style="thin"):
    s = Side(style=style, color=hex_color)
    return Border(left=s, right=s, top=s, bottom=s)

def border_bottom(hex_color=BORDER_C):
    s = Side(style="thin", color=hex_color)
    return Border(bottom=s)

def apply(ws, row, col, value=None, bg=None, fg=WHITE, bold=False,
          size=10, h="center", v="center", num_fmt=None, wrap=False):
    c = ws.cell(row=row, column=col)
    if value is not None:
        c.value = value
    if bg:
        c.fill = fill(bg)
    c.font  = font(fg, bold=bold, size=size)
    c.alignment = align(h, v, wrap)
    if num_fmt:
        c.number_format = num_fmt
    return c

def merge_apply(ws, r1, c1, r2, c2, value=None, bg=None, fg=WHITE,
                bold=False, size=10, h="center", v="center", num_fmt=None, wrap=False):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    c = apply(ws, r1, c1, value, bg, fg, bold, size, h, v, num_fmt, wrap)
    # Fill merged area bg
    if bg:
        for row in range(r1, r2+1):
            for col in range(c1, c2+1):
                ws.cell(row=row, column=col).fill = fill(bg)
    return c

# ── Load data ───────────────────────────────────────────────────────────────
with open(HTML_PATH, "r") as f:
    html = f.read()
m = re.search(r"const RAW_DATA = (\[.*?\]);", html, re.DOTALL)
records = json.loads(m.group(1))
cvr_month = re.search(r"CVR Month: ([^<\"]+)", html).group(1).strip()

# Aggregate totals
total_eac = sum(r["su_eac"]+r["oc_eac"]+r["ma_eac"] for r in records)
total_ctd = sum(r["su_ctd"]+r["oc_ctd"]+r["ma_ctd"] for r in records)
total_rem = sum(r["su_rem"]+r["oc_rem"]+r["ma_rem"] for r in records)
total_act = sum(r["su_act"]+r["oc_act"]+r["ma_act"] for r in records)
total_acc = sum(r["su_acc"]+r["oc_acc"]+r["ma_acc"] for r in records)
total_com = sum(r["su_com"]+r["oc_com"]+r["ma_com"] for r in records)
over_count  = sum(1 for r in records if r["ou_status"]=="over")
under_count = sum(1 for r in records if r["ou_status"]=="under")

# Per-category totals
su_eac = sum(r["su_eac"] for r in records)
oc_eac = sum(r["oc_eac"] for r in records)
ma_eac = sum(r["ma_eac"] for r in records)
su_rem = sum(r["su_rem"] for r in records)
oc_rem = sum(r["oc_rem"] for r in records)
ma_rem = sum(r["ma_rem"] for r in records)
su_ctd = sum(r["su_ctd"] for r in records)
oc_ctd = sum(r["oc_ctd"] for r in records)
ma_ctd = sum(r["ma_ctd"] for r in records)

# Per cost-code aggregates (collapse SU/OC/MA)
by_cc = defaultdict(lambda: dict(eac=0,ctd=0,rem=0,act=0,item=""))
for r in records:
    cc = r["cost_code"]
    by_cc[cc]["eac"] += r["su_eac"]+r["oc_eac"]+r["ma_eac"]
    by_cc[cc]["ctd"] += r["su_ctd"]+r["oc_ctd"]+r["ma_ctd"]
    by_cc[cc]["rem"] += r["su_rem"]+r["oc_rem"]+r["ma_rem"]
    by_cc[cc]["act"] += r["su_act"]+r["oc_act"]+r["ma_act"]
    by_cc[cc]["item"] = r["item"]

top12_eac = sorted(by_cc.items(), key=lambda x: x[1]["eac"], reverse=True)[:12]
top10_rem = sorted(by_cc.items(), key=lambda x: x[1]["rem"], reverse=True)[:10]
all_sorted = sorted(by_cc.items(), key=lambda x: x[1]["eac"], reverse=True)

# ── Open workbook ────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(EXCEL_IN)

# Remove existing Dashboard if present, recreate
if "Dashboard" in wb.sheetnames:
    del wb["Dashboard"]
ws = wb.create_sheet("Dashboard", 0)

# ── Sheet-wide defaults ──────────────────────────────────────────────────────
ws.sheet_view.showGridLines = False
ws.sheet_properties.tabColor = ACCENT

# Column widths
col_widths = {
    1: 1.5,   # A – thin left margin
    2: 14,    # B
    3: 11,
    4: 11,
    5: 11,
    6: 11,
    7: 11,
    8: 11,
    9: 11,
    10: 11,
    11: 11,
    12: 11,
    13: 11,
    14: 11,
    15: 11,
    16: 11,
    17: 11,
    18: 11,
    19: 11,
    20: 11,
    21: 1.5,  # Z – thin right margin
}
for col, w in col_widths.items():
    ws.column_dimensions[get_column_letter(col)].width = w

# Fill entire used area with BG
for row in range(1, 80):
    for col in range(1, 22):
        ws.cell(row=row, column=col).fill = fill(BG)

# ── ROW 1-3: Header ──────────────────────────────────────────────────────────
ws.row_dimensions[1].height = 6
ws.row_dimensions[2].height = 32
ws.row_dimensions[3].height = 6

# Accent left bar col A rows 1-3
for row in (1,2,3):
    ws.cell(row=row, column=1).fill = fill(ACCENT)

# Header background cols B-T
for col in range(2, 21):
    for row in (1,2,3):
        ws.cell(row=row, column=col).fill = fill(SURFACE)

# Title
merge_apply(ws, 2, 2, 2, 14,
    value="Budget & Commitments  Dashboard",
    bg=SURFACE, fg=WHITE, bold=True, size=18, h="left")

# CVR badge
merge_apply(ws, 2, 15, 2, 20,
    value=f"CVR Month:  {cvr_month}",
    bg=ACCENT, fg=WHITE, bold=True, size=11, h="right")

# ── ROW 4: spacer ─────────────────────────────────────────────────────────────
ws.row_dimensions[4].height = 8

# ── ROWS 5-13: KPI cards ──────────────────────────────────────────────────────
kpis = [
    ("Total EAC",        total_eac,    ACCENT,  "$#,##0"),
    ("Cost to Date",     total_ctd,    TEAL,    "$#,##0"),
    ("Remaining",        total_rem,    GREEN if total_rem>=0 else RED, "$#,##0"),
    ("Actuals",          total_act,    YELLOW,  "$#,##0"),
    ("Accrued",          total_acc,    PURPLE,  "$#,##0"),
    ("Committed",        total_com,    ORANGE,  "$#,##0"),
    ("Over Budget",      over_count,   RED,     "0"),
    ("Under Budget",     under_count,  GREEN,   "0"),
]

ws.row_dimensions[5].height = 4    # top accent bar
ws.row_dimensions[6].height = 18   # label
ws.row_dimensions[7].height = 4    # padding
ws.row_dimensions[8].height = 24   # value
ws.row_dimensions[9].height = 4    # padding
ws.row_dimensions[10].height = 18  # sub-label row
ws.row_dimensions[11].height = 6   # bottom border
ws.row_dimensions[12].height = 6   # spacer

card_cols = [2, 4, 6, 8, 10, 12, 14, 16]  # start col for each card (2 cols wide)

for i, (label, val, color, nfmt) in enumerate(kpis):
    c1 = card_cols[i]
    c2 = c1 + 1

    # Card bg
    for row in range(5, 12):
        for col in range(c1, c2+1):
            ws.cell(row=row, column=col).fill = fill(SURFACE)

    # Top color bar
    for col in range(c1, c2+1):
        ws.cell(row=5, column=col).fill = fill(color)

    # Label
    merge_apply(ws, 6, c1, 6, c2, value=label,
                bg=SURFACE, fg=SUBTEXT, bold=False, size=9)

    # Value
    c = merge_apply(ws, 8, c1, 8, c2, value=val,
                    bg=SURFACE, fg=color, bold=True, size=15)
    c.number_format = nfmt

    # Bottom border line
    for col in range(c1, c2+1):
        ws.cell(row=11, column=col).fill = fill(SURFACE)

# ── CHART DATA SHEET (hidden) ─────────────────────────────────────────────────
# We write chart data to a hidden helper sheet and reference it for charts
CHART_SHEET = "_ChartData"
if CHART_SHEET in wb.sheetnames:
    del wb[CHART_SHEET]
wc = wb.create_sheet(CHART_SHEET)
wc.sheet_state = "hidden"

# ── Chart 1: EAC vs CTD by top cost codes ────────────────────────────────────
# Write to wc cols A-C
wc["A1"] = "Cost Code"; wc["B1"] = "EAC"; wc["C1"] = "CTD"
for i, (cc, d) in enumerate(top12_eac, 2):
    wc.cell(i, 1).value = cc
    wc.cell(i, 2).value = round(d["eac"])
    wc.cell(i, 3).value = round(d["ctd"])
ch1_rows = len(top12_eac) + 1

chart1 = BarChart()
chart1.type = "col"
chart1.grouping = "clustered"
chart1.title = "CVR EAC vs Cost to Date"
chart1.style = 10
chart1.shape = 4
chart1.height = 12
chart1.width  = 22

labels1 = Reference(wc, min_col=1, min_row=2, max_row=ch1_rows)
eac_ref = Reference(wc, min_col=2, min_row=1, max_row=ch1_rows)
ctd_ref = Reference(wc, min_col=3, min_row=1, max_row=ch1_rows)
chart1.add_data(eac_ref, titles_from_data=True)
chart1.add_data(ctd_ref, titles_from_data=True)
chart1.set_categories(labels1)
chart1.series[0].graphicalProperties.solidFill = ACCENT
chart1.series[0].graphicalProperties.line.solidFill = ACCENT
chart1.series[1].graphicalProperties.solidFill = TEAL
chart1.series[1].graphicalProperties.line.solidFill = TEAL

# Data labels
for s in chart1.series:
    s.dLbls = DataLabelList()
    s.dLbls.showVal = True
    s.dLbls.showLegendKey = False
    s.dLbls.showCatName = False
    s.dLbls.showSerName = False
    s.dLbls.numFmt = "$#,##0"

chart1.plot_area.layout = Layout()
ws.add_chart(chart1, "B13")

# ── Chart 2: Remaining by Category ───────────────────────────────────────────
# Write to wc cols E-G
wc["E1"] = "Category"; wc["F1"] = "EAC"; wc["G1"] = "Remaining"
cats = [("Subcontract", su_eac, su_rem), ("Other Costs", oc_eac, oc_rem), ("Materials", ma_eac, ma_rem)]
for i, (cat, eac, rem) in enumerate(cats, 2):
    wc.cell(i, 5).value = cat
    wc.cell(i, 6).value = round(eac)
    wc.cell(i, 7).value = round(rem)

chart2 = BarChart()
chart2.type = "col"
chart2.grouping = "clustered"
chart2.title = "Remaining Budget by Category"
chart2.style = 10
chart2.height = 12
chart2.width  = 22

labels2 = Reference(wc, min_col=5, min_row=2, max_row=4)
eac2    = Reference(wc, min_col=6, min_row=1, max_row=4)
rem2    = Reference(wc, min_col=7, min_row=1, max_row=4)
chart2.add_data(eac2, titles_from_data=True)
chart2.add_data(rem2, titles_from_data=True)
chart2.set_categories(labels2)
chart2.series[0].graphicalProperties.solidFill = PURPLE
chart2.series[0].graphicalProperties.line.solidFill = PURPLE
chart2.series[1].graphicalProperties.solidFill = GREEN
chart2.series[1].graphicalProperties.line.solidFill = GREEN

for s in chart2.series:
    s.dLbls = DataLabelList()
    s.dLbls.showVal = True
    s.dLbls.showLegendKey = False
    s.dLbls.showCatName = False
    s.dLbls.showSerName = False
    s.dLbls.numFmt = "$#,##0"

ws.add_chart(chart2, "K13")

# ── Chart 3: Top 10 Items by Remaining (horizontal bar) ──────────────────────
wc["I1"] = "Cost Code"; wc["J1"] = "Remaining"
for i, (cc, d) in enumerate(top10_rem, 2):
    wc.cell(i, 9).value  = cc
    wc.cell(i, 10).value = round(d["rem"])
ch3_rows = len(top10_rem) + 1

chart3 = BarChart()
chart3.type = "bar"   # horizontal
chart3.grouping = "clustered"
chart3.title = "Top 10 Items – Remaining Budget"
chart3.style = 10
chart3.height = 12
chart3.width  = 22

labels3 = Reference(wc, min_col=9,  min_row=2, max_row=ch3_rows)
rem3    = Reference(wc, min_col=10, min_row=1, max_row=ch3_rows)
chart3.add_data(rem3, titles_from_data=True)
chart3.set_categories(labels3)
chart3.series[0].graphicalProperties.solidFill = ACCENT
chart3.series[0].graphicalProperties.line.solidFill = ACCENT
chart3.series[0].dLbls = DataLabelList()
chart3.series[0].dLbls.showVal = True
chart3.series[0].dLbls.showLegendKey = False
chart3.series[0].dLbls.numFmt = "$#,##0"

ws.add_chart(chart3, "B31")

# ── Chart 4: Actuals / Accrued / Committed (pie) ─────────────────────────────
wc["L1"] = "Type"; wc["M1"] = "Amount"
wc["L2"] = "Actuals";   wc["M2"] = round(total_act)
wc["L3"] = "Accrued";   wc["M3"] = round(total_acc)
wc["L4"] = "Committed"; wc["M4"] = round(total_com)

chart4 = PieChart()
chart4.title  = "Actuals · Accrued · Committed"
chart4.style  = 10
chart4.height = 12
chart4.width  = 22

labels4 = Reference(wc, min_col=12, min_row=2, max_row=4)
data4   = Reference(wc, min_col=13, min_row=1, max_row=4)
chart4.add_data(data4, titles_from_data=True)
chart4.set_categories(labels4)
chart4.series[0].dLbls = DataLabelList()
chart4.series[0].dLbls.showVal = True
chart4.series[0].dLbls.showPercent = True
chart4.series[0].dLbls.showLegendKey = False
chart4.series[0].dLbls.showCatName = True
# Set slice colors
from openpyxl.chart.data_source import NumDataSource, NumRef
from openpyxl.drawing.fill import PatternFillProperties
for idx, color in enumerate([YELLOW, PURPLE, ORANGE]):
    try:
        pt = chart4.series[0].dPt
        # Color each data point
        from openpyxl.chart.series import DataPoint
        dp = DataPoint(idx=idx)
        dp.spPr = openpyxl.drawing.spreadsheet_drawing.SpPr()
    except:
        pass

ws.add_chart(chart4, "K31")

# ── ROWS 49+: Data table ──────────────────────────────────────────────────────
tbl_start = 50

ws.row_dimensions[48].height = 8
ws.row_dimensions[49].height = 8

# Section header
merge_apply(ws, 49, 2, 49, 20,
    value="  Cost Code Summary",
    bg=SURFACE2, fg=WHITE, bold=True, size=11, h="left")

# Table header row
ws.row_dimensions[tbl_start].height = 20
tbl_headers = [
    (2,  "Cost Code",  14, "left"),
    (4,  "Item",       14, "left"),
    (8,  "EAC",        11, "right"),
    (10, "CTD",        11, "right"),
    (12, "Remaining",  11, "right"),
    (14, "Actuals",    11, "right"),
    (16, "O / U",      11, "center"),
]
for col, label, width, h_align in tbl_headers:
    c = ws.cell(row=tbl_start, column=col)
    c.value = label
    c.fill  = fill(SURFACE2)
    c.font  = font(WHITE, bold=True, size=9)
    c.alignment = align(h_align)
    ws.column_dimensions[get_column_letter(col)].width = width

# Data rows
for i, (cc, d) in enumerate(all_sorted):
    row = tbl_start + 1 + i
    ws.row_dimensions[row].height = 16
    row_bg = SURFACE if i % 2 == 0 else BG
    ou_color = RED if d["rem"] < 0 else GREEN
    ou_label = "▲ Over" if d["rem"] < 0 else "✓ Under"

    cells = [
        (2,  cc,           WHITE,    "left",  None),
        (4,  d["item"],    SUBTEXT,  "left",  None),
        (8,  d["eac"],     WHITE,    "right", "$#,##0"),
        (10, d["ctd"],     TEAL,     "right", "$#,##0"),
        (12, d["rem"],     GREEN if d["rem"]>=0 else RED, "right", "$#,##0"),
        (14, d["act"],     YELLOW,   "right", "$#,##0"),
        (16, ou_label,     ou_color, "center", None),
    ]
    # Fill background
    for col in range(2, 18):
        ws.cell(row=row, column=col).fill = fill(row_bg)

    for col, val, fg_c, h_align, nfmt in cells:
        c = ws.cell(row=row, column=col)
        c.value     = val
        c.fill      = fill(row_bg)
        c.font      = font(fg_c, size=9)
        c.alignment = align(h_align)
        if nfmt:
            c.number_format = nfmt

    # Span item across 3 cols
    try:
        ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=7)
    except:
        pass

# ── Freeze panes below header ─────────────────────────────────────────────────
ws.freeze_panes = "B13"

# ── Save ─────────────────────────────────────────────────────────────────────
print("Saving...")
wb.save(EXCEL_OUT)
print(f"Saved: {EXCEL_OUT}")
