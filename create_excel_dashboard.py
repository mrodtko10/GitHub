"""
create_excel_dashboard.py
-------------------------
Builds a fully interactive Excel dashboard inside Weekly_Commitments_Review.xlsm
with dark theme, KPI cards, slicers, and charts — mirroring the HTML dashboard.

Requirements: pywin32  (pip install pywin32)
Run once from the same folder as the Excel file.
"""

import os, sys, time

try:
    import win32com.client as win32
    import pywintypes
except ImportError:
    input("ERROR: pywin32 not installed.\nRun:  pip install pywin32\n\nPress Enter.")
    sys.exit(1)

HERE      = os.path.dirname(os.path.abspath(__file__))
XLSM      = os.path.join(HERE, "Weekly_Commitments_Review.xlsm")
XLSX      = os.path.join(HERE, "Weekly_Commitments_Review.xlsx")
SRC       = XLSM if os.path.exists(XLSM) else XLSX
SAVE_AS   = XLSM  # always save as .xlsm

# ---------------------------------------------------------------------------
# Color helpers  (Excel interior color = B*65536 + G*256 + R)
# ---------------------------------------------------------------------------
def ec(r, g, b):   return b * 65536 + g * 256 + r

BG       = ec(15,  17,  23)   # #0f1117
SURFACE  = ec(26,  29,  39)   # #1a1d27
SURFACE2 = ec(34,  38,  58)   # #22263a
BORDER   = ec(46,  51,  80)   # #2e3350
ACCENT   = ec(79,  142, 247)  # #4f8ef7  blue
PURPLE   = ec(124, 92,  252)  # #7c5cfc
GREEN    = ec(34,  197, 94)   # #22c55e
RED      = ec(239, 68,  68)   # #ef4444
YELLOW   = ec(245, 158, 11)   # #f59e0b
TEAL     = ec(20,  184, 166)  # #14b8a6
WHITE    = ec(226, 232, 240)  # #e2e8f0
SUBTEXT  = ec(148, 163, 184)  # #94a3b8
DARK_TXT = ec(100, 110, 140)

# ---------------------------------------------------------------------------
# Column mapping in Category Table (1-based for VBA/COM, but we read 0-based)
# ---------------------------------------------------------------------------
COL_CC  = 0;  COL_ITEM = 1
SU_EAC, OC_EAC, MA_EAC   = 14,15,16
SU_ACC, OC_ACC, MA_ACC   = 20,21,22
SU_ACT, OC_ACT, MA_ACT   = 23,24,25
SU_COM, OC_COM, MA_COM   = 26,27,28
SU_CTD, OC_CTD, MA_CTD   = 32,33,34
SU_REM, OC_REM, MA_REM   = 38,39,40
SU_OU,  OC_OU,  MA_OU    = 41,42,43

def safe(v):
    try:    return float(v) if v not in (None,"","-") else 0.0
    except: return 0.0


# ---------------------------------------------------------------------------
# Read source data from Category Table
# ---------------------------------------------------------------------------
def read_data(wb):
    ws = None
    for sh in wb.Sheets:
        if "category" in sh.Name.lower() and "table" in sh.Name.lower():
            ws = sh; break
    if not ws:
        for sh in wb.Sheets:
            if "dashboard" not in sh.Name.lower() and "po" not in sh.Name.lower():
                ws = sh; break
    if not ws:
        ws = wb.Sheets(1)

    print(f"  Reading sheet: {ws.Name}")
    last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # xlUp = -4162

    # CVR month from row 1
    cvr_month = "May, 2026"
    for c in range(1, 8):
        v = str(ws.Cells(1, c).Value or "").strip()
        v = v.replace("CVR Month:", "").replace("CVR Month :", "").strip()
        if v and v != "0":
            cvr_month = v; break

    records = []
    for i in range(4, last_row + 1):  # skip header rows 1-3
        cc   = str(ws.Cells(i, COL_CC+1).Value or "").strip()
        item = str(ws.Cells(i, COL_ITEM+1).Value or "").strip()
        if not cc or cc in ("Cost Code","Total","0"):
            continue

        def g(c): return safe(ws.Cells(i, c+1).Value)

        su_ou_v = g(SU_OU); oc_ou_v = g(OC_OU); ma_ou_v = g(MA_OU)
        ou_status = "Over" if (su_ou_v+oc_ou_v+ma_ou_v) < 0 else "Under"

        # Store one row per category (long format for slicers)
        for cat, e,acc,act,com,ctd,rem,ou_v in [
            ("Subcontract", g(SU_EAC),g(SU_ACC),g(SU_ACT),g(SU_COM),g(SU_CTD),g(SU_REM),su_ou_v),
            ("Other Costs", g(OC_EAC),g(OC_ACC),g(OC_ACT),g(OC_COM),g(OC_CTD),g(OC_REM),oc_ou_v),
            ("Materials",   g(MA_EAC),g(MA_ACC),g(MA_ACT),g(MA_COM),g(MA_CTD),g(MA_REM),ma_ou_v),
        ]:
            records.append({
                "cc": cc, "item": item, "category": cat,
                "eac": e, "acc": acc, "act": act, "com": com,
                "ctd": ctd, "rem": rem,
                "ou_status": ou_status,
            })
    return records, cvr_month


# ---------------------------------------------------------------------------
# Sheet helpers
# ---------------------------------------------------------------------------
def delete_sheet(wb, name):
    xl = wb.Application
    xl.DisplayAlerts = False
    for sh in list(wb.Sheets):
        if sh.Name == name:
            sh.Delete()
    xl.DisplayAlerts = True

def add_sheet(wb, name, after_sheet=None):
    delete_sheet(wb, name)
    if after_sheet:
        return wb.Sheets.Add(After=after_sheet)
    return wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))

def fmt_cell(c, bg=None, fg=WHITE, bold=False, size=11, h_align=None, v_align=None, num_fmt=None, border_color=None):
    if bg is not None:    c.Interior.Color = bg
    c.Font.Color  = fg
    c.Font.Bold   = bold
    c.Font.Size   = size
    if h_align:           c.HorizontalAlignment = h_align
    if v_align:           c.VerticalAlignment   = v_align
    if num_fmt:           c.NumberFormat = num_fmt
    if border_color is not None:
        for side in (7,8,9,10):  # xlEdgeLeft/Right/Top/Bottom
            c.Borders(side).Color = border_color
            c.Borders(side).Weight = 2

def fill_range(ws, r1, c1, r2, c2, bg=None, fg=WHITE):
    rng = ws.Range(ws.Cells(r1,c1), ws.Cells(r2,c2))
    if bg is not None: rng.Interior.Color = bg
    rng.Font.Color = fg

XL_LEFT   = -4131
XL_CENTER = -4108
XL_RIGHT  = -4152
XL_MIDDLE = -4108


# ---------------------------------------------------------------------------
# Write the _Data sheet
# ---------------------------------------------------------------------------
def build_data_sheet(wb, records):
    ws = add_sheet(wb, "_Data")
    ws.Name = "_Data"
    ws.Tab.Color = SURFACE2

    headers = ["Cost Code","Item","Category","EAC","Accrued","Actuals",
               "Committed","CTD","Remaining","O/U Status"]
    for c, h in enumerate(headers, 1):
        cell = ws.Cells(1, c)
        cell.Value = h
        fmt_cell(cell, bg=SURFACE2, fg=WHITE, bold=True, size=10)

    for r, rec in enumerate(records, 2):
        vals = [rec["cc"], rec["item"], rec["category"],
                rec["eac"], rec["acc"], rec["act"],
                rec["com"], rec["ctd"], rec["rem"], rec["ou_status"]]
        for c, v in enumerate(vals, 1):
            ws.Cells(r, c).Value = v

    # Format as Excel Table
    last_row = len(records) + 1
    tbl_rng  = ws.Range(ws.Cells(1,1), ws.Cells(last_row, 10))
    tbl = ws.ListObjects.Add(1, tbl_rng, None, 1)  # xlSrcRange, HasHeaders=True
    tbl.Name = "BudgetData"
    tbl.TableStyle = "TableStyleMedium2"

    ws.Columns("A:J").AutoFit()
    ws.Visible = 2  # xlSheetVeryHidden
    return ws, last_row


# ---------------------------------------------------------------------------
# Build the Dashboard sheet
# ---------------------------------------------------------------------------
CHART_TYPE = {
    "bar_cluster": 57,   # xlColumnClustered
    "bar_stack":   58,   # xlColumnStacked
    "bar_horiz":   -4100, # xlBarClustered
    "line":        65,   # xlLine
    "pie":         5,    # xlPie
    "donut":       -4120, # xlDoughnut
}

def add_chart(ws, ch_type, left, top, w, h, title, src_sheet, data_range_addr,
              series_colors=None, has_legend=True, x_labels_range=None):
    """Add a chart to ws with given dimensions and source data."""
    co = ws.Shapes.AddChart2(-1, ch_type, left, top, w, h)
    ch = co.Chart
    ch.HasTitle = True
    ch.ChartTitle.Text = title
    ch.ChartTitle.Font.Color = WHITE
    ch.ChartTitle.Font.Size  = 11
    ch.ChartTitle.Font.Bold  = True

    # Source data
    data_rng = src_sheet.Range(data_range_addr)
    ch.SetSourceData(data_rng)

    # Dark background
    ch.ChartArea.Format.Fill.ForeColor.RGB  = SURFACE
    ch.ChartArea.Format.Fill.Visible = True
    ch.ChartArea.Format.Fill.Solid()
    ch.ChartArea.Border.LineStyle = 0  # none
    ch.PlotArea.Format.Fill.ForeColor.RGB = SURFACE2
    ch.PlotArea.Format.Fill.Solid()
    ch.PlotArea.Border.LineStyle = 0

    # Axes
    try:
        ax_cat = ch.Axes(1)  # xlCategory=1
        ax_val = ch.Axes(2)  # xlValue=2
        for ax in (ax_cat, ax_val):
            ax.TickLabels.Font.Color = SUBTEXT
            ax.TickLabels.Font.Size  = 8
            ax.MajorGridlines.Format.Line.ForeColor.RGB = BORDER
            ax.Format.Line.ForeColor.RGB = BORDER
    except: pass

    # Legend
    if has_legend:
        ch.HasLegend = True
        ch.Legend.Font.Color = SUBTEXT
        ch.Legend.Font.Size  = 8
        ch.Legend.Position   = 3  # xlLegendPositionBottom = -4107, right = 3
    else:
        ch.HasLegend = False

    # Series colors
    if series_colors:
        for idx, color in enumerate(series_colors):
            try:
                s = ch.SeriesCollection(idx + 1)
                s.Format.Fill.ForeColor.RGB = color
                s.Format.Line.ForeColor.RGB = color
            except: pass

    # Data labels
    try:
        for idx in range(1, ch.SeriesCollection().Count + 1):
            s = ch.SeriesCollection(idx)
            s.HasDataLabels = True
            dl = s.DataLabels()
            dl.Font.Color = WHITE
            dl.Font.Size  = 7
            dl.NumberFormat = "$#,##0"
            dl.ShowValue = True
    except: pass

    return co, ch


def build_dashboard(wb, ws_data, records, cvr_month):
    # Try to find existing Dashboard sheet, else create one
    dash_ws = None
    for sh in wb.Sheets:
        if sh.Name.lower() == "dashboard":
            dash_ws = sh
            break

    if dash_ws is None:
        dash_ws = add_sheet(wb, "Dashboard")
        dash_ws.Name = "Dashboard"
    else:
        # Clear its content/charts but keep the sheet
        dash_ws.Cells.Clear()
        for sh in list(dash_ws.Shapes):
            sh.Delete()

    ws = dash_ws
    ws.Tab.Color = ACCENT

    # Grid setup
    xl = wb.Application
    xl.ScreenUpdating = False

    # Set row heights and column widths
    ws.Cells.RowHeight    = 18
    ws.Cells.ColumnWidth  = 10
    ws.Cells.Interior.Color = BG
    ws.Cells.Font.Color     = WHITE
    ws.Cells.Font.Name      = "Segoe UI"
    ws.Cells.Font.Size      = 10

    # Hide gridlines
    xl.ActiveWindow.DisplayGridlines = False

    # -------------------------------------------------------------------------
    # Row 1-2: Header
    # -------------------------------------------------------------------------
    ws.Rows("1:2").RowHeight = 30
    ws.Range("A1:T2").Merge()
    hdr = ws.Cells(1,1)
    hdr.Value = f"  Budget & Commitments Dashboard  ·  CVR Month: {cvr_month}"
    hdr.Font.Bold  = True
    hdr.Font.Size  = 16
    hdr.Font.Color = WHITE
    hdr.HorizontalAlignment = XL_LEFT
    hdr.VerticalAlignment   = XL_MIDDLE
    hdr.Interior.Color      = SURFACE
    # Accent bar on left
    ws.Columns("A").ColumnWidth = 1
    ws.Range("A1:A2").Interior.Color = ACCENT

    # -------------------------------------------------------------------------
    # Row 3: spacer
    # -------------------------------------------------------------------------
    ws.Rows("3").RowHeight = 8
    ws.Rows("3").Interior.Color = BG

    # -------------------------------------------------------------------------
    # Compute aggregates for KPIs
    # -------------------------------------------------------------------------
    total_eac = sum(r["eac"] for r in records)
    total_ctd = sum(r["ctd"] for r in records)
    total_rem = sum(r["rem"] for r in records)
    total_acc = sum(r["acc"] for r in records)
    total_act = sum(r["act"] for r in records)
    total_com = sum(r["com"] for r in records)
    over_count  = len(set(r["cc"] for r in records if r["ou_status"]=="Over"))
    under_count = len(set(r["cc"] for r in records if r["ou_status"]=="Under"))
    unique_ccs  = len(set(r["cc"] for r in records))

    kpis = [
        ("Total EAC",         total_eac,     ACCENT,  "$#,##0"),
        ("Cost to Date",      total_ctd,     TEAL,    "$#,##0"),
        ("Remaining Budget",  total_rem,     GREEN if total_rem >= 0 else RED, "$#,##0"),
        ("Actuals",           total_act,     YELLOW,  "$#,##0"),
        ("Accrued",           total_acc,     PURPLE,  "$#,##0"),
        ("Committed",         total_com,     ec(255,165,0), "$#,##0"),
        ("Cost Codes",        unique_ccs,    SUBTEXT, "0"),
        ("Over Budget",       over_count,    RED,     "0"),
        ("Under Budget",      under_count,   GREEN,   "0"),
    ]

    # -------------------------------------------------------------------------
    # Rows 4-10: KPI cards  (9 cards, 2 cols each + spacer = 20 cols)
    # -------------------------------------------------------------------------
    ws.Rows("4:10").RowHeight = 18
    kpi_start_col = 2  # B
    card_w = 2         # each card spans 2 columns
    gap    = 0

    for k, (label, value, color, nfmt) in enumerate(kpis):
        col = kpi_start_col + k * (card_w + gap)
        # Card background
        card_rng = ws.Range(ws.Cells(4, col), ws.Cells(10, col+card_w-1))
        card_rng.Interior.Color = SURFACE
        # Color bar top
        bar_rng = ws.Range(ws.Cells(4, col), ws.Cells(4, col+card_w-1))
        bar_rng.Interior.Color = color
        ws.Rows("4").RowHeight = 4
        # Label
        lbl_rng = ws.Range(ws.Cells(5, col), ws.Cells(6, col+card_w-1))
        lbl_rng.Merge()
        lbl_cell = ws.Cells(5, col)
        lbl_cell.Value = label
        lbl_cell.Font.Size  = 9
        lbl_cell.Font.Color = SUBTEXT
        lbl_cell.HorizontalAlignment = XL_CENTER
        lbl_cell.VerticalAlignment   = XL_MIDDLE
        lbl_cell.Interior.Color      = SURFACE
        # Value
        val_rng = ws.Range(ws.Cells(7, col), ws.Cells(10, col+card_w-1))
        val_rng.Merge()
        val_cell = ws.Cells(7, col)
        val_cell.Value          = value
        val_cell.NumberFormat   = nfmt
        val_cell.Font.Size      = 14
        val_cell.Font.Bold      = True
        val_cell.Font.Color     = color
        val_cell.HorizontalAlignment = XL_CENTER
        val_cell.VerticalAlignment   = XL_MIDDLE
        val_cell.Interior.Color      = SURFACE

    ws.Rows("11").RowHeight = 10  # spacer

    # -------------------------------------------------------------------------
    # Build chart-data ranges on _Data sheet (aggregated)
    # -------------------------------------------------------------------------

    # --- Chart data on a helper sheet ---
    try:
        ws_calc = wb.Sheets("_Calc")
        ws_calc.Cells.Clear()
    except:
        ws_calc = add_sheet(wb, "_Calc")
        ws_calc.Name = "_Calc"
        ws_calc.Visible = 2  # very hidden

    # Aggregate by cost code (sum all 3 categories)
    from collections import defaultdict
    by_cc = defaultdict(lambda: dict(eac=0,ctd=0,rem=0,act=0,acc=0,com=0,item=""))
    by_cat = defaultdict(lambda: dict(eac=0,ctd=0,rem=0,act=0,acc=0,com=0))

    for r in records:
        cc  = r["cc"]
        cat = r["category"]
        by_cc[cc]["eac"] += r["eac"];  by_cc[cc]["ctd"] += r["ctd"]
        by_cc[cc]["rem"] += r["rem"];  by_cc[cc]["act"] += r["act"]
        by_cc[cc]["acc"] += r["acc"];  by_cc[cc]["com"] += r["com"]
        by_cc[cc]["item"] = r["item"]
        by_cat[cat]["eac"] += r["eac"]; by_cat[cat]["ctd"] += r["ctd"]
        by_cat[cat]["rem"] += r["rem"]; by_cat[cat]["act"] += r["act"]
        by_cat[cat]["acc"] += r["acc"]; by_cat[cat]["com"] += r["com"]

    # Sort by EAC descending, top 15 for bar chart
    sorted_ccs = sorted(by_cc.items(), key=lambda x: x[1]["eac"], reverse=True)[:15]
    sorted_rem = sorted(by_cc.items(), key=lambda x: x[1]["rem"], reverse=True)[:10]

    # Write Chart 1 data: EAC vs CTD by top 15 cost codes  (rows 1-17, cols A-D)
    ws_calc.Cells(1,1).Value = "Item"
    ws_calc.Cells(1,2).Value = "EAC"
    ws_calc.Cells(1,3).Value = "CTD"
    ws_calc.Cells(1,4).Value = "Remaining"
    for i,(cc,d) in enumerate(sorted_ccs, 2):
        ws_calc.Cells(i,1).Value = cc
        ws_calc.Cells(i,2).Value = d["eac"]
        ws_calc.Cells(i,3).Value = d["ctd"]
        ws_calc.Cells(i,4).Value = d["rem"]
    ch1_last = len(sorted_ccs) + 1

    # Write Chart 2 data: Remaining by Category  (rows 1-5, cols F-H)
    categories = ["Subcontract","Other Costs","Materials"]
    ws_calc.Cells(1,6).Value = "Category"
    ws_calc.Cells(1,7).Value = "EAC"
    ws_calc.Cells(1,8).Value = "Remaining"
    for i, cat in enumerate(categories, 2):
        ws_calc.Cells(i,6).Value = cat
        ws_calc.Cells(i,7).Value = by_cat[cat]["eac"]
        ws_calc.Cells(i,8).Value = by_cat[cat]["rem"]

    # Write Chart 3 data: Top 10 by Remaining  (rows 1-12, cols J-L)
    ws_calc.Cells(1,10).Value = "Cost Code"
    ws_calc.Cells(1,11).Value = "Remaining"
    for i,(cc,d) in enumerate(sorted_rem, 2):
        ws_calc.Cells(i,10).Value = cc
        ws_calc.Cells(i,11).Value = d["rem"]
    ch3_last = len(sorted_rem) + 1

    # Write Chart 4 data: Actuals/Accrued/Committed breakdown  (rows 1-4, cols M-N)
    ws_calc.Cells(1,13).Value = "Category"
    ws_calc.Cells(1,14).Value = "Amount"
    ws_calc.Cells(2,13).Value = "Actuals"
    ws_calc.Cells(3,13).Value = "Accrued"
    ws_calc.Cells(4,13).Value = "Committed"
    ws_calc.Cells(2,14).Value = total_act
    ws_calc.Cells(3,14).Value = total_acc
    ws_calc.Cells(4,14).Value = total_com

    # -------------------------------------------------------------------------
    # Charts on Dashboard
    # Use points: 1 point ≈ 0.75 px; Excel default col ~48pt, row ~12pt
    # Approx col width (cols are 10 chars wide ≈ 64pt each)
    # Layout:  2 charts side by side in rows 12-32, then 2 more in 33-53
    # -------------------------------------------------------------------------

    # Rough pixel positions (Excel uses points)
    col_px   = 64   # pts per column (approx for col width 10)
    row_px   = 18   # pts per row

    left1  = col_px * 1       # col B
    left2  = col_px * 10 + 10 # col K
    top1   = row_px * 12      # row 12
    top2   = row_px * 34      # row 34
    cw     = col_px * 9 - 10  # chart width
    ch_h   = row_px * 20      # chart height

    # Chart 1: EAC vs CTD (grouped bar)
    ch1_range = f"A1:C{ch1_last}"
    co1, c1 = add_chart(ws, CHART_TYPE["bar_cluster"], left1, top1, cw, ch_h,
                         "CVR EAC vs Cost to Date", ws_calc, ch1_range,
                         series_colors=[ACCENT, TEAL], has_legend=True)

    # Chart 2: Remaining by Category (grouped bar)
    co2, c2 = add_chart(ws, CHART_TYPE["bar_cluster"], left2, top1, cw, ch_h,
                         "Remaining Budget by Category", ws_calc, "F1:H4",
                         series_colors=[PURPLE, GREEN], has_legend=True)

    # Chart 3: Top 10 Remaining (horizontal bar)
    ch3_range = f"J1:K{ch3_last}"
    co3, c3 = add_chart(ws, CHART_TYPE["bar_horiz"], left1, top2, cw, ch_h,
                         "Top 10 Items by Remaining Budget", ws_calc, ch3_range,
                         series_colors=[ACCENT], has_legend=False)

    # Chart 4: Actuals / Accrued / Committed (donut)
    co4, c4 = add_chart(ws, CHART_TYPE["donut"], left2, top2, cw, ch_h,
                         "Actuals · Accrued · Committed", ws_calc, "M1:N4",
                         series_colors=[YELLOW, PURPLE, ec(255,140,0)], has_legend=True)

    # Donut hole
    try:
        c4.SeriesCollection(1).Explosion = 0
        c4.SeriesCollection(1).Parent.HoleSize = 55
    except: pass

    # -------------------------------------------------------------------------
    # Slicers connected to the BudgetData table
    # -------------------------------------------------------------------------
    try:
        # Slicer cache for Category
        sc_cat = wb.SlicerCaches.Add2(ws_data.ListObjects("BudgetData"), "Category")
        sl_cat = sc_cat.Slicers.Add(
            ws,
            Type=1,
            Top=top1,
            Left=left2 + cw + 10,
            Width=120,
            Height=ch_h // 2 - 10,
        )
        sl_cat.Caption = "Category"
        sl_cat.Style   = "SlicerStyleDark1"

        # Slicer for O/U Status
        sc_ou = wb.SlicerCaches.Add2(ws_data.ListObjects("BudgetData"), "O/U Status")
        sl_ou = sc_ou.Slicers.Add(
            ws,
            Type=1,
            Top=top1 + ch_h // 2,
            Left=left2 + cw + 10,
            Width=120,
            Height=ch_h // 2 - 10,
        )
        sl_ou.Caption = "Over / Under"
        sl_ou.Style   = "SlicerStyleDark1"
    except Exception as e:
        print(f"  (Slicers skipped – {e})")

    # -------------------------------------------------------------------------
    # Mini data table at bottom (rows 55+)
    # -------------------------------------------------------------------------
    tbl_top = 56
    ws.Rows(f"{tbl_top}:{tbl_top}").RowHeight = 20
    tbl_headers = ["Cost Code","Item","EAC","CTD","Remaining","O/U"]
    tbl_colors  = [SURFACE2]*6

    for c, h in enumerate(tbl_headers, 2):
        cell = ws.Cells(tbl_top, c)
        cell.Value              = h
        cell.Font.Bold          = True
        cell.Font.Size          = 9
        cell.Font.Color         = WHITE
        cell.Interior.Color     = SURFACE2
        cell.HorizontalAlignment = XL_CENTER

    for i, (cc, d) in enumerate(sorted(by_cc.items(), key=lambda x: x[1]["eac"], reverse=True), 1):
        row = tbl_top + i
        vals = [cc, d["item"], d["eac"], d["ctd"], d["rem"]]
        for c, v in enumerate(vals, 2):
            cell = ws.Cells(row, c)
            cell.Value          = v
            cell.Font.Size      = 9
            cell.Interior.Color = SURFACE if i % 2 == 0 else BG
            cell.Font.Color     = WHITE
            if c >= 4:
                cell.NumberFormat = "$#,##0"
        # O/U status
        ou_val = "Over" if d["rem"] < 0 else "Under"
        cell_ou = ws.Cells(row, 7)
        cell_ou.Value          = ou_val
        cell_ou.Font.Size      = 9
        cell_ou.Font.Bold      = True
        cell_ou.Font.Color     = RED if ou_val == "Over" else GREEN
        cell_ou.Interior.Color = SURFACE if i % 2 == 0 else BG
        cell_ou.HorizontalAlignment = XL_CENTER

    xl.ScreenUpdating = True
    return ws


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    if not os.path.exists(SRC):
        input(f"ERROR: Excel file not found in:\n  {HERE}\n\nPress Enter.")
        sys.exit(1)

    print(f"Opening: {SRC}")
    xl = win32.Dispatch("Excel.Application")
    xl.Visible      = False
    xl.DisplayAlerts = False

    try:
        wb = xl.Workbooks.Open(SRC)

        print("Reading data...")
        records, cvr_month = read_data(wb)
        print(f"  {len(records)} rows ({len(records)//3} cost codes), CVR: {cvr_month}")

        print("Building data sheet...")
        ws_data, _ = build_data_sheet(wb, records)

        print("Building dashboard...")
        build_dashboard(wb, ws_data, records, cvr_month)

        # Make Dashboard the active sheet
        for sh in wb.Sheets:
            if sh.Name.lower() == "dashboard":
                sh.Activate()
                break

        print(f"Saving as: {SAVE_AS}")
        wb.SaveAs(SAVE_AS, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled
        wb.Close(False)

    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback; traceback.print_exc()
        try: xl.Quit()
        except: pass
        input("\nPress Enter to close.")
        sys.exit(1)
    finally:
        try: xl.Quit()
        except: pass

    input(
        f"\nDone!\nExcel dashboard created: {SAVE_AS}\n\n"
        "Open the file and go to the Dashboard tab.\n"
        "Use the slicers on the right to filter all charts by Category or Over/Under.\n\n"
        "Press Enter to close."
    )


if __name__ == "__main__":
    main()
