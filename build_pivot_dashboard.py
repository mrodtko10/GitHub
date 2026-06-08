"""
build_pivot_dashboard.py
Creates a full pivot-table dashboard inside Excel with:
  - _Data sheet: Excel Table "BudgetData" (auto-expands as rows are added)
  - 4 PivotTables (hidden sheets) all sharing one PivotCache
  - 4 PivotCharts embedded on the Dashboard sheet
  - Category + O/U Slicers connected to ALL 4 PivotTables
  - VBA "Refresh Dashboard" button that re-reads Category Table → updates data → refreshes all pivots
  - Dark-themed Dashboard KPI cards

Requirements: pip install pywin32
Run from the same folder as Weekly_Commitments_Review.xlsx
"""

import os, sys, time
try:
    import win32com.client as win32
    from win32com.client import constants as xlc
except ImportError:
    input("ERROR: pywin32 not installed.\nRun:  pip install pywin32\nPress Enter.")
    sys.exit(1)

HERE      = os.path.dirname(os.path.abspath(__file__))
SRC       = os.path.join(HERE, "Weekly_Commitments_Review.xlsx")
if not os.path.exists(SRC):
    SRC   = os.path.join(HERE, "Weekly_Commitments_Review.xlsm")
SAVE_AS   = os.path.join(HERE, "Weekly_Commitments_Review.xlsm")

# ── Excel constants ──────────────────────────────────────────────────────────
xlDatabase          = 1
xlRowField          = 1
xlColumnField       = 2
xlPageField         = 3
xlDataField         = 4
xlSum               = -4157
xlCount             = -4112
xlUp                = -4162
xlColumnClustered   = 57
xlBarClustered      = -4100
xlDoughnut          = -4120
xlPie               = 5
xlLocationAsObject  = 2
xlLocationAsNewSheet= 1
xlVeryHidden        = 2
xlSheetHidden       = 0

# ── Color helper (BGR int for Excel) ────────────────────────────────────────
def rgb(r, g, b): return b*65536 + g*256 + r
BG      = rgb(15,  17,  23)
SURFACE = rgb(26,  29,  39)
SRF2    = rgb(34,  38,  58)
BORDER  = rgb(46,  51,  80)
ACCENT  = rgb(79,  142, 247)
PURPLE  = rgb(124, 92,  252)
GREEN   = rgb(34,  197, 94)
RED     = rgb(239, 68,  68)
YELLOW  = rgb(245, 158, 11)
TEAL    = rgb(20,  184, 166)
WHITE   = rgb(226, 232, 240)
SUB     = rgb(148, 163, 184)
ORANGE  = rgb(249, 115, 22)

# ── VBA refresh macro ────────────────────────────────────────────────────────
VBA = r'''
Option Explicit

Sub RefreshDashboard()
    Dim wsData  As Worksheet
    Dim wsSrc   As Worksheet
    Dim tbl     As ListObject
    Dim lastRow As Long
    Dim i       As Long
    Dim newRow  As ListRow
    Dim cc      As String

    Application.ScreenUpdating = False
    Application.Calculation   = xlCalculationAutomatic

    ' Find source sheet
    For Each wsSrc In ThisWorkbook.Sheets
        If InStr(LCase(wsSrc.Name), "category") > 0 And _
           InStr(LCase(wsSrc.Name), "table") > 0 Then Exit For
    Next wsSrc
    If wsSrc Is Nothing Then
        MsgBox "Category Table sheet not found.", vbCritical: Exit Sub
    End If

    ' Force recalculation of source sheet
    wsSrc.Calculate

    ' Get _Data sheet and BudgetData table
    Set wsData = ThisWorkbook.Sheets("_Data")
    Set tbl    = wsData.ListObjects("BudgetData")

    ' Clear existing data rows (keep header)
    On Error Resume Next
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
    On Error GoTo 0

    ' Find last row in Category Table (col A)
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    ' Write new rows (data starts row 4 = index 4, skip headers 1-3)
    Dim catNames(1 To 3) As String
    catNames(1) = "Subcontract"
    catNames(2) = "Other Costs"
    catNames(3) = "Materials"

    ' Column offsets for SU(1), OC(2), MA(3)
    '  EAC: 15,16,17  CTD: 33,34,35  Rem: 39,40,41
    '  Act: 24,25,26  Acc: 21,22,23  Com: 27,28,29
    Dim eacCols(1 To 3) As Integer
    Dim ctdCols(1 To 3) As Integer
    Dim remCols(1 To 3) As Integer
    Dim actCols(1 To 3) As Integer
    Dim accCols(1 To 3) As Integer
    Dim comCols(1 To 3) As Integer
    eacCols(1)=15: eacCols(2)=16: eacCols(3)=17
    ctdCols(1)=33: ctdCols(2)=34: ctdCols(3)=35
    remCols(1)=39: remCols(2)=40: remCols(3)=41
    actCols(1)=24: actCols(2)=25: actCols(3)=26
    accCols(1)=21: accCols(2)=22: accCols(3)=23
    comCols(1)=27: comCols(2)=28: comCols(3)=29

    For i = 4 To lastRow
        cc = Trim(CStr(wsSrc.Cells(i, 1).Value))
        If cc = "" Or cc = "Cost Code" Or cc = "Total" Then GoTo NextRow

        Dim itemName As String
        itemName = Trim(CStr(wsSrc.Cells(i, 2).Value))

        Dim k As Integer
        For k = 1 To 3
            Dim eacV As Double, ctdV As Double, remV As Double
            Dim actV As Double, accV As Double, comV As Double
            eacV = SafeNum(wsSrc.Cells(i, eacCols(k)).Value)
            ctdV = SafeNum(wsSrc.Cells(i, ctdCols(k)).Value)
            remV = SafeNum(wsSrc.Cells(i, remCols(k)).Value)
            actV = SafeNum(wsSrc.Cells(i, actCols(k)).Value)
            accV = SafeNum(wsSrc.Cells(i, accCols(k)).Value)
            comV = SafeNum(wsSrc.Cells(i, comCols(k)).Value)

            Dim ouStat As String
            If remV < 0 Then ouStat = "Over" Else ouStat = "Under"

            Set newRow = tbl.ListRows.Add
            With newRow.Range
                .Cells(1, 1).Value  = cc
                .Cells(1, 2).Value  = itemName
                .Cells(1, 3).Value  = catNames(k)
                .Cells(1, 4).Value  = eacV
                .Cells(1, 5).Value  = ctdV
                .Cells(1, 6).Value  = remV
                .Cells(1, 7).Value  = actV
                .Cells(1, 8).Value  = accV
                .Cells(1, 9).Value  = comV
                .Cells(1, 10).Value = ouStat
            End With
        Next k
NextRow:
    Next i

    ' Refresh all PivotTables
    ThisWorkbook.RefreshAll
    Application.ScreenUpdating = True
    MsgBox "Dashboard refreshed! " & tbl.DataBodyRange.Rows.Count & _
           " rows loaded.", vbInformation, "Done"
End Sub

Private Function SafeNum(v As Variant) As Double
    If IsNumeric(v) Then SafeNum = CDbl(v) Else SafeNum = 0
End Function
'''

# ── Helpers ──────────────────────────────────────────────────────────────────
def delete_sheet(wb, name):
    xl = wb.Application
    xl.DisplayAlerts = False
    for sh in list(wb.Sheets):
        if sh.Name == name:
            sh.Delete()
    xl.DisplayAlerts = True

def get_or_create(wb, name, after=None):
    delete_sheet(wb, name)
    if after:
        return wb.Sheets.Add(After=after)
    return wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))

def fmt_cells(ws, r1, c1, r2, c2, bg=None, fg=WHITE, bold=False, sz=10,
              halign=-4108, valign=-4108, nfmt=None):
    rng = ws.Range(ws.Cells(r1,c1), ws.Cells(r2,c2))
    if bg is not None: rng.Interior.Color = bg
    rng.Font.Color = fg
    rng.Font.Bold  = bold
    rng.Font.Size  = sz
    rng.Font.Name  = "Segoe UI"
    rng.HorizontalAlignment = halign
    rng.VerticalAlignment   = valign
    if nfmt: rng.NumberFormat = nfmt
    return rng

def merge_fmt(ws, r1, c1, r2, c2, val=None, bg=None, fg=WHITE,
              bold=False, sz=10, halign=-4108, valign=-4108, nfmt=None):
    rng = ws.Range(ws.Cells(r1,c1), ws.Cells(r2,c2))
    rng.Merge()
    if val is not None: ws.Cells(r1,c1).Value = val
    if bg is not None: rng.Interior.Color = bg
    rng.Font.Color = fg
    rng.Font.Bold  = bold
    rng.Font.Size  = sz
    rng.Font.Name  = "Segoe UI"
    rng.HorizontalAlignment = halign
    rng.VerticalAlignment   = valign
    if nfmt: rng.NumberFormat = nfmt
    return rng

def style_chart(chart, title_text):
    """Apply dark theme to a chart object."""
    chart.HasTitle = True
    chart.ChartTitle.Text = title_text
    chart.ChartTitle.Font.Color = WHITE
    chart.ChartTitle.Font.Size  = 11
    chart.ChartTitle.Font.Bold  = True
    chart.ChartTitle.Font.Name  = "Segoe UI"
    # Background
    chart.ChartArea.Format.Fill.Solid()
    chart.ChartArea.Format.Fill.ForeColor.RGB  = SURFACE
    chart.ChartArea.Border.LineStyle = 0
    chart.PlotArea.Format.Fill.Solid()
    chart.PlotArea.Format.Fill.ForeColor.RGB   = SRF2
    chart.PlotArea.Border.LineStyle = 0
    # Legend
    try:
        chart.HasLegend = True
        chart.Legend.Font.Color = SUB
        chart.Legend.Font.Size  = 8
        chart.Legend.Font.Name  = "Segoe UI"
        chart.Legend.Position   = -4107  # xlLegendPositionBottom
    except: pass

def style_axes(chart, x_title=None, y_title=None, y_fmt="$#,##0"):
    try:
        ax_cat = chart.Axes(1)   # xlCategory
        ax_val = chart.Axes(2)   # xlValue
        for ax in (ax_cat, ax_val):
            ax.TickLabels.Font.Color = SUB
            ax.TickLabels.Font.Size  = 8
            ax.TickLabels.Font.Name  = "Segoe UI"
            ax.Format.Line.ForeColor.RGB = BORDER
        ax_val.MajorGridlines.Format.Line.ForeColor.RGB = BORDER
        ax_val.NumberFormat = y_fmt
        if x_title:
            ax_cat.HasTitle = True
            ax_cat.AxisTitle.Text = x_title
            ax_cat.AxisTitle.Font.Color = SUB
            ax_cat.AxisTitle.Font.Size  = 8
        if y_title:
            ax_val.HasTitle = True
            ax_val.AxisTitle.Text = y_title
            ax_val.AxisTitle.Font.Color = SUB
            ax_val.AxisTitle.Font.Size  = 8
    except: pass

def color_series(chart, series_colors):
    for i, clr in enumerate(series_colors):
        try:
            s = chart.SeriesCollection(i+1)
            s.Format.Fill.Solid()
            s.Format.Fill.ForeColor.RGB = clr
            s.Format.Line.ForeColor.RGB = clr
        except: pass

def add_data_labels(chart, num_fmt="$#,##0"):
    try:
        for i in range(1, chart.SeriesCollection().Count + 1):
            s = chart.SeriesCollection(i)
            s.HasDataLabels = True
            dl = s.DataLabels()
            dl.ShowValue = True
            dl.NumberFormat = num_fmt
            dl.Font.Color = WHITE
            dl.Font.Size  = 7
            dl.Font.Name  = "Segoe UI"
    except: pass


# ── Main ────────────────────────────────────────────────────────────────────
def main():
    print(f"Opening: {SRC}")
    xl = win32.Dispatch("Excel.Application")
    xl.Visible       = False
    xl.DisplayAlerts = False

    try:
        wb = xl.Workbooks.Open(SRC)
        xl.Calculate()

        # ── Find Category Table sheet ────────────────────────────────────────
        ws_cat = None
        for sh in wb.Sheets:
            if "category" in sh.Name.lower() and "table" in sh.Name.lower():
                ws_cat = sh; break
        if not ws_cat:
            for sh in wb.Sheets:
                if "dashboard" not in sh.Name.lower() and "po" not in sh.Name.lower():
                    ws_cat = sh; break
        print(f"  Category Table: {ws_cat.Name}")

        last_src_row = ws_cat.Cells(ws_cat.Rows.Count, 1).End(xlUp).Row
        print(f"  Source rows: {last_src_row}")

        # ── Read data ────────────────────────────────────────────────────────
        cats   = ["Subcontract", "Other Costs", "Materials"]
        # col indices (1-based): EAC 15/16/17, CTD 33/34/35, Rem 39/40/41
        #                        Act 24/25/26, Acc 21/22/23, Com 27/28/29
        eac_c = [15,16,17]; ctd_c = [33,34,35]; rem_c = [39,40,41]
        act_c = [24,25,26]; acc_c = [21,22,23]; com_c = [27,28,29]

        def sn(v):
            try: return float(v) if v not in (None,"") else 0.0
            except: return 0.0

        records = []
        for i in range(4, last_src_row + 1):
            cc   = str(ws_cat.Cells(i,1).Value or "").strip()
            item = str(ws_cat.Cells(i,2).Value or "").strip()
            if not cc or cc in ("Cost Code","Total","0"):
                continue
            for k in range(3):
                eac = sn(ws_cat.Cells(i, eac_c[k]).Value)
                ctd = sn(ws_cat.Cells(i, ctd_c[k]).Value)
                rem = sn(ws_cat.Cells(i, rem_c[k]).Value)
                act = sn(ws_cat.Cells(i, act_c[k]).Value)
                acc = sn(ws_cat.Cells(i, acc_c[k]).Value)
                com = sn(ws_cat.Cells(i, com_c[k]).Value)
                ou  = "Over" if rem < 0 else "Under"
                records.append((cc, item, cats[k], eac, ctd, rem, act, acc, com, ou))

        print(f"  Records: {len(records)}")

        # ── _Data sheet ──────────────────────────────────────────────────────
        ws_data = get_or_create(wb, "_Data")
        ws_data.Name = "_Data"

        HEADERS = ["Cost Code","Item","Category","EAC","CTD",
                   "Remaining","Actuals","Accrued","Committed","O/U Status"]
        for c, h in enumerate(HEADERS, 1):
            ws_data.Cells(1, c).Value = h

        for r, rec in enumerate(records, 2):
            for c, v in enumerate(rec, 1):
                ws_data.Cells(r, c).Value = v

        # Format money columns
        for col_idx in [4,5,6,7,8,9]:
            col_letter = chr(ord('A') + col_idx - 1)
            ws_data.Range(f"{col_letter}2:{col_letter}{len(records)+1}").NumberFormat = "$#,##0.00"

        # Create Excel Table
        tbl_rng = ws_data.Range(ws_data.Cells(1,1), ws_data.Cells(len(records)+1, 10))
        tbl = ws_data.ListObjects.Add(xlDatabase, tbl_rng, None, 1)
        tbl.Name = "BudgetData"
        tbl.TableStyle = "TableStyleMedium9"
        ws_data.Columns("A:J").AutoFit()
        ws_data.Visible = xlVeryHidden

        print("  _Data sheet built.")

        # ── PivotCache from BudgetData table ─────────────────────────────────
        # Reference by table object range so new rows auto-include
        pc = wb.PivotCaches().Create(xlDatabase, tbl.Range)
        print("  PivotCache created.")

        # ── Create PivotTables ───────────────────────────────────────────────
        def new_pt_sheet(name):
            ws = get_or_create(wb, name)
            ws.Name = name
            ws.Visible = xlVeryHidden
            return ws

        def make_pt(ws, pt_name, row_fields, value_fields):
            """
            row_fields: list of field names for rows
            value_fields: list of (field_name, caption, num_fmt)
            """
            pt = pc.CreatePivotTable(
                TableDestination=ws.Range("A1"),
                TableName=pt_name
            )
            pos = 1
            for f in row_fields:
                pf = pt.PivotFields(f)
                pf.Orientation = xlRowField
                pf.Position     = pos
                pos += 1
            for fname, caption, nfmt in value_fields:
                df = pt.AddDataField(pt.PivotFields(fname), caption, xlSum)
                df.NumberFormat = nfmt
            pt.ShowTableStyleRowStripes = False
            pt.TableStyle2 = ""
            return pt

        # PT1: EAC vs CTD by Cost Code
        ws_pt1 = new_pt_sheet("_PT1")
        pt1 = make_pt(ws_pt1, "PT_EAC_CTD",
                      ["Cost Code"],
                      [("EAC","CVR EAC","$#,##0"),
                       ("CTD","Cost to Date","$#,##0")])

        # PT2: EAC vs Remaining by Category
        ws_pt2 = new_pt_sheet("_PT2")
        pt2 = make_pt(ws_pt2, "PT_CAT_REM",
                      ["Category"],
                      [("EAC","EAC","$#,##0"),
                       ("Remaining","Remaining","$#,##0")])

        # PT3: Remaining by Cost Code (for horizontal bar top items)
        ws_pt3 = new_pt_sheet("_PT3")
        pt3 = make_pt(ws_pt3, "PT_TOP_REM",
                      ["Cost Code"],
                      [("Remaining","Remaining","$#,##0")])

        # PT4: Actuals/Accrued/Committed breakdown (no rows = grand total only)
        ws_pt4 = new_pt_sheet("_PT4")
        pt4 = make_pt(ws_pt4, "PT_BREAKDOWN",
                      ["Category"],
                      [("Actuals","Actuals","$#,##0"),
                       ("Accrued","Accrued","$#,##0"),
                       ("Committed","Committed","$#,##0")])

        print("  PivotTables created.")

        # ── Dashboard sheet ──────────────────────────────────────────────────
        # Find or create Dashboard as first visible sheet
        dash_ws = None
        for sh in wb.Sheets:
            if sh.Name.lower() == "dashboard":
                dash_ws = sh; break
        if dash_ws is None:
            dash_ws = wb.Sheets.Add(Before=wb.Sheets(1))
            dash_ws.Name = "Dashboard"
        else:
            dash_ws.Cells.Clear()
            for shp in list(dash_ws.Shapes):
                shp.Delete()
            for sc in list(wb.SlicerCaches):
                try: sc.Delete()
                except: pass

        dash_ws.Activate()
        xl.ActiveWindow.DisplayGridlines = False
        dash_ws.Tab.Color = ACCENT

        # Column widths
        #  A=margin, B-J=left zone (9 cols), K=gap, L-T=right zone (9 cols), U=margin
        dash_ws.Columns("A").ColumnWidth = 1.5
        dash_ws.Columns("K").ColumnWidth = 1.5
        dash_ws.Columns("U").ColumnWidth = 1.5
        for c in range(2, 11):   # B-J
            dash_ws.Columns(c).ColumnWidth = 11.5
        for c in range(12, 21):  # L-T
            dash_ws.Columns(c).ColumnWidth = 11.5

        # Fill canvas dark
        fmt_cells(dash_ws, 1, 1, 95, 21, bg=BG)

        # ── HEADER rows 1-3 ──────────────────────────────────────────────────
        dash_ws.Rows(1).RowHeight = 5
        dash_ws.Rows(2).RowHeight = 34
        dash_ws.Rows(3).RowHeight = 5
        fmt_cells(dash_ws, 1, 1, 3, 21, bg=SURFACE)
        dash_ws.Range("A1:A3").Interior.Color = ACCENT   # left stripe

        merge_fmt(dash_ws, 2, 2, 2, 13,
                  val="Budget & Commitments  Dashboard",
                  bg=SURFACE, fg=WHITE, bold=True, sz=18, halign=-4131)

        # Get CVR month from row 1 of Category Table
        cvr = ""
        for c in range(1, 6):
            v = str(ws_cat.Cells(1, c).Value or "").strip()
            v = v.replace("CVR Month:","").replace("CVR Month :","").strip()
            if v and v != "0": cvr = v; break

        merge_fmt(dash_ws, 2, 14, 2, 20,
                  val=f"CVR Month:  {cvr}",
                  bg=ACCENT, fg=WHITE, bold=True, sz=11, halign=-4152)

        dash_ws.Rows(4).RowHeight = 10

        # ── KPI CARDS rows 5-11 ──────────────────────────────────────────────
        # Totals
        tot_eac = sum(r[3] for r in records)
        tot_ctd = sum(r[4] for r in records)
        tot_rem = sum(r[5] for r in records)
        tot_act = sum(r[6] for r in records)
        tot_acc = sum(r[7] for r in records)
        tot_com = sum(r[8] for r in records)
        over_n  = len(set(r[0] for r in records if r[9]=="Over"))
        under_n = len(set(r[0] for r in records if r[9]=="Under"))

        KPIS = [
            ("Total EAC",     tot_eac, ACCENT, "$#,##0"),
            ("Cost to Date",  tot_ctd, TEAL,   "$#,##0"),
            ("Remaining",     tot_rem, GREEN if tot_rem>=0 else RED, "$#,##0"),
            ("Actuals",       tot_act, YELLOW, "$#,##0"),
            ("Accrued",       tot_acc, PURPLE, "$#,##0"),
            ("Committed",     tot_com, ORANGE, "$#,##0"),
            ("Over Budget",   over_n,  RED,    "0"),
            ("Under Budget",  under_n, GREEN,  "0"),
        ]
        for row in (5,6,7,8,9,10,11):
            dash_ws.Rows(row).RowHeight = {5:4,6:18,7:4,8:26,9:4,10:4,11:10}.get(row,16)

        for i,(lbl,val,clr,nfmt) in enumerate(KPIS):
            c1 = 2 + i*2
            c2 = c1 + 1
            fmt_cells(dash_ws, 5, c1, 11, c2, bg=SURFACE)
            fmt_cells(dash_ws, 5, c1, 5,  c2, bg=clr)      # accent top bar
            merge_fmt(dash_ws, 6, c1, 6, c2, val=lbl,
                      bg=SURFACE, fg=SUB, sz=9)
            merge_fmt(dash_ws, 8, c1, 8, c2, val=val,
                      bg=SURFACE, fg=clr, bold=True, sz=14, nfmt=nfmt)

        dash_ws.Rows(12).RowHeight = 10

        # ── CHART ZONES ──────────────────────────────────────────────────────
        # Set uniform row heights for chart areas
        for r in range(13, 34):  dash_ws.Rows(r).RowHeight = 18
        dash_ws.Rows(34).RowHeight = 12   # gap
        for r in range(35, 56):  dash_ws.Rows(r).RowHeight = 18
        dash_ws.Rows(56).RowHeight = 14   # gap before table

        # Chart positioning (points: 1 col ≈ 73pt, 1 row ≈ 18pt)
        # Left zone:  col B(2) →  L_left = col A width (~11pt)
        # Right zone: col L(12) → R_left = L_left + 9*73
        col_w   = 73   # pt per col
        row_h   = 18   # pt per row
        left_L  = col_w * 1         # left of left zone (col B)
        left_R  = col_w * 11        # left of right zone (col L)
        ch_w    = col_w * 9 - 5     # chart width (9 cols minus gap)
        ch_h    = row_h * 20        # chart height (20 rows)
        top1    = row_h * 12        # top of first chart row (row 13)
        top2    = row_h * 34 + 12   # top of second chart row (row 35)

        def add_pivotchart(pt_source, ch_type, left, top, w, h, title,
                           x_title, y_title, series_colors, y_fmt="$#,##0"):
            co = dash_ws.Shapes.AddChart2(-1, ch_type, left, top, w, h)
            ch = co.Chart
            ch.SetSourceData(pt_source.TableRange2)
            style_chart(ch, title)
            style_axes(ch, x_title=x_title, y_title=y_title, y_fmt=y_fmt)
            color_series(ch, series_colors)
            add_data_labels(ch, y_fmt)
            return co, ch

        # Chart 1: EAC vs CTD by Cost Code (left, row 1)
        co1, ch1 = add_pivotchart(
            pt1, xlColumnClustered, left_L, top1, ch_w, ch_h,
            "CVR EAC vs Cost to Date",
            x_title="Cost Code", y_title="Amount ($)",
            series_colors=[ACCENT, TEAL]
        )

        # Chart 2: EAC vs Remaining by Category (right, row 1)
        co2, ch2 = add_pivotchart(
            pt2, xlColumnClustered, left_R, top1, ch_w, ch_h,
            "Remaining Budget by Category",
            x_title="Category", y_title="Amount ($)",
            series_colors=[PURPLE, GREEN]
        )

        # Chart 3: Top Items by Remaining (left, row 2) — horizontal bar
        co3, ch3 = add_pivotchart(
            pt3, xlBarClustered, left_L, top2, ch_w, ch_h,
            "Items by Remaining Budget",
            x_title=None, y_title="Remaining ($)",
            series_colors=[ACCENT]
        )

        # Chart 4: Actuals/Accrued/Committed by Category (right, row 2)
        co4, ch4 = add_pivotchart(
            pt4, xlColumnClustered, left_R, top2, ch_w, ch_h,
            "Actuals · Accrued · Committed by Category",
            x_title="Category", y_title="Amount ($)",
            series_colors=[YELLOW, PURPLE, ORANGE]
        )

        print("  PivotCharts created.")

        # ── SLICERS ──────────────────────────────────────────────────────────
        # Slicers go to the right of the charts, spanning both chart rows
        slicer_left  = left_R + ch_w + 6
        slicer_w     = 130
        slicer_h_cat = ch_h           # full height for category slicer (row 1 height)
        slicer_h_ou  = ch_h           # O/U slicer (row 2 height)

        # Slicer: Category (connected to all 4 PTs)
        sc_cat = wb.SlicerCaches.Add2(pt1, "Category")
        sc_cat.Slicers.Add(
            SlicerDestination=dash_ws,
            Top=top1, Left=slicer_left,
            Width=slicer_w, Height=slicer_h_cat
        )
        sc_cat.PivotTables.AddPivotTable(pt2)
        sc_cat.PivotTables.AddPivotTable(pt3)
        sc_cat.PivotTables.AddPivotTable(pt4)

        # Slicer: O/U Status (connected to all 4 PTs)
        sc_ou = wb.SlicerCaches.Add2(pt1, "O/U Status")
        sc_ou.Slicers.Add(
            SlicerDestination=dash_ws,
            Top=top2, Left=slicer_left,
            Width=slicer_w, Height=slicer_h_ou
        )
        sc_ou.PivotTables.AddPivotTable(pt2)
        sc_ou.PivotTables.AddPivotTable(pt3)
        sc_ou.PivotTables.AddPivotTable(pt4)

        # Style slicers
        for sc in (sc_cat, sc_ou):
            for sl in sc.Slicers:
                sl.Style = "SlicerStyleDark1"

        print("  Slicers created.")

        # ── DATA TABLE (rows 57+) ─────────────────────────────────────────────
        tbl_start = 57
        dash_ws.Rows(tbl_start).RowHeight = 20
        merge_fmt(dash_ws, tbl_start, 2, tbl_start, 20,
                  val="  Cost Code Summary", bg=SRF2, fg=WHITE, bold=True, sz=11,
                  halign=-4131)

        hdr_row = tbl_start + 1
        dash_ws.Rows(hdr_row).RowHeight = 20
        HDR = [
            (2,"Cost Code",7,"center"),
            (3,"Item",     9,"left"),
            (8,"EAC",      9,"right"),
            (10,"CTD",     9,"right"),
            (12,"Remaining",9,"right"),
            (14,"Actuals", 9,"right"),
            (16,"O/U",     8,"center"),
        ]
        HA = {"center":-4108,"left":-4131,"right":-4152}
        for col,lbl,sz,ha in HDR:
            c = dash_ws.Cells(hdr_row, col)
            c.Value = lbl
            c.Interior.Color = SRF2
            c.Font.Color = WHITE; c.Font.Bold = True
            c.Font.Size = sz; c.Font.Name = "Segoe UI"
            c.HorizontalAlignment = HA[ha]

        # Aggregate by cost code for table
        from collections import defaultdict
        by_cc = defaultdict(lambda: {"eac":0,"ctd":0,"rem":0,"act":0,"item":""})
        for r in records:
            cc = r[0]
            by_cc[cc]["eac"] += r[3]; by_cc[cc]["ctd"] += r[4]
            by_cc[cc]["rem"] += r[5]; by_cc[cc]["act"] += r[6]
            by_cc[cc]["item"] = r[1]
        all_s = sorted(by_cc.items(), key=lambda x: x[1]["eac"], reverse=True)

        for i,(cc,d) in enumerate(all_s):
            row = hdr_row + 1 + i
            dash_ws.Rows(row).RowHeight = 16
            rb = SURFACE if i%2==0 else BG
            ou = d["rem"] < 0
            ou_lbl = "▲ Over" if ou else "✓ Under"
            ou_clr = RED if ou else GREEN

            fmt_cells(dash_ws, row, 2, row, 17, bg=rb)

            def wc(col, val, fg_, ha, nfmt=None):
                c = dash_ws.Cells(row, col)
                c.Value = val
                c.Interior.Color = rb
                c.Font.Color = fg_; c.Font.Size = 9; c.Font.Name = "Segoe UI"
                c.HorizontalAlignment = HA[ha]
                if nfmt: c.NumberFormat = nfmt

            wc(2,  cc,       WHITE,  "center")
            wc(3,  d["item"],SUB,    "left")
            wc(8,  d["eac"], WHITE,  "right", "$#,##0")
            wc(10, d["ctd"], TEAL,   "right", "$#,##0")
            wc(12, d["rem"], ou_clr, "right", "$#,##0")
            wc(14, d["act"], YELLOW, "right", "$#,##0")
            wc(16, ou_lbl,   ou_clr, "center")
            dash_ws.Cells(row, 16).Font.Bold = True
            # Span item cols 3-7
            try: dash_ws.Range(dash_ws.Cells(row,3), dash_ws.Cells(row,7)).Merge()
            except: pass

        # ── REFRESH BUTTON ───────────────────────────────────────────────────
        btn = dash_ws.Buttons().Add(
            slicer_left, top1 - 30, slicer_w, 24
        )
        btn.Caption  = "🔄  Refresh Dashboard"
        btn.OnAction = "RefreshDashboard"
        btn.Font.Size = 9
        btn.Font.Name = "Segoe UI"

        # ── VBA MODULE ───────────────────────────────────────────────────────
        xl.DisplayAlerts = False
        # Remove existing module if present
        try:
            for comp in wb.VBProject.VBComponents:
                if comp.Name == "DashboardRefresh":
                    wb.VBProject.VBComponents.Remove(comp)
                    break
        except: pass
        mod = wb.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule = 1
        mod.Name = "DashboardRefresh"
        mod.CodeModule.AddFromString(VBA)

        print("  VBA macro added.")

        # ── Freeze & activate ────────────────────────────────────────────────
        dash_ws.Activate()
        xl.ActiveWindow.FreezePanes = False
        dash_ws.Cells(13, 2).Select()
        xl.ActiveWindow.FreezePanes = False

        print(f"Saving as: {SAVE_AS}")
        wb.SaveAs(SAVE_AS, 52)   # 52 = xlOpenXMLWorkbookMacroEnabled
        wb.Close(False)
        print("Done.")

    except Exception as e:
        import traceback; traceback.print_exc()
        try: xl.Quit()
        except: pass
        input(f"\nERROR: {e}\nPress Enter.")
        sys.exit(1)
    finally:
        try: xl.Quit()
        except: pass

    input(
        f"\nSuccess!\n"
        f"File saved: {SAVE_AS}\n\n"
        f"→ Open the file and go to the Dashboard tab\n"
        f"→ Use the Category / O/U slicers to filter all 4 charts at once\n"
        f"→ Click '🔄 Refresh Dashboard' after updating the Excel data\n\n"
        "Press Enter to close."
    )

if __name__ == "__main__":
    main()
