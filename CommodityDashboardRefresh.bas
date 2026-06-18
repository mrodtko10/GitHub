Attribute VB_Name = "CommodityDashboardRefresh"
Option Explicit

' ============================================================
' PASTE YOUR DASHBOARD FILE PATH HERE (or leave "" to use
' Dashboard sheet cell B4, or auto-detect from workbook folder)
' ============================================================
Private Const HTML_PATH_OVERRIDE As String = ""

' ============================================================
' MACROS TO ASSIGN TO BUTTONS:
'
'   RefreshCommodityDashboard  -> "Refresh Data" button
'        Reads all tracking sheets, rebuilds the full data block
'        in the HTML, and writes it back.  Does NOT open browser.
'
'   OpenCommodityDashboard     -> "Open Dashboard" button
'        Opens the HTML file in your default browser.
'
'   RefreshAndOpenCommodity    -> calls both in sequence
'
' Path priority:
'   1. HTML_PATH_OVERRIDE constant (if set above)
'   2. Dashboard sheet cell B4 (full path)
'   3. Workbook folder + "commodity_curve_dashboard.html"
'
' Required sheets (exact names, case-insensitive):
'   Summary          : header row 12, data rows 13+
'   BL % Complete    : header row 1, data rows 2+
'   Weekly Data      : data rows 2+ (4 rows per activity, col E = type A/B/C/D)
'   Time Sheet Roll Up: header row 2, data rows 3+
'   BL and Schedule  : data rows 2+
'   Dashboard        : B4 = HTML file path
'
' HTML file must contain:
'   // ===COMMODITY_DATA_START===
'   // ===COMMODITY_DATA_END===
' ============================================================

' ── Column indices (1-based) for each sheet ──────────────────
' Summary
Private Const SUM_DATA_START  As Long = 13   ' first data row (1-based)
Private Const SUM_COL_ACT     As Long = 3    ' C  = Activity name
Private Const SUM_COL_WHCSUB  As Long = 4    ' D  = WHC / Sub
Private Const SUM_COL_PDNP    As Long = 5    ' E  = PD or NP flag
Private Const SUM_COL_CAT     As Long = 6    ' F  = Category (contains Civil/SS/HV/PV)
Private Const SUM_COL_BLHRS   As Long = 11   ' K  = BL man-hours
Private Const SUM_COL_BLPCT   As Long = 15   ' O  = BL % complete (fraction 0-1)
Private Const SUM_COL_EAPCT   As Long = 16   ' P  = Earned % complete (fraction 0-1)
Private Const SUM_COL_VAR     As Long = 20   ' T  = Variance (BL Hrs - Actual Hrs)
Private Const SUM_DATE_ROW    As Long = 4    ' Row 4
Private Const SUM_DATE_COL    As Long = 3    ' Col C = report date

' BL % Complete
Private Const BLC_DATA_START     As Long = 2   ' first data row
Private Const BLC_COL_BUDGET     As Long = 2   ' B = total budget hrs
Private Const BLC_COL_ACT        As Long = 4   ' D = activity name
Private Const BLC_COL_FIRST_DATE As Long = 5   ' E = first date column

' Weekly Data  (4 rows per activity: type A / B / C / D)
Private Const WD_COL_ACT        As Long = 1   ' A = activity name (first row of group)
Private Const WD_COL_TYPE       As Long = 5   ' E = row type letter (A/B/C/D)
Private Const WD_COL_FIRST_VAL  As Long = 8   ' H = first weekly value column

' BL and Schedule
Private Const BLS_COL_ACT    As Long = 2    ' B = activity name
Private Const BLS_COL_FCFIN  As Long = 4    ' D = FC Finish date
Private Const BLS_COL_BLFIN  As Long = 12   ' L = BL Finish date

' Time Sheet Roll Up
Private Const TS_DATA_START  As Long = 3    ' first data row
Private Const TS_COL_DATE    As Long = 6    ' F = date
Private Const TS_COL_HRS     As Long = 9    ' I = hours
Private Const TS_COL_CC      As Long = 10   ' J = CC code (matches activity code)

Private Const HRS_PER_WEEK As Double = 60.0

' ══════════════════════════════════════════════════════════════
' REFRESH
' ══════════════════════════════════════════════════════════════
Sub RefreshCommodityDashboard()
    Dim htmlPath As String
    Dim html     As String
    Dim startPos As Long
    Dim endPos   As Long
    Dim fNum     As Integer
    Dim oneLine  As String

    ' ── Resolve HTML path ────────────────────────────────────
    htmlPath = GetHtmlPath()
    If htmlPath = "" Then
        MsgBox "No HTML path found." & vbCrLf & vbCrLf & _
               "Set cell B4 on the Dashboard sheet to the full path " & _
               "of the dashboard HTML file.", vbExclamation, "Path Not Set"
        Exit Sub
    End If

    ' ── Safety checks ────────────────────────────────────────
    Dim ext As String
    ext = LCase(Right(Trim(htmlPath), 5))
    If ext <> ".html" And LCase(Right(Trim(htmlPath), 4)) <> ".htm" Then
        MsgBox "SAFETY STOP - path does not end in .html:" & vbCrLf & htmlPath, _
               vbCritical, "Wrong File Type"
        Exit Sub
    End If
    If LCase(Trim(htmlPath)) = LCase(ThisWorkbook.FullName) Then
        MsgBox "SAFETY STOP - path points to this workbook!", vbCritical, "Wrong File"
        Exit Sub
    End If
    If Dir(htmlPath) = "" Then
        MsgBox "HTML file not found:" & vbCrLf & htmlPath, vbCritical, "File Not Found"
        Exit Sub
    End If

    ' ── Find required sheets ─────────────────────────────────
    Dim wsSum  As Worksheet, wsBLC As Worksheet
    Dim wsWD   As Worksheet, wsTS  As Worksheet
    Dim wsBLS  As Worksheet
    Set wsSum = FindSheet("summary")
    Set wsBLC = FindSheet("bl % complete")
    Set wsWD  = FindSheet("weekly data")
    Set wsTS  = FindSheet("time sheet roll up")
    Set wsBLS = FindSheet("bl and schedule")

    If wsSum Is Nothing Or wsBLC Is Nothing Then
        MsgBox "Required sheets not found.  Need 'Summary' and 'BL % Complete'.", _
               vbCritical, "Sheet Missing"
        Exit Sub
    End If

    Application.StatusBar = "Commodity Dashboard: reading BL % Complete..."
    Application.ScreenUpdating = False

    ' ══════════════════════════════════════════════════════════
    ' 1. BL % COMPLETE ─ dates, activity names, budget hrs, BL fractions
    ' ══════════════════════════════════════════════════════════
    Dim lastBlcCol As Long, lastBlcRow As Long
    lastBlcCol = wsBLC.Cells(1, wsBLC.Columns.Count).End(xlToLeft).Column
    lastBlcRow = wsBLC.Cells(wsBLC.Rows.Count, BLC_COL_ACT).End(xlUp).Row

    Dim numWeeks As Long
    numWeeks = lastBlcCol - BLC_COL_FIRST_DATE + 1
    If numWeeks < 1 Then numWeeks = 1

    Dim allDates()  As String
    Dim dateToIdx   As Object       ' Scripting.Dictionary: date string -> Long index
    Set dateToIdx = CreateObject("Scripting.Dictionary")

    ' Read date headers
    Dim lastNzIdx As Long
    lastNzIdx = 0
    Dim ci As Long, wi As Long

    ReDim allDates(0 To numWeeks - 1)
    For ci = 0 To numWeeks - 1
        Dim dv As Variant
        dv = wsBLC.Cells(1, BLC_COL_FIRST_DATE + ci).Value
        Dim ds As String
        ds = DateToISO(dv)
        allDates(ci) = ds
        If ds <> "" And Not dateToIdx.Exists(ds) Then
            dateToIdx.Add ds, ci
        End If
    Next ci

    ' Build activity arrays from BL % Complete
    ' Use Dictionary: actName -> index in actNames()
    Dim actNames()   As String
    Dim actBudget()  As Double          ' budget hours
    Dim actBLInc()   As Double          ' [actIdx, weekIdx] incremental BL hrs
    Dim actSection() As String
    Dim actIsPD()    As Boolean
    Dim actWhcSub()  As String          ' "WHC" or "Sub"
    Dim actCategory() As String         ' "PD-WHC" / "NP-WHC" / "PD-SUB"

    Dim numActs As Long
    numActs = lastBlcRow - BLC_DATA_START + 1
    If numActs < 1 Then numActs = 1

    ReDim actNames(0 To numActs - 1)
    ReDim actBudget(0 To numActs - 1)
    ReDim actBLInc(0 To numActs - 1, 0 To numWeeks - 1)
    ReDim actSection(0 To numActs - 1)
    ReDim actIsPD(0 To numActs - 1)
    ReDim actWhcSub(0 To numActs - 1)
    ReDim actCategory(0 To numActs - 1)

    Dim actIdx As Object    ' actName -> Long index
    Set actIdx = CreateObject("Scripting.Dictionary")

    Dim ai As Long
    ai = 0
    Dim r As Long
    For r = BLC_DATA_START To lastBlcRow
        Dim actName As String
        actName = Trim(CStr(wsBLC.Cells(r, BLC_COL_ACT).Value))
        If actName = "" Or actName = "0" Then GoTo NextBLCRow

        Dim budget As Double
        budget = N(wsBLC.Cells(r, BLC_COL_BUDGET))

        actNames(ai) = actName
        actBudget(ai) = budget
        If Not actIdx.Exists(actName) Then actIdx.Add actName, ai

        For ci = 0 To numWeeks - 1
            Dim frac As Double
            frac = N(wsBLC.Cells(r, BLC_COL_FIRST_DATE + ci))
            Dim blInc As Double
            blInc = frac * budget
            actBLInc(ai, ci) = blInc
            If blInc > 0 Then
                If ci > lastNzIdx Then lastNzIdx = ci
            End If
        Next ci

        ai = ai + 1
        If ai >= numActs Then Exit For
NextBLCRow:
    Next r
    numActs = ai

    ' Trim ALL_DATES to last non-zero BL week + 5 buffer
    Dim N_dates As Long
    N_dates = lastNzIdx + 6
    If N_dates > numWeeks Then N_dates = numWeeks

    ' ══════════════════════════════════════════════════════════
    ' 2. SUMMARY ─ section, PD/NP, category, percentages
    ' ══════════════════════════════════════════════════════════
    Application.StatusBar = "Commodity Dashboard: reading Summary..."

    Dim lastSumRow As Long
    lastSumRow = wsSum.Cells(wsSum.Rows.Count, SUM_COL_ACT).End(xlUp).Row

    ' Maps for Summary data: actName -> various fields
    Dim sumSection  As Object : Set sumSection  = CreateObject("Scripting.Dictionary")
    Dim sumIsPD     As Object : Set sumIsPD     = CreateObject("Scripting.Dictionary")
    Dim sumWhcSub   As Object : Set sumWhcSub   = CreateObject("Scripting.Dictionary")
    Dim sumCategory As Object : Set sumCategory = CreateObject("Scripting.Dictionary")
    Dim sumBLPct    As Object : Set sumBLPct    = CreateObject("Scripting.Dictionary")
    Dim sumEAPct    As Object : Set sumEAPct    = CreateObject("Scripting.Dictionary")
    Dim sumVar      As Object : Set sumVar      = CreateObject("Scripting.Dictionary")
    Dim sumBLHrs    As Object : Set sumBLHrs    = CreateObject("Scripting.Dictionary")

    For r = SUM_DATA_START To lastSumRow
        Dim sAct As String
        sAct = Trim(CStr(wsSum.Cells(r, SUM_COL_ACT).Value))
        If sAct = "" Then GoTo NextSumRow

        Dim sCat As String
        sCat = Trim(CStr(wsSum.Cells(r, SUM_COL_CAT).Value))
        Dim sSec As String
        sSec = CatToSection(sCat)
        If sSec = "" Then GoTo NextSumRow

        Dim sPdNp As String
        sPdNp = UCase(Trim(CStr(wsSum.Cells(r, SUM_COL_PDNP).Value)))
        Dim bIsPD As Boolean
        bIsPD = (sPdNp = "PD")

        Dim sWS As String
        sWS = LCase(Trim(CStr(wsSum.Cells(r, SUM_COL_WHCSUB).Value)))
        Dim sWhcSub As String
        If sWS = "sub" Or sWS = "subcontractor" Then
            sWhcSub = "Sub"
        Else
            sWhcSub = "WHC"
        End If

        Dim sCategoryOut As String
        If bIsPD And sWhcSub = "Sub" Then
            sCategoryOut = "PD-SUB"
        ElseIf bIsPD Then
            sCategoryOut = "PD-WHC"
        Else
            sCategoryOut = "NP-WHC"
        End If

        If Not sumSection.Exists(sAct) Then
            sumSection.Add  sAct, sSec
            sumIsPD.Add     sAct, bIsPD
            sumWhcSub.Add   sAct, sWhcSub
            sumCategory.Add sAct, sCategoryOut
            sumBLPct.Add    sAct, N(wsSum.Cells(r, SUM_COL_BLPCT))
            sumEAPct.Add    sAct, N(wsSum.Cells(r, SUM_COL_EAPCT))
            sumVar.Add      sAct, N(wsSum.Cells(r, SUM_COL_VAR))
            sumBLHrs.Add    sAct, N(wsSum.Cells(r, SUM_COL_BLHRS))
        End If
NextSumRow:
    Next r

    ' Apply Summary section/PD info back to act arrays
    For ai = 0 To numActs - 1
        Dim an As String
        an = actNames(ai)
        If sumSection.Exists(an) Then
            actSection(ai)  = CStr(sumSection(an))
            actIsPD(ai)     = CBool(sumIsPD(an))
            actWhcSub(ai)   = CStr(sumWhcSub(an))
            actCategory(ai) = CStr(sumCategory(an))
        Else
            ' Fallback: derive section from activity code
            actSection(ai)  = SectionFromCode(an)
            actIsPD(ai)     = False
            actWhcSub(ai)   = "WHC"
            actCategory(ai) = "NP-WHC"
        End If
    Next ai

    ' ══════════════════════════════════════════════════════════
    ' 3. Section PD budget totals
    ' ══════════════════════════════════════════════════════════
    Dim budOverall As Double, budCivil As Double
    Dim budPV As Double,     budSS As Double
    budOverall = 0 : budCivil = 0 : budPV = 0 : budSS = 0

    For ai = 0 To numActs - 1
        If Not actIsPD(ai) Then GoTo NextBudAct
        Dim bh As Double
        bh = actBudget(ai)
        budOverall = budOverall + bh
        Select Case actSection(ai)
            Case "Civil" : budCivil = budCivil + bh
            Case "PV"    : budPV    = budPV    + bh
            Case "SS"    : budSS    = budSS    + bh
        End Select
NextBudAct:
    Next ai
    If budOverall = 0 Then budOverall = 1
    If budCivil   = 0 Then budCivil   = 1
    If budPV      = 0 Then budPV      = 1
    If budSS      = 0 Then budSS      = 1

    ' ══════════════════════════════════════════════════════════
    ' 4. WEEKLY DATA ─ earned & spent increments per activity
    ' ══════════════════════════════════════════════════════════
    Application.StatusBar = "Commodity Dashboard: reading Weekly Data..."

    ' actEarnedInc(ai, wi) = earned hrs this week for activity ai at week wi
    Dim actEarnedInc() As Double
    ReDim actEarnedInc(0 To numActs - 1, 0 To numWeeks - 1)

    ' CC spent hours by activity name: ccSpent(actName -> array[0..N-1])
    Dim ccSpent As Object
    Set ccSpent = CreateObject("Scripting.Dictionary")

    If Not wsWD Is Nothing Then
        Dim lastWDRow As Long, lastWDCol As Long
        lastWDRow = wsWD.Cells(wsWD.Rows.Count, WD_COL_ACT).End(xlUp).Row
        lastWDCol = wsWD.Cells(1, wsWD.Columns.Count).End(xlToLeft).Column

        ' Read date headers from Weekly Data row 1
        Dim wdDateOffset As Long   ' index into ALL_DATES for first WD column
        wdDateOffset = -1
        Dim wdNumCols As Long
        wdNumCols = lastWDCol - WD_COL_FIRST_VAL + 1
        If wdNumCols < 1 Then wdNumCols = 0

        ' Map WD column offsets to ALL_DATES indices
        Dim wdColToWeek() As Long
        ReDim wdColToWeek(0 To wdNumCols - 1)
        Dim cwi As Long
        For cwi = 0 To wdNumCols - 1
            wdColToWeek(cwi) = -1
            Dim wdDv As Variant
            wdDv = wsWD.Cells(1, WD_COL_FIRST_VAL + cwi).Value
            Dim wdDs As String
            wdDs = DateToISO(wdDv)
            If wdDs <> "" And dateToIdx.Exists(wdDs) Then
                wdColToWeek(cwi) = CLng(dateToIdx(wdDs))
            End If
        Next cwi

        Dim curActName As String
        curActName = ""
        For r = 2 To lastWDRow
            ' Try to read activity name from col A
            Dim wdAct As String
            wdAct = Trim(CStr(wsWD.Cells(r, WD_COL_ACT).Value))
            If wdAct <> "" And wdAct <> "0" Then curActName = wdAct

            If curActName = "" Then GoTo NextWDRow

            Dim rowType As String
            rowType = UCase(Trim(CStr(wsWD.Cells(r, WD_COL_TYPE).Value)))

            ' Only process type C (earned) rows
            If rowType <> "C" Then GoTo NextWDRow

            Dim targetAI As Long
            targetAI = -1
            If actIdx.Exists(curActName) Then targetAI = CLng(actIdx(curActName))

            For cwi = 0 To wdNumCols - 1
                Dim wdWi As Long
                wdWi = wdColToWeek(cwi)
                If wdWi < 0 Or wdWi >= numWeeks Then GoTo NextWDCol
                Dim wdVal As Double
                wdVal = N(wsWD.Cells(r, WD_COL_FIRST_VAL + cwi))
                If wdVal <> 0 And targetAI >= 0 Then
                    actEarnedInc(targetAI, wdWi) = actEarnedInc(targetAI, wdWi) + wdVal
                End If
NextWDCol:
            Next cwi
NextWDRow:
        Next r
    End If

    ' ══════════════════════════════════════════════════════════
    ' 5. TIME SHEET ROLL UP ─ spent hours per CC code per week
    ' ══════════════════════════════════════════════════════════
    Application.StatusBar = "Commodity Dashboard: reading Time Sheet Roll Up..."

    If Not wsTS Is Nothing Then
        Dim lastTSRow As Long
        lastTSRow = wsTS.Cells(wsTS.Rows.Count, TS_COL_CC).End(xlUp).Row

        For r = TS_DATA_START To lastTSRow
            Dim tsD As Variant
            tsD = wsTS.Cells(r, TS_COL_DATE).Value
            Dim tsDateStr As String
            tsDateStr = DateToISO(tsD)
            If tsDateStr = "" Then GoTo NextTSRow

            Dim tsHrs As Double
            tsHrs = N(wsTS.Cells(r, TS_COL_HRS))
            If tsHrs <= 0 Then GoTo NextTSRow

            Dim tsCC As String
            tsCC = Trim(CStr(wsTS.Cells(r, TS_COL_CC).Value))
            If tsCC = "" Or tsCC = "-" Then GoTo NextTSRow

            ' Map date to week-ending Sunday
            Dim tsSunday As Date
            If IsDate(tsD) Then
                tsSunday = WeekEndingSunday(CDate(tsD))
            ElseIf IsNumeric(tsD) And CDbl(tsD) > 40000 Then
                tsSunday = WeekEndingSunday(CDate(CLng(tsD)))
            Else
                GoTo NextTSRow
            End If

            Dim tsWS As String
            tsWS = Format(tsSunday, "YYYY-MM-DD")
            If Not dateToIdx.Exists(tsWS) Then GoTo NextTSRow

            Dim tsWi As Long
            tsWi = CLng(dateToIdx(tsWS))
            If tsWi >= numWeeks Then GoTo NextTSRow

            ' Accumulate into ccSpent dictionary (keyed by CC code)
            If Not ccSpent.Exists(tsCC) Then
                Dim newArr() As Double
                ReDim newArr(0 To numWeeks - 1)
                ccSpent.Add tsCC, newArr
            End If
            Dim spArr() As Double
            spArr = ccSpent(tsCC)
            spArr(tsWi) = spArr(tsWi) + tsHrs
            ccSpent(tsCC) = spArr

            ' Also accumulate per-activity spent (match CC code to activity)
            If actIdx.Exists(tsCC) Then
                Dim tsAI As Long
                tsAI = CLng(actIdx(tsCC))
                ' Stored separately - we'll use ccSpent lookup in output
            End If
NextTSRow:
        Next r
    End If

    ' Build per-activity spent array (lookup from ccSpent by extracting activity code from name)
    ' actSpentInc(ai, wi): spent hrs for activity ai at week wi
    Dim actSpentInc() As Double
    ReDim actSpentInc(0 To numActs - 1, 0 To numWeeks - 1)
    For ai = 0 To numActs - 1
        Dim aCode As String
        aCode = ExtractNumericCode(actNames(ai))
        If ccSpent.Exists(aCode) Then
            Dim spA() As Double
            spA = ccSpent(aCode)
            For wi = 0 To numWeeks - 1
                actSpentInc(ai, wi) = spA(wi)
            Next wi
        ElseIf ccSpent.Exists(actNames(ai)) Then
            Dim spB() As Double
            spB = ccSpent(actNames(ai))
            For wi = 0 To numWeeks - 1
                actSpentInc(ai, wi) = spB(wi)
            Next wi
        End If
    Next ai

    ' ══════════════════════════════════════════════════════════
    ' 6. Compute cumulative BL / Earned / Spent curves
    ' ══════════════════════════════════════════════════════════
    Application.StatusBar = "Commodity Dashboard: computing S-curves..."

    ' Curves: [0..N_dates-1] for Overall, Civil, PV, SS
    Dim blOv()  As Double : ReDim blOv(0 To N_dates - 1)
    Dim eaOv()  As Double : ReDim eaOv(0 To N_dates - 1)
    Dim spOv()  As Double : ReDim spOv(0 To N_dates - 1)
    Dim blCiv() As Double : ReDim blCiv(0 To N_dates - 1)
    Dim eaCiv() As Double : ReDim eaCiv(0 To N_dates - 1)
    Dim spCiv() As Double : ReDim spCiv(0 To N_dates - 1)
    Dim blPV()  As Double : ReDim blPV(0 To N_dates - 1)
    Dim eaPV()  As Double : ReDim eaPV(0 To N_dates - 1)
    Dim spPV()  As Double : ReDim spPV(0 To N_dates - 1)
    Dim blSS()  As Double : ReDim blSS(0 To N_dates - 1)
    Dim eaSS()  As Double : ReDim eaSS(0 To N_dates - 1)
    Dim spSS()  As Double : ReDim spSS(0 To N_dates - 1)

    ' Running totals
    Dim runBLOv  As Double, runEAOv  As Double, runSPOv  As Double
    Dim runBLCiv As Double, runEACiv As Double, runSPCiv As Double
    Dim runBLPV  As Double, runEAPV  As Double, runSPPV  As Double
    Dim runBLSS  As Double, runEASS  As Double, runSPSS  As Double

    For wi = 0 To N_dates - 1
        Dim incBLOv  As Double, incEAOv  As Double, incSPOv  As Double
        Dim incBLCiv As Double, incEACiv As Double, incSPCiv As Double
        Dim incBLPV  As Double, incEAPV  As Double, incSPPV  As Double
        Dim incBLSS  As Double, incEASS  As Double, incSPSS  As Double
        incBLOv = 0 : incEAOv = 0 : incSPOv = 0
        incBLCiv = 0 : incEACiv = 0 : incSPCiv = 0
        incBLPV = 0 : incEAPV = 0 : incSPPV = 0
        incBLSS = 0 : incEASS = 0 : incSPSS = 0

        For ai = 0 To numActs - 1
            If Not actIsPD(ai) Then GoTo NextCurveAct
            Dim bl_i As Double, ea_i As Double, sp_i As Double
            If wi < numWeeks Then
                bl_i = actBLInc(ai, wi)
                ea_i = actEarnedInc(ai, wi)
                sp_i = actSpentInc(ai, wi)
            End If
            incBLOv = incBLOv + bl_i
            incEAOv = incEAOv + ea_i
            incSPOv = incSPOv + sp_i
            Select Case actSection(ai)
                Case "Civil"
                    incBLCiv = incBLCiv + bl_i
                    incEACiv = incEACiv + ea_i
                    incSPCiv = incSPCiv + sp_i
                Case "PV"
                    incBLPV = incBLPV + bl_i
                    incEAPV = incEAPV + ea_i
                    incSPPV = incSPPV + sp_i
                Case "SS"
                    incBLSS = incBLSS + bl_i
                    incEASS = incEASS + ea_i
                    incSPSS = incSPSS + sp_i
            End Select
NextCurveAct:
        Next ai

        runBLOv  = runBLOv  + incBLOv  : blOv(wi)  = RoundTo4(runBLOv  / budOverall * 100)
        runEAOv  = runEAOv  + incEAOv  : eaOv(wi)  = RoundTo4(runEAOv  / budOverall * 100)
        runSPOv  = runSPOv  + incSPOv  : spOv(wi)  = RoundTo4(runSPOv  / budOverall * 100)
        runBLCiv = runBLCiv + incBLCiv : blCiv(wi) = RoundTo4(runBLCiv / budCivil   * 100)
        runEACiv = runEACiv + incEACiv : eaCiv(wi) = RoundTo4(runEACiv / budCivil   * 100)
        runSPCiv = runSPCiv + incSPCiv : spCiv(wi) = RoundTo4(runSPCiv / budCivil   * 100)
        runBLPV  = runBLPV  + incBLPV  : blPV(wi)  = RoundTo4(runBLPV  / budPV      * 100)
        runEAPV  = runEAPV  + incEAPV  : eaPV(wi)  = RoundTo4(runEAPV  / budPV      * 100)
        runSPPV  = runSPPV  + incSPPV  : spPV(wi)  = RoundTo4(runSPPV  / budPV      * 100)
        runBLSS  = runBLSS  + incBLSS  : blSS(wi)  = RoundTo4(runBLSS  / budSS      * 100)
        runEASS  = runEASS  + incEASS  : eaSS(wi)  = RoundTo4(runEASS  / budSS      * 100)
        runSPSS  = runSPSS  + incSPSS  : spSS(wi)  = RoundTo4(runSPSS  / budSS      * 100)
    Next wi

    ' ══════════════════════════════════════════════════════════
    ' 7. Current week index
    ' ══════════════════════════════════════════════════════════
    ' Try report date from Summary first, fallback to today
    Dim reportDate As String
    Dim rd As Variant
    rd = wsSum.Cells(SUM_DATE_ROW, SUM_DATE_COL).Value
    Dim curSunday As Date
    If IsDate(rd) Then
        curSunday = WeekEndingSunday(CDate(rd))
    ElseIf IsNumeric(rd) And CDbl(rd) > 40000 Then
        curSunday = WeekEndingSunday(CDate(CLng(rd)))
    Else
        curSunday = WeekEndingSunday(Date)
    End If
    reportDate = Format(curSunday, "YYYY-MM-DD")

    Dim currentWeekIdx As Long
    currentWeekIdx = -1
    If dateToIdx.Exists(reportDate) Then
        currentWeekIdx = CLng(dateToIdx(reportDate))
    End If
    If currentWeekIdx < 0 Or currentWeekIdx >= N_dates Then
        ' Fallback: last week with earned increment
        Dim prevEA As Double
        prevEA = 0
        For wi = 0 To N_dates - 1
            If eaOv(wi) > prevEA Then
                currentWeekIdx = wi
                prevEA = eaOv(wi)
            End If
        Next wi
    End If
    If currentWeekIdx < 0 Then currentWeekIdx = 0

    Dim weekEndStr As String
    weekEndStr = allDates(currentWeekIdx)

    Dim pwIdx As Long
    pwIdx = currentWeekIdx - 1
    If pwIdx < 0 Then pwIdx = 0

    ' ══════════════════════════════════════════════════════════
    ' 8. BL AND SCHEDULE ─ bl_finish / fc_finish per activity
    ' ══════════════════════════════════════════════════════════
    Dim blsFinBL As Object : Set blsFinBL = CreateObject("Scripting.Dictionary")
    Dim blsFinFC As Object : Set blsFinFC = CreateObject("Scripting.Dictionary")

    If Not wsBLS Is Nothing Then
        Dim lastBLSRow As Long
        lastBLSRow = wsBLS.Cells(wsBLS.Rows.Count, BLS_COL_ACT).End(xlUp).Row
        For r = 2 To lastBLSRow
            Dim blsAct As String
            blsAct = Trim(CStr(wsBLS.Cells(r, BLS_COL_ACT).Value))
            If blsAct = "" Then GoTo NextBLSRow
            Dim blsFin As Variant, fcFin As Variant
            blsFin = wsBLS.Cells(r, BLS_COL_BLFIN).Value
            fcFin  = wsBLS.Cells(r, BLS_COL_FCFIN).Value
            If Not blsFinBL.Exists(blsAct) Then
                blsFinBL.Add blsAct, DateToISO(blsFin)
                blsFinFC.Add blsAct, DateToISO(fcFin)
            End If
NextBLSRow:
        Next r
    End If

    ' ══════════════════════════════════════════════════════════
    ' 9. Build WEEKLY DATA for ea_week from Weekly Data type C
    '    (current week earned increment per activity)
    ' ══════════════════════════════════════════════════════════
    ' actEarnedInc(ai, currentWeekIdx) already holds it

    ' ══════════════════════════════════════════════════════════
    ' BUILD JSON
    ' ══════════════════════════════════════════════════════════
    Application.StatusBar = "Commodity Dashboard: building JSON..."

    Dim jb As String    ' JSON data block string

    ' ── let WEEK_ENDING ──────────────────────────────────────
    jb = "let WEEK_ENDING = '" & weekEndStr & "';" & vbLf

    ' ── let SC_CURRENT_IDX ───────────────────────────────────
    jb = jb & "let SC_CURRENT_IDX = " & CStr(currentWeekIdx) & ";" & vbLf

    ' ── const ALL_DATES ──────────────────────────────────────
    Dim adParts() As String
    ReDim adParts(0 To N_dates - 1)
    For wi = 0 To N_dates - 1
        adParts(wi) = Q(allDates(wi))
    Next wi
    jb = jb & "const ALL_DATES = [" & Join(adParts, ",") & "];" & vbLf & vbLf

    ' ── const SCURVE ─────────────────────────────────────────
    jb = jb & "const SCURVE = {" & vbLf
    jb = jb & "  dates: ALL_DATES," & vbLf
    jb = jb & "  overall: {" & vbLf
    jb = jb & "    bl:     " & ArrJ(blOv,  N_dates) & "," & vbLf
    jb = jb & "    earned: " & ArrJ(eaOv,  N_dates) & "," & vbLf
    jb = jb & "    spent:  " & ArrJ(spOv,  N_dates) & vbLf
    jb = jb & "  }," & vbLf
    jb = jb & "  sections: {" & vbLf
    jb = jb & "    Civil: {" & vbLf
    jb = jb & "      bl:     " & ArrJ(blCiv, N_dates) & "," & vbLf
    jb = jb & "      earned: " & ArrJ(eaCiv, N_dates) & "," & vbLf
    jb = jb & "      spent:  " & ArrJ(spCiv, N_dates) & vbLf
    jb = jb & "    }," & vbLf
    jb = jb & "    PV: {" & vbLf
    jb = jb & "      bl:     " & ArrJ(blPV, N_dates) & "," & vbLf
    jb = jb & "      earned: " & ArrJ(eaPV, N_dates) & "," & vbLf
    jb = jb & "      spent:  " & ArrJ(spPV, N_dates) & vbLf
    jb = jb & "    }," & vbLf
    jb = jb & "    SS: {" & vbLf
    jb = jb & "      bl:     " & ArrJ(blSS, N_dates) & "," & vbLf
    jb = jb & "      earned: " & ArrJ(eaSS, N_dates) & "," & vbLf
    jb = jb & "      spent:  " & ArrJ(spSS, N_dates) & vbLf
    jb = jb & "    }" & vbLf
    jb = jb & "  }," & vbLf
    jb = jb & "  pf: {}," & vbLf
    jb = jb & "  fc: {}" & vbLf
    jb = jb & "};" & vbLf & vbLf

    ' ── const STATS_BY_DATE ──────────────────────────────────
    Application.StatusBar = "Commodity Dashboard: building STATS_BY_DATE..."
    jb = jb & "const STATS_BY_DATE = {" & vbLf
    For wi = 0 To N_dates - 1
        Dim sEAov As Double, sSPov As Double, sCPF As String
        sEAov = eaOv(wi) : sSPov = spOv(wi)
        If sSPov > 0 Then
            sCPF = J(RoundTo4(sEAov / sSPov))
        Else
            sCPF = "null"
        End If

        Dim civEA As Double, civSP As Double, civCPF As String
        civEA = eaCiv(wi) : civSP = spCiv(wi)
        If civSP > 0 Then civCPF = J(RoundTo4(civEA / civSP)) Else civCPF = "null"

        Dim pvEA As Double, pvSP As Double, pvCPF As String
        pvEA = eaPV(wi) : pvSP = spPV(wi)
        If pvSP > 0 Then pvCPF = J(RoundTo4(pvEA / pvSP)) Else pvCPF = "null"

        Dim ssEA As Double, ssSP As Double, ssCPF As String
        ssEA = eaSS(wi) : ssSP = spSS(wi)
        If ssSP > 0 Then ssCPF = J(RoundTo4(ssEA / ssSP)) Else ssCPF = "null"

        Dim comma As String
        If wi < N_dates - 1 Then comma = "," Else comma = ""

        jb = jb & "  " & Q(allDates(wi)) & ":{" & vbLf
        jb = jb & "    " & Q("Overall") & ":{bl:" & J(blOv(wi))  & ",earned:" & J(sEAov)  & ",spent:" & J(sSPov)  & ",cumu_pf:" & sCPF  & "}," & vbLf
        jb = jb & "    " & Q("Civil")   & ":{bl:" & J(blCiv(wi)) & ",earned:" & J(civEA)  & ",spent:" & J(civSP)  & ",cumu_pf:" & civCPF & "}," & vbLf
        jb = jb & "    " & Q("PV")      & ":{bl:" & J(blPV(wi))  & ",earned:" & J(pvEA)   & ",spent:" & J(pvSP)   & ",cumu_pf:" & pvCPF  & "}," & vbLf
        jb = jb & "    " & Q("SS")      & ":{bl:" & J(blSS(wi))  & ",earned:" & J(ssEA)   & ",spent:" & J(ssSP)   & ",cumu_pf:" & ssCPF  & "}," & vbLf
        jb = jb & "  }" & comma & vbLf
    Next wi
    jb = jb & "};" & vbLf & vbLf

    ' ── const FTE_BY_ACTIVITY ────────────────────────────────
    Application.StatusBar = "Commodity Dashboard: building FTE_BY_ACTIVITY..."
    jb = jb & "const FTE_BY_ACTIVITY = {" & vbLf

    Dim zeroArr() As Double
    ReDim zeroArr(0 To N_dates - 1)
    Dim zeroJson As String
    zeroJson = ArrJ(zeroArr, N_dates)

    Dim firstFteAct As Boolean
    firstFteAct = True

    For ai = 0 To numActs - 1
        If actSection(ai) = "" Then GoTo NextFteAct

        Dim fteBL()  As Double : ReDim fteBL(0 To N_dates - 1)
        Dim fteEA()  As Double : ReDim fteEA(0 To N_dates - 1)
        Dim fteSP()  As Double : ReDim fteSP(0 To N_dates - 1)
        For wi = 0 To N_dates - 1
            If wi < numWeeks Then
                fteBL(wi) = RoundTo4(actBLInc(ai, wi)     / HRS_PER_WEEK)
                fteEA(wi) = RoundTo4(actEarnedInc(ai, wi) / HRS_PER_WEEK)
                fteSP(wi) = RoundTo4(actSpentInc(ai, wi)  / HRS_PER_WEEK)
            End If
        Next wi

        Dim fteComma As String
        If Not firstFteAct Then fteComma = "," Else fteComma = ""
        firstFteAct = False

        Dim safeActName As String
        safeActName = Replace(actNames(ai), """", "\""")

        jb = jb & fteComma & Q(safeActName) & ":{" & vbLf
        jb = jb & "  " & Q("section") & ":" & Q(actSection(ai)) & "," & vbLf
        jb = jb & "  " & Q("bl")      & ":" & ArrJ(fteBL, N_dates) & "," & vbLf
        jb = jb & "  " & Q("earned")  & ":" & ArrJ(fteEA, N_dates) & "," & vbLf
        jb = jb & "  " & Q("spent")   & ":" & ArrJ(fteSP, N_dates) & "," & vbLf
        jb = jb & "  " & Q("fc")      & ":" & zeroJson & vbLf
        jb = jb & "}" & vbLf
NextFteAct:
    Next ai
    jb = jb & "};" & vbLf & vbLf

    ' ── const FTE (aggregated by section) ────────────────────
    Application.StatusBar = "Commodity Dashboard: building FTE..."
    ' Aggregate BL FTE, earned FTE, spent FTE by section
    Dim fteOvBL()  As Double : ReDim fteOvBL(0 To N_dates - 1)
    Dim fteOvEA()  As Double : ReDim fteOvEA(0 To N_dates - 1)
    Dim fteOvSP()  As Double : ReDim fteOvSP(0 To N_dates - 1)
    Dim fteCivBL() As Double : ReDim fteCivBL(0 To N_dates - 1)
    Dim fteCivEA() As Double : ReDim fteCivEA(0 To N_dates - 1)
    Dim fteCivSP() As Double : ReDim fteCivSP(0 To N_dates - 1)
    Dim ftePVBL()  As Double : ReDim ftePVBL(0 To N_dates - 1)
    Dim ftePVEA()  As Double : ReDim ftePVEA(0 To N_dates - 1)
    Dim ftePVSP()  As Double : ReDim ftePVSP(0 To N_dates - 1)
    Dim fteSSBL()  As Double : ReDim fteSSBL(0 To N_dates - 1)
    Dim fteSSEA()  As Double : ReDim fteSSEA(0 To N_dates - 1)
    Dim fteSSESP() As Double : ReDim fteSSESP(0 To N_dates - 1)

    For ai = 0 To numActs - 1
        If actSection(ai) = "" Then GoTo NextFteAgg
        For wi = 0 To N_dates - 1
            If wi >= numWeeks Then GoTo NextFteWi
            Dim fbl As Double, fea As Double, fsp As Double
            fbl = actBLInc(ai, wi)     / HRS_PER_WEEK
            fea = actEarnedInc(ai, wi) / HRS_PER_WEEK
            fsp = actSpentInc(ai, wi)  / HRS_PER_WEEK
            fteOvBL(wi) = fteOvBL(wi) + fbl
            fteOvEA(wi) = fteOvEA(wi) + fea
            fteOvSP(wi) = fteOvSP(wi) + fsp
            Select Case actSection(ai)
                Case "Civil"
                    fteCivBL(wi) = fteCivBL(wi) + fbl
                    fteCivEA(wi) = fteCivEA(wi) + fea
                    fteCivSP(wi) = fteCivSP(wi) + fsp
                Case "PV"
                    ftePVBL(wi)  = ftePVBL(wi)  + fbl
                    ftePVEA(wi)  = ftePVEA(wi)  + fea
                    ftePVSP(wi)  = ftePVSP(wi)  + fsp
                Case "SS"
                    fteSSBL(wi)  = fteSSBL(wi)  + fbl
                    fteSSEA(wi)  = fteSSEA(wi)  + fea
                    fteSSESP(wi) = fteSSESP(wi) + fsp
            End Select
NextFteWi:
        Next wi
NextFteAgg:
    Next ai

    jb = jb & "const FTE = {" & vbLf
    jb = jb & "  Overall: {bl:" & ArrJRound(fteOvBL, N_dates, 4)  & ",earned:" & ArrJRound(fteOvEA, N_dates, 4)  & ",spent:" & ArrJRound(fteOvSP, N_dates, 4)  & ",fc:" & zeroJson & "}," & vbLf
    jb = jb & "  Civil:   {bl:" & ArrJRound(fteCivBL, N_dates, 4) & ",earned:" & ArrJRound(fteCivEA, N_dates, 4) & ",spent:" & ArrJRound(fteCivSP, N_dates, 4) & ",fc:" & zeroJson & "}," & vbLf
    jb = jb & "  PV:      {bl:" & ArrJRound(ftePVBL, N_dates, 4)  & ",earned:" & ArrJRound(ftePVEA, N_dates, 4)  & ",spent:" & ArrJRound(ftePVSP, N_dates, 4)  & ",fc:" & zeroJson & "}," & vbLf
    jb = jb & "  SS:      {bl:" & ArrJRound(fteSSBL, N_dates, 4)  & ",earned:" & ArrJRound(fteSSEA, N_dates, 4)  & ",spent:" & ArrJRound(fteSSESP, N_dates, 4) & ",fc:" & zeroJson & "}" & vbLf
    jb = jb & "};" & vbLf & vbLf

    ' ── const PD_DONUT ───────────────────────────────────────
    Application.StatusBar = "Commodity Dashboard: building PD_DONUT..."
    jb = jb & "const PD_DONUT = {" & vbLf
    For wi = 0 To N_dates - 1
        Dim comPD As String
        If wi < N_dates - 1 Then comPD = "," Else comPD = ""

        Dim pdCivBL As Double, pdCivEA As Double, pdCivSP As Double
        pdCivBL = RoundTo0(blCiv(wi) / 100 * budCivil)
        pdCivEA = RoundTo0(eaCiv(wi) / 100 * budCivil)
        pdCivSP = RoundTo0(spCiv(wi) / 100 * budCivil)

        Dim pdPVBL As Double, pdPVEA As Double, pdPVSP As Double
        pdPVBL = RoundTo0(blPV(wi) / 100 * budPV)
        pdPVEA = RoundTo0(eaPV(wi) / 100 * budPV)
        pdPVSP = RoundTo0(spPV(wi) / 100 * budPV)

        Dim pdSSBL As Double, pdSSEA As Double, pdSSSP As Double
        pdSSBL = RoundTo0(blSS(wi) / 100 * budSS)
        pdSSEA = RoundTo0(eaSS(wi) / 100 * budSS)
        pdSSSP = RoundTo0(spSS(wi) / 100 * budSS)

        jb = jb & "  " & Q(allDates(wi)) & ":{"
        jb = jb & "Civil:{bl:" & J(pdCivBL) & ",earned:" & J(pdCivEA) & ",spent:" & J(pdCivSP) & ",pct_rem:" & J(RoundTo2(100 - blCiv(wi))) & "},"
        jb = jb & "PV:{bl:" & J(pdPVBL) & ",earned:" & J(pdPVEA) & ",spent:" & J(pdPVSP) & ",pct_rem:" & J(RoundTo2(100 - blPV(wi))) & "},"
        jb = jb & "SS:{bl:" & J(pdSSBL) & ",earned:" & J(pdSSEA) & ",spent:" & J(pdSSSP) & ",pct_rem:" & J(RoundTo2(100 - blSS(wi))) & "}"
        jb = jb & "}" & comPD & vbLf
    Next wi
    jb = jb & "};" & vbLf

    ' ── let MANPOWER ─────────────────────────────────────────
    Application.StatusBar = "Commodity Dashboard: building MANPOWER..."
    jb = jb & "let MANPOWER = ["
    Dim firstMP As Boolean
    firstMP = True

    For ai = 0 To numActs - 1
        If actSection(ai) = "" Then GoTo NextMPAct

        Dim mpBLPW  As Double, mpSPPW  As Double
        Dim mpBLCW  As Double, mpSPCW  As Double
        If pwIdx < numWeeks Then
            mpBLPW = RoundTo2(actBLInc(ai, pwIdx)     / HRS_PER_WEEK)
            mpSPPW = RoundTo2(actSpentInc(ai, pwIdx)  / HRS_PER_WEEK)
        End If
        If currentWeekIdx < numWeeks Then
            mpBLCW = RoundTo2(actBLInc(ai, currentWeekIdx)    / HRS_PER_WEEK)
            mpSPCW = RoundTo2(actSpentInc(ai, currentWeekIdx) / HRS_PER_WEEK)
        End If

        If Not firstMP Then jb = jb & ","
        firstMP = False

        Dim mpAct As String
        mpAct = Replace(actNames(ai), """", "\""")

        jb = jb & "{" & Q("activity") & ":" & Q(mpAct) & ","
        jb = jb & Q("whc_sub")    & ":" & Q(actWhcSub(ai)) & ","
        jb = jb & Q("category")   & ":" & Q(actCategory(ai)) & ","
        jb = jb & Q("section")    & ":" & Q(actSection(ai)) & ","
        jb = jb & Q("pw_bl_fte")  & ":" & J(mpBLPW) & ","
        jb = jb & Q("pw_act_fte") & ":" & J(mpSPPW) & ","
        jb = jb & Q("pw_avg_day") & ":null,"
        jb = jb & Q("cw_bl_fte")  & ":" & J(mpBLCW) & ","
        jb = jb & Q("cw_act_fte") & ":" & J(mpSPCW) & ","
        jb = jb & Q("cw_avg_day") & ":null}"
NextMPAct:
    Next ai
    jb = jb & "];" & vbLf

    ' ── let ACTIVITIES ───────────────────────────────────────
    Application.StatusBar = "Commodity Dashboard: building ACTIVITIES..."
    jb = jb & "let ACTIVITIES = ["
    Dim firstAct As Boolean
    firstAct = True

    For ai = 0 To numActs - 1
        Dim aName As String
        aName = actNames(ai)
        If actSection(ai) = "" Then GoTo NextActRow
        If Not sumSection.Exists(aName) Then GoTo NextActRow

        Dim aBlPct   As Double, aEaPct   As Double, aVar As Double
        aBlPct = RoundTo2(CDbl(sumBLPct(aName)) * 100)
        aEaPct = RoundTo2(CDbl(sumEAPct(aName)) * 100)
        aVar   = RoundTo1(CDbl(sumVar(aName)))

        Dim aEaWeek As Double
        If currentWeekIdx < numWeeks Then
            aEaWeek = RoundTo1(actEarnedInc(ai, currentWeekIdx))
        End If

        Dim aBlFin As String, aFcFin As String, aFinVar As String
        aBlFin = "null" : aFcFin = "null" : aFinVar = "null"
        If blsFinBL.Exists(aName) Then
            Dim blFinStr As String
            blFinStr = CStr(blsFinBL(aName))
            If blFinStr <> "" Then aBlFin = Q(blFinStr)
            Dim fcFinStr As String
            fcFinStr = CStr(blsFinFC(aName))
            If fcFinStr <> "" Then aFcFin = Q(fcFinStr)
            If aBlFin <> "null" And aFcFin <> "null" Then
                Dim blDate As Date, fcDate As Date
                On Error Resume Next
                blDate = CDate(blFinStr)
                fcDate = CDate(fcFinStr)
                On Error GoTo 0
                If IsDate(blDate) And IsDate(fcDate) Then
                    aFinVar = CStr(DateDiff("d", blDate, fcDate))
                End If
            End If
        End If

        If Not firstAct Then jb = jb & ","
        firstAct = False

        Dim acEsc As String
        acEsc = Replace(aName, """", "\""")

        jb = jb & "{" & Q("activity") & ":" & Q(acEsc) & ","
        jb = jb & Q("whc_sub")    & ":" & Q(actWhcSub(ai)) & ","
        jb = jb & Q("category")   & ":" & Q(actCategory(ai)) & ","
        jb = jb & Q("section")    & ":" & Q(actSection(ai)) & ","
        jb = jb & Q("bl_finish")  & ":" & aBlFin & ","
        jb = jb & Q("fc_finish")  & ":" & aFcFin & ","
        jb = jb & Q("fin_var")    & ":" & aFinVar & ","
        jb = jb & Q("bl_pct")     & ":" & J(aBlPct) & ","
        jb = jb & Q("ea_pct")     & ":" & J(aEaPct) & ","
        jb = jb & Q("bl_act_var") & ":" & J(aVar) & ","
        jb = jb & Q("ea_week")    & ":" & J(aEaWeek) & "}"
NextActRow:
    Next ai
    jb = jb & "];"

    ' ══════════════════════════════════════════════════════════
    ' REPLACE DATA BLOCK IN HTML
    ' ══════════════════════════════════════════════════════════
    Application.StatusBar = "Commodity Dashboard: reading HTML..."

    fNum = FreeFile
    html = ""
    Open htmlPath For Input As #fNum
    Dim oneLine2 As String
    Do While Not EOF(fNum)
        Line Input #fNum, oneLine2
        html = html & oneLine2 & vbLf
    Loop
    Close #fNum

    Dim startMark As String, endMark As String
    startMark = "// ===COMMODITY_DATA_START==="
    endMark   = "// ===COMMODITY_DATA_END==="

    startPos = InStr(html, startMark)
    If startPos = 0 Then
        MsgBox "Marker '" & startMark & "' not found in the HTML file." & vbCrLf & _
               htmlPath, vbCritical, "Marker Not Found"
        GoTo CleanExit
    End If

    ' Find end of the startMark line
    Dim afterStart As Long
    afterStart = InStr(startPos, html, vbLf)
    If afterStart = 0 Then afterStart = Len(html)
    afterStart = afterStart + 1  ' first char after the start marker line

    endPos = InStr(afterStart, html, endMark)
    If endPos = 0 Then
        MsgBox "Marker '" & endMark & "' not found in the HTML file.", vbCritical, "Marker Not Found"
        GoTo CleanExit
    End If

    ' Build replacement: keep start marker line + new data + end marker line
    Dim beforeBlock As String, afterBlock As String
    beforeBlock = Left(html, afterStart - 1)  ' up through the \n after start marker
    afterBlock  = Mid(html, endPos)           ' from "// ===COMMODITY_DATA_END===" to end

    html = beforeBlock & jb & vbLf & afterBlock

    Application.StatusBar = "Commodity Dashboard: writing HTML..."
    fNum = FreeFile
    Open htmlPath For Output As #fNum
    Print #fNum, html
    Close #fNum

    Dim actCount As Long
    actCount = 0
    For ai = 0 To numActs - 1
        If actSection(ai) <> "" And sumSection.Exists(actNames(ai)) Then actCount = actCount + 1
    Next ai

    MsgBox "Commodity Dashboard refreshed!" & vbCrLf & vbCrLf & _
           "Week ending : " & weekEndStr & vbCrLf & _
           "Week index  : " & currentWeekIdx & vbCrLf & _
           "Activities  : " & actCount & vbCrLf & _
           "Date range  : " & allDates(0) & " -> " & allDates(N_dates - 1) & vbCrLf & vbCrLf & _
           "Click 'Open Dashboard' or refresh your browser.", _
           vbInformation, "Refresh Complete"

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

' ── OPEN ──────────────────────────────────────────────────────
Sub OpenCommodityDashboard()
    Dim htmlPath As String
    htmlPath = GetHtmlPath()
    If htmlPath = "" Or Dir(htmlPath) = "" Then
        MsgBox "HTML file not found." & vbCrLf & _
               "Check that cell B4 on the Dashboard sheet has the correct path.", _
               vbCritical, "File Not Found"
        Exit Sub
    End If
    Dim url As String
    url = "file:///" & Replace(Replace(htmlPath, "\", "/"), " ", "%20")
    On Error Resume Next
    ThisWorkbook.FollowHyperlink Address:=url, NewWindow:=True
    On Error GoTo 0
End Sub

' ── REFRESH + OPEN ────────────────────────────────────────────
Sub RefreshAndOpenCommodity()
    RefreshCommodityDashboard
    OpenCommodityDashboard
End Sub

' ══════════════════════════════════════════════════════════════
' PATH RESOLVER
' ══════════════════════════════════════════════════════════════
Private Function GetHtmlPath() As String
    If HTML_PATH_OVERRIDE <> "" Then
        GetHtmlPath = HTML_PATH_OVERRIDE
        Exit Function
    End If

    Dim dashWs As Worksheet
    Dim p As String
    On Error Resume Next
    Set dashWs = FindSheet("dashboard")
    On Error GoTo 0
    If Not dashWs Is Nothing Then
        On Error Resume Next
        p = Trim(CStr(dashWs.Cells(4, 2).Value))
        On Error GoTo 0
        If p <> "" And p <> "0" Then
            GetHtmlPath = p
            Exit Function
        End If
    End If

    If ThisWorkbook.Path <> "" Then
        GetHtmlPath = ThisWorkbook.Path & "\commodity_curve_dashboard.html"
    End If
End Function

' ══════════════════════════════════════════════════════════════
' HELPERS
' ══════════════════════════════════════════════════════════════
Private Function FindSheet(nameLower As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If LCase(ws.Name) = nameLower Then
            Set FindSheet = ws
            Exit Function
        End If
    Next ws
    Set FindSheet = Nothing
End Function

Private Function N(cell As Object) As Double
    On Error Resume Next
    If IsNumeric(cell.Value) Then N = CDbl(cell.Value) Else N = 0
    On Error GoTo 0
End Function

Private Function Q(s As String) As String
    Q = """" & s & """"
End Function

Private Function J(v As Double) As String
    J = Replace(CStr(CDec(v)), ",", ".")
End Function

Private Function RoundTo4(v As Double) As Double
    RoundTo4 = Int(v * 10000 + 0.5) / 10000
End Function

Private Function RoundTo2(v As Double) As Double
    RoundTo2 = Int(v * 100 + 0.5) / 100
End Function

Private Function RoundTo1(v As Double) As Double
    RoundTo1 = Int(v * 10 + 0.5) / 10
End Function

Private Function RoundTo0(v As Double) As Double
    RoundTo0 = Int(v + 0.5)
End Function

Private Function ArrJ(vals() As Double, n As Long) As String
    Dim parts() As String
    ReDim parts(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1
        parts(i) = J(vals(i))
    Next i
    ArrJ = "[" & Join(parts, ",") & "]"
End Function

Private Function ArrJRound(vals() As Double, n As Long, decimals As Long) As String
    Dim parts() As String
    ReDim parts(0 To n - 1)
    Dim i As Long
    Dim factor As Double
    factor = 10 ^ decimals
    For i = 0 To n - 1
        parts(i) = J(Int(vals(i) * factor + 0.5) / factor)
    Next i
    ArrJRound = "[" & Join(parts, ",") & "]"
End Function

Private Function DateToISO(v As Variant) As String
    On Error GoTo BadDate
    If IsEmpty(v) Or IsNull(v) Or v = "" Or v = 0 Then DateToISO = "" : Exit Function
    Dim d As Date
    If IsDate(v) Then
        d = CDate(v)
    ElseIf IsNumeric(v) And CDbl(v) > 40000 Then
        d = CDate(CLng(v))
    Else
        DateToISO = "" : Exit Function
    End If
    DateToISO = Format(d, "YYYY-MM-DD")
    Exit Function
BadDate:
    DateToISO = ""
End Function

Private Function WeekEndingSunday(d As Date) As Date
    ' Return the Sunday on or after d (Sunday=1, Monday=2 ... Saturday=7 with vbSunday)
    Dim dow As Integer
    dow = Weekday(d, vbSunday)   ' 1=Sun, 2=Mon, ..., 7=Sat
    WeekEndingSunday = d + ((8 - dow) Mod 7)
End Function

Private Function CatToSection(cat As String) As String
    ' Derive section from the category / description text in Summary col F
    If InStr(cat, "Civil") > 0 Then
        CatToSection = "Civil"
    ElseIf InStr(UCase(cat), "SS") > 0 Or InStr(cat, "HV") > 0 Then
        CatToSection = "SS"
    ElseIf Left(UCase(Trim(cat)), 2) = "PV" Then
        CatToSection = "PV"
    Else
        CatToSection = ""
    End If
End Function

Private Function SectionFromCode(actName As String) As String
    ' Derive section from the activity code prefix as a fallback
    Dim code As String
    code = actName
    If Left(code, 2) = "2-" Then
        SectionFromCode = "SS"
    ElseIf InStr(code, "-41") > 0 Then
        SectionFromCode = "Civil"
    Else
        SectionFromCode = "PV"
    End If
End Function

Private Function ExtractNumericCode(actName As String) As String
    ' Extract leading numeric code: "0-4100-Silt Fence" -> "0-4100"
    ' "0-4005-00 PreCon FOE" -> "0-4005-00"
    Dim firstToken As String
    Dim spPos As Long
    spPos = InStr(actName, " ")
    If spPos > 0 Then
        firstToken = Left(actName, spPos - 1)
    Else
        firstToken = actName
    End If

    Dim parts() As String
    parts = Split(firstToken, "-")
    Dim codeParts() As String
    ReDim codeParts(UBound(parts))
    Dim ci As Long, cnt As Long
    cnt = 0
    For ci = 0 To UBound(parts)
        If IsNumeric(parts(ci)) Then
            codeParts(cnt) = parts(ci)
            cnt = cnt + 1
        Else
            Exit For
        End If
    Next ci
    If cnt >= 2 Then
        ReDim Preserve codeParts(0 To cnt - 1)
        ExtractNumericCode = Join(codeParts, "-")
    Else
        ExtractNumericCode = actName
    End If
End Function
