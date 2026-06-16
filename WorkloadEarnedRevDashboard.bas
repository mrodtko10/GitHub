Attribute VB_Name = "WorkloadEarnedRevDashboard"
Option Explicit

' ============================================================
' PASTE YOUR DASHBOARD FILE PATH HERE (include the filename):
'   Example: "C:\Users\You\Documents\Earned_Revenue_Dashboard.html"
' Leave it as "" to auto-detect from Dashboard sheet cell B4.
' ============================================================
Private Const HTML_PATH_OVERRIDE As String = ""

' ============================================================
' MACROS TO ASSIGN TO BUTTONS:
'
'   RefreshWorkloadDashboard  -> "Refresh Data" button
'        Reads Sheet1 and rewrites the Earned Revenue Dashboard HTML
'        with the latest ER data. Does NOT open the browser.
'
'   OpenWorkloadDashboard     -> "Open Dashboard" button
'        Opens the HTML file in your default browser.
'
'   RefreshAndOpenWorkload    -> calls both in sequence
'
' Path priority:
'   1. HTML_PATH_OVERRIDE constant above (if set)
'   2. Dashboard sheet cell B4 (full path)
'   3. Workbook folder + "Earned_Revenue_Dashboard.html"
' ============================================================

' ── REFRESH ──────────────────────────────────────────────────────────────────
Sub RefreshWorkloadDashboard()

    Dim wsData   As Worksheet
    Dim htmlPath As String
    Dim html     As String
    Dim startPos As Long
    Dim endPos   As Long
    Dim i        As Long
    Dim col      As Long

    ' ── Resolve HTML path ────────────────────────────────────────────────────
    htmlPath = GetHtmlPath()
    If htmlPath = "" Then
        MsgBox "No HTML path found." & vbCrLf & vbCrLf & _
               "Please paste the full path to the dashboard HTML file" & vbCrLf & _
               "into cell B4 on the Dashboard sheet.", _
               vbExclamation, "Path Not Set"
        Exit Sub
    End If

    ' ── Safety checks ────────────────────────────────────────────────────────
    If LCase(Right(Trim(htmlPath), 5)) <> ".html" And _
       LCase(Right(Trim(htmlPath), 4)) <> ".htm" Then
        MsgBox "SAFETY STOP – path does not end in .html:" & vbCrLf & htmlPath, _
               vbCritical, "Wrong File Type"
        Exit Sub
    End If

    If LCase(Trim(htmlPath)) = LCase(ThisWorkbook.FullName) Then
        MsgBox "SAFETY STOP – path points to this workbook!", vbCritical, "Wrong File"
        Exit Sub
    End If

    If Dir(htmlPath) = "" Then
        MsgBox "HTML file not found at:" & vbCrLf & htmlPath, vbCritical, "File Not Found"
        Exit Sub
    End If

    ' ── Find Sheet1 ──────────────────────────────────────────────────────────
    Set wsData = Nothing
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If LCase(ws.Name) = "sheet1" Then
            Set wsData = ws
            Exit For
        End If
    Next ws
    If wsData Is Nothing Then
        MsgBox "Could not find 'Sheet1'.", vbCritical
        Exit Sub
    End If

    wsData.Calculate

    ' ── Build months JSON (row 2 = type, row 3 = label, cols 10-57 = 48 months) ─
    Const FIRST_MONTH_COL As Long = 10
    Const LAST_MONTH_COL  As Long = 57
    Const NUM_MONTHS      As Long = 48

    Dim monthLabels(47) As String
    Dim monthTypes(47)  As String

    For col = FIRST_MONTH_COL To LAST_MONTH_COL
        Dim idx As Long
        idx = col - FIRST_MONTH_COL
        monthLabels(idx) = Trim(CStr(wsData.Cells(3, col).Value))
        Dim rawType As String
        rawType = Trim(CStr(wsData.Cells(2, col).Value))
        If InStr(LCase(rawType), "actual") > 0 Then
            monthTypes(idx) = "Actual"
        Else
            monthTypes(idx) = "Forecast"
        End If
    Next col

    ' Build months JSON array
    Dim monthsJson As String
    monthsJson = "["
    Dim mi As Long
    For mi = 0 To NUM_MONTHS - 1
        If mi > 0 Then monthsJson = monthsJson & ","
        monthsJson = monthsJson & "{"
        monthsJson = monthsJson & Q("label") & ":" & Q(monthLabels(mi)) & ","
        monthsJson = monthsJson & Q("type")  & ":" & Q(monthTypes(mi))
        monthsJson = monthsJson & "}"
    Next mi
    monthsJson = monthsJson & "]"

    ' ── Build projects JSON (rows 4 to 25001) ────────────────────────────────
    ' Column mapping (Sheet1):
    '   A(1)  = Project Number (code)
    '   B(2)  = Client
    '   C(3)  = Status ("Secured" / "Anticipated" / "Possible")
    '   D(4)  = Total Contract Value (tcv)
    '   E(5)  = Earned to Date (earned)
    '   F(6)  = Backlog Amount (backlog)
    '   J(10) - BE(57) = monthly[0..47] earned revenue values (48 months)

    Dim projectsJson As String
    projectsJson = "["
    Dim firstRec As Boolean
    firstRec = True

    For i = 4 To 25001
        Dim projCode As String
        projCode = Trim(CStr(wsData.Cells(i, 1).Value))

        ' Skip blank, header-echo, total, placeholder rows
        If projCode = "" Then GoTo SkipRow
        If projCode = "Project Number" Then GoTo SkipRow
        If projCode = "Total" Then GoTo SkipRow
        If projCode = "New" Or projCode = "0" Then GoTo SkipRow

        Dim vClient  As String
        Dim vStatus  As String
        vClient = Trim(CStr(wsData.Cells(i, 2).Value))
        vStatus = Trim(CStr(wsData.Cells(i, 3).Value))

        Dim vTcv     As Double
        Dim vEarned  As Double
        Dim vBacklog As Double
        vTcv     = N(wsData.Cells(i, 4))
        vEarned  = N(wsData.Cells(i, 5))
        vBacklog = N(wsData.Cells(i, 6))

        ' Read 48 monthly values and compute startIdx / endIdx
        Dim monthly(47) As Double
        Dim firstNZ     As Long
        Dim lastNZ      As Long
        firstNZ = -1
        lastNZ  = -1

        Dim mc As Long
        For mc = 0 To 47
            monthly(mc) = N(wsData.Cells(i, FIRST_MONTH_COL + mc))
            If Abs(monthly(mc)) > 0.001 Then
                If firstNZ = -1 Then firstNZ = mc
                lastNZ = mc
            End If
        Next mc

        ' Escape strings for JSON
        projCode = Replace(projCode, """", "\""")
        vClient  = Replace(vClient,  """", "\""")
        vStatus  = Replace(vStatus,  """", "\""")

        ' startIdx / endIdx (null if no non-zero values)
        Dim startIdxStr As String
        Dim endIdxStr   As String
        If firstNZ = -1 Then
            startIdxStr = "null"
            endIdxStr   = "null"
        Else
            startIdxStr = CStr(firstNZ)
            endIdxStr   = CStr(lastNZ)
        End If

        ' Build monthly array JSON
        Dim mArr As String
        mArr = "["
        For mc = 0 To 47
            If mc > 0 Then mArr = mArr & ","
            mArr = mArr & J(monthly(mc))
        Next mc
        mArr = mArr & "]"

        ' Build record
        Dim rec As String
        rec = "{"
        rec = rec & Q("code")     & ":" & Q(projCode)   & ","
        rec = rec & Q("client")   & ":" & Q(vClient)    & ","
        rec = rec & Q("status")   & ":" & Q(vStatus)    & ","
        rec = rec & Q("tcv")      & ":" & J(vTcv)       & ","
        rec = rec & Q("earned")   & ":" & J(vEarned)    & ","
        rec = rec & Q("backlog")  & ":" & J(vBacklog)   & ","
        rec = rec & Q("startIdx") & ":" & startIdxStr   & ","
        rec = rec & Q("endIdx")   & ":" & endIdxStr     & ","
        rec = rec & Q("monthly")  & ":" & mArr
        rec = rec & "}"

        If Not firstRec Then projectsJson = projectsJson & ","
        projectsJson = projectsJson & rec
        firstRec = False
SkipRow:
    Next i
    projectsJson = projectsJson & "]"

    ' ── Assemble full ER object ───────────────────────────────────────────────
    Dim erJson As String
    erJson = "{"
    erJson = erJson & Q("months")   & ":" & monthsJson   & ","
    erJson = erJson & Q("projects") & ":" & projectsJson
    erJson = erJson & "}"

    ' ── Read HTML file ───────────────────────────────────────────────────────
    Dim fNum    As Integer
    Dim oneLine As String
    fNum = FreeFile
    html = ""
    Open htmlPath For Input As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, oneLine
        html = html & oneLine & vbLf
    Loop
    Close #fNum

    ' ── Replace const ER = { ... } ───────────────────────────────────────────
    startPos = InStr(html, "const ER = {")
    If startPos = 0 Then
        MsgBox "Marker 'const ER = {' not found in the HTML file." & vbCrLf & _
               "Is this the correct file?" & vbCrLf & htmlPath, _
               vbCritical, "Marker Not Found"
        Exit Sub
    End If

    ' End marker: the next 'const ' declaration after the ER block
    endPos = InStr(startPos + 12, html, vbLf & "const ")
    If endPos = 0 Then
        MsgBox "Could not find the end of the ER data block.", vbCritical
        Exit Sub
    End If

    html = Left(html, startPos - 1) & "const ER = " & erJson & ";" & Mid(html, endPos)

    ' ── Write updated HTML ───────────────────────────────────────────────────
    fNum = FreeFile
    Open htmlPath For Output As #fNum
    Print #fNum, html
    Close #fNum

    Dim recCount As Long
    recCount = (Len(projectsJson) - Len(Replace(projectsJson, "},{", ""))) + 1
    If projectsJson = "[]" Then recCount = 0

    MsgBox "Earned Revenue Dashboard refreshed!" & vbCrLf & _
           recCount & " projects written." & vbCrLf & vbCrLf & _
           "Click 'Open Dashboard' or refresh your browser.", _
           vbInformation, "Refresh Complete"
End Sub

' ── OPEN ─────────────────────────────────────────────────────────────────────
Sub OpenWorkloadDashboard()
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

' ── REFRESH + OPEN ───────────────────────────────────────────────────────────
Sub RefreshAndOpenWorkload()
    RefreshWorkloadDashboard
    OpenWorkloadDashboard
End Sub

' ── Path resolver ─────────────────────────────────────────────────────────────
' Priority: 1) HTML_PATH_OVERRIDE, 2) Dashboard!B4, 3) workbook folder
Private Function GetHtmlPath() As String
    If HTML_PATH_OVERRIDE <> "" Then
        GetHtmlPath = HTML_PATH_OVERRIDE
        Exit Function
    End If

    Dim dashWs As Worksheet
    Dim p      As String
    On Error Resume Next
    Set dashWs = ThisWorkbook.Sheets("Dashboard")
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
        GetHtmlPath = ThisWorkbook.Path & "\Earned_Revenue_Dashboard.html"
    End If
End Function

' ── Helpers ──────────────────────────────────────────────────────────────────
Private Function N(cell As Object) As Double
    If IsNumeric(cell.Value) Then N = CDbl(cell.Value) Else N = 0
End Function

Private Function Q(s As String) As String
    Q = """" & s & """"
End Function

Private Function J(v As Double) As String
    J = Replace(CStr(CDec(v)), ",", ".")
End Function
