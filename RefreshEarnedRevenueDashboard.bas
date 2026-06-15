Attribute VB_Name = "EarnedRevenueDashboard"
Option Explicit

' ============================================================
' PASTE YOUR DASHBOARD FILE PATH HERE (include the filename):
'   Example: "C:\Users\You\Downloads\Earned_Revenue_Dashboard 20260614.html"
' Leave it as "" to auto-detect from Charts sheet cell B4.
' ============================================================
Private Const HTML_PATH_OVERRIDE As String = ""

' ============================================================
' TWO MACROS TO ASSIGN TO BUTTONS:
'
'   RefreshEarnedRevenueDashboard  -> "Refresh & Open Dashboard" button
'        Reads the "Dashboard" sheet and rewrites the HTML
'        const ER = {...} block with the latest project data,
'        then opens the file in your default browser.
'
'   OpenEarnedRevenueDashboard     -> optional "Open" button
'        Opens the HTML file in your default browser only.
'
' Path priority:
'   1. HTML_PATH_OVERRIDE constant above (if set)
'   2. Charts sheet cell B4
'   3. Charts sheet cell B2 (filename) appended to workbook folder
' ============================================================

' ── REFRESH THEN OPEN (assign this to the "Refresh & Open Dashboard" button) ─
Sub RefreshAndOpenEarnedRevenueDashboard()
    RefreshEarnedRevenueDashboard
    OpenEarnedRevenueDashboard
End Sub

' ── REFRESH ONLY ─────────────────────────────────────────────────────────────
Sub RefreshEarnedRevenueDashboard()

    Dim wsData   As Worksheet
    Dim htmlPath As String
    Dim html     As String

    ' ── Resolve HTML path ────────────────────────────────────────────────────
    htmlPath = GetHtmlPath()
    If htmlPath = "" Then
        MsgBox "No HTML path found." & vbCrLf & vbCrLf & _
               "Paste the full path to the dashboard HTML file" & vbCrLf & _
               "into cell B4 on the Charts sheet.", _
               vbExclamation, "Path Not Set"
        Exit Sub
    End If

    ' ── Safety checks ────────────────────────────────────────────────────────
    If LCase(Right(Trim(htmlPath), 5)) <> ".html" And _
       LCase(Right(Trim(htmlPath), 4)) <> ".htm" Then
        MsgBox "SAFETY STOP – path does not end in .html:" & vbCrLf & htmlPath & vbCrLf & vbCrLf & _
               "Check that cell B4 on the Charts sheet contains the HTML file path.", _
               vbCritical, "Wrong File Type"
        Exit Sub
    End If

    If LCase(Trim(htmlPath)) = LCase(ThisWorkbook.FullName) Then
        MsgBox "SAFETY STOP – path points to this workbook!" & vbCrLf & _
               "Update cell B4 on the Charts sheet with the path to the HTML dashboard.", _
               vbCritical, "Wrong File"
        Exit Sub
    End If

    If Dir(htmlPath) = "" Then
        MsgBox "HTML file not found at:" & vbCrLf & htmlPath & vbCrLf & vbCrLf & _
               "Update cell B4 on the Charts sheet with the correct path.", _
               vbCritical, "File Not Found"
        Exit Sub
    End If

    ' ── Find Dashboard sheet ─────────────────────────────────────────────────
    Set wsData = Nothing
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If LCase(ws.Name) = "dashboard" Then
            Set wsData = ws
            Exit For
        End If
    Next ws
    If wsData Is Nothing Then
        MsgBox "Could not find the 'Dashboard' sheet.", vbCritical
        Exit Sub
    End If

    wsData.Calculate

    ' ── Build months JSON array ───────────────────────────────────────────────
    ' 48 months: Jan 2026 (col J=10) through Dec 2029 (col BE=57)
    ' Actuals = Jan-Apr 2026 (indices 0-3); remainder = Forecast
    Dim monthAbbr(11) As String
    monthAbbr(0) = "Jan": monthAbbr(1) = "Feb": monthAbbr(2)  = "Mar"
    monthAbbr(3) = "Apr": monthAbbr(4) = "May": monthAbbr(5)  = "Jun"
    monthAbbr(6) = "Jul": monthAbbr(7) = "Aug": monthAbbr(8)  = "Sep"
    monthAbbr(9) = "Oct": monthAbbr(10) = "Nov": monthAbbr(11) = "Dec"

    Dim monthsJSON As String
    Dim mi As Long, yr As Long, mo As Long
    mi = 0
    monthsJSON = "["
    For yr = 2026 To 2029
        For mo = 0 To 11
            If mi > 0 Then monthsJSON = monthsJSON & ","
            monthsJSON = monthsJSON & _
                "{" & Q("label") & ":" & Q(monthAbbr(mo) & " " & CStr(yr)) & _
                "," & Q("type")  & ":" & Q(IIf(mi < 4, "Actual", "Forecast")) & "}"
            mi = mi + 1
        Next mo
    Next yr
    monthsJSON = monthsJSON & "]"

    ' ── Build projects JSON array ─────────────────────────────────────────────
    ' Dashboard sheet column mapping (row 4 = headers, row 5+ = data):
    '   A(1)  Project Number    D(4)  Total Contract Value
    '   B(2)  Client            E(5)  Earned to Date
    '   C(3)  Status            F(6)  Backlog
    '   J(10) Jan 2026 (idx 0) … BE(57) Dec 2029 (idx 47)

    Dim projectsJSON As String
    Dim firstRec As Boolean
    Dim projCount As Long
    projectsJSON = "["
    firstRec = True
    projCount = 0

    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 5 To lastRow

        Dim projCode As String
        projCode = Trim(CStr(wsData.Cells(r, 1).Value))
        If projCode = "" Or projCode = "Project" Then GoTo SkipRow

        Dim projStatus As String
        projStatus = Trim(CStr(wsData.Cells(r, 3).Value))
        If projStatus <> "Secured" And projStatus <> "Anticipated" Then GoTo SkipRow

        Dim clientName As String
        clientName = Trim(CStr(wsData.Cells(r, 2).Value))

        Dim tcv        As Double
        Dim earnedTD   As Double
        Dim backlogAmt As Double
        tcv        = N(wsData.Cells(r, 4))
        earnedTD   = N(wsData.Cells(r, 5))
        backlogAmt = N(wsData.Cells(r, 6))

        ' Read 48 monthly values (columns J=10 through BE=57)
        Dim monthly(47) As Double
        Dim i As Long
        For i = 0 To 47
            monthly(i) = N(wsData.Cells(r, 10 + i))
        Next i

        ' startIdx / endIdx: first and last non-zero value in forecast range (idx 4+)
        Dim startIdx As Long, endIdx As Long, hasRange As Boolean
        startIdx = -1: endIdx = -1: hasRange = False
        For i = 4 To 47
            If monthly(i) <> 0 Then
                If Not hasRange Then startIdx = i: hasRange = True
                endIdx = i
            End If
        Next i

        ' Monthly array JSON
        Dim monthlyArr As String
        monthlyArr = "["
        For i = 0 To 47
            If i > 0 Then monthlyArr = monthlyArr & ","
            monthlyArr = monthlyArr & J(monthly(i))
        Next i
        monthlyArr = monthlyArr & "]"

        ' Assemble project record
        Dim rec As String
        rec = "{"
        rec = rec & Q("code")     & ":" & Q(EscJ(projCode))    & ","
        rec = rec & Q("client")   & ":" & Q(EscJ(clientName))  & ","
        rec = rec & Q("status")   & ":" & Q(EscJ(projStatus))  & ","
        rec = rec & Q("tcv")      & ":" & J(tcv)               & ","
        rec = rec & Q("earned")   & ":" & J(earnedTD)          & ","
        rec = rec & Q("backlog")  & ":" & J(backlogAmt)        & ","
        rec = rec & Q("startIdx") & ":" & IIf(hasRange, CStr(startIdx), "null") & ","
        rec = rec & Q("endIdx")   & ":" & IIf(hasRange, CStr(endIdx),   "null") & ","
        rec = rec & Q("monthly")  & ":" & monthlyArr
        rec = rec & "}"

        If Not firstRec Then projectsJSON = projectsJSON & ","
        projectsJSON = projectsJSON & rec
        firstRec = False
        projCount = projCount + 1

SkipRow:
    Next r
    projectsJSON = projectsJSON & "]"

    If projCount = 0 Then
        MsgBox "No project rows found on the Dashboard sheet." & vbCrLf & _
               "Expected data from row 5 down with Status = Secured or Anticipated.", _
               vbExclamation, "No Data"
        Exit Sub
    End If

    ' ── Assemble full ER JSON ─────────────────────────────────────────────────
    Dim erJSON As String
    erJSON = "{" & Q("months") & ":" & monthsJSON & "," & Q("projects") & ":" & projectsJSON & "}"

    ' ── Read HTML file ────────────────────────────────────────────────────────
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

    ' ── Replace const ER = {...} block ───────────────────────────────────────
    Dim startPos As Long
    startPos = InStr(html, "const ER = ")
    If startPos = 0 Then
        MsgBox "Marker 'const ER = ' not found in the HTML file." & vbCrLf & _
               "Is this the correct Earned Revenue Dashboard HTML?" & vbCrLf & vbCrLf & htmlPath, _
               vbCritical, "Marker Not Found"
        Exit Sub
    End If

    ' Walk braces to find the closing "}" of the ER object
    Dim depth As Long, pos As Long
    depth = 0
    Dim objStart As Long
    objStart = InStr(startPos, html, "{")
    For pos = objStart To Len(html)
        Dim ch As String
        ch = Mid(html, pos, 1)
        If ch = "{" Then depth = depth + 1
        If ch = "}" Then
            depth = depth - 1
            If depth = 0 Then Exit For
        End If
    Next pos

    ' Scan forward to include any trailing ";" (may be on same or next line)
    Dim endPos As Long
    endPos = pos
    Dim scanPos As Long
    For scanPos = pos + 1 To pos + 20
        If scanPos > Len(html) Then Exit For
        Dim sc As String
        sc = Mid(html, scanPos, 1)
        If sc = ";" Then endPos = scanPos: Exit For
        If sc <> vbLf And sc <> Chr(13) And sc <> " " Then Exit For
    Next scanPos

    html = Left(html, startPos - 1) & _
           "const ER = " & erJSON & ";" & vbLf & _
           Mid(html, endPos + 1)

    ' ── Write updated HTML ───────────────────────────────────────────────────
    fNum = FreeFile
    Open htmlPath For Output As #fNum
    Print #fNum, html;
    Close #fNum

    MsgBox "Dashboard refreshed — " & projCount & " project(s) written." & vbCrLf & vbCrLf & _
           "The browser will open with the latest data.", _
           vbInformation, "Refresh Complete"
End Sub

' ── OPEN ONLY: launch HTML in default browser ─────────────────────────────────
Sub OpenEarnedRevenueDashboard()
    Dim htmlPath As String
    htmlPath = GetHtmlPath()

    If htmlPath = "" Or Dir(htmlPath) = "" Then
        MsgBox "HTML file not found." & vbCrLf & _
               "Make sure cell B4 on the Charts sheet has the correct path.", _
               vbCritical, "File Not Found"
        Exit Sub
    End If

    Dim url As String
    url = "file:///" & Replace(Replace(htmlPath, "\", "/"), " ", "%20")
    On Error Resume Next
    ThisWorkbook.FollowHyperlink Address:=url, NewWindow:=True
    On Error GoTo 0
End Sub

' ── Path resolver ─────────────────────────────────────────────────────────────
' Priority: 1) HTML_PATH_OVERRIDE constant  2) Charts!B4  3) workbook folder + Charts!B2
Private Function GetHtmlPath() As String

    ' 1. Hardcoded override (set at top of module if needed)
    If HTML_PATH_OVERRIDE <> "" Then
        GetHtmlPath = HTML_PATH_OVERRIDE
        Exit Function
    End If

    ' 2. Charts sheet cell B4 (full path)
    Dim chartsWs As Worksheet
    Dim p As String
    On Error Resume Next
    Set chartsWs = ThisWorkbook.Sheets("Charts")
    On Error GoTo 0

    If Not chartsWs Is Nothing Then
        On Error Resume Next
        p = Trim(CStr(chartsWs.Cells(4, 2).Value))   ' B4
        On Error GoTo 0
        If p <> "" And p <> "0" And LCase(Right(p, 5)) = ".html" Then
            GetHtmlPath = p
            Exit Function
        End If

        ' 3. Workbook folder + filename from Charts!B2
        Dim fName As String
        On Error Resume Next
        fName = Trim(CStr(chartsWs.Cells(2, 2).Value))   ' B2
        On Error GoTo 0
        If fName <> "" And ThisWorkbook.Path <> "" Then
            GetHtmlPath = ThisWorkbook.Path & "\" & fName
            Exit Function
        End If
    End If

    ' 4. Last resort: workbook folder + default name
    If ThisWorkbook.Path <> "" Then
        GetHtmlPath = ThisWorkbook.Path & "\Earned_Revenue_Dashboard.html"
    End If
End Function

' ── Helpers ───────────────────────────────────────────────────────────────────
Private Function N(cell As Object) As Double
    If IsNumeric(cell.Value) Then N = CDbl(cell.Value) Else N = 0
End Function

Private Function Q(s As String) As String
    Q = """" & s & """"
End Function

' Converts Double to JSON number (dot decimal, no thousand separators)
Private Function J(v As Double) As String
    If v = 0 Then J = "0": Exit Function
    J = Replace(CStr(CDec(v)), ",", ".")
End Function

' Escapes a string for embedding in a JSON string literal
Private Function EscJ(s As String) As String
    s = Replace(s, "\",   "\\")
    s = Replace(s, """",  "\""")
    s = Replace(s, Chr(13), "")
    s = Replace(s, Chr(10), "\n")
    EscJ = s
End Function
