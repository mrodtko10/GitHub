Attribute VB_Name = "PowerDashboard"
Option Explicit

' ============================================================
' PASTE YOUR DASHBOARD FILE PATH HERE (include the filename):
'   Example: "C:\Users\You\Documents\executive_dashboard.html"
' Leave it as "" to auto-detect from Dashboard sheet cell B4.
' ============================================================
Private Const HTML_PATH_OVERRIDE As String = ""

' ============================================================
' TWO MACROS TO ASSIGN TO BUTTONS:
'
'   RefreshPowerDashboard  -> "Refresh Data" button
'        Reads the "Finalized Projects" sheet and rewrites
'        executive_dashboard.html with the latest ALL_DATA.
'        Does NOT open the browser.
'
'   OpenPowerDashboard     -> "Open Dashboard" button
'        Opens the HTML file in your default browser.
'
' Path priority:
'   1. HTML_PATH_OVERRIDE constant above (if set)
'   2. Dashboard sheet cell B4
'   3. Workbook folder + "executive_dashboard.html"
' ============================================================

' ── REFRESH ─────────────────────────────────────────────────────────────────
Sub RefreshPowerDashboard()

    Dim wsData   As Worksheet
    Dim htmlPath As String
    Dim html     As String
    Dim jsonArr  As String
    Dim startPos As Long
    Dim endPos   As Long
    Dim lastRow  As Long
    Dim i        As Long

    ' ── Resolve HTML path ────────────────────────────────────────────────────
    htmlPath = GetHtmlPath()
    If htmlPath = "" Then
        MsgBox "No HTML path found." & vbCrLf & vbCrLf & _
               "Please paste the full path to executive_dashboard.html" & vbCrLf & _
               "into cell B4 on the Dashboard sheet.", _
               vbExclamation, "Path Not Set"
        Exit Sub
    End If

    ' ── Safety checks ────────────────────────────────────────────────────────
    If LCase(Right(Trim(htmlPath), 5)) <> ".html" And _
       LCase(Right(Trim(htmlPath), 4)) <> ".htm" Then
        MsgBox "SAFETY STOP – path does not end in .html:" & vbCrLf & htmlPath & vbCrLf & vbCrLf & _
               "Check that cell B4 contains the HTML file path.", _
               vbCritical, "Wrong File Type"
        Exit Sub
    End If

    If LCase(Trim(htmlPath)) = LCase(ThisWorkbook.FullName) Then
        MsgBox "SAFETY STOP – path points to this workbook!" & vbCrLf & _
               "Update cell B4 with the path to executive_dashboard.html.", _
               vbCritical, "Wrong File"
        Exit Sub
    End If

    If Dir(htmlPath) = "" Then
        MsgBox "HTML file not found at:" & vbCrLf & htmlPath & vbCrLf & vbCrLf & _
               "Make sure executive_dashboard.html exists at that path.", _
               vbCritical, "File Not Found"
        Exit Sub
    End If

    ' ── Find "Finalized Projects" sheet ──────────────────────────────────────
    Set wsData = Nothing
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If LCase(ws.Name) = "finalized projects" Then
            Set wsData = ws
            Exit For
        End If
    Next ws
    If wsData Is Nothing Then
        MsgBox "Could not find the 'Finalized Projects' sheet.", vbCritical
        Exit Sub
    End If

    wsData.Calculate

    ' ── Build ALL_DATA JSON ───────────────────────────────────────────────────
    ' Column mapping (row 1 = headers, data starts row 2):
    '   A(1)  = File Path          B(2)  = Report Month     C(3)  = Project (id)
    '   D(4)  = Status             E(5)  = Client/Project   F(6)  = PM
    '   G(7)  = Business Unit      H(8)  = Total Rev        I(9)  = Pending COs
    '   J(10) = Pending COs Clmd   K(11) = % Claimed        L(12) = Warranty
    '   M(13) = Contingency        N(14) = JTD Cost         O(15) = ETC
    '   P(16) = EAC                Q(17) = Prior Mo Profit  R(18) = Bid Profit
    '   S(19) = Curr Mo Profit     T(20) = Margin on Cost   U(21) = Margin on Rev
    '   V(22) = Movement           W(23) = Amt Billed       X(24) = Amt Paid
    '   Y(25) = Retention Status   Z(26) = Amt Aging        AA(27)= % Billed
    '   AB(28)= % Paid             AC(29)= Cash Position    AD(30)= Fcst Rem Cash
    '   AE(31)= Coins JTD          AF(32)= Delta            AG(33)= Booked
    '   AH(34)= Forecasted         AI(35)= %               AJ(36)= Report Status
    '   AK(37)= PC Comments        AL(38)= Cost Rpt Match   AM(39)= % Complete
    '   AN(40)= Earned Margin      AO(41)= Remaining Margin AP(42)= Earned Revenue
    '   AQ(43)= Chg Monthly Costs  AR(44)= Chg Contingency  AS(45)= Chg Earned Margin
    '   AT(46)= % Margin           AU(47)= Chg Earned Rev

    lastRow  = wsData.UsedRange.Row + wsData.UsedRange.Rows.Count - 1
    jsonArr  = "["
    Dim firstRec As Boolean
    firstRec = True

    For i = 2 To lastRow
        ' ── Read key identifier columns ──────────────────────────────────────
        Dim projId   As String
        Dim fullName As String
        projId   = Trim(CStr(wsData.Cells(i, 3).Value))   ' C = Project id
        fullName = Trim(CStr(wsData.Cells(i, 5).Value))   ' E = Client/Project name
        If projId = "" Or projId = "Project" Or projId = "0" Then GoTo SkipRow

        ' ── Report Month → "YYYY-MM" ─────────────────────────────────────────
        Dim monthStr As String
        Dim rawMonth As Variant
        rawMonth = wsData.Cells(i, 2).Value
        If IsDate(rawMonth) Then
            monthStr = Format(CDate(rawMonth), "yyyy-mm")
        Else
            Dim rawStr As String
            rawStr = Trim(CStr(rawMonth))
            If rawStr = "" Or rawStr = "0" Then GoTo SkipRow
            ' Already "YYYY-MM-DD" or "YYYY-MM" format
            monthStr = Left(rawStr, 7)
        End If

        ' ── Extract client (text before "/" in name) ─────────────────────────
        Dim clientStr As String
        Dim slashPos  As Long
        slashPos = InStr(fullName, "/")
        If slashPos > 1 Then
            clientStr = Trim(Left(fullName, slashPos - 1))
        Else
            clientStr = fullName
        End If

        ' ── Read numeric fields ───────────────────────────────────────────────
        Dim vRev       As Double, vProfit    As Double, vPriorPro  As Double
        Dim vBidPro    As Double, vMarginRev As Double, vMarginCst As Double
        Dim vPctCmp    As Double, vBilled    As Double, vPaid      As Double
        Dim vJtdCost   As Double, vEac       As Double, vEtc       As Double
        Dim vCont      As Double, vWarranty  As Double, vMovement  As Double
        Dim vCashPos   As Double, vEarnedMgn As Double, vRemMgn    As Double
        Dim vEarnedRev As Double, vPctMargin As Double, vChgCont   As Double

        vRev       = N(wsData.Cells(i, 8))    ' H  Total Rev
        vProfit    = N(wsData.Cells(i, 19))   ' S  Current Month Profit
        vPriorPro  = N(wsData.Cells(i, 17))   ' Q  Prior Month Profit
        vBidPro    = N(wsData.Cells(i, 18))   ' R  Bid Profit
        vMarginCst = N(wsData.Cells(i, 20))   ' T  Margin on Cost
        vMarginRev = N(wsData.Cells(i, 21))   ' U  Margin on Rev
        vMovement  = N(wsData.Cells(i, 22))   ' V  Movement
        vBilled    = N(wsData.Cells(i, 23))   ' W  Amount Billed
        vPaid      = N(wsData.Cells(i, 24))   ' X  Amount Paid
        vCashPos   = N(wsData.Cells(i, 29))   ' AC Cash Position
        vWarranty  = N(wsData.Cells(i, 12))   ' L  Warranty
        vCont      = N(wsData.Cells(i, 13))   ' M  Contingency
        vJtdCost   = N(wsData.Cells(i, 14))   ' N  JTD Cost
        vEtc       = N(wsData.Cells(i, 15))   ' O  ETC
        vEac       = N(wsData.Cells(i, 16))   ' P  EAC
        vPctCmp    = N(wsData.Cells(i, 39))   ' AM % Complete
        vEarnedMgn = N(wsData.Cells(i, 40))   ' AN Earned Margin
        vRemMgn    = N(wsData.Cells(i, 41))   ' AO Remaining Margin
        vEarnedRev = N(wsData.Cells(i, 42))   ' AP Earned Revenue
        vChgCont   = N(wsData.Cells(i, 44))   ' AR Change in Contingency
        vPctMargin = N(wsData.Cells(i, 46))   ' AT % Margin

        ' ── Read string fields ────────────────────────────────────────────────
        Dim vStatus As String, vPm As String, vBu As String
        vStatus = Trim(CStr(wsData.Cells(i, 4).Value))    ' D  Status
        vPm     = Trim(CStr(wsData.Cells(i, 6).Value))    ' F  PM
        vBu     = Trim(CStr(wsData.Cells(i, 7).Value))    ' G  Business Unit

        ' ── Escape strings for JSON ───────────────────────────────────────────
        projId    = Replace(projId,    """", "\""")
        fullName  = Replace(fullName,  """", "\""")
        clientStr = Replace(clientStr, """", "\""")
        vStatus   = Replace(vStatus,   """", "\""")
        vPm       = Replace(vPm,       """", "\""")
        vBu       = Replace(vBu,       """", "\""")

        ' ── Build record (split to stay under VBA's 24-continuation limit) ────
        Dim rec As String
        rec = "{"
        rec = rec & Q("month")             & ":" & Q(monthStr)   & ","
        rec = rec & Q("id")                & ":" & Q(projId)     & ","
        rec = rec & Q("name")              & ":" & Q(fullName)   & ","
        rec = rec & Q("client")            & ":" & Q(clientStr)  & ","
        rec = rec & Q("pm")                & ":" & Q(vPm)        & ","
        rec = rec & Q("bu")                & ":" & Q(vBu)        & ","
        rec = rec & Q("status")            & ":" & Q(vStatus)    & ","
        rec = rec & Q("rev")               & ":" & J(vRev)       & ","
        rec = rec & Q("profit")            & ":" & J(vProfit)    & ","
        rec = rec & Q("prior_profit")      & ":" & J(vPriorPro)  & ","
        rec = rec & Q("bid_profit")        & ":" & J(vBidPro)    & ","
        rec = rec & Q("margin_rev")        & ":" & J(vMarginRev) & ","
        rec = rec & Q("margin_cost")       & ":" & J(vMarginCst) & ","
        rec = rec & Q("pct_complete")      & ":" & J(vPctCmp)    & ","
        rec = rec & Q("billed")            & ":" & J(vBilled)    & ","
        rec = rec & Q("paid")              & ":" & J(vPaid)      & ","
        rec = rec & Q("jtd_cost")          & ":" & J(vJtdCost)   & ","
        rec = rec & Q("eac")               & ":" & J(vEac)       & ","
        rec = rec & Q("etc")               & ":" & J(vEtc)       & ","
        rec = rec & Q("contingency")       & ":" & J(vCont)      & ","
        rec = rec & Q("warranty")          & ":" & J(vWarranty)  & ","
        rec = rec & Q("movement")          & ":" & J(vMovement)  & ","
        rec = rec & Q("cash_pos")          & ":" & J(vCashPos)   & ","
        rec = rec & Q("earned_margin")     & ":" & J(vEarnedMgn) & ","
        rec = rec & Q("remaining_margin")  & ":" & J(vRemMgn)    & ","
        rec = rec & Q("earned_rev")        & ":" & J(vEarnedRev) & ","
        rec = rec & Q("pct_margin")        & ":" & J(vPctMargin) & ","
        rec = rec & Q("change_contingency")& ":" & J(vChgCont)
        rec = rec & "}"

        If Not firstRec Then jsonArr = jsonArr & ","
        jsonArr  = jsonArr & rec
        firstRec = False
SkipRow:
    Next i
    jsonArr = jsonArr & "]"

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

    ' ── Replace ALL_DATA ─────────────────────────────────────────────────────
    startPos = InStr(html, "const ALL_DATA = [")
    If startPos = 0 Then
        MsgBox "ALL_DATA marker not found in the HTML file." & vbCrLf & _
               "Is this the correct executive_dashboard.html?" & vbCrLf & vbCrLf & htmlPath, _
               vbCritical, "Marker Not Found"
        Exit Sub
    End If
    endPos = InStr(startPos, html, "];")
    If endPos = 0 Then
        MsgBox "Could not find the closing ]; of ALL_DATA.", vbCritical
        Exit Sub
    End If
    html = Left(html, startPos - 1) & "const ALL_DATA = " & jsonArr & ";" & Mid(html, endPos + 2)

    ' ── Write updated HTML (file closed before any browser interaction) ───────
    fNum = FreeFile
    Open htmlPath For Output As #fNum
    Print #fNum, html
    Close #fNum

    MsgBox "Dashboard data refreshed!" & vbCrLf & _
           (lastRow - 1) & " rows scanned, " & _
           (Len(jsonArr) - Len(Replace(jsonArr, "},{", ""))) + 1 & " records written." & vbCrLf & vbCrLf & _
           "Click 'Open Dashboard' or refresh your browser to see the latest data.", _
           vbInformation, "Refresh Complete"
End Sub

' ── OPEN: launch HTML in default browser ────────────────────────────────────
Sub OpenPowerDashboard()
    Dim htmlPath As String
    htmlPath = GetHtmlPath()

    If htmlPath = "" Or Dir(htmlPath) = "" Then
        MsgBox "HTML file not found." & vbCrLf & _
               "Make sure cell B4 on the Dashboard sheet has the correct path.", _
               vbCritical, "File Not Found"
        Exit Sub
    End If

    Dim url As String
    url = "file:///" & Replace(Replace(htmlPath, "\", "/"), " ", "%20")
    On Error Resume Next
    ThisWorkbook.FollowHyperlink Address:=url, NewWindow:=True
    On Error GoTo 0
End Sub

' ── COMBINED: refresh then open ──────────────────────────────────────────────
Sub RefreshAndOpenDashboard()
    RefreshPowerDashboard
    OpenPowerDashboard
End Sub

' ── Path resolver ────────────────────────────────────────────────────────────
' Priority: 1) HTML_PATH_OVERRIDE constant, 2) Dashboard!B4, 3) workbook folder
Private Function GetHtmlPath() As String
    ' 1. Hardcoded override
    If HTML_PATH_OVERRIDE <> "" Then
        GetHtmlPath = HTML_PATH_OVERRIDE
        Exit Function
    End If

    ' 2. Dashboard sheet B4
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

    ' 3. Same folder as workbook
    If ThisWorkbook.Path <> "" Then
        GetHtmlPath = ThisWorkbook.Path & "\executive_dashboard.html"
    End If
End Function

' ── Helpers ──────────────────────────────────────────────────────────────────
Private Function N(cell As Object) As Double
    If IsNumeric(cell.Value) Then N = CDbl(cell.Value) Else N = 0
End Function

Private Function Q(s As String) As String
    Q = """" & s & """"
End Function

' Converts a Double to a JSON number string (no thousand separators, dot decimal)
Private Function J(v As Double) As String
    J = Replace(CStr(CDec(v)), ",", ".")
End Function
