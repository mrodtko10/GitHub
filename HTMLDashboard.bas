Attribute VB_Name = "HTMLDashboard"
Option Explicit

' ============================================================
' PASTE YOUR DASHBOARD FILE PATH HERE (include the filename):
'   Example: "C:\Users\You\Documents\budget_dashboard.html"
' Leave it as "" to auto-detect from Dashboard cell B4.
' ============================================================
Private Const HTML_PATH_OVERRIDE As String = ""

' ============================================================
' TWO MACROS TO ASSIGN TO BUTTONS:
'
'   RefreshHTMLDashboard  -> "Refresh Data" button
'        Reads Category Table and rewrites budget_dashboard.html
'        Does NOT open the browser (safe – no file interference)
'
'   OpenDashboardButton   -> "Open Dashboard" button
'        Opens the HTML file in your browser
'
' Path is read from: HTML_PATH_OVERRIDE constant (if set),
'   otherwise Dashboard sheet cell B4,
'   otherwise workbook folder + "budget_dashboard.html".
' ============================================================

' ── REFRESH: write updated data to HTML file ─────────────────────────────────
Sub RefreshHTMLDashboard()

    Dim wsData    As Worksheet
    Dim htmlPath  As String
    Dim html      As String
    Dim jsonArray As String
    Dim startPos  As Long
    Dim endPos    As Long
    Dim cvrMonth  As String
    Dim lastRow   As Long
    Dim i         As Long

    ' ── Resolve path ────────────────────────────────────────────────────────
    htmlPath = GetHtmlPath()
    If htmlPath = "" Then
        MsgBox "No HTML path found. Please paste the full path to" & vbCrLf & _
               "budget_dashboard.html into cell B4 on the Dashboard sheet.", _
               vbExclamation, "Path Not Set"
        Exit Sub
    End If

    ' ── Safety: must end in .html ────────────────────────────────────────────
    If LCase(Right(Trim(htmlPath), 5)) <> ".html" And _
       LCase(Right(Trim(htmlPath), 4)) <> ".htm" Then
        MsgBox "SAFETY STOP – path does not end in .html:" & vbCrLf & htmlPath & vbCrLf & vbCrLf & _
               "Check that cell B4 contains the HTML file path, not the Excel file path.", _
               vbCritical, "Wrong File Type"
        Exit Sub
    End If

    If LCase(htmlPath) = LCase(ThisWorkbook.FullName) Then
        MsgBox "SAFETY STOP – path points to this workbook!" & vbCrLf & _
               "Update cell B4 with the path to budget_dashboard.html.", _
               vbCritical, "Wrong File"
        Exit Sub
    End If

    If Dir(htmlPath) = "" Then
        MsgBox "HTML file not found at:" & vbCrLf & htmlPath & vbCrLf & vbCrLf & _
               "Make sure budget_dashboard.html exists at that path.", _
               vbCritical, "File Not Found"
        Exit Sub
    End If

    ' ── Find Category Table ──────────────────────────────────────────────────
    Set wsData = Nothing
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If InStr(LCase(ws.Name), "category") > 0 And _
           InStr(LCase(ws.Name), "table") > 0 Then
            Set wsData = ws
            Exit For
        End If
    Next ws
    If wsData Is Nothing Then
        MsgBox "Could not find the 'Category Table' sheet.", vbCritical
        Exit Sub
    End If

    wsData.Calculate

    ' ── CVR Month ────────────────────────────────────────────────────────────
    cvrMonth = ""
    Dim c As Long
    For c = 1 To 8
        Dim cv As String
        cv = Trim(CStr(wsData.Cells(1, c).Value))
        If cv <> "" And cv <> "0" Then
            cv = Replace(Replace(cv, "CVR Month:", ""), "CVR Month :", "")
            cvrMonth = Trim(cv)
            Exit For
        End If
    Next c

    ' ── Build JSON ───────────────────────────────────────────────────────────
    lastRow   = wsData.UsedRange.Row + wsData.UsedRange.Rows.Count - 1
    jsonArray = "["
    Dim firstRec As Boolean
    firstRec = True

    For i = 4 To lastRow
        Dim cc As String, itemName As String
        cc       = Trim(CStr(wsData.Cells(i, 1).Value))
        itemName = Trim(CStr(wsData.Cells(i, 2).Value))
        If cc = "" Or cc = "Cost Code" Or cc = "Total" Or cc = "0" Then GoTo Skip

        Dim suEac As Double, ocEac As Double, maEac As Double
        Dim suAcc As Double, ocAcc As Double, maAcc As Double
        Dim suAct As Double, ocAct As Double, maAct As Double
        Dim suCom As Double, ocCom As Double, maCom As Double
        Dim suCtd As Double, ocCtd As Double, maCtd As Double
        Dim suRem As Double, ocRem As Double, maRem As Double
        Dim suOu  As Double, ocOu  As Double, maOu  As Double

        suEac = N(wsData.Cells(i,15)): ocEac = N(wsData.Cells(i,16)): maEac = N(wsData.Cells(i,17))
        suAcc = N(wsData.Cells(i,21)): ocAcc = N(wsData.Cells(i,22)): maAcc = N(wsData.Cells(i,23))
        suAct = N(wsData.Cells(i,24)): ocAct = N(wsData.Cells(i,25)): maAct = N(wsData.Cells(i,26))
        suCom = N(wsData.Cells(i,27)): ocCom = N(wsData.Cells(i,28)): maCom = N(wsData.Cells(i,29))
        suCtd = N(wsData.Cells(i,33)): ocCtd = N(wsData.Cells(i,34)): maCtd = N(wsData.Cells(i,35))
        suRem = N(wsData.Cells(i,39)): ocRem = N(wsData.Cells(i,40)): maRem = N(wsData.Cells(i,41))
        suOu  = N(wsData.Cells(i,42)): ocOu  = N(wsData.Cells(i,43)): maOu  = N(wsData.Cells(i,44))

        Dim ouStatus As String
        ouStatus = IIf((suOu + ocOu + maOu) < 0, "over", "under")

        cc       = Replace(cc,       """", "\""")
        itemName = Replace(itemName, """", "\""")

        Dim rec As String
        rec = "{" & _
            Q("cost_code") & ":" & Q(cc)       & "," & _
            Q("item")      & ":" & Q(itemName) & "," & _
            Q("su_eac") & ":" & J(suEac) & "," & Q("oc_eac") & ":" & J(ocEac) & "," & Q("ma_eac") & ":" & J(maEac) & "," & _
            Q("su_acc") & ":" & J(suAcc) & "," & Q("oc_acc") & ":" & J(ocAcc) & "," & Q("ma_acc") & ":" & J(maAcc) & "," & _
            Q("su_act") & ":" & J(suAct) & "," & Q("oc_act") & ":" & J(ocAct) & "," & Q("ma_act") & ":" & J(maAct) & "," & _
            Q("su_com") & ":" & J(suCom) & "," & Q("oc_com") & ":" & J(ocCom) & "," & Q("ma_com") & ":" & J(maCom) & "," & _
            Q("su_ctd") & ":" & J(suCtd) & "," & Q("oc_ctd") & ":" & J(ocCtd) & "," & Q("ma_ctd") & ":" & J(maCtd) & "," & _
            Q("su_rem") & ":" & J(suRem) & "," & Q("oc_rem") & ":" & J(ocRem) & "," & Q("ma_rem") & ":" & J(maRem) & "," & _
            Q("ou_status") & ":" & Q(ouStatus) & "}"

        If Not firstRec Then jsonArray = jsonArray & ","
        jsonArray = jsonArray & rec
        firstRec = False
Skip:
    Next i
    jsonArray = jsonArray & "]"

    ' ── Read HTML ────────────────────────────────────────────────────────────
    Dim fNum As Integer
    Dim oneLine As String
    fNum = FreeFile
    html = ""
    Open htmlPath For Input As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, oneLine
        html = html & oneLine & vbLf
    Loop
    Close #fNum

    ' ── Replace RAW_DATA ─────────────────────────────────────────────────────
    startPos = InStr(html, "const RAW_DATA = [")
    If startPos = 0 Then
        MsgBox "RAW_DATA not found in the HTML. Is this the right file?" & vbCrLf & htmlPath, vbCritical
        Exit Sub
    End If
    endPos = InStr(startPos, html, "];")
    If endPos = 0 Then
        MsgBox "Could not find the closing ]; of RAW_DATA.", vbCritical
        Exit Sub
    End If
    html = Left(html, startPos - 1) & "const RAW_DATA = " & jsonArray & ";" & Mid(html, endPos + 2)

    ' ── Update CVR Month ─────────────────────────────────────────────────────
    If cvrMonth <> "" Then
        Dim bPos As Long, bEnd As Long
        bPos = InStr(html, "CVR Month: ")
        If bPos > 0 Then
            bEnd = InStr(bPos + 11, html, "<")
            If bEnd > bPos Then
                html = Left(html, bPos + 10) & cvrMonth & Mid(html, bEnd)
            End If
        End If
    End If

    ' ── Write HTML (file closed before any browser interaction) ──────────────
    fNum = FreeFile
    Open htmlPath For Output As #fNum
    Print #fNum, html
    Close #fNum

    MsgBox "Dashboard data refreshed!" & vbCrLf & _
           (lastRow - 3) & " rows scanned." & vbCrLf & vbCrLf & _
           "Now click the 'Open Dashboard' button (or refresh your browser).", _
           vbInformation, "Refresh Complete"
End Sub

' ── OPEN: launch HTML in browser (separate from refresh) ─────────────────────
Sub OpenDashboardButton()
    Dim htmlPath As String
    htmlPath = GetHtmlPath()

    If htmlPath = "" Or Dir(htmlPath) = "" Then
        MsgBox "HTML file not found." & vbCrLf & _
               "Make sure cell B4 on the Dashboard sheet has the correct path.", _
               vbCritical, "File Not Found"
        Exit Sub
    End If

    ' Use WScript.Shell to open as a file:/// URL
    ' This routes to the browser rather than PDF viewer
    Dim url As String
    url = "file:///" & Replace(Replace(htmlPath, "\", "/"), " ", "%20")

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    On Error Resume Next
    wsh.Run url
    If Err.Number <> 0 Then
        ' Fallback: use FollowHyperlink
        Err.Clear
        ThisWorkbook.FollowHyperlink Address:=htmlPath, NewWindow:=True
    End If
    On Error GoTo 0
End Sub

' ── Shared path resolver ─────────────────────────────────────────────────────
Private Function GetHtmlPath() As String
    If HTML_PATH_OVERRIDE <> "" Then
        GetHtmlPath = HTML_PATH_OVERRIDE: Exit Function
    End If
    Dim wsDash As Worksheet
    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    On Error GoTo 0
    If Not wsDash Is Nothing Then
        Dim p As String
        p = Trim(CStr(wsDash.Cells(4, 2).Value))
        If p <> "" And p <> "0" Then
            GetHtmlPath = p: Exit Function
        End If
    End If
    If ThisWorkbook.Path <> "" Then
        GetHtmlPath = ThisWorkbook.Path & "\budget_dashboard.html"
    End If
End Function

Private Function N(cell As Object) As Double
    If IsNumeric(cell.Value) Then N = CDbl(cell.Value) Else N = 0
End Function
Private Function Q(s As String) As String: Q = """" & s & """": End Function
Private Function J(v As Double) As String: J = Replace(CStr(CDec(v)), ",", "."): End Function
