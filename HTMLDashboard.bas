Attribute VB_Name = "HTMLDashboard"
Option Explicit

' ============================================================
' PASTE YOUR DASHBOARD FILE PATH HERE (include the filename):
'   Example: "C:\Users\You\Documents\budget_dashboard.html"
' Leave it as "" to auto-detect (workbook folder).
' ============================================================
Private Const HTML_PATH_OVERRIDE As String = ""

' ============================================================
' RefreshHTMLDashboard
' Reads Category Table, rebuilds RAW_DATA in budget_dashboard.html
'
' INSTALL:
'   Alt+F11 > (if old module exists: right-click it > Remove)
'   File > Import File > select this .bas file
'   Assign macro RefreshHTMLDashboard to a button
' ============================================================

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

    ' ── Resolve HTML file path ───────────────────────────────────────────────
    ' Priority: 1) HTML_PATH_OVERRIDE constant  2) Dashboard!B4  3) workbook folder
    If HTML_PATH_OVERRIDE <> "" Then
        htmlPath = HTML_PATH_OVERRIDE
    Else
        ' Try to read path from Dashboard sheet cell B4
        Dim wsDash As Worksheet
        On Error Resume Next
        Set wsDash = ThisWorkbook.Sheets("Dashboard")
        On Error GoTo 0
        If Not wsDash Is Nothing Then
            Dim cellPath As String
            cellPath = Trim(CStr(wsDash.Cells(4, 2).Value))   ' row 4, col B
            If cellPath <> "" And cellPath <> "0" Then
                htmlPath = cellPath
            End If
        End If
        ' Fall back to workbook folder
        If htmlPath = "" Then
            If ThisWorkbook.Path <> "" Then
                htmlPath = ThisWorkbook.Path & "\budget_dashboard.html"
            Else
                MsgBox "No file path found. Please either:" & vbCrLf & _
                       "  1. Paste the full HTML path in Dashboard cell B4, or" & vbCrLf & _
                       "  2. Paste it in HTML_PATH_OVERRIDE at the top of this module.", _
                       vbExclamation, "Path Unknown"
                Exit Sub
            End If
        End If
    End If

    If Dir(htmlPath) = "" Then
        MsgBox "HTML file not found at:" & vbCrLf & htmlPath & vbCrLf & vbCrLf & _
               "To fix: open this module (Alt+F11) and paste the correct path" & vbCrLf & _
               "into the HTML_PATH_OVERRIDE constant at the top.", _
               vbCritical, "File Not Found"
        Exit Sub
    End If

    ' ── Find Category Table sheet ────────────────────────────────────────────
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

    wsData.Calculate   ' force formula recalculation

    ' ── CVR Month (row 1) ────────────────────────────────────────────────────
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

    ' ── Build JSON (data starts row 4) ──────────────────────────────────────
    ' Column map (1-based):
    '  EAC  SU=15 OC=16 MA=17 | Accrued SU=21 OC=22 MA=23
    '  Act  SU=24 OC=25 MA=26 | Committed SU=27 OC=28 MA=29
    '  CTD  SU=33 OC=34 MA=35 | Rem  SU=39 OC=40 MA=41
    '  O/U  SU=42 OC=43 MA=44

    lastRow = wsData.UsedRange.Row + wsData.UsedRange.Rows.Count - 1
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
    Dim line As String
    fNum = FreeFile
    html = ""
    Open htmlPath For Input As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, line
        html = html & line & vbLf
    Loop
    Close #fNum

    ' ── Replace RAW_DATA block ───────────────────────────────────────────────
    startPos = InStr(html, "const RAW_DATA = [")
    If startPos = 0 Then
        MsgBox "RAW_DATA not found in the HTML file. Are you pointing at the right file?", vbCritical
        Exit Sub
    End If
    endPos = InStr(startPos, html, "];")
    If endPos = 0 Then
        MsgBox "Could not find the closing ]; of RAW_DATA.", vbCritical: Exit Sub
    End If

    html = Left(html, startPos - 1) & _
           "const RAW_DATA = " & jsonArray & ";" & _
           Mid(html, endPos + 2)

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

    ' ── Write HTML ───────────────────────────────────────────────────────────
    fNum = FreeFile
    Open htmlPath For Output As #fNum
    Print #fNum, html
    Close #fNum

    ' Open dashboard in browser after refresh
    OpenDashboard htmlPath

    MsgBox "Done!  " & (lastRow - 3) & " rows scanned." & vbCrLf & _
           "Dashboard opened in your browser.", vbInformation, "Dashboard Updated"
End Sub

' ============================================================
' OpenDashboard  –  assign this to your hyperlink button.
' Forces the file to open in a browser instead of PDF viewer.
' ============================================================
Sub OpenDashboardButton()
    OpenDashboard GetHtmlPath()
End Sub

' ── Internal: open a file path in a browser ─────────────────────────────────
Private Sub OpenDashboard(htmlPath As String)
    If htmlPath = "" Or Dir(htmlPath) = "" Then
        MsgBox "HTML file not found:" & vbCrLf & htmlPath, vbCritical
        Exit Sub
    End If
    ' Convert backslashes to forward slashes for file:/// URL
    Dim url As String
    url = "file:///" & Replace(htmlPath, "\", "/")

    ' Try browsers in order: Edge (always on Win10/11), Chrome, Firefox
    Dim launched As Boolean
    launched = False

    ' Microsoft Edge
    If Not launched Then
        On Error Resume Next
        Shell "cmd /c start msedge """ & url & """", vbHide
        If Err.Number = 0 Then launched = True
        On Error GoTo 0
    End If

    ' Google Chrome
    If Not launched Then
        On Error Resume Next
        Shell "cmd /c start chrome """ & url & """", vbHide
        If Err.Number = 0 Then launched = True
        On Error GoTo 0
    End If

    ' Firefox
    If Not launched Then
        On Error Resume Next
        Shell "cmd /c start firefox """ & url & """", vbHide
        If Err.Number = 0 Then launched = True
        On Error GoTo 0
    End If

    ' Last resort: let Windows decide (may still be PDF viewer)
    If Not launched Then
        Shell "cmd /c start """" """ & htmlPath & """", vbHide
    End If
End Sub

' ── Internal: resolve the HTML path (same logic as RefreshHTMLDashboard) ─────
Private Function GetHtmlPath() As String
    If HTML_PATH_OVERRIDE <> "" Then
        GetHtmlPath = HTML_PATH_OVERRIDE
        Exit Function
    End If
    Dim wsDash As Worksheet
    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    On Error GoTo 0
    If Not wsDash Is Nothing Then
        Dim cellPath As String
        cellPath = Trim(CStr(wsDash.Cells(4, 2).Value))
        If cellPath <> "" And cellPath <> "0" Then
            GetHtmlPath = cellPath
            Exit Function
        End If
    End If
    If ThisWorkbook.Path <> "" Then
        GetHtmlPath = ThisWorkbook.Path & "\budget_dashboard.html"
    End If
End Function

Private Function N(cell As Object) As Double
    If IsNumeric(cell.Value) Then N = CDbl(cell.Value) Else N = 0
End Function

Private Function Q(s As String) As String
    Q = """" & s & """"
End Function

Private Function J(v As Double) As String
    J = Replace(CStr(CDec(v)), ",", ".")
End Function
