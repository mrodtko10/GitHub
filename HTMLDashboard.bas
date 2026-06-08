Attribute VB_Name = "HTMLDashboard"
Option Explicit

' ============================================================
' RefreshHTMLDashboard
' Reads the Category Table, rebuilds the RAW_DATA block in
' budget_dashboard.html, and saves the file.
'
' HOW TO INSTALL:
'   1. In Excel press Alt+F11 to open the VBA editor
'   2. File > Import File > select this .bas file
'      (or Insert > Module, then paste everything below)
'   3. Close the VBA editor
'   4. Assign the macro to a button on your Dashboard sheet
'
' budget_dashboard.html must be in the same folder as this workbook.
' ============================================================

Sub RefreshHTMLDashboard()

    Dim wsData      As Worksheet
    Dim htmlPath    As String
    Dim html        As String
    Dim jsonArray   As String
    Dim startPos    As Long
    Dim endPos      As Long
    Dim cvrMonth    As String
    Dim lastRow     As Long
    Dim i           As Long

    ' ── Locate HTML file ────────────────────────────────────────────────────
    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the workbook first before running this macro.", _
               vbExclamation, "Workbook Not Saved"
        Exit Sub
    End If
    htmlPath = ThisWorkbook.Path & "\budget_dashboard.html"
    If Dir(htmlPath) = "" Then
        MsgBox "Could not find budget_dashboard.html in:" & vbCrLf & _
               ThisWorkbook.Path & vbCrLf & vbCrLf & _
               "Make sure budget_dashboard.html is in the same folder as this workbook.", _
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

    ' Force recalculation so formula cells return current values
    wsData.Calculate

    ' ── Read CVR Month from row 1 ────────────────────────────────────────────
    cvrMonth = ""
    Dim c As Long
    For c = 1 To 8
        Dim cv As String
        cv = Trim(CStr(wsData.Cells(1, c).Value))
        If cv <> "" And cv <> "0" Then
            cv = Replace(cv, "CVR Month:", "")
            cv = Replace(cv, "CVR Month :", "")
            cvrMonth = Trim(cv)
            Exit For
        End If
    Next c

    ' ── Build JSON array (data rows start at row 4) ──────────────────────────
    ' Column map (1-based):
    '   EAC       : SU=15  OC=16  MA=17
    '   Accrued   : SU=21  OC=22  MA=23
    '   Actuals   : SU=24  OC=25  MA=26
    '   Committed : SU=27  OC=28  MA=29
    '   CTD       : SU=33  OC=34  MA=35
    '   Rem EAC   : SU=39  OC=40  MA=41
    '   Over/Under: SU=42  OC=43  MA=44

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    jsonArray = "["
    Dim first As Boolean
    first = True

    For i = 4 To lastRow
        Dim cc As String
        Dim itemName As String
        cc       = Trim(CStr(wsData.Cells(i, 1).Value))
        itemName = Trim(CStr(wsData.Cells(i, 2).Value))

        ' Skip blank, header, or total rows
        If cc = "" Or cc = "Cost Code" Or cc = "Total" Or cc = "0" Then
            GoTo NextRow
        End If

        ' Read all numeric fields
        Dim suEac As Double, ocEac As Double, maEac As Double
        Dim suAcc As Double, ocAcc As Double, maAcc As Double
        Dim suAct As Double, ocAct As Double, maAct As Double
        Dim suCom As Double, ocCom As Double, maCom As Double
        Dim suCtd As Double, ocCtd As Double, maCtd As Double
        Dim suRem As Double, ocRem As Double, maRem As Double
        Dim suOu  As Double, ocOu  As Double, maOu  As Double

        suEac = N(wsData.Cells(i, 15)): ocEac = N(wsData.Cells(i, 16)): maEac = N(wsData.Cells(i, 17))
        suAcc = N(wsData.Cells(i, 21)): ocAcc = N(wsData.Cells(i, 22)): maAcc = N(wsData.Cells(i, 23))
        suAct = N(wsData.Cells(i, 24)): ocAct = N(wsData.Cells(i, 25)): maAct = N(wsData.Cells(i, 26))
        suCom = N(wsData.Cells(i, 27)): ocCom = N(wsData.Cells(i, 28)): maCom = N(wsData.Cells(i, 29))
        suCtd = N(wsData.Cells(i, 33)): ocCtd = N(wsData.Cells(i, 34)): maCtd = N(wsData.Cells(i, 35))
        suRem = N(wsData.Cells(i, 39)): ocRem = N(wsData.Cells(i, 40)): maRem = N(wsData.Cells(i, 41))
        suOu  = N(wsData.Cells(i, 42)): ocOu  = N(wsData.Cells(i, 43)): maOu  = N(wsData.Cells(i, 44))

        Dim ouStatus As String
        If (suOu + ocOu + maOu) < 0 Then ouStatus = "over" Else ouStatus = "under"

        ' Escape any quotes in strings
        cc       = Replace(cc,       """", "\""")
        itemName = Replace(itemName, """", "\""")

        Dim rec As String
        rec = "{" & _
            Q("cost_code") & ":" & Q(cc)       & "," & _
            Q("item")      & ":" & Q(itemName) & "," & _
            Q("su_eac") & ":" & F(suEac) & "," & Q("oc_eac") & ":" & F(ocEac) & "," & Q("ma_eac") & ":" & F(maEac) & "," & _
            Q("su_acc") & ":" & F(suAcc) & "," & Q("oc_acc") & ":" & F(ocAcc) & "," & Q("ma_acc") & ":" & F(maAcc) & "," & _
            Q("su_act") & ":" & F(suAct) & "," & Q("oc_act") & ":" & F(ocAct) & "," & Q("ma_act") & ":" & F(maAct) & "," & _
            Q("su_com") & ":" & F(suCom) & "," & Q("oc_com") & ":" & F(ocCom) & "," & Q("ma_com") & ":" & F(maCom) & "," & _
            Q("su_ctd") & ":" & F(suCtd) & "," & Q("oc_ctd") & ":" & F(ocCtd) & "," & Q("ma_ctd") & ":" & F(maCtd) & "," & _
            Q("su_rem") & ":" & F(suRem) & "," & Q("oc_rem") & ":" & F(ocRem) & "," & Q("ma_rem") & ":" & F(maRem) & "," & _
            Q("ou_status") & ":" & Q(ouStatus) & _
            "}"

        If Not first Then jsonArray = jsonArray & ","
        jsonArray = jsonArray & rec
        first = False

NextRow:
    Next i
    jsonArray = jsonArray & "]"

    ' ── Read HTML file (native VBA file I/O — no FSO encoding issues) ────────
    Dim fileNum As Integer
    fileNum = FreeFile
    Dim oneLine As String
    html = ""
    Open htmlPath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, oneLine
        html = html & oneLine & vbLf
    Loop
    Close #fileNum

    ' ── Replace const RAW_DATA = [...]; ─────────────────────────────────────
    startPos = InStr(html, "const RAW_DATA = [")
    If startPos = 0 Then
        MsgBox "Could not locate RAW_DATA in the HTML file." & vbCrLf & _
               "Make sure you are using the correct budget_dashboard.html.", vbCritical
        Exit Sub
    End If
    endPos = InStr(startPos, html, "];")
    If endPos = 0 Then
        MsgBox "Could not find the end of the RAW_DATA block (];).", vbCritical
        Exit Sub
    End If
    endPos = endPos + 1  ' include the semicolon

    html = Left(html, startPos - 1) & _
           "const RAW_DATA = " & jsonArray & ";" & _
           Mid(html, endPos + 1)

    ' ── Update CVR Month in title and badge ─────────────────────────────────
    If cvrMonth <> "" Then
        ' <title>Budget Dashboard – ...</title>
        Dim p1 As Long, p2 As Long
        p1 = InStr(html, Chr(8211))          ' en-dash –
        If p1 = 0 Then p1 = InStr(html, " - ")
        If p1 > 0 Then
            p2 = InStr(p1 + 2, html, "<")
            If p2 > p1 Then
                html = Left(html, p1) & Chr(8211) & " " & cvrMonth & Mid(html, p2)
            End If
        End If
        ' CVR Month: badge
        Dim bPos As Long
        bPos = InStr(html, "CVR Month: ")
        If bPos > 0 Then
            Dim bEnd As Long
            bEnd = InStr(bPos + 11, html, "<")
            If bEnd > bPos Then
                html = Left(html, bPos + 10) & cvrMonth & Mid(html, bEnd)
            End If
        End If
    End If

    ' ── Write updated HTML ──────────────────────────────────────────────────
    fileNum = FreeFile
    Open htmlPath For Output As #fileNum
    Print #fileNum, html
    Close #fileNum

    MsgBox "Dashboard updated!" & vbCrLf & _
           (lastRow - 3) & " rows scanned  ·  " & _
           "Refresh your browser to see the changes.", _
           vbInformation, "Done"
End Sub

' ── Helper: safely convert cell value to Double ─────────────────────────────
Private Function N(cell As Object) As Double
    If IsNumeric(cell.Value) Then N = CDbl(cell.Value) Else N = 0
End Function

' ── Helper: JSON-quote a string ──────────────────────────────────────────────
Private Function Q(s As String) As String
    Q = """" & s & """"
End Function

' ── Helper: format number for JSON (period as decimal, no locale) ────────────
Private Function F(v As Double) As String
    ' Use CDec to avoid scientific notation on large numbers,
    ' then replace any locale comma with a period.
    F = Replace(CStr(CDec(v)), ",", ".")
End Function
