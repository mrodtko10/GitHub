Attribute VB_Name = "RefreshEarnedRevenueDashboard"
Option Explicit

'=============================================================================
' RefreshEarnedRevenueDashboard
'
' Reads project data from the "Dashboard" sheet and regenerates the const ER
' data block inside Earned_Revenue_Dashboard_<date>.html.
'
' HOW TO INSTALL:
'   1. In Excel, press Alt+F11 to open the VBA editor.
'   2. In the Project pane right-click the workbook > Import File > select
'      this .bas file.  (Or paste the code into a new Module.)
'   3. Assign the macro to a button on the Charts sheet, or run it via
'      Developer > Macros > RefreshEarnedRevenueDashboard.
'
' HOW IT WORKS:
'   - Charts sheet cell B2 = dashboard file name  (e.g. Earned_Revenue_Dashboard 20260614.html)
'   - Charts sheet cell B4 = full file path       (e.g. C:\Users\...\Earned_Revenue_Dashboard 20260614.html)
'   - The macro reads B4 as the primary path. If B4 holds only a folder it
'     will append the filename from B2 automatically.
'   - Update B4 (and B2 if the filename changes) whenever you move the file.
'   - If the path is wrong or missing you will be prompted to browse; B4 is
'     updated automatically so the next run works without a prompt.
'   - Column mapping on the Dashboard sheet (row 4 = headers, row 5+ = data):
'       A  Project Number     D  Total Contract Value
'       B  Client             E  Earned to Date
'       C  Status             F  Backlog
'       J  Jan 2026 (idx 0)  … through …  BE  Dec 2029 (idx 47)
'   - Actuals = Jan-Apr 2026 (indices 0-3); all other months = Forecast.
'=============================================================================

Public Sub RefreshEarnedRevenueDashboard()

    Const DATA_SHEET    As String = "Dashboard"
    Const CHARTS_SHEET  As String = "Charts"
    Const FNAME_ROW     As Long   = 2    ' row 2 = dashboard file name
    Const PATH_ROW      As Long   = 4    ' row 4 = full file path
    Const PATH_COL      As Long   = 2    ' fallback column B if dynamic search fails
    Const HEADER_ROW    As Long   = 4    ' Row 4 = column headers on Dashboard sheet
    Const DATA_ROW      As Long   = 5    ' First data row
    Const FIRST_MO_COL  As Long   = 10   ' Column J  = Jan 2026
    Const NUM_MONTHS    As Long   = 48   ' Jan 2026 – Dec 2029
    Const ACTUAL_MOS    As Long   = 4    ' Jan – Apr 2026 are Actual

    '--- Worksheet references ------------------------------------------------
    Dim wsData   As Worksheet
    Dim wsCharts As Worksheet
    On Error Resume Next
    Set wsData   = ThisWorkbook.Sheets(DATA_SHEET)
    Set wsCharts = ThisWorkbook.Sheets(CHARTS_SHEET)
    On Error GoTo 0

    If wsData Is Nothing Then
        MsgBox "Sheet """ & DATA_SHEET & """ not found.", vbCritical, "Refresh Error"
        Exit Sub
    End If
    If wsCharts Is Nothing Then
        MsgBox "Sheet """ & CHARTS_SHEET & """ not found.", vbCritical, "Refresh Error"
        Exit Sub
    End If

    '--- Locate HTML dashboard file ------------------------------------------
    '    The path is searched dynamically: scan rows 2 and 4 of the Charts
    '    sheet across all columns for a cell value that ends in ".html".
    '    This is robust to merged cells and varying column positions.
    '    Fallback: look for the cell adjacent to a label containing "path".
    Dim htmlFileName As String
    Dim htmlPath     As String
    Dim pathCell     As Range

    ' 1) Scan row 4 (Dashboard Path row) for a .html value
    Set pathCell = FindHtmlCell(wsCharts, FNAME_ROW, PATH_ROW)

    If Not pathCell Is Nothing Then
        htmlPath = Trim(CStr(pathCell.Value))
    End If

    ' 2) If still empty or just a filename, also check row 2 for a full path
    If LCase(Right(htmlPath, 5)) <> ".html" Or InStr(htmlPath, "\") = 0 Then
        Dim fnCell As Range
        Set fnCell = FindHtmlCell(wsCharts, FNAME_ROW, FNAME_ROW)
        If Not fnCell Is Nothing Then
            htmlFileName = Trim(CStr(fnCell.Value))
        End If
        ' If row 4 value is a folder, append filename from row 2
        If htmlPath <> "" And LCase(Right(htmlPath, 5)) <> ".html" Then
            If Right(htmlPath, 1) <> "\" Then htmlPath = htmlPath & "\"
            htmlPath = htmlPath & htmlFileName
        ElseIf htmlPath = "" Then
            htmlPath = htmlFileName  ' may still be just a name, browse will handle it
        End If
    End If

    If htmlPath = "" Or Dir(htmlPath) = "" Then
        Dim msg As String
        If htmlPath = "" Then
            msg = "Dashboard path not found on the Charts sheet (rows 2 or 4)." & vbCrLf & _
                  "Please browse to the HTML file."
        Else
            msg = "HTML file not found at:" & vbCrLf & htmlPath & vbCrLf & vbCrLf & _
                  "Please browse to the correct file." & vbCrLf & _
                  "(After browsing, the path in row 4 will be updated automatically.)"
        End If
        MsgBox msg, vbExclamation, "File Not Found"
        htmlPath = Application.GetOpenFilename( _
            "HTML Files (*.html),*.html", , _
            "Locate the Earned Revenue Dashboard HTML file")
        If CStr(htmlPath) = "False" Then Exit Sub
        ' Write the resolved path back into the cell the macro found (or B4 if none found)
        If Not pathCell Is Nothing Then
            pathCell.Value = htmlPath
        Else
            wsCharts.Cells(PATH_ROW, PATH_COL).Value = htmlPath
        End If
    End If

    '--- Build months JSON ---------------------------------------------------
    Dim monthAbbr(11) As String
    monthAbbr(0)  = "Jan": monthAbbr(1)  = "Feb": monthAbbr(2)  = "Mar"
    monthAbbr(3)  = "Apr": monthAbbr(4)  = "May": monthAbbr(5)  = "Jun"
    monthAbbr(6)  = "Jul": monthAbbr(7)  = "Aug": monthAbbr(8)  = "Sep"
    monthAbbr(9)  = "Oct": monthAbbr(10) = "Nov": monthAbbr(11) = "Dec"

    Dim monthsJSON As String
    Dim mi As Long, yr As Long, mo As Long
    mi = 0
    monthsJSON = "["
    For yr = 2026 To 2029
        For mo = 0 To 11
            If mi > 0 Then monthsJSON = monthsJSON & ","
            monthsJSON = monthsJSON & _
                "{""label"":""" & monthAbbr(mo) & " " & CStr(yr) & """" & _
                ",""type"":""" & IIf(mi < ACTUAL_MOS, "Actual", "Forecast") & """}"
            mi = mi + 1
        Next mo
    Next yr
    monthsJSON = monthsJSON & "]"

    '--- Build projects JSON -------------------------------------------------
    Dim projectsJSON As String
    projectsJSON = "["
    Dim firstProj As Boolean
    firstProj = True
    Dim projCount As Long
    projCount = 0

    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = DATA_ROW To lastRow

        Dim projCode As String
        projCode = Trim(CStr(wsData.Cells(r, 1).Value))
        If projCode = "" Then GoTo NextRow

        Dim clientName As String
        clientName = Trim(CStr(wsData.Cells(r, 2).Value))

        Dim projStatus As String
        projStatus = Trim(CStr(wsData.Cells(r, 3).Value))
        If projStatus <> "Secured" And projStatus <> "Anticipated" Then GoTo NextRow

        Dim tcv        As Double
        Dim earnedTD   As Double
        Dim backlogAmt As Double
        tcv       = SafeDouble(wsData.Cells(r, 4).Value)
        earnedTD  = SafeDouble(wsData.Cells(r, 5).Value)
        backlogAmt = SafeDouble(wsData.Cells(r, 6).Value)

        '--- 48 monthly values (columns J through BE) -----------------------
        Dim monthly(47) As Double
        Dim i As Long
        For i = 0 To NUM_MONTHS - 1
            monthly(i) = SafeDouble(wsData.Cells(r, FIRST_MO_COL + i).Value)
        Next i

        '--- startIdx / endIdx: first and last non-zero forecast month -------
        Dim startIdx As Long, endIdx As Long, hasRange As Boolean
        startIdx = -1: endIdx = -1: hasRange = False
        For i = ACTUAL_MOS To NUM_MONTHS - 1
            If monthly(i) <> 0 Then
                If Not hasRange Then startIdx = i: hasRange = True
                endIdx = i
            End If
        Next i

        '--- Monthly array JSON ---------------------------------------------
        Dim monthlyArr As String
        monthlyArr = "["
        For i = 0 To NUM_MONTHS - 1
            If i > 0 Then monthlyArr = monthlyArr & ","
            monthlyArr = monthlyArr & DblToJson(monthly(i))
        Next i
        monthlyArr = monthlyArr & "]"

        '--- Assemble project object ----------------------------------------
        Dim startIdxStr As String, endIdxStr As String
        startIdxStr = IIf(hasRange, CStr(startIdx), "null")
        endIdxStr   = IIf(hasRange, CStr(endIdx),   "null")

        Dim projJSON As String
        projJSON = "{""code"":""" & EscJson(projCode) & """" & _
                   ",""client"":""" & EscJson(clientName) & """" & _
                   ",""status"":""" & EscJson(projStatus) & """" & _
                   ",""tcv"":" & DblToJson(tcv) & _
                   ",""earned"":" & DblToJson(earnedTD) & _
                   ",""backlog"":" & DblToJson(backlogAmt) & _
                   ",""startIdx"":" & startIdxStr & _
                   ",""endIdx"":" & endIdxStr & _
                   ",""monthly"":" & monthlyArr & "}"

        If Not firstProj Then projectsJSON = projectsJSON & ","
        projectsJSON = projectsJSON & projJSON
        firstProj = False
        projCount = projCount + 1

NextRow:
    Next r

    projectsJSON = projectsJSON & "]"

    If projCount = 0 Then
        MsgBox "No project rows found on the Dashboard sheet (expected data from row " & _
               DATA_ROW & " down, with Status = Secured or Anticipated).", _
               vbExclamation, "No Data"
        Exit Sub
    End If

    '--- Assemble full ER JSON -----------------------------------------------
    Dim erJSON As String
    erJSON = "{""months"":" & monthsJSON & ",""projects"":" & projectsJSON & "}"

    '--- Read HTML file ------------------------------------------------------
    Dim fNum As Integer
    fNum = FreeFile
    Dim htmlContent As String
    Dim lineIn As String
    Open htmlPath For Input As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, lineIn
        htmlContent = htmlContent & lineIn & Chr(10)
    Loop
    Close #fNum

    '--- Locate "const ER = " block in the HTML ------------------------------
    Dim erStart As Long
    erStart = InStr(htmlContent, "const ER = ")
    If erStart = 0 Then
        MsgBox "Could not find 'const ER = ' in the HTML file." & vbCrLf & _
               "Make sure you are pointing to the correct dashboard HTML.", _
               vbCritical, "Parse Error"
        Exit Sub
    End If

    '--- Walk braces to find the closing "}" of the ER object ----------------
    Dim depth As Long, pos As Long, ch As String
    depth = 0
    Dim objStart As Long
    objStart = InStr(erStart, htmlContent, "{")

    For pos = objStart To Len(htmlContent)
        ch = Mid(htmlContent, pos, 1)
        If ch = "{" Then
            depth = depth + 1
        ElseIf ch = "}" Then
            depth = depth - 1
            If depth = 0 Then Exit For
        End If
    Next pos
    ' pos now points to the closing "}" of the ER JSON object

    '--- Find the trailing ";" (may be on the same line or the next line) ----
    Dim erEnd As Long
    erEnd = pos
    Dim scanPos As Long
    For scanPos = pos + 1 To pos + 20
        If scanPos > Len(htmlContent) Then Exit For
        Dim sc As String
        sc = Mid(htmlContent, scanPos, 1)
        If sc = ";" Then
            erEnd = scanPos
            Exit For
        ElseIf sc = Chr(10) Or sc = Chr(13) Or sc = " " Then
            ' whitespace between "}" and ";" — keep scanning
        Else
            Exit For   ' something unexpected; stop here and keep original ";"
        End If
    Next scanPos

    '--- Write the updated HTML ----------------------------------------------
    Dim newHtml As String
    newHtml = Left(htmlContent, erStart - 1) & _
              "const ER = " & erJSON & ";" & Chr(10) & _
              Mid(htmlContent, erEnd + 1)

    fNum = FreeFile
    Open htmlPath For Output As #fNum
    Print #fNum, newHtml;   ' semicolon suppresses extra trailing newline
    Close #fNum

    MsgBox "Dashboard refreshed — " & projCount & " project(s) exported." & vbCrLf & vbCrLf & _
           htmlPath, vbInformation, "Earned Revenue Dashboard"

End Sub

'=============================================================================
' FindHtmlCell
' Scans the given sheet row(s) and returns the first cell whose value ends in
' ".html". Checks the target row first, then the fallback row.
' Handles merged cells by reading the MergeArea's top-left cell value.
'=============================================================================
Private Function FindHtmlCell(ws As Worksheet, _
                               fallbackRow As Long, _
                               targetRow As Long) As Range
    Dim c As Range
    Dim val As String
    ' Scan target row across used columns
    Dim lastCol As Long
    lastCol = ws.Cells(targetRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 2 Then lastCol = 20   ' minimum scan width

    Dim r As Long
    For r = 1 To 2   ' pass 1 = targetRow, pass 2 = fallbackRow
        Dim scanRow As Long
        scanRow = IIf(r = 1, targetRow, fallbackRow)
        Dim col As Long
        For col = 1 To lastCol
            Set c = ws.Cells(scanRow, col)
            ' If this cell is part of a merge, read from the top-left
            If c.MergeCells Then Set c = c.MergeArea.Cells(1, 1)
            val = Trim(CStr(c.Value))
            If LCase(Right(val, 5)) = ".html" Then
                Set FindHtmlCell = c
                Exit Function
            End If
        Next col
    Next r
    Set FindHtmlCell = Nothing
End Function

'=============================================================================
' SafeDouble  — returns 0 if the cell is empty or non-numeric
'=============================================================================
Private Function SafeDouble(v As Variant) As Double
    If IsNumeric(v) Then SafeDouble = CDbl(v) Else SafeDouble = 0
End Function

'=============================================================================
' DblToJson  — converts a Double to a JSON-safe number string.
'              Uses Str() so the decimal separator is always "." regardless
'              of Windows regional settings.
'=============================================================================
Private Function DblToJson(d As Double) As String
    If d = 0 Then DblToJson = "0": Exit Function
    ' Str() always uses "." as the decimal point (locale-independent)
    DblToJson = LTrim(Str(d))
End Function

'=============================================================================
' EscJson  — escapes a string for safe embedding in a JSON string literal
'=============================================================================
Private Function EscJson(s As String) As String
    s = Replace(s, "\",  "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, Chr(13), "")
    s = Replace(s, Chr(10), "\n")
    EscJson = s
End Function
