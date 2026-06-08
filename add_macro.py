"""
add_macro.py
------------
Run this ONCE to inject the "Refresh Dashboard" macro into your Excel file
and save it as Weekly_Commitments_Review.xlsm

Requirements: pywin32  (pip install pywin32)
Place this script in the same folder as the Excel file, then double-click it.
"""

import os
import sys
import shutil

try:
    import win32com.client
except ImportError:
    input(
        "ERROR: pywin32 is not installed.\n"
        "Run:  pip install pywin32\n\n"
        "Press Enter to close."
    )
    sys.exit(1)

HERE       = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH  = os.path.join(HERE, "Weekly_Commitments_Review.xlsx")
XLSM_PATH  = os.path.join(HERE, "Weekly_Commitments_Review.xlsm")

# ---------------------------------------------------------------------------
# VBA macro code
# ---------------------------------------------------------------------------
VBA_CODE = r'''
Sub RefreshDashboard()
    Dim wsData      As Worksheet
    Dim htmlPath    As String
    Dim fso         As Object
    Dim ts          As Object
    Dim htmlContent As String
    Dim newData     As String
    Dim i           As Long
    Dim lastRow     As Long
    Dim rec         As String
    Dim records()   As String
    Dim recCount    As Long
    Dim cvrMonth    As String

    ' ---- Locate files in same folder as this workbook ----
    htmlPath = ThisWorkbook.Path & "\budget_dashboard.html"

    If Dir(htmlPath) = "" Then
        MsgBox "Could not find budget_dashboard.html in:" & vbCrLf & ThisWorkbook.Path, vbCritical
        Exit Sub
    End If

    ' ---- Find Category Table sheet ----
    Dim wsName As String
    wsName = ""
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If InStr(LCase(ws.Name), "category") > 0 And InStr(LCase(ws.Name), "table") > 0 Then
            wsName = ws.Name
            Exit For
        End If
    Next ws
    If wsName = "" Then
        MsgBox "Could not find 'Category Table' sheet.", vbCritical
        Exit Sub
    End If
    Set wsData = ThisWorkbook.Sheets(wsName)

    ' ---- Read CVR Month from row 1 ----
    cvrMonth = ""
    Dim c As Long
    For c = 1 To 10
        Dim cellVal As String
        cellVal = Trim(CStr(wsData.Cells(1, c).Value))
        If cellVal <> "" And cellVal <> "0" Then
            cellVal = Replace(cellVal, "CVR Month:", "")
            cellVal = Replace(cellVal, "CVR Month :", "")
            cvrMonth = Trim(cellVal)
            Exit For
        End If
    Next c

    ' ---- Build JSON records (data starts row 4, skip header rows 1-3) ----
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    recCount = 0
    ReDim records(1 To lastRow)

    For i = 4 To lastRow
        Dim costCode As String
        Dim item     As String
        costCode = Trim(CStr(wsData.Cells(i, 1).Value))
        item     = Trim(CStr(wsData.Cells(i, 2).Value))

        If costCode = "" Or costCode = "Cost Code" Or costCode = "Total" Or costCode = "0" Then
            GoTo NextRow
        End If

        ' Helper: read numeric cell, return 0 for blank/text
        Dim suEac As Double, ocEac As Double, maEac As Double
        Dim suAcc As Double, ocAcc As Double, maAcc As Double
        Dim suAct As Double, ocAct As Double, maAct As Double
        Dim suCom As Double, ocCom As Double, maCom As Double
        Dim suCtd As Double, ocCtd As Double, maCtd As Double
        Dim suRem As Double, ocRem As Double, maRem As Double
        Dim suOu  As Double, ocOu  As Double, maOu  As Double

        suEac = SafeNum(wsData.Cells(i, 15).Value)
        ocEac = SafeNum(wsData.Cells(i, 16).Value)
        maEac = SafeNum(wsData.Cells(i, 17).Value)
        suAcc = SafeNum(wsData.Cells(i, 21).Value)
        ocAcc = SafeNum(wsData.Cells(i, 22).Value)
        maAcc = SafeNum(wsData.Cells(i, 23).Value)
        suAct = SafeNum(wsData.Cells(i, 24).Value)
        ocAct = SafeNum(wsData.Cells(i, 25).Value)
        maAct = SafeNum(wsData.Cells(i, 26).Value)
        suCom = SafeNum(wsData.Cells(i, 27).Value)
        ocCom = SafeNum(wsData.Cells(i, 28).Value)
        maCom = SafeNum(wsData.Cells(i, 29).Value)
        suCtd = SafeNum(wsData.Cells(i, 33).Value)
        ocCtd = SafeNum(wsData.Cells(i, 34).Value)
        maCtd = SafeNum(wsData.Cells(i, 35).Value)
        suRem = SafeNum(wsData.Cells(i, 39).Value)
        ocRem = SafeNum(wsData.Cells(i, 40).Value)
        maRem = SafeNum(wsData.Cells(i, 41).Value)
        suOu  = SafeNum(wsData.Cells(i, 42).Value)
        ocOu  = SafeNum(wsData.Cells(i, 43).Value)
        maOu  = SafeNum(wsData.Cells(i, 44).Value)

        Dim ouStatus As String
        If (suOu + ocOu + maOu) < 0 Then
            ouStatus = "over"
        Else
            ouStatus = "under"
        End If

        ' Escape quotes in strings
        costCode = Replace(costCode, """", "\""")
        item     = Replace(item, """", "\""")

        rec = "{" & _
            """cost_code"":""" & costCode & """," & _
            """item"":"""     & item     & """," & _
            """su_eac"":"  & Fmt(suEac) & "," & _
            """oc_eac"":"  & Fmt(ocEac) & "," & _
            """ma_eac"":"  & Fmt(maEac) & "," & _
            """su_acc"":"  & Fmt(suAcc) & "," & _
            """oc_acc"":"  & Fmt(ocAcc) & "," & _
            """ma_acc"":"  & Fmt(maAcc) & "," & _
            """su_act"":"  & Fmt(suAct) & "," & _
            """oc_act"":"  & Fmt(ocAct) & "," & _
            """ma_act"":"  & Fmt(maAct) & "," & _
            """su_com"":"  & Fmt(suCom) & "," & _
            """oc_com"":"  & Fmt(ocCom) & "," & _
            """ma_com"":"  & Fmt(maCom) & "," & _
            """su_ctd"":"  & Fmt(suCtd) & "," & _
            """oc_ctd"":"  & Fmt(ocCtd) & "," & _
            """ma_ctd"":"  & Fmt(maCtd) & "," & _
            """su_rem"":"  & Fmt(suRem) & "," & _
            """oc_rem"":"  & Fmt(ocRem) & "," & _
            """ma_rem"":"  & Fmt(maRem) & "," & _
            """ou_status"":""" & ouStatus & """" & _
            "}"

        recCount = recCount + 1
        records(recCount) = rec

NextRow:
    Next i

    ' ---- Build full JSON array ----
    newData = "const RAW_DATA = ["
    Dim j As Long
    For j = 1 To recCount
        If j > 1 Then newData = newData & ","
        newData = newData & records(j)
    Next j
    newData = newData & "];"

    ' ---- Read current HTML ----
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts  = fso.OpenTextFile(htmlPath, 1, False, -2)  ' -2 = default system encoding
    htmlContent = ts.ReadAll
    ts.Close

    ' ---- Replace RAW_DATA line ----
    Dim startPos As Long, endPos As Long
    startPos = InStr(htmlContent, "const RAW_DATA = [")
    If startPos = 0 Then
        MsgBox "Could not locate RAW_DATA in the HTML file.", vbCritical
        Exit Sub
    End If
    endPos = InStr(startPos, htmlContent, "];")
    If endPos = 0 Then
        MsgBox "Could not find end of RAW_DATA block.", vbCritical
        Exit Sub
    End If
    endPos = endPos + 1  ' include the ";" character

    htmlContent = Left(htmlContent, startPos - 1) & newData & Mid(htmlContent, endPos + 1)

    ' ---- Replace CVR Month if found ----
    If cvrMonth <> "" Then
        ' Title tag
        Dim titleStart As Long
        titleStart = InStr(htmlContent, "Budget Dashboard")
        If titleStart > 0 Then
            ' Replace text between "– " and "<"
            Dim dashPos As Long
            dashPos = InStr(titleStart, htmlContent, Chr(8211))  ' en-dash
            If dashPos = 0 Then dashPos = InStr(titleStart, htmlContent, "-")
            If dashPos > 0 Then
                Dim afterDash As Long, ltPos As Long
                afterDash = dashPos + 2
                ltPos = InStr(afterDash, htmlContent, "<")
                If ltPos > afterDash Then
                    htmlContent = Left(htmlContent, afterDash - 1) & cvrMonth & Mid(htmlContent, ltPos)
                End If
            End If
        End If
        ' Badge span  "CVR Month: "
        Dim badgePos As Long
        badgePos = InStr(htmlContent, "CVR Month: ")
        If badgePos > 0 Then
            Dim afterBadge As Long, ltPos2 As Long
            afterBadge = badgePos + Len("CVR Month: ")
            ltPos2 = InStr(afterBadge, htmlContent, "<")
            If ltPos2 > afterBadge Then
                htmlContent = Left(htmlContent, afterBadge - 1) & cvrMonth & Mid(htmlContent, ltPos2)
            End If
        End If
    End If

    ' ---- Write updated HTML ----
    Set ts = fso.CreateTextFile(htmlPath, True, False)
    ts.Write htmlContent
    ts.Close

    MsgBox "Dashboard updated!" & vbCrLf & recCount & " cost code records written." & vbCrLf & vbCrLf & "Refresh your browser to see the changes.", vbInformation, "Done"
End Sub

' ---- Helper: safely convert cell value to Double ----
Private Function SafeNum(v As Variant) As Double
    If IsNumeric(v) Then
        SafeNum = CDbl(v)
    Else
        SafeNum = 0
    End If
End Function

' ---- Helper: format number for JSON (no locale comma) ----
Private Function Fmt(v As Double) As String
    Fmt = CStr(v)
    ' Replace locale decimal separator if it isn't a period
    Fmt = Replace(Fmt, ",", ".")
End Function
'''

# ---------------------------------------------------------------------------
# Inject VBA and save as xlsm
# ---------------------------------------------------------------------------
def main():
    src = XLSX_PATH if os.path.exists(XLSX_PATH) else XLSM_PATH
    if not os.path.exists(src):
        input(f"ERROR: Could not find Excel file at:\n  {HERE}\n\nPress Enter to close.")
        sys.exit(1)

    print(f"Opening: {src}")
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False

    try:
        wb = xl.Workbooks.Open(src)

        # Add a new VBA module
        module = wb.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        module.Name = "DashboardRefresh"
        module.CodeModule.AddFromString(VBA_CODE)

        # Add a button on the Dashboard sheet if it exists
        dash_sheet = None
        for sh in wb.Sheets:
            if "dashboard" in sh.Name.lower():
                dash_sheet = sh
                break

        if dash_sheet:
            # Place button near top-right of the dashboard sheet
            btn = dash_sheet.Buttons().Add(10, 10, 160, 28)
            btn.Caption = "🔄 Refresh Dashboard"
            btn.OnAction = "RefreshDashboard"

        # Save as xlsm (52 = xlOpenXMLWorkbookMacroEnabled)
        wb.SaveAs(XLSM_PATH, 52)
        wb.Close(False)
        print(f"Saved: {XLSM_PATH}")

        # Remove old xlsx if different path
        if src != XLSM_PATH and os.path.exists(XLSX_PATH):
            os.remove(XLSX_PATH)
            print(f"Removed old .xlsx file.")

    finally:
        xl.Quit()

    input(
        "\nDone!\n"
        f"Your macro-enabled workbook is ready:\n  {XLSM_PATH}\n\n"
        "The '🔄 Refresh Dashboard' button on the Dashboard tab will\n"
        "update budget_dashboard.html whenever you click it.\n\n"
        "Press Enter to close."
    )


if __name__ == "__main__":
    main()
