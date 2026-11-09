Attribute VB_Name = "FPAForecast"
' ===== Vantora Software - FY27 Rolling Forecast - macro module =====
' Import via the VBE (File > Import File), then assign the sheet buttons:
'   Refresh Actuals -> RefreshActuals_Click      Roll Forward Month -> RollForwardMonth_Click
'   Snapshot Version -> SnapshotVersion_Click    Submit for Approval -> SubmitForApproval_Click
' Written by FP&A, last touched v7. Do not edit without telling Priya.
Option Explicit

' "Refresh Actuals" button: reload the Actuals tab from the GL warehouse via ODBC
Sub RefreshActuals_Click()
    Dim cn As Object, rs As Object, ws As Worksheet, tbl As ListObject
    Dim sql As String

    On Error GoTo RefreshFailed
    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets("Actuals")
    Set tbl = ws.ListObjects("tblActuals")

    sql = "SELECT posting_date, period_label, line_item, region, product_line, amount " & _
          "FROM GL.ACTUALS_MONTHLY WHERE fiscal_year >= 2026 ORDER BY posting_date"

    Set cn = CreateObject("ADODB.Connection")
    cn.Open "DSN=VANTORA_GL;UID=fpa_reader;"
    Set rs = cn.Execute(sql)

    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.ClearContents
    ws.Range("A2").CopyFromRecordset rs
    rs.Close
    cn.Close

    Application.Calculate
    Application.ScreenUpdating = True
    MsgBox "Actuals refreshed from the GL warehouse.", vbInformation
    Exit Sub

RefreshFailed:
    Application.ScreenUpdating = True
    MsgBox "Could not refresh actuals: " & Err.Description, vbExclamation
End Sub

' Scenario switch: copy the chosen column of tblDrivers into the live monthly rate column
Sub SetScenario(Optional ByVal scenarioName As String = "Base")
    Dim ws As Worksheet, tbl As ListObject, lr As ListRow, srcCol As Long

    Set ws = ThisWorkbook.Worksheets("Assumptions")
    Set tbl = ws.ListObjects("tblDrivers")

    Select Case scenarioName
        Case "Downside"
            srcCol = 4
        Case "Upside"
            srcCol = 6
        Case Else
            srcCol = 5
    End Select

    For Each lr In tbl.ListRows
        lr.Range(1, 2).Value = lr.Range(1, srcCol).Value
    Next lr

    ws.Range("B2").Value = scenarioName
    Application.Calculate
End Sub

' "Roll Forward Month" button: lock the first forecast month in as an actual
Sub RollForwardMonth_Click()
    Dim fc As Worksheet, act As ListObject, newRow As ListRow
    Dim firstFcCol As Long, r As Long, monthLabel As String

    Set fc = ThisWorkbook.Worksheets("Forecast")
    Set act = ThisWorkbook.Worksheets("Actuals").ListObjects("tblActuals")
    firstFcCol = 20                                  ' Jul-27
    monthLabel = fc.Cells(4, firstFcCol).Value

    Application.ScreenUpdating = False
    For r = 6 To 45
        If Len(fc.Cells(r, 1).Value) > 0 And IsNumeric(fc.Cells(r, firstFcCol).Value) Then
            If Left(fc.Cells(r, 1).Value, 5) <> "Total" And fc.Cells(r, 1).Value <> "Gross Margin" Then
                Set newRow = act.ListRows.Add
                newRow.Range(1, 1).Value = DateSerial(2027, 7, 1)
                newRow.Range(1, 2).Value = monthLabel
                newRow.Range(1, 3).Value = fc.Cells(r, 1).Value
                newRow.Range(1, 6).Value = fc.Cells(r, firstFcCol).Value
            End If
        End If
    Next r

    Application.CalculateFull
    Application.ScreenUpdating = True
    MsgBox monthLabel & " rolled into actuals.", vbInformation
End Sub

' "Snapshot Version" button: freeze the forecast as values and log a row in tblVersions
Sub SnapshotVersion_Click()
    Dim src As Range, vs As Worksheet, tbl As ListObject, newRow As ListRow
    Dim nextRow As Long

    Set vs = ThisWorkbook.Worksheets("Versions")
    Set tbl = vs.ListObjects("tblVersions")
    Set src = ThisWorkbook.Worksheets("Forecast").Range("A4:Z30")

    Set newRow = tbl.ListRows.Add
    newRow.Range(1, 1).Value = "v" & tbl.ListRows.Count
    newRow.Range(1, 2).Value = Now
    newRow.Range(1, 3).Value = ThisWorkbook.Worksheets("Assumptions").Range("B2").Value
    newRow.Range(1, 4).Value = Application.UserName

    nextRow = vs.Cells(vs.Rows.Count, 9).End(xlUp).Row + 1
    src.Copy
    vs.Cells(nextRow, 9).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
End Sub

' "Submit for Approval" button: lock the forecast, log it, and mail the CFO
Sub SubmitForApproval_Click()
    Dim fc As Worksheet, tbl As ListObject, newRow As ListRow
    Dim olApp As Object, olMail As Object

    Set fc = ThisWorkbook.Worksheets("Forecast")
    Set tbl = ThisWorkbook.Worksheets("Approvals").ListObjects("tblApprovals")

    On Error GoTo SubmitFailed
    fc.Protect UserInterfaceOnly:=True

    Set newRow = tbl.ListRows.Add
    newRow.Range(1, 1).Value = "FY27 v7"
    newRow.Range(1, 2).Value = Application.UserName
    newRow.Range(1, 3).Value = "CFO"
    newRow.Range(1, 4).Value = "Pending"
    newRow.Range(1, 5).Value = Now

    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    olMail.To = "cfo@vantora.example"
    olMail.Subject = "FY27 rolling forecast v7 - ready for approval"
    olMail.Body = "Hi," & vbCrLf & vbCrLf & _
                  "The FY27 rolling forecast (v7) is locked and ready for your review." & _
                  vbCrLf & vbCrLf & "Thanks, FP&A"
    olMail.Send
    Exit Sub

SubmitFailed:
    fc.Unprotect
    MsgBox "Submit failed: " & Err.Description, vbExclamation
End Sub

' Full recalculation helper - run after pasting new actuals
Sub RecalcAll()
    Application.ScreenUpdating = False
    Application.CalculateFull
    ThisWorkbook.Worksheets(1).Calculate
    Application.ScreenUpdating = True
End Sub

' Report formatting helper - run before printing to PDF
Sub FormatReport()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Forecast")
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Range("A4:Z4").Font.Bold = True
    ws.Range("A4:Z4").Interior.Color = RGB(31, 56, 100)
    ws.Range("A4:Z4").Font.Color = RGB(255, 255, 255)
    ws.Range("C5:Z40").NumberFormat = "#,##0;(#,##0)"
    ws.Columns("A:Z").AutoFit
End Sub
