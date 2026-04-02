Attribute VB_Name = "modReport"
Option Explicit

Public Sub PublishValidationReport(ByVal documentId As String, ByVal issues As Collection)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIdx As Long
    Dim i As Long
    Dim issue As clsValidationIssue

    Set ws = ThisWorkbook.Worksheets(SHEET_VALIDATION)

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowIdx = lastRow To 2 Step -1
        If CStr(ws.Cells(rowIdx, 1).Value) = documentId Then
            ws.Rows(rowIdx).Delete
        End If
    Next rowIdx

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 1 To issues.Count
        Set issue = issues(i)
        ws.Cells(lastRow + i, 1).Value = documentId
        ws.Cells(lastRow + i, 2).Value = issue.Severity
        ws.Cells(lastRow + i, 3).Value = issue.Code
        ws.Cells(lastRow + i, 4).Value = issue.Message
        ws.Cells(lastRow + i, 5).Value = Now
    Next i
End Sub
