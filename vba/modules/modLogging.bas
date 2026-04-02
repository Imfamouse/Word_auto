Attribute VB_Name = "modLogging"
Option Explicit

Public Sub LogAction(ByVal documentId As String, ByVal actionName As String, ByVal resultValue As String, ByVal messageText As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = Environ$("Username")
    ws.Cells(nextRow, 3).Value = documentId
    ws.Cells(nextRow, 4).Value = actionName
    ws.Cells(nextRow, 5).Value = resultValue
    ws.Cells(nextRow, 6).Value = messageText
End Sub
