Attribute VB_Name = "modConfig"
Option Explicit

Public Function GetConfigValue(ByVal keyName As String, Optional ByVal defaultValue As String = "") As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIdx As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_CFG_APP)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For rowIdx = 2 To lastRow
        If Trim$(CStr(ws.Cells(rowIdx, 1).Value)) = keyName Then
            GetConfigValue = Trim$(CStr(ws.Cells(rowIdx, 2).Value))
            Exit Function
        End If
    Next rowIdx

    GetConfigValue = defaultValue
End Function

Public Sub EnsureRequiredConfig()
    Dim templatesPath As String
    Dim outputPath As String

    templatesPath = GetConfigValue("templates_path")
    outputPath = GetConfigValue("output_path")

    If Len(templatesPath) = 0 Then Err.Raise vbObjectError + 1000, "EnsureRequiredConfig", "Missing config: templates_path"
    If Len(outputPath) = 0 Then Err.Raise vbObjectError + 1001, "EnsureRequiredConfig", "Missing config: output_path"
End Sub
