Attribute VB_Name = "modExcelData"
Option Explicit

Public Function GetDocumentCard(ByVal documentId As String) As clsDocumentCard
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIdx As Long
    Dim card As clsDocumentCard

    Set ws = ThisWorkbook.Worksheets(SHEET_DOC_CARDS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For rowIdx = 2 To lastRow
        If CStr(ws.Cells(rowIdx, 1).Value) = documentId Then
            Set card = New clsDocumentCard
            card.LoadFromRow ws, rowIdx
            Set GetDocumentCard = card
            Exit Function
        End If
    Next rowIdx

    Err.Raise vbObjectError + 1100, "GetDocumentCard", "Document ID not found: " & documentId
End Function

Public Sub SaveDocumentCard(ByVal card As clsDocumentCard)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIdx As Long
    Dim targetRow As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_DOC_CARDS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    targetRow = 0
    For rowIdx = 2 To lastRow
        If CStr(ws.Cells(rowIdx, 1).Value) = card.DocumentID Then
            targetRow = rowIdx
            Exit For
        End If
    Next rowIdx

    If targetRow = 0 Then targetRow = lastRow + 1
    card.SaveToRow ws, targetRow
End Sub

Public Function GetTemplatePathByDocumentType(ByVal documentType As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIdx As Long
    Dim templateFile As String
    Dim basePath As String
    Dim normalizedInput As String
    Dim normalizedRowType As String

    Set ws = ThisWorkbook.Worksheets(SHEET_REF_TEMPLATES)
    basePath = GetConfigValue("templates_path")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    normalizedInput = NormalizeDocumentType(documentType)

    For rowIdx = 2 To lastRow
        normalizedRowType = NormalizeDocumentType(CStr(ws.Cells(rowIdx, 1).Value))
        If normalizedRowType = normalizedInput Then
            templateFile = CStr(ws.Cells(rowIdx, 2).Value)
            GetTemplatePathByDocumentType = basePath & Application.PathSeparator & templateFile
            Exit Function
        End If
    Next rowIdx

    Err.Raise vbObjectError + 1101, "GetTemplatePathByDocumentType", "Template not found for document type: " & documentType
End Function

Private Function NormalizeDocumentType(ByVal valueText As String) As String
    Dim norm As String
    norm = UCase$(Trim$(valueText))

    Select Case norm
        Case UCase$(DOC_TYPE_RI), "RI", "РI", "РИ", "REPAIR INSTRUCTION"
            NormalizeDocumentType = UCase$(DOC_TYPE_RI)
        Case UCase$(DOC_TYPE_EA), "EA", "ИНЖЕНЕРНЫЙ АНАЛИЗ", "ENGINEERING ANALYSIS"
            NormalizeDocumentType = UCase$(DOC_TYPE_EA)
        Case Else
            NormalizeDocumentType = norm
    End Select
End Function
