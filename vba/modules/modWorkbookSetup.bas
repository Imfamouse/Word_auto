Attribute VB_Name = "modWorkbookSetup"
Option Explicit

Public Sub EnsureWorkbookStructure()
    EnsureSheetWithHeaders SHEET_UI_DASHBOARD, Array("action", "description", "macro_name")
    EnsureSheetWithHeaders SHEET_DOC_CARDS, Array( _
        "document_id", "document_type", "title", "aircraft_model", "aircraft_number", "msn", "assembly_number", _
        "part_number", "component_name", "applicability", "revision", "date", "author", "checker", "approver", _
        "related_analysis_number", "related_instruction_number", "references", "attachments", "remarks", "status", _
        "word_doc_path", "pdf_path")
    EnsureSheetWithHeaders SHEET_CFG_APP, Array("key", "value")
    EnsureSheetWithHeaders SHEET_REF_DOCUMENT_TYPES, Array("document_type", "description")
    EnsureSheetWithHeaders SHEET_REF_TEMPLATES, Array("document_type", "template_file")
    EnsureSheetWithHeaders SHEET_REF_STATUSES, Array("status", "description")
    EnsureSheetWithHeaders SHEET_REF_USERS, Array("user_name", "role", "active_flag")
    EnsureSheetWithHeaders SHEET_REQUIRED_FIELDS, Array("document_type", "field_name", "mandatory_flag")
    EnsureSheetWithHeaders SHEET_REQUIRED_SECTIONS, Array("document_type", "section_title", "mandatory_flag", "condition_rule")
    EnsureSheetWithHeaders SHEET_RULES_FILENAME, Array("document_type", "pattern", "description")
    EnsureSheetWithHeaders SHEET_EA_MATRIX, Array("document_id", "clause_id", "clause_title", "applicability_flag", "means_of_compliance", "covered_in_section", "evidence_reference", "status", "comment")
    EnsureSheetWithHeaders SHEET_RI_MATRIX, Array("document_id", "section_code", "section_title", "mandatory_flag", "condition_rule", "present_flag", "comment")
    EnsureSheetWithHeaders SHEET_VALIDATION, Array("document_id", "severity", "code", "message", "timestamp")
    EnsureSheetWithHeaders SHEET_LOG, Array("timestamp", "user", "document_id", "action", "result", "message")

    SeedReferenceData
    SeedDefaultConfig
End Sub

Private Sub EnsureSheetWithHeaders(ByVal sheetName As String, ByVal headers As Variant)
    Dim ws As Worksheet
    Dim i As Long

    Set ws = GetOrCreateSheet(sheetName)

    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = CStr(headers(i))
        ws.Cells(1, i + 1).Font.Bold = True
    Next i

    ws.Rows(1).Interior.Color = RGB(230, 230, 230)
    ws.Rows(1).WrapText = False
    ws.Cells.EntireColumn.AutoFit
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            Set GetOrCreateSheet = ws
            Exit Function
        End If
    Next ws

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
    Set GetOrCreateSheet = ws
End Function

Private Sub SeedReferenceData()
    Dim wsTypes As Worksheet
    Dim wsTemplates As Worksheet
    Dim wsStatuses As Worksheet

    Set wsTypes = ThisWorkbook.Worksheets(SHEET_REF_DOCUMENT_TYPES)
    Set wsTemplates = ThisWorkbook.Worksheets(SHEET_REF_TEMPLATES)
    Set wsStatuses = ThisWorkbook.Worksheets(SHEET_REF_STATUSES)

    If Len(CStr(wsTypes.Cells(2, 1).Value)) = 0 Then
        wsTypes.Cells(2, 1).Value = DOC_TYPE_RI
        wsTypes.Cells(2, 2).Value = "Repair action instruction"
        wsTypes.Cells(3, 1).Value = DOC_TYPE_EA
        wsTypes.Cells(3, 2).Value = "Engineering analysis record"
    End If

    If Len(CStr(wsTemplates.Cells(2, 1).Value)) = 0 Then
        wsTemplates.Cells(2, 1).Value = DOC_TYPE_RI
        wsTemplates.Cells(2, 2).Value = "RepairInstruction.dotx"
        wsTemplates.Cells(3, 1).Value = DOC_TYPE_EA
        wsTemplates.Cells(3, 2).Value = "EngineeringAnalysis.dotx"
    End If

    If Len(CStr(wsStatuses.Cells(2, 1).Value)) = 0 Then
        wsStatuses.Cells(2, 1).Value = STATUS_DRAFT
        wsStatuses.Cells(2, 2).Value = "Editable draft"
        wsStatuses.Cells(3, 1).Value = STATUS_IN_REVIEW
        wsStatuses.Cells(3, 2).Value = "Under checking"
        wsStatuses.Cells(4, 1).Value = STATUS_RELEASED
        wsStatuses.Cells(4, 2).Value = "Released for use"
    End If
End Sub

Private Sub SeedDefaultConfig()
    Dim ws As Worksheet
    Dim basePath As String

    Set ws = ThisWorkbook.Worksheets(SHEET_CFG_APP)
    basePath = ThisWorkbook.Path

    If Len(CStr(ws.Cells(2, 1).Value)) = 0 Then
        ws.Cells(2, 1).Value = "templates_path"
        ws.Cells(2, 2).Value = basePath & Application.PathSeparator & "templates"
    End If

    If Len(CStr(ws.Cells(3, 1).Value)) = 0 Then
        ws.Cells(3, 1).Value = "output_path"
        ws.Cells(3, 2).Value = basePath & Application.PathSeparator & "output"
    End If
End Sub
