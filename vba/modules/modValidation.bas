Attribute VB_Name = "modValidation"
Option Explicit

Public Function ValidateBeforeRelease(ByVal card As clsDocumentCard) As Collection
    Dim issues As Collection
    Set issues = New Collection

    ValidateRequiredFields card, issues
    ValidateWordDocument card, issues

    If card.DocumentType = DOC_TYPE_EA Then
        ValidateEAClauseMatrix card, issues
    ElseIf card.DocumentType = DOC_TYPE_RI Then
        ValidateRISectionMatrix card, issues
    End If

    Set ValidateBeforeRelease = issues
End Function

Private Sub ValidateRequiredFields(ByVal card As clsDocumentCard, ByRef issues As Collection)
    If Len(card.DocumentID) = 0 Then AddIssue issues, ISSUE_SEVERITY_ERROR, "CARD_REQUIRED", "document_id is required"
    If Len(card.DocumentType) = 0 Then AddIssue issues, ISSUE_SEVERITY_ERROR, "CARD_REQUIRED", "document_type is required"
    If Len(card.Title) = 0 Then AddIssue issues, ISSUE_SEVERITY_ERROR, "CARD_REQUIRED", "title is required"
    If Len(card.Author) = 0 Then AddIssue issues, ISSUE_SEVERITY_ERROR, "CARD_REQUIRED", "author is required"
    If Len(card.WordDocPath) = 0 Then AddIssue issues, ISSUE_SEVERITY_ERROR, "WORD_DOC", "word_doc_path is empty"
End Sub

Private Sub ValidateWordDocument(ByVal card As clsDocumentCard, ByRef issues As Collection)
    Dim textBody As String

    If Len(card.WordDocPath) = 0 Then Exit Sub
    If Dir$(card.WordDocPath) = "" Then
        AddIssue issues, ISSUE_SEVERITY_ERROR, "WORD_DOC", "Word file not found: " & card.WordDocPath
        Exit Sub
    End If

    textBody = ReadDocText(card.WordDocPath)

    If InStr(1, textBody, "{{", vbTextCompare) > 0 Then
        AddIssue issues, ISSUE_SEVERITY_ERROR, "UNRESOLVED_MARKER", "Document contains unresolved markers {{...}}"
    End If

    If ContainsTrashPlaceholder(textBody) Then
        AddIssue issues, ISSUE_SEVERITY_WARNING, "TRASH_PLACEHOLDER", "Document contains TBD/XXX/???/sample/draft"
    End If

    ValidateRequiredSectionsInText card, textBody, issues
End Sub

Private Function ReadDocText(ByVal docxPath As String) As String
    Dim wordApp As Object
    Dim wordDoc As Object

    On Error GoTo CleanUp

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Open(docxPath, False, True)
    ReadDocText = CStr(wordDoc.Content.Text)

CleanUp:
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Function

Private Function ContainsTrashPlaceholder(ByVal textBody As String) As Boolean
    Dim lowered As String
    lowered = LCase$(textBody)

    ContainsTrashPlaceholder = (InStr(lowered, "tbd") > 0) _
        Or (InStr(lowered, "xxx") > 0) _
        Or (InStr(lowered, "???") > 0) _
        Or (InStr(lowered, "sample") > 0) _
        Or (InStr(lowered, "draft") > 0)
End Function

Private Sub ValidateRequiredSectionsInText(ByVal card As clsDocumentCard, ByVal textBody As String, ByRef issues As Collection)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIdx As Long
    Dim targetType As String
    Dim sectionTitle As String
    Dim mandatoryFlag As String

    Set ws = ThisWorkbook.Worksheets(SHEET_REQUIRED_SECTIONS)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For rowIdx = 2 To lastRow
        targetType = CStr(ws.Cells(rowIdx, 1).Value)
        sectionTitle = CStr(ws.Cells(rowIdx, 2).Value)
        mandatoryFlag = UCase$(CStr(ws.Cells(rowIdx, 3).Value))

        If targetType = card.DocumentType And mandatoryFlag = "YES" Then
            If InStr(1, textBody, sectionTitle, vbTextCompare) = 0 Then
                AddIssue issues, ISSUE_SEVERITY_ERROR, "MISSING_SECTION", "Missing required section: " & sectionTitle
            End If
        End If
    Next rowIdx
End Sub

Private Sub ValidateEAClauseMatrix(ByVal card As clsDocumentCard, ByRef issues As Collection)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIdx As Long
    Dim applicability As String
    Dim statusValue As String
    Dim means As String
    Dim coveredSection As String

    Set ws = ThisWorkbook.Worksheets(SHEET_EA_MATRIX)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For rowIdx = 2 To lastRow
        If CStr(ws.Cells(rowIdx, 1).Value) = card.DocumentID Then
            applicability = UCase$(CStr(ws.Cells(rowIdx, 4).Value))
            statusValue = CStr(ws.Cells(rowIdx, 8).Value)
            means = CStr(ws.Cells(rowIdx, 5).Value)
            coveredSection = CStr(ws.Cells(rowIdx, 6).Value)

            If applicability = "YES" Then
                If Len(statusValue) = 0 Then AddIssue issues, ISSUE_SEVERITY_ERROR, "EA_MATRIX", "Applicable clause without status in row " & rowIdx
                If Len(means) = 0 Then AddIssue issues, ISSUE_SEVERITY_ERROR, "EA_MATRIX", "Applicable clause without means_of_compliance in row " & rowIdx
                If Len(coveredSection) = 0 Then AddIssue issues, ISSUE_SEVERITY_ERROR, "EA_MATRIX", "Applicable clause without covered_in_section in row " & rowIdx
            End If
        End If
    Next rowIdx
End Sub

Private Sub ValidateRISectionMatrix(ByVal card As clsDocumentCard, ByRef issues As Collection)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIdx As Long
    Dim mandatoryFlag As String
    Dim presentFlag As String
    Dim sectionTitle As String

    Set ws = ThisWorkbook.Worksheets(SHEET_RI_MATRIX)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For rowIdx = 2 To lastRow
        If CStr(ws.Cells(rowIdx, 1).Value) = card.DocumentID Then
            sectionTitle = CStr(ws.Cells(rowIdx, 3).Value)
            mandatoryFlag = UCase$(CStr(ws.Cells(rowIdx, 4).Value))
            presentFlag = UCase$(CStr(ws.Cells(rowIdx, 6).Value))

            If mandatoryFlag = "YES" And presentFlag <> "YES" Then
                AddIssue issues, ISSUE_SEVERITY_ERROR, "RI_MATRIX", "Mandatory RI section missing: " & sectionTitle
            End If
        End If
    Next rowIdx
End Sub

Private Sub AddIssue(ByRef issues As Collection, ByVal severity As String, ByVal code As String, ByVal message As String)
    Dim issue As clsValidationIssue
    Set issue = New clsValidationIssue

    issue.Severity = severity
    issue.Code = code
    issue.Message = message

    issues.Add issue
End Sub
