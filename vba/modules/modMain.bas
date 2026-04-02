Attribute VB_Name = "modMain"
Option Explicit

Public Sub AppInitialize()
    On Error GoTo ErrHandler
    EnsureWorkbookStructure
    EnsureRequiredConfig
    LogAction "", "AppInitialize", "OK", "Workbook structure and configuration validated"
    MsgBox "Initialization completed", vbInformation
    Exit Sub

ErrHandler:
    LogAction "", "AppInitialize", "ERROR", Err.Description
    MsgBox "Initialization failed: " & Err.Description, vbCritical
End Sub

Public Sub OpenDocumentCard()
    frmDocumentCard.Show
End Sub

Public Sub CreateWordDocument()
    Dim card As clsDocumentCard
    Dim docPath As String

    On Error GoTo ErrHandler

    Set card = frmDocumentCard.ReadCardFromForm
    SaveDocumentCard card

    docPath = CreateDocumentFromTemplate(card)
    If Len(docPath) = 0 Then Err.Raise vbObjectError + 1400, "CreateWordDocument", "Failed to create Word document"

    card.WordDocPath = docPath
    SaveDocumentCard card
    LogAction card.DocumentID, "CreateWordDocument", "OK", docPath

    MsgBox "Word document created: " & docPath, vbInformation
    Exit Sub

ErrHandler:
    If Not card Is Nothing Then
        LogAction card.DocumentID, "CreateWordDocument", "ERROR", Err.Description
    Else
        LogAction "", "CreateWordDocument", "ERROR", Err.Description
    End If
    MsgBox "CreateWordDocument failed: " & Err.Description, vbCritical
End Sub

Public Sub ValidateCurrentDocument()
    Dim card As clsDocumentCard
    Dim issues As Collection

    On Error GoTo ErrHandler

    Set card = frmDocumentCard.ReadCardFromForm
    Set issues = ValidateBeforeRelease(card)

    PublishValidationReport card.DocumentID, issues
    frmValidationReport.LoadIssues card.DocumentID

    LogAction card.DocumentID, "ValidateCurrentDocument", "OK", "Issues count: " & CStr(issues.Count)
    frmValidationReport.Show
    Exit Sub

ErrHandler:
    LogAction "", "ValidateCurrentDocument", "ERROR", Err.Description
    MsgBox "Validation failed: " & Err.Description, vbCritical
End Sub

Public Sub ExportCurrentToPdf()
    Dim card As clsDocumentCard
    Dim pdfPath As String

    On Error GoTo ErrHandler

    Set card = frmDocumentCard.ReadCardFromForm
    If Len(card.WordDocPath) = 0 Then Err.Raise vbObjectError + 1401, "ExportCurrentToPdf", "word_doc_path is empty"

    pdfPath = ExportDocumentToPdf(card.WordDocPath)
    If Len(pdfPath) = 0 Then Err.Raise vbObjectError + 1402, "ExportCurrentToPdf", "PDF export failed"

    card.PdfPath = pdfPath
    SaveDocumentCard card

    LogAction card.DocumentID, "ExportCurrentToPdf", "OK", pdfPath
    MsgBox "PDF exported: " & pdfPath, vbInformation
    Exit Sub

ErrHandler:
    If Not card Is Nothing Then
        LogAction card.DocumentID, "ExportCurrentToPdf", "ERROR", Err.Description
    Else
        LogAction "", "ExportCurrentToPdf", "ERROR", Err.Description
    End If
    MsgBox "PDF export failed: " & Err.Description, vbCritical
End Sub
