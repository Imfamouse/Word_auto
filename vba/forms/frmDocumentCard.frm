VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDocumentCard
   Caption         =   "Document Card"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8400
   StartUpPosition =   1
End
Attribute VB_Name = "frmDocumentCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TField
    Key As String
    Caption As String
End Type

Private mFields() As TField
Private mUiBuilt As Boolean
Private WithEvents mBtnSave As MSForms.CommandButton
Private WithEvents mBtnCreateDoc As MSForms.CommandButton
Private WithEvents mBtnValidate As MSForms.CommandButton
Private WithEvents mBtnExportPdf As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton

Private Sub UserForm_Initialize()
    BuildFields
    BuildDynamicLayout
    LoadActiveRowIntoControls
End Sub

Public Function ReadCardFromForm() As clsDocumentCard
    Dim card As clsDocumentCard

    If Not mUiBuilt Then
        BuildFields
        BuildDynamicLayout
    End If

    Set card = New clsDocumentCard
    card.DocumentID = GetControlText("tb_document_id")
    card.DocumentType = GetControlText("tb_document_type")
    card.Title = GetControlText("tb_title")
    card.AircraftModel = GetControlText("tb_aircraft_model")
    card.AircraftNumber = GetControlText("tb_aircraft_number")
    card.MSN = GetControlText("tb_msn")
    card.AssemblyNumber = GetControlText("tb_assembly_number")
    card.PartNumber = GetControlText("tb_part_number")
    card.ComponentName = GetControlText("tb_component_name")
    card.Applicability = GetControlText("tb_applicability")
    card.Revision = GetControlText("tb_revision")
    card.DocDate = GetControlText("tb_date")
    card.Author = GetControlText("tb_author")
    card.Checker = GetControlText("tb_checker")
    card.Approver = GetControlText("tb_approver")
    card.RelatedAnalysisNumber = GetControlText("tb_related_analysis_number")
    card.RelatedInstructionNumber = GetControlText("tb_related_instruction_number")
    card.References = GetControlText("tb_references")
    card.Attachments = GetControlText("tb_attachments")
    card.Remarks = GetControlText("tb_remarks")
    card.Status = GetControlText("tb_status")
    card.WordDocPath = GetControlText("tb_word_doc_path")
    card.PdfPath = GetControlText("tb_pdf_path")

    Set ReadCardFromForm = card
End Function

Private Sub BuildFields()
    ReDim mFields(1 To 23)

    SetField 1, "document_id", "Document ID"
    SetField 2, "document_type", "Document Type"
    SetField 3, "title", "Title"
    SetField 4, "aircraft_model", "Aircraft Model"
    SetField 5, "aircraft_number", "Aircraft Number"
    SetField 6, "msn", "MSN"
    SetField 7, "assembly_number", "Assembly Number"
    SetField 8, "part_number", "Part Number"
    SetField 9, "component_name", "Component Name"
    SetField 10, "applicability", "Applicability"
    SetField 11, "revision", "Revision"
    SetField 12, "date", "Date"
    SetField 13, "author", "Author"
    SetField 14, "checker", "Checker"
    SetField 15, "approver", "Approver"
    SetField 16, "related_analysis_number", "Related Analysis #"
    SetField 17, "related_instruction_number", "Related Instruction #"
    SetField 18, "references", "References"
    SetField 19, "attachments", "Attachments"
    SetField 20, "remarks", "Remarks"
    SetField 21, "status", "Status"
    SetField 22, "word_doc_path", "Word Doc Path"
    SetField 23, "pdf_path", "PDF Path"
End Sub

Private Sub SetField(ByVal idx As Long, ByVal fieldKey As String, ByVal fieldCaption As String)
    mFields(idx).Key = fieldKey
    mFields(idx).Caption = fieldCaption
End Sub

Private Sub BuildDynamicLayout()
    Const LEFT_LABEL As Single = 12
    Const LEFT_INPUT As Single = 170
    Const TOP_START As Single = 12
    Const ROW_H As Single = 24
    Const INPUT_W As Single = 650
    Const BTN_TOP_GAP As Single = 10

    Dim i As Long
    Dim lbl As MSForms.Label
    Dim tb As MSForms.TextBox
    Dim topPos As Single
    Dim btnTop As Single

    For i = 1 To UBound(mFields)
        topPos = TOP_START + (i - 1) * ROW_H

        Set lbl = Me.Controls.Add("Forms.Label.1", "lbl_" & mFields(i).Key, True)
        lbl.Caption = mFields(i).Caption
        lbl.Left = LEFT_LABEL
        lbl.Top = topPos + 3
        lbl.Width = 150

        Set tb = Me.Controls.Add("Forms.TextBox.1", "tb_" & mFields(i).Key, True)
        tb.Left = LEFT_INPUT
        tb.Top = topPos
        tb.Width = INPUT_W
        tb.Height = 18
    Next i

    btnTop = TOP_START + UBound(mFields) * ROW_H + BTN_TOP_GAP

    Set mBtnSave = Me.Controls.Add("Forms.CommandButton.1", "btn_save", True)
    mBtnSave.Caption = "Save Card"
    mBtnSave.Left = LEFT_INPUT
    mBtnSave.Top = btnTop
    mBtnSave.Width = 100

    Set mBtnCreateDoc = Me.Controls.Add("Forms.CommandButton.1", "btn_create_doc", True)
    mBtnCreateDoc.Caption = "Create DOCX"
    mBtnCreateDoc.Left = LEFT_INPUT + 110
    mBtnCreateDoc.Top = btnTop
    mBtnCreateDoc.Width = 100

    Set mBtnValidate = Me.Controls.Add("Forms.CommandButton.1", "btn_validate", True)
    mBtnValidate.Caption = "Validate"
    mBtnValidate.Left = LEFT_INPUT + 220
    mBtnValidate.Top = btnTop
    mBtnValidate.Width = 100

    Set mBtnExportPdf = Me.Controls.Add("Forms.CommandButton.1", "btn_export_pdf", True)
    mBtnExportPdf.Caption = "Export PDF"
    mBtnExportPdf.Left = LEFT_INPUT + 330
    mBtnExportPdf.Top = btnTop
    mBtnExportPdf.Width = 100

    Set mBtnClose = Me.Controls.Add("Forms.CommandButton.1", "btn_close", True)
    mBtnClose.Caption = "Close"
    mBtnClose.Left = LEFT_INPUT + 440
    mBtnClose.Top = btnTop
    mBtnClose.Width = 100

    Me.Caption = "Document Card"
    Me.Width = LEFT_INPUT + INPUT_W + 24
    Me.Height = btnTop + 52
    mUiBuilt = True
End Sub

Private Sub mBtnSave_Click()
    Dim card As clsDocumentCard

    On Error GoTo ErrHandler
    Set card = ReadCardFromForm()
    SaveDocumentCard card
    LogAction card.DocumentID, "SaveCard", "OK", "Card saved from form"
    MsgBox "Document card saved", vbInformation
    Exit Sub

ErrHandler:
    LogAction "", "SaveCard", "ERROR", Err.Description
    MsgBox "Save failed: " & Err.Description, vbCritical
End Sub

Private Sub mBtnCreateDoc_Click()
    Dim card As clsDocumentCard
    Dim docPath As String

    On Error GoTo ErrHandler
    Set card = ReadCardFromForm()
    SaveDocumentCard card

    docPath = CreateDocumentFromTemplate(card)
    If Len(docPath) = 0 Then Err.Raise vbObjectError + 1601, "mBtnCreateDoc_Click", "DOCX creation failed"

    card.WordDocPath = docPath
    SaveDocumentCard card
    SetControlText "tb_word_doc_path", docPath

    LogAction card.DocumentID, "CreateWordDocument", "OK", docPath
    MsgBox "DOCX created: " & docPath, vbInformation
    Exit Sub

ErrHandler:
    LogAction "", "CreateWordDocument", "ERROR", Err.Description
    MsgBox "Create DOCX failed: " & Err.Description, vbCritical
End Sub

Private Sub mBtnValidate_Click()
    Dim card As clsDocumentCard
    Dim issues As Collection

    On Error GoTo ErrHandler

    Set card = ReadCardFromForm()
    Set issues = ValidateBeforeRelease(card)

    PublishValidationReport card.DocumentID, issues
    frmValidationReport.LoadIssues card.DocumentID
    frmValidationReport.Show

    LogAction card.DocumentID, "ValidateCurrentDocument", "OK", "Issues: " & CStr(issues.Count)
    Exit Sub

ErrHandler:
    LogAction "", "ValidateCurrentDocument", "ERROR", Err.Description
    MsgBox "Validation failed: " & Err.Description, vbCritical
End Sub

Private Sub mBtnExportPdf_Click()
    Dim card As clsDocumentCard
    Dim pdfPath As String

    On Error GoTo ErrHandler

    Set card = ReadCardFromForm()
    If Len(card.WordDocPath) = 0 Then Err.Raise vbObjectError + 1602, "mBtnExportPdf_Click", "Word Doc Path is empty"

    pdfPath = ExportDocumentToPdf(card.WordDocPath)
    If Len(pdfPath) = 0 Then Err.Raise vbObjectError + 1603, "mBtnExportPdf_Click", "PDF export failed"

    card.PdfPath = pdfPath
    SaveDocumentCard card
    SetControlText "tb_pdf_path", pdfPath

    LogAction card.DocumentID, "ExportCurrentToPdf", "OK", pdfPath
    MsgBox "PDF exported: " & pdfPath, vbInformation
    Exit Sub

ErrHandler:
    LogAction "", "ExportCurrentToPdf", "ERROR", Err.Description
    MsgBox "Export PDF failed: " & Err.Description, vbCritical
End Sub

Private Sub mBtnClose_Click()
    Unload Me
End Sub

Private Sub LoadActiveRowIntoControls()
    Dim ws As Worksheet
    Dim rowIdx As Long
    Dim i As Long

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(SHEET_DOC_CARDS)
    rowIdx = ActiveCell.Row
    If rowIdx < 2 Then rowIdx = 2

    For i = 1 To UBound(mFields)
        SetControlText "tb_" & mFields(i).Key, CStr(ws.Cells(rowIdx, i).Value)
    Next i

SafeExit:
End Sub

Private Sub SetControlText(ByVal controlName As String, ByVal valueText As String)
    Me.Controls(controlName).Text = valueText
End Sub

Private Function GetControlText(ByVal controlName As String) As String
    GetControlText = Trim$(CStr(Me.Controls(controlName).Text))
End Function
