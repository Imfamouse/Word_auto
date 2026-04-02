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
Private WithEvents mBtnHelp As MSForms.CommandButton

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
    card.DocumentID               = GetControlText("tb_document_id")
    card.DocumentType             = GetControlText("tb_document_type")
    card.Title                    = GetControlText("tb_title")
    card.AircraftModel            = GetControlText("tb_aircraft_model")
    card.AircraftVariant          = GetControlText("tb_aircraft_variant")
    card.AircraftNumber           = GetControlText("tb_aircraft_number")
    card.MSN                      = GetControlText("tb_msn")
    card.AircraftManufactureDate  = GetControlText("tb_aircraft_manufacture_date")
    card.AircraftHours            = GetControlText("tb_aircraft_hours")
    card.AircraftCycles           = GetControlText("tb_aircraft_cycles")
    card.AssemblyNumber           = GetControlText("tb_assembly_number")
    card.PartNumber               = GetControlText("tb_part_number")
    card.ComponentName            = GetControlText("tb_component_name")
    card.ComponentSN              = GetControlText("tb_component_sn")
    card.ComponentHours           = GetControlText("tb_component_hours")
    card.ComponentCycles          = GetControlText("tb_component_cycles")
    card.ComponentManufactureDate = GetControlText("tb_component_manufacture_date")
    card.Applicability            = GetControlText("tb_applicability")
    card.Revision                 = GetControlText("tb_revision")
    card.DocDate                  = GetControlText("tb_date")
    card.Author                   = GetControlText("tb_author")
    card.Checker                  = GetControlText("tb_checker")
    card.Approver                 = GetControlText("tb_approver")
    card.RelatedAnalysisNumber    = GetControlText("tb_related_analysis_number")
    card.RelatedInstructionNumber = GetControlText("tb_related_instruction_number")
    card.References               = GetControlText("tb_references")
    card.Attachments              = GetControlText("tb_attachments")
    card.Remarks                  = GetControlText("tb_remarks")
    card.Status                   = GetControlText("tb_status")
    card.WordDocPath              = GetControlText("tb_word_doc_path")
    card.PdfPath                  = GetControlText("tb_pdf_path")

    Set ReadCardFromForm = card
End Function

Private Sub BuildFields()
    ReDim mFields(1 To 31)

    SetField 1,  "document_id",               "Document ID"
    SetField 2,  "document_type",             "Document Type"
    SetField 3,  "title",                     "Title"
    SetField 4,  "aircraft_model",            "Aircraft Model (Type)"
    SetField 5,  "aircraft_variant",          "Aircraft Variant/Model"
    SetField 6,  "aircraft_number",           "Aircraft Reg. Number"
    SetField 7,  "msn",                       "MSN"
    SetField 8,  "aircraft_manufacture_date", "Aircraft Manuf. Date"
    SetField 9,  "aircraft_hours",            "Aircraft Hours (FH)"
    SetField 10, "aircraft_cycles",           "Aircraft Cycles (FC)"
    SetField 11, "assembly_number",           "Assembly Number"
    SetField 12, "part_number",               "Part Number"
    SetField 13, "component_name",            "Component Name"
    SetField 14, "component_sn",              "Component S/N"
    SetField 15, "component_hours",           "Component Hours (FH)"
    SetField 16, "component_cycles",          "Component Cycles (FC)"
    SetField 17, "component_manufacture_date","Component Manuf. Date"
    SetField 18, "applicability",             "Applicability"
    SetField 19, "revision",                  "Revision"
    SetField 20, "date",                      "Date"
    SetField 21, "author",                    "Author"
    SetField 22, "checker",                   "Checker"
    SetField 23, "approver",                  "Approver"
    SetField 24, "related_analysis_number",   "Related Analysis #"
    SetField 25, "related_instruction_number","Related Instruction #"
    SetField 26, "references",                "References"
    SetField 27, "attachments",               "Attachments"
    SetField 28, "remarks",                   "Remarks"
    SetField 29, "status",                    "Status"
    SetField 30, "word_doc_path",             "Word Doc Path"
    SetField 31, "pdf_path",                  "PDF Path"
End Sub

Private Sub SetField(ByVal idx As Long, ByVal fieldKey As String, ByVal fieldCaption As String)
    mFields(idx).Key = fieldKey
    mFields(idx).Caption = fieldCaption
End Sub

Private Sub BuildDynamicLayout()
    Const LEFT_LABEL As Single = 12
    Const LEFT_INPUT As Single = 200
    Const TOP_START As Single = 12
    Const ROW_H As Single = 24
    Const INPUT_W As Single = 620
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
        lbl.Width = 180

        Set tb = Me.Controls.Add("Forms.TextBox.1", "tb_" & mFields(i).Key, True)
        tb.Left = LEFT_INPUT
        tb.Top = topPos
        tb.Width = INPUT_W
        tb.Height = 18
        tb.ControlTipText = GetFieldHint(mFields(i).Key)
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

    Set mBtnHelp = Me.Controls.Add("Forms.CommandButton.1", "btn_help", True)
    mBtnHelp.Caption = "Field Help"
    mBtnHelp.Left = LEFT_INPUT + 550
    mBtnHelp.Top = btnTop
    mBtnHelp.Width = 100

    Me.Caption = "Document Card"
    Me.Width = LEFT_INPUT + INPUT_W + 140
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

Private Sub mBtnHelp_Click()
    MsgBox BuildFieldHelpText(), vbInformation, "Field Help"
End Sub

Private Function BuildFieldHelpText() As String
    Dim i As Long
    Dim textOut As String

    For i = 1 To UBound(mFields)
        textOut = textOut & mFields(i).Caption & ": " & GetFieldHint(mFields(i).Key) & vbCrLf
    Next i

    textOut = textOut & vbCrLf & "Buttons:" & vbCrLf
    textOut = textOut & "- Save Card: save card to doc_cards sheet" & vbCrLf
    textOut = textOut & "- Create DOCX: create Word document from template" & vbCrLf
    textOut = textOut & "- Validate: run checks and open report" & vbCrLf
    textOut = textOut & "- Export PDF: export current DOCX to PDF" & vbCrLf
    textOut = textOut & "- Close: close form"

    BuildFieldHelpText = textOut
End Function

Private Function GetFieldHint(ByVal fieldKey As String) As String
    Select Case fieldKey
        Case "document_id":               GetFieldHint = "Unique document number, e.g. RI-2026-001"
        Case "document_type":             GetFieldHint = "Repair Instruction or Engineering Analysis"
        Case "title":                     GetFieldHint = "Short technical title of the repair"
        Case "aircraft_model":            GetFieldHint = "Aircraft type, e.g. RRJ-95, A320"
        Case "aircraft_variant":          GetFieldHint = "Aircraft model/variant, e.g. RRJ-95B100"
        Case "aircraft_number":           GetFieldHint = "Registration number, e.g. RA-89001"
        Case "msn":                       GetFieldHint = "Manufacturer serial number"
        Case "aircraft_manufacture_date": GetFieldHint = "Aircraft manufacture date, e.g. 2015-06-01"
        Case "aircraft_hours":            GetFieldHint = "Total aircraft flight hours"
        Case "aircraft_cycles":           GetFieldHint = "Total aircraft flight cycles"
        Case "assembly_number":           GetFieldHint = "Component assembly number (top-level p/n)"
        Case "part_number":               GetFieldHint = "Damaged part number"
        Case "component_name":            GetFieldHint = "Component name, e.g. PIVOTING DOOR RH"
        Case "component_sn":              GetFieldHint = "Component serial number"
        Case "component_hours":           GetFieldHint = "Component total flight hours"
        Case "component_cycles":          GetFieldHint = "Component total flight cycles"
        Case "component_manufacture_date":GetFieldHint = "Component manufacture date"
        Case "applicability":             GetFieldHint = "Applicability limits / conditions"
        Case "revision":                  GetFieldHint = "Revision index, use – for initial issue"
        Case "date":                      GetFieldHint = "Document date DD.MM.YYYY"
        Case "author":                    GetFieldHint = "Engineer author (Last I.O.)"
        Case "checker":                   GetFieldHint = "Checker name"
        Case "approver":                  GetFieldHint = "Approver name"
        Case "related_analysis_number":   GetFieldHint = "Linked EA document number"
        Case "related_instruction_number":GetFieldHint = "Linked RI document number"
        Case "references":                GetFieldHint = "Referenced documents (SDR, CMM, AMM...)"
        Case "attachments":               GetFieldHint = "Attached files list"
        Case "remarks":                   GetFieldHint = "Operational comments"
        Case "status":                    GetFieldHint = "Draft / In Review / Released"
        Case "word_doc_path":             GetFieldHint = "Generated DOCX path (auto-filled)"
        Case "pdf_path":                  GetFieldHint = "Generated PDF path (auto-filled)"
        Case Else:                        GetFieldHint = ""
    End Select
End Function

Private Sub LoadActiveRowIntoControls()
    Dim ws As Worksheet
    Dim rowIdx As Long
    Dim colMap As Object
    Dim i As Long
    Dim colIdx As Long

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets(SHEET_DOC_CARDS)
    rowIdx = ActiveCell.Row
    If rowIdx < 2 Then rowIdx = 2

    ' Build column index map from header row (column name -> column index)
    Set colMap = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        Dim hdr As String
        hdr = Trim$(LCase$(CStr(ws.Cells(1, i).Value)))
        If Len(hdr) > 0 Then colMap(hdr) = i
    Next i

    ' Load each field by name
    For i = 1 To UBound(mFields)
        Dim key As String
        key = mFields(i).Key
        If colMap.Exists(key) Then
            colIdx = colMap(key)
            SetControlText "tb_" & key, CStr(ws.Cells(rowIdx, colIdx).Value)
        End If
    Next i

SafeExit:
End Sub

Private Sub SetControlText(ByVal controlName As String, ByVal valueText As String)
    Me.Controls(controlName).Text = valueText
End Sub

Private Function GetControlText(ByVal controlName As String) As String
    GetControlText = Trim$(CStr(Me.Controls(controlName).Text))
End Function
