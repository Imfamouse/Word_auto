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

    Dim i As Long
    Dim lbl As MSForms.Label
    Dim tb As MSForms.TextBox
    Dim topPos As Single

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

    Me.Caption = "Document Card"
    Me.Width = LEFT_INPUT + INPUT_W + 24
    Me.Height = TOP_START + UBound(mFields) * ROW_H + 42
    mUiBuilt = True
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
