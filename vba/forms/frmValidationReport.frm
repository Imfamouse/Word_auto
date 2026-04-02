VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmValidationReport
   Caption         =   "Validation Report"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8400
   StartUpPosition =   1
End
Attribute VB_Name = "frmValidationReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mList As MSForms.ListBox
Private mHeader As MSForms.Label
Private WithEvents mBtnClose As MSForms.CommandButton
Private mUiBuilt As Boolean

Private Sub UserForm_Initialize()
    BuildDynamicLayout
End Sub

Public Sub LoadIssues(ByVal documentId As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim errCount As Long
    Dim warnCount As Long

    If Not mUiBuilt Then BuildDynamicLayout

    mList.Clear
    errCount = 0
    warnCount = 0

    Set ws = ThisWorkbook.Worksheets(SHEET_VALIDATION)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If CStr(ws.Cells(r, 1).Value) = documentId Then
            Dim sev As String
            sev = CStr(ws.Cells(r, 2).Value)
            If sev = ISSUE_SEVERITY_ERROR Then errCount = errCount + 1
            If sev = ISSUE_SEVERITY_WARNING Then warnCount = warnCount + 1
            mList.AddItem sev & " | " & CStr(ws.Cells(r, 3).Value) & " | " & CStr(ws.Cells(r, 4).Value)
        End If
    Next r

    If mList.ListCount = 0 Then
        mList.AddItem "No issues found — document passed validation"
    End If

    mHeader.Caption = "Validation Report: " & documentId & _
        "    Errors: " & CStr(errCount) & "   Warnings: " & CStr(warnCount)
End Sub

Private Sub BuildDynamicLayout()
    Const FORM_W As Single = 900
    Const LIST_H As Single = 300

    Set mHeader = Me.Controls.Add("Forms.Label.1", "lbl_header", True)
    mHeader.Left = 12
    mHeader.Top = 12
    mHeader.Width = FORM_W - 24
    mHeader.Caption = "Validation Report"

    Set mList = Me.Controls.Add("Forms.ListBox.1", "lst_issues", True)
    mList.Left = 12
    mList.Top = 36
    mList.Width = FORM_W - 24
    mList.Height = LIST_H

    Set mBtnClose = Me.Controls.Add("Forms.CommandButton.1", "btn_close", True)
    mBtnClose.Caption = "Close"
    mBtnClose.Left = FORM_W - 120
    mBtnClose.Top = 36 + LIST_H + 8
    mBtnClose.Width = 100

    Me.Width = FORM_W + 20
    Me.Height = 36 + LIST_H + 52
    mUiBuilt = True
End Sub

Private Sub mBtnClose_Click()
    Unload Me
End Sub
