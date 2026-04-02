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
Private mUiBuilt As Boolean

Private Sub UserForm_Initialize()
    BuildDynamicLayout
End Sub

Public Sub LoadIssues(ByVal documentId As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    If Not mUiBuilt Then BuildDynamicLayout

    mList.Clear
    mHeader.Caption = "Validation Report: " & documentId

    Set ws = ThisWorkbook.Worksheets(SHEET_VALIDATION)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If CStr(ws.Cells(r, 1).Value) = documentId Then
            mList.AddItem CStr(ws.Cells(r, 2).Value) & " | " & CStr(ws.Cells(r, 3).Value) & " | " & CStr(ws.Cells(r, 4).Value)
        End If
    Next r

    If mList.ListCount = 0 Then
        mList.AddItem "No issues found"
    End If
End Sub

Private Sub BuildDynamicLayout()
    Set mHeader = Me.Controls.Add("Forms.Label.1", "lbl_header", True)
    mHeader.Left = 12
    mHeader.Top = 12
    mHeader.Width = 780
    mHeader.Caption = "Validation Report"

    Set mList = Me.Controls.Add("Forms.ListBox.1", "lst_issues", True)
    mList.Left = 12
    mList.Top = 36
    mList.Width = 780
    mList.Height = 220

    Me.Width = 820
    Me.Height = 320
    mUiBuilt = True
End Sub
