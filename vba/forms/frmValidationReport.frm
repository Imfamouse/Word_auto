VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmValidationReport
   Caption         =   "Validation Report"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5000
   StartUpPosition =   1
End
Attribute VB_Name = "frmValidationReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub LoadIssues(ByVal documentId As String)
    Me.Caption = "Validation Report: " & documentId
End Sub
