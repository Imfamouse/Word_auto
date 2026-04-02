VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDocumentCard
   Caption         =   "Document Card"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4500
   StartUpPosition =   1
End
Attribute VB_Name = "frmDocumentCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function ReadCardFromForm() As clsDocumentCard
    Dim card As clsDocumentCard
    Dim activeRow As Long
    Dim ws As Worksheet

    Set card = New clsDocumentCard
    Set ws = ThisWorkbook.Worksheets(SHEET_DOC_CARDS)

    activeRow = ActiveCell.Row
    If activeRow < 2 Then activeRow = 2

    card.LoadFromRow ws, activeRow
    Set ReadCardFromForm = card
End Function
