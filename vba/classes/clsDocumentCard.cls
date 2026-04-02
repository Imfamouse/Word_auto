VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsDocumentCard"
Option Explicit

Public DocumentID As String
Public DocumentType As String
Public Title As String
Public AircraftModel As String
Public AircraftNumber As String
Public MSN As String
Public AssemblyNumber As String
Public PartNumber As String
Public ComponentName As String
Public Applicability As String
Public Revision As String
Public DocDate As String
Public Author As String
Public Checker As String
Public Approver As String
Public RelatedAnalysisNumber As String
Public RelatedInstructionNumber As String
Public References As String
Public Attachments As String
Public Remarks As String
Public Status As String
Public WordDocPath As String
Public PdfPath As String

Public Sub LoadFromRow(ByVal ws As Worksheet, ByVal rowIdx As Long)
    DocumentID = CStr(ws.Cells(rowIdx, 1).Value)
    DocumentType = CStr(ws.Cells(rowIdx, 2).Value)
    Title = CStr(ws.Cells(rowIdx, 3).Value)
    AircraftModel = CStr(ws.Cells(rowIdx, 4).Value)
    AircraftNumber = CStr(ws.Cells(rowIdx, 5).Value)
    MSN = CStr(ws.Cells(rowIdx, 6).Value)
    AssemblyNumber = CStr(ws.Cells(rowIdx, 7).Value)
    PartNumber = CStr(ws.Cells(rowIdx, 8).Value)
    ComponentName = CStr(ws.Cells(rowIdx, 9).Value)
    Applicability = CStr(ws.Cells(rowIdx, 10).Value)
    Revision = CStr(ws.Cells(rowIdx, 11).Value)
    DocDate = CStr(ws.Cells(rowIdx, 12).Value)
    Author = CStr(ws.Cells(rowIdx, 13).Value)
    Checker = CStr(ws.Cells(rowIdx, 14).Value)
    Approver = CStr(ws.Cells(rowIdx, 15).Value)
    RelatedAnalysisNumber = CStr(ws.Cells(rowIdx, 16).Value)
    RelatedInstructionNumber = CStr(ws.Cells(rowIdx, 17).Value)
    References = CStr(ws.Cells(rowIdx, 18).Value)
    Attachments = CStr(ws.Cells(rowIdx, 19).Value)
    Remarks = CStr(ws.Cells(rowIdx, 20).Value)
    Status = CStr(ws.Cells(rowIdx, 21).Value)
    WordDocPath = CStr(ws.Cells(rowIdx, 22).Value)
    PdfPath = CStr(ws.Cells(rowIdx, 23).Value)
End Sub

Public Sub SaveToRow(ByVal ws As Worksheet, ByVal rowIdx As Long)
    ws.Cells(rowIdx, 1).Value = DocumentID
    ws.Cells(rowIdx, 2).Value = DocumentType
    ws.Cells(rowIdx, 3).Value = Title
    ws.Cells(rowIdx, 4).Value = AircraftModel
    ws.Cells(rowIdx, 5).Value = AircraftNumber
    ws.Cells(rowIdx, 6).Value = MSN
    ws.Cells(rowIdx, 7).Value = AssemblyNumber
    ws.Cells(rowIdx, 8).Value = PartNumber
    ws.Cells(rowIdx, 9).Value = ComponentName
    ws.Cells(rowIdx, 10).Value = Applicability
    ws.Cells(rowIdx, 11).Value = Revision
    ws.Cells(rowIdx, 12).Value = DocDate
    ws.Cells(rowIdx, 13).Value = Author
    ws.Cells(rowIdx, 14).Value = Checker
    ws.Cells(rowIdx, 15).Value = Approver
    ws.Cells(rowIdx, 16).Value = RelatedAnalysisNumber
    ws.Cells(rowIdx, 17).Value = RelatedInstructionNumber
    ws.Cells(rowIdx, 18).Value = References
    ws.Cells(rowIdx, 19).Value = Attachments
    ws.Cells(rowIdx, 20).Value = Remarks
    ws.Cells(rowIdx, 21).Value = Status
    ws.Cells(rowIdx, 22).Value = WordDocPath
    ws.Cells(rowIdx, 23).Value = PdfPath
End Sub
