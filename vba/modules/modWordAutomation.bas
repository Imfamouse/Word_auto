Attribute VB_Name = "modWordAutomation"
Option Explicit

Public Function CreateDocumentFromTemplate(ByVal card As clsDocumentCard) As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim templatePath As String
    Dim outputDir As String
    Dim outputDocx As String

    On Error GoTo ErrHandler

    templatePath = GetTemplatePathByDocumentType(card.DocumentType)
    If Dir$(templatePath) = "" Then Err.Raise vbObjectError + 1200, "CreateDocumentFromTemplate", "Template file not found: " & templatePath

    outputDir = GetConfigValue("output_path")
    If Dir$(outputDir, vbDirectory) = "" Then MkDir outputDir

    outputDocx = outputDir & Application.PathSeparator & BuildOutputFileName(card, "docx")

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Add(templatePath)

    ReplaceAllMarkers wordDoc, card
    wordDoc.SaveAs2 outputDocx, WORD_FORMAT_DOCX

    CreateDocumentFromTemplate = outputDocx

CleanUp:
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Exit Function

ErrHandler:
    CreateDocumentFromTemplate = ""
    Resume CleanUp
End Function

Public Function ExportDocumentToPdf(ByVal docxPath As String) As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim pdfPath As String

    On Error GoTo ErrHandler

    pdfPath = Left$(docxPath, Len(docxPath) - 5) & ".pdf"

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Open(docxPath, False, True)

    wordDoc.ExportAsFixedFormat pdfPath, WORD_FORMAT_PDF
    ExportDocumentToPdf = pdfPath

CleanUp:
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Exit Function

ErrHandler:
    ExportDocumentToPdf = ""
    Resume CleanUp
End Function

Private Sub ReplaceAllMarkers(ByVal wordDoc As Object, ByVal card As clsDocumentCard)
    ReplaceText wordDoc, "{{DocumentID}}", card.DocumentID
    ReplaceText wordDoc, "{{DocumentType}}", card.DocumentType
    ReplaceText wordDoc, "{{Title}}", card.Title
    ReplaceText wordDoc, "{{AircraftNumber}}", card.AircraftNumber
    ReplaceText wordDoc, "{{MSN}}", card.MSN
    ReplaceText wordDoc, "{{AssemblyNumber}}", card.AssemblyNumber
    ReplaceText wordDoc, "{{PartNumber}}", card.PartNumber
    ReplaceText wordDoc, "{{ComponentName}}", card.ComponentName
    ReplaceText wordDoc, "{{Revision}}", card.Revision
    ReplaceText wordDoc, "{{Date}}", card.DocDate
    ReplaceText wordDoc, "{{Author}}", card.Author
    ReplaceText wordDoc, "{{Checker}}", card.Checker
    ReplaceText wordDoc, "{{Approver}}", card.Approver
    ReplaceText wordDoc, "{{Applicability}}", card.Applicability
End Sub

Private Sub ReplaceText(ByVal wordDoc As Object, ByVal findText As String, ByVal replaceText As String)
    With wordDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = 1
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Execute Replace:=2
    End With
End Sub

Private Function BuildOutputFileName(ByVal card As clsDocumentCard, ByVal extensionWithoutDot As String) As String
    BuildOutputFileName = card.DocumentID & "_Rev" & card.Revision & "." & extensionWithoutDot
End Function
