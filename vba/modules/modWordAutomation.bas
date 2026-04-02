Attribute VB_Name = "modWordAutomation"
Option Explicit

Public Function CreateDocumentFromTemplate(ByVal card As clsDocumentCard) As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim templatePath As String
    Dim outputDir As String
    Dim outputDocx As String
    Dim errText As String

    On Error GoTo ErrHandler

    templatePath = GetTemplatePathByDocumentType(card.DocumentType)
    If Dir$(templatePath) = "" Then Err.Raise vbObjectError + 1200, "CreateDocumentFromTemplate", "Template file not found: " & templatePath

    outputDir = GetConfigValue("output_path")
    EnsureDirectoryExists outputDir

    outputDocx = outputDir & Application.PathSeparator & BuildOutputFileName(card, "docx")

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False

    ' All templates are proper OOXML (.dotx/.docx). Open as template to get a new document.
    If IsOoxmlFile(templatePath) Then
        Set wordDoc = wordApp.Documents.Add(templatePath)
    Else
        Err.Raise vbObjectError + 1201, "CreateDocumentFromTemplate", "Template is not a valid OOXML file: " & templatePath
    End If

    ReplaceAllMarkers wordDoc, card
    wordDoc.SaveAs2 outputDocx, WORD_FORMAT_DOCX

    CreateDocumentFromTemplate = outputDocx

CleanUp:
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
    If Len(errText) > 0 Then Err.Raise vbObjectError + 1299, "CreateDocumentFromTemplate", errText
    Exit Function

ErrHandler:
    errText = "Failed to create DOCX. " & Err.Description
    Resume CleanUp
End Function

Public Function ExportDocumentToPdf(ByVal docxPath As String) As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim pdfPath As String
    Dim errText As String

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
    If Len(errText) > 0 Then Err.Raise vbObjectError + 1300, "ExportDocumentToPdf", errText
    Exit Function

ErrHandler:
    errText = "PDF export failed. " & Err.Description
    Resume CleanUp
End Function

Private Sub ReplaceAllMarkers(ByVal wordDoc As Object, ByVal card As clsDocumentCard)
    ' Document identity
    ReplaceText wordDoc, "{{DocumentID}}",              card.DocumentID
    ReplaceText wordDoc, "{{Title}}",                   card.Title

    ' Aircraft data
    ReplaceText wordDoc, "{{AircraftModel}}",            card.AircraftModel
    ReplaceText wordDoc, "{{AircraftVariant}}",          card.AircraftVariant
    ReplaceText wordDoc, "{{AircraftNumber}}",           card.AircraftNumber
    ReplaceText wordDoc, "{{MSN}}",                      card.MSN
    ReplaceText wordDoc, "{{AircraftManufactureDate}}",  card.AircraftManufactureDate
    ReplaceText wordDoc, "{{AircraftHours}}",            card.AircraftHours
    ReplaceText wordDoc, "{{AircraftCycles}}",           card.AircraftCycles

    ' Component data
    ReplaceText wordDoc, "{{AssemblyNumber}}",           card.AssemblyNumber
    ReplaceText wordDoc, "{{PartNumber}}",               card.PartNumber
    ReplaceText wordDoc, "{{ComponentName}}",            card.ComponentName
    ReplaceText wordDoc, "{{ComponentSN}}",              card.ComponentSN
    ReplaceText wordDoc, "{{ComponentHours}}",           card.ComponentHours
    ReplaceText wordDoc, "{{ComponentCycles}}",          card.ComponentCycles
    ReplaceText wordDoc, "{{ComponentManufactureDate}}", card.ComponentManufactureDate

    ' Document metadata
    ReplaceText wordDoc, "{{Revision}}",                 card.Revision
    ReplaceText wordDoc, "{{Date}}",                     card.DocDate
    ReplaceText wordDoc, "{{DocDate}}",                  card.DocDate
    ReplaceText wordDoc, "{{Applicability}}",            card.Applicability

    ' Responsible persons
    ReplaceText wordDoc, "{{Author}}",                   card.Author
    ReplaceText wordDoc, "{{Checker}}",                  card.Checker
    ReplaceText wordDoc, "{{Approver}}",                 card.Approver
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
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .Execute Replace:=2
    End With
End Sub

Private Function BuildOutputFileName(ByVal card As clsDocumentCard, ByVal extensionWithoutDot As String) As String
    BuildOutputFileName = card.DocumentID & "_Rev" & card.Revision & "." & extensionWithoutDot
End Function

Private Sub EnsureDirectoryExists(ByVal folderPath As String)
    Dim fso As Object

    If Len(folderPath) = 0 Then Err.Raise vbObjectError + 1301, "EnsureDirectoryExists", "output_path is empty"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub

Private Function IsOoxmlFile(ByVal filePath As String) As Boolean
    Dim ff As Integer
    Dim sig As String

    ff = FreeFile
    Open filePath For Binary Access Read As #ff
    sig = Space$(2)
    Get #ff, 1, sig
    Close #ff

    IsOoxmlFile = (sig = "PK")
End Function
