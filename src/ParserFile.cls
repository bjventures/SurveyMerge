VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParserFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Parsers")
'
' class module: ParserFile
'
Option Explicit

'
' This function parses and prints the SurveyRuns in a single file.
' Errors in individual SurveyRuns are handled by printing the error, they are not propagated.
'
Public Function parse(ByVal filePath As String, ByVal fileName As String, ByRef printer As IPrinter, ByVal startLine As Long, isFirstFile As Boolean) As Long

    On Error GoTo Catch
    
    Dim fileText As String
    Dim surveyRunCount As Double
    Dim runCounter As Long
    Dim firstRow As Long
    Dim parser As ParserSurveyRun
    Dim surveyName As String
    Dim subjectId As String
    Dim lineArray() As String

    Set parser = New ParserSurveyRun

    #If Mac Then
        If isFirstFile Then
            MsgBox "SurveyMerge is not fully supported on MacOS since some letters of languages with extended characters may not be correctly imported. You should check the imported data carefully.", vbOKOnly, ProjectName
ProjectName
        End If
        fileText = cleanFileText(getAsciiFile(filePath, fileName))
    #Else
        fileText = cleanFileText(getUtfFile(filePath, fileName))
    #End If

    Dim dataFile As ModelDataFile
    Set dataFile = New ModelDataFile
    dataFile.fileContents = fileText

    For runCounter = 1 To dataFile.surveyRunCount
        parseAndPrintSurveyRun dataFile.surveyName, dataFile.subjectId, dataFile.surveyRunLines(runCounter), parser, printer, startLine + runCounter - 1
    Next runCounter
    parse = dataFile.surveyRunCount + startLine
    Exit Function
Catch:
    ' Need to close file before error gets propagated.
    Err.Raise Err.number, Err.Source, "There is an error in the file '" & fileName & "'. " & Err.description

End Function

Private Function getAsciiFile(ByVal filePath As String, ByVal fileName As String) As String
    Dim fileId As Integer
    Dim fullFilePath As String

    fullFilePath = filePath & fileName
    fileId = FreeFile

    Open fullFilePath For Input As fileId
    getAsciiFile = Input(LOF(fileId), fileId)
    Close fileId

End Function

Private Function getUtfFile(ByVal filePath As String, ByVal fileName As String) As String
    ' Ensure reference is set to Microsoft ActiveX DataObjects library.
    ' Under "tools/references". References travel with the excel file, so once added, no need to worry.
    Dim objStream As Variant
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile (filePath & fileName)
    getUtfFile = objStream.ReadText()

End Function

' The middle lines in the survey run are combined so each SurveyRun will have 3 lines.
Private Function cleanFileText(ByRef fileText As String) As String
    Dim cleanText As String
    Dim questionArray() As String
    Dim lineArray() As String
    Dim counter As Long

    Dim n As Long
    
    ' Make sure lines are separated by vbLf
    fileText = Replace(fileText, vbCrLf, vbLf)
    ' Just in case there are Mac line breaks
    fileText = Replace(fileText, vbCr, vbNullString)
    questionArray = Split(fileText, Chr$(34) & "Start Time")

    cleanText = questionArray(0)
    
    For counter = 1 To UBound(questionArray)
        cleanText = cleanText & Chr$(34) & "Start Time"
        lineArray = Split(questionArray(counter), vbLf)
        cleanText = cleanText + lineArray(0) + vbLf
        For n = 1 To UBound(lineArray) - 2
            cleanText = cleanText + lineArray(n)
        Next n
        cleanText = cleanText & vbLf & lineArray(UBound(lineArray) - 1) & vbLf
    Next counter

    cleanFileText = cleanText
    
End Function

'
' This function handles errors in the SurveyRun by printing the error.
'
Private Sub parseAndPrintSurveyRun(ByRef surveyName As String, ByRef subjectId As String, ByRef runLines As ModelDataLines, ByRef parser As ParserSurveyRun, ByRef printer As IPrinter, ByRef lineNo As Long)
    
    On Error GoTo Catch
    Dim surveyRun As ModelSurveyRun
    Set surveyRun = parser.parse(surveyName, subjectId, runLines)
    printer.printData surveyRun, lineNo

Finally:
    Exit Sub
Catch:
    Dim msg As String
    Select Case Err.number
    Case CustomError.InvalidQuestionType
        msg = "The question type is not recognised."
    Case CustomError.SurveyRunError
        msg = "The number of questions is inconsistent."
    Case Else
        msg = "Error " & Err.number & ": " & Err.description
    End Select
    
    printer.printError msg, lineNo
    Resume Finally
        
End Sub

