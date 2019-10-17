VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParserFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Data.Parsers")
'
' class module: ParserFile
'
Option Explicit

'
' This function parses and prints the SurveyRuns in a single file.
' Errors in individual SurveyRuns are handled by printing the error, they are not propagated.
'
Public Function parse(ByVal filePath As String, ByVal fileName As String, ByRef printer As IPrinter, ByVal startLine As Long) As Long

    On Error GoTo Catch
    
    Dim fileText As String
    Dim surveyRunCount As Double
    Dim i As Long
    Dim firstRow As Long
    Dim runLines As Variant
    Dim parser As ParserSurveyRun
    Dim surveyRun As ModelSurveyRun
    Dim surveyName As String
    Dim subjectId As String
    Dim lineArray() As String
    
    Set parser = New ParserSurveyRun
    
    #If Mac Then
        fileText = cleanFileText(getAsciFile(filePath, fileName))
    #Else
        fileText = cleanFileText(getUtfFile(filePath, fileName))
    #End If

    lineArray = Split(fileText, vbLf)
    
    surveyRunCount = (UBound(lineArray) - FileRow.firstAnswer) / 3
    
    surveyName = getMetaData(lineArray, "Survey Name", 0)
    subjectId = getMetaData(lineArray, "Subject ID", 2)
    
    For i = 0 To surveyRunCount - 1
        firstRow = i * 3 + FileRow.firstAnswer
        ' Create an array of the 3 survey run lines
        runLines = Array(lineArray(firstRow), lineArray(firstRow + 1), lineArray(firstRow + 2))
        parseAndPrintSurveyRun surveyName, subjectId, runLines, parser, surveyRun, printer, startLine + i
    Next i
    parse = surveyRunCount + startLine
    
    Exit Function

Catch:
    ' Need to close file before error gets propagated.
    Err.Raise Err.number, Err.Source, "There is an error in the file '" & fileName & "'. " & Err.description

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
    Dim i As Long
    Dim n As Long
    
    ' Make sure lines are separated by vbLf
    fileText = Replace(fileText, vbCrLf, vbLf)
    ' Just in case there are Mac line breaks
    fileText = Replace(fileText, vbCr, vbNullString)
    questionArray = Split(fileText, Chr$(34) & "Start Time")

    cleanText = questionArray(0)
    
    For i = 1 To UBound(questionArray)
        cleanText = cleanText & Chr$(34) & "Start Time"
        lineArray = Split(questionArray(i), vbLf)
        cleanText = cleanText + lineArray(0) + vbLf
        For n = 1 To UBound(lineArray) - 2
            cleanText = cleanText + lineArray(n)
        Next n
        cleanText = cleanText & vbLf & lineArray(UBound(lineArray) - 1) & vbLf
    Next i

    cleanFileText = cleanText
    
End Function

'
' This function handles errors in the SurveyRun by printing the error.
'
Private Sub parseAndPrintSurveyRun(ByRef surveyName As String, ByRef subjectId As String, ByRef runLines As Variant, ByRef parser As ParserSurveyRun, ByRef surveyRun As ModelSurveyRun, ByRef printer As IPrinter, ByRef lineNo As Long)
    
    On Error GoTo Catch
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

Private Function getMetaData(ByRef fileLineArray() As String, ByVal dataName As String, ByRef lineNo As Long) As String

    Dim dataArray() As String
    dataArray = Split(fileLineArray(lineNo), ",")
    
    If validateMetaData(dataName, dataArray) = False Then
        Err.Raise CustomError.IncorrectDataFormat, "ParserFile.getMetaData", "The value '" & dataName & "' was not found on line " & lineNo & "."
    End If
    getMetaData = Mid$(dataArray(FileCol.metadata - 1), 2, Len(dataArray(1)) - 2)
 
End Function

Private Function validateMetaData(ByVal dataName As Variant, ByRef dataArray As Variant) As Boolean

    ' Check there is a value
    If UBound(dataArray) <> 1 Then
        validateMetaData = False
        Exit Function
    End If
    ' Check the name of the value is correct.
    ' Here need "And dataName <> dataArray(0)" since an earlier version of the Android app does not add quotes to the keyword.
    ' Long term can remove this.
    If Chr$(34) & dataName & Chr$(34) <> dataArray(0) And dataName <> dataArray(0) Then
        validateMetaData = False
        Exit Function
    End If
    ' Check the value is in parentheses
    If Left$(dataArray(FileCol.metadata - 1), 1) <> Chr$(34) Or Right$(dataArray(1), 1) <> Chr$(34) Then
        validateMetaData = False
        Exit Function
    End If
    validateMetaData = True
End Function

