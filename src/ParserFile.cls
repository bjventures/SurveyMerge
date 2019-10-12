VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParserFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' class module: ParserFile
'
Option Explicit

'
' This function parses and prints the SurveyRuns in a single file.
' Note errors in individual SurveyRuns are handled by printing the error, they are not propagated.
'
Public Function parse(filePath As String, fileName As String, printer As IPrinter, startLine As Integer) As Integer

On Error GoTo Catch
    
    Dim fileText As String
    Dim surveyRunCount As Double
    Dim i As Integer
    Dim firstRow As Integer
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
        Call parseAndPrintSurveyRun(surveyName, subjectId, runLines, parser, surveyRun, printer, startLine + i)
    Next i
    parse = surveyRunCount + startLine
    
Finally:
    Exit Function

Catch:
    ' Need to close file before error gets propagated.
    Call Err.Raise(Err.number, Err.Source, "There is an error in the file '" & fileName & "'. " & Err.description)

End Function

Private Function getUtfFile(filePath As String, fileName As String) As String
    ' Ensure reference is set to Microsoft ActiveX DataObjects library.
    ' Under "tools/references". References travel with the excel file, so once added, no need to worry.
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile (filePath & fileName)
    getUtfFile = objStream.ReadText()

End Function

Private Function getAsciFile(filePath As String, fileName As String) As String
    Dim inputFile As Integer
    Dim file As String
    Dim fileText As String
    inputFile = FreeFile

     file = filePath + fileName
    Open file For Input As #inputFile
    fileText = Input(LOF(inputFile), #inputFile)
    Close #inputFile
    getAsciFile = fileText
    
Finally:
    Exit Function
    
Catch:
    ' Need to close file before error gets propagated.
    Close #inputFile
    Call Err.Raise(Err.number, Err.Source, Err.description)

End Function

'
' This function is temporarily needed as older data files may have line breaks in the middle of text questions.
' The middle lines in the survey run are combined so each SurveyRun will have 3 lines.
' TODO: Since May 2019, line breaks are removed from the PIEL App data file. In late 2020 can remove this.
'
Private Function cleanFileText(fileText As String) As String
    Dim cleanText As String
    Dim questionArray() As String
    Dim lineArray() As String
    Dim i As Integer
    Dim n As Integer
    
    ' Make sure lines are separated by vbLf
    fileText = Replace(fileText, vbCrLf, vbLf)
    ' Just in case there are Mac line breaks
    fileText = Replace(fileText, vbCr, "")
    questionArray = Split(fileText, Chr(34) & "Start Time")

    cleanText = questionArray(0)
    
    For i = 1 To UBound(questionArray)
        cleanText = cleanText & Chr(34) & "Start Time"
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
Private Sub parseAndPrintSurveyRun(surveyName As String, subjectId As String, runLines As Variant, parser As ParserSurveyRun, surveyRun As ModelSurveyRun, printer As IPrinter, lineNo As Integer)
    
    On Error GoTo Catch
    Set surveyRun = parser.parse(surveyName, subjectId, runLines)
    Call printer.printData(surveyRun, lineNo)
    
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
    
    Call printer.printError(msg, lineNo)
    Resume Finally
        
End Sub

Private Function getMetaData(fileLineArray() As String, dataName As String, lineNo As Integer) As String

    Dim dataArray() As String
    dataArray = Split(fileLineArray(lineNo), ",")
    
    If validateMetaData(dataName, dataArray) = False Then
        Call Err.Raise(CustomError.IncorrectDataFormat, "ProjectName.getMetaData", "The value '" & dataName & "' was not found on line " & lineNo & ".")
    End If
    getMetaData = Mid(dataArray(FileCol.metadata - 1), 2, Len(dataArray(1)) - 2)
 
End Function

Private Function validateMetaData(dataName, dataArray) As Boolean

   ' Check there is a value
   If UBound(dataArray) <> 1 Then
        validateMetaData = False
        Exit Function
   End If
   ' Check the name of the value is correct.
    ' Here need "And dataName <> dataArray(0)" since an earlier version of the Android app does not add quotes to the keyword.
    ' Long term can remove this.
   If Chr(34) & dataName & Chr(34) <> dataArray(0) And dataName <> dataArray(0) Then
        validateMetaData = False
        Exit Function
   End If
   ' Check the value is in parentheses
   If Left(dataArray(FileCol.metadata - 1), 1) <> Chr(34) Or Right(dataArray(1), 1) <> Chr(34) Then
        validateMetaData = False
        Exit Function
   End If
   validateMetaData = True
End Function

Private Function getQuestionLineNo(lineNo As Integer) As Integer
    getQuestionLineNo = ((lineNo - FileRow.firstAnswer) Mod 3) + 1
End Function

