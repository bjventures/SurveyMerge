VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelDataFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Models")
'
' class module: ModelDataFile
'
Option Explicit
Private mSurveyName As String
Private mSubjectId As String
Private Const FirstAnswerLine = 4
Private Const SurveyNameLine = 1
Private Const SubjectIdLine = 3

Public Property Let fileContents(ByVal fileText As String)
    Dim lineArray() As String
    lineArray = Split(fileText, vbLf)
    mSurveyName = getMetaData(lineArray(SurveyNameLine - 1), "Survey Name")
    mSubjectId = getMetaData(lineArray(SubjectIdLine - 1), "Subject ID")
    
End Property

Private Function getMetaData(ByVal lineText As String, ByVal dataName As String) As String
    Dim dataArray() As String
    dataArray = Split(lineText, ",")
    If dataArray(0) <> dataName And UBound(dataArray) < 2 Then
        Err.Raise CustomError.IncorrectDataFormat, "ModelDataFile.getMetaData", "The value '" & dataName & "' was not found."
    End If
    getMetaData = removeQuotes(dataArray(1))
End Function

Public Property Get surveyName() As String
    surveyName = mSurveyName
End Property

Public Property Get subjectId() As String
    subjectId = mSubjectId
End Property




Private Function removeQuotes(originalString As String) As String
    Dim modifiedString As String
    If Left$(originalString, 1) = Chr$(34) Then modifiedString = Right$(originalString, Len(originalString) - 1)
    If Right$(modifiedString, 1) = Chr$(34) Then modifiedString = Left$(modifiedString, Len(modifiedString) - 1)
    removeQuotes = modifiedString
End Function

'    lineArray = Split(fileText, vbLf)
'
'    surveyRunCount = (UBound(lineArray) - FileRow.firstAnswer) / 3
'
'    surveyName = getMetaData(lineArray, "Survey Name", 0)
'    subjectId = getMetaData(lineArray, "Subject ID", 2)
'
'    For i = 0 To surveyRunCount - 1
'        firstRow = i * 3 + FileRow.firstAnswer
'        ' Create an array of the 3 survey run lines
'        runLines = Array(lineArray(firstRow), lineArray(firstRow + 1), lineArray(firstRow + 2))
'        parseAndPrintSurveyRun surveyName, subjectId, runLines, parser, surveyRun, printer, startLine + i
'    Next i
'    parse = surveyRunCount + startLine
    
    


