VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelDataFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Models.DataFile")
'
' class module: ModelDataFile
'
Option Explicit
Private lineArray() As String
Private mSurveyName As String
Private mSubjectId As String
Private mSurveyRunCount As Long
Private Const FirstAnswerLine = 4
Private Const SurveyNameLine = 1
Private Const SubjectIdLine = 3
Private Const LinesPerSurveyRun = 3

Public Property Let fileContents(ByVal fileText As String)
    lineArray = Split(fileText, vbLf)
    mSurveyName = getMetaData(lineArray(SurveyNameLine - 1), "Survey Name")
    mSubjectId = getMetaData(lineArray(SubjectIdLine - 1), "Subject ID")
    mSurveyRunCount = CLng((UBound(lineArray) - FirstAnswerLine - 1) / LinesPerSurveyRun)
End Property

Private Function getMetaData(ByVal lineText As String, ByVal dataName As String) As String
    Dim lineArray() As String
    lineArray = Split(lineText, ",")
    If lineArray(0) <> dataName And UBound(lineArray) < 2 Then
        Err.Raise CustomError.IncorrectDataFormat, "ModelDataFile.getMetaData", "The value '" & dataName & "' was not found."
    End If
    getMetaData = removeQuotes(lineArray(1))
End Function

Public Property Get surveyName() As String
    surveyName = mSurveyName
End Property

Public Property Get subjectId() As String
    subjectId = mSubjectId
End Property

Public Property Get surveyRunCount() As Long
    surveyRunCount = mSurveyRunCount
End Property

Public Function surveyRunLines(ByVal runNumber As String) As ModelDataLines
    If runNumber < 1 Or runNumber > mSurveyRunCount Then Err.Raise CustomError.IncorrectDataFormat, "ModelDataFile.surveyRunLines", "The value for 'runNumber' is not valid."
    Dim startLineNumber As Long
    Dim rows As New ModelDataLines
    startLineNumber = FirstAnswerLine + ((runNumber - 1) * 3)
    rows.header = lineArray(startLineNumber)
    rows.answer = lineArray(startLineNumber + 1)
    rows.timeStamp = lineArray(startLineNumber + 2)
    Set surveyRunLines = rows
End Function

Private Function removeQuotes(originalString As String) As String
    Dim modifiedString As String
    If Left$(originalString, 1) = Chr$(34) Then modifiedString = Right$(originalString, Len(originalString) - 1)
    If Right$(modifiedString, 1) = Chr$(34) Then modifiedString = Left$(modifiedString, Len(modifiedString) - 1)
    removeQuotes = modifiedString
End Function
