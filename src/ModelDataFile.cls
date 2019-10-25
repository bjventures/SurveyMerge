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
    mSurveyName = getMetaData(lineArray, "Survey Name", 0)
    mSubjectId = getMetaData(lineArray, "Subject ID", 2)
    mSurveyRunCount = CLng((UBound(lineArray) - FirstAnswerLine - 1) / LinesPerSurveyRun)
End Property

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
    ' TODO: Here need "And dataName <> dataArray(0)" since an earlier version of the Android app does not add quotes to the keyword.
    '       Consider removing in a future version.
    If Chr$(34) & dataName & Chr$(34) <> dataArray(0) And dataName <> dataArray(0) Then
        validateMetaData = False
        Exit Function
    End If
    ' TODO: This check is disable for now, a very early version did not have parentheses. Consider enabling in a future version.
    'If Left$(dataArray(FileCol.metadata - 1), 1) <> Chr$(34) Or Right$(dataArray(1), 1) <> Chr$(34) Then
    '    validateMetaData = False
    '    Exit Function
    'End If
    validateMetaData = True
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

