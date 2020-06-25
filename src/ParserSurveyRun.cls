VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParserSurveyRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Parsers")
'
' class module: ParserSurveyRun
'
Option Explicit

Private singleSurveyRun As ModelSurveyRun

Private Sub Class_Initialize()
    Set singleSurveyRun = New ModelSurveyRun
End Sub

Public Function parse(ByVal surveyName As String, ByVal participantId As String, ByRef runLines As ModelDataLines) As ModelSurveyRun
    Dim parser As ParserAnswers
    setSurveyRunMetaData surveyName, participantId, runLines
    
    Set parser = New ParserAnswers
    singleSurveyRun.answerCollection = parser.parse(runLines)
    
    If singleSurveyRun.answerCollection.count <> singleSurveyRun.questionCount _
       Then Err.Raise CustomError.SurveyRunError, "ParserSurveyRun.parse", "The question count does not match the question numbers."
    Set parse = singleSurveyRun
    
End Function

Private Sub setSurveyRunMetaData(ByVal surveyName As String, ByVal participantId As String, ByVal dataArray As ModelDataLines)
    Dim splitArray() As String
    Dim startDateString As String
    Dim endDateString As String
    Dim dateLength As Integer
        
    On Error GoTo Catch
    singleSurveyRun.surveyName = surveyName
    singleSurveyRun.participantId = participantId
        
    splitArray = Split(dataArray.answer, ",")
    startDateString = splitArray(0)
    endDateString = splitArray(1)
    
    If Len(startDateString) = 23 Then
        singleSurveyRun.startTime = DateValue(Mid$(startDateString, 1, 10)) + TimeValue("0" + Mid$(startDateString, 12, 7))
    Else
        singleSurveyRun.startTime = DateValue(Mid$(startDateString, 1, 10)) + TimeValue(Mid$(startDateString, 12, 8))
    End If
   
    If Len(endDateString) = 23 Then
        singleSurveyRun.endTime = DateValue(Mid$(endDateString, 1, 10)) + TimeValue("0" + Mid$(endDateString, 12, 7))
    Else
        singleSurveyRun.endTime = DateValue(Mid$(endDateString, 1, 10)) + TimeValue(Mid$(endDateString, 12, 8))
    End If

    ' The first 2 elements are not question numbers and need to adjust from index (ie not "- 2")
    singleSurveyRun.questionCount = UBound(Split(dataArray.header, ",")) - 1

    Exit Sub

Catch:
    ' Catch if the date string is not valid.
     If Err.number = 13 Then Err.Raise CustomError.IncorrectDataFormat, "ParserSurveyRun.setSurveyRunMetaData", "There is an error in the survey run timestamp data."

End Sub


