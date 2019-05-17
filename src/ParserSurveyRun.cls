VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParserSurveyRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' class module: ParserSurveyRun
'
Option Explicit

Private singleSurveyRun As ModelSurveyRun

Private Sub Class_Initialize()

    Set singleSurveyRun = New ModelSurveyRun

End Sub

Public Function parse(surveyName As String, participantId As String, dataArray As Variant) As ModelSurveyRun
    
    Dim parser As ParserAnswers
    
    Call setSurveyRunMetaData(surveyName, participantId, dataArray)
    
    Set parser = New ParserAnswers
    Set singleSurveyRun.answerCollection = parser.parse(dataArray)
    If singleSurveyRun.answerCollection.count <> singleSurveyRun.questionCount Then
        Call Err.Raise(CustomError.SurveyRunError)
    End If

    Set parse = singleSurveyRun
    
End Function

Private Function setSurveyRunMetaData(surveyName As String, participantId As String, dataArray As Variant)
        
    Dim dateString As String
        
On Error GoTo Catch
    singleSurveyRun.surveyName = surveyName
    singleSurveyRun.participantId = participantId
            
    dateString = Split(dataArray(1), ",")(0)
    singleSurveyRun.startTime = DateValue(Mid(dateString, 1, 10)) + TimeValue(Mid(dateString, 12, 8))
    dateString = Split(dataArray(1), ",")(1)
    singleSurveyRun.endTime = DateValue(Mid(dateString, 1, 10)) + TimeValue(Mid(dateString, 12, 8))
    
    ' Note that first 2 elements are not question numbers and need to adjust from index (ie not "- 2")
    singleSurveyRun.questionCount = UBound(Split(dataArray(0), ",")) - 1

Finally:
    Exit Function
Catch:
    Call Err.Raise(CustomError.IncorrectDataFormat)
    Resume Finally

End Function
