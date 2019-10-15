VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelSurveyRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Model")
'
' class module: ModelSurveyRun
'
Option Explicit


Private mQuestionCount As Long
Private mSurveyName As String
Private mParticipantId As String
Private mStartTime As Date
Private mEndTime As Date
Private mAnswerCollection As Answers

Public Property Get questionCount() As Long
    questionCount = mQuestionCount
End Property

Public Property Let questionCount(ByVal value As Long)
    mQuestionCount = value
End Property

Public Property Get surveyName() As String
    surveyName = mSurveyName
End Property

Public Property Let surveyName(ByVal value As String)
    mSurveyName = value
End Property

Public Property Get participantId() As String
    participantId = mParticipantId
End Property

Public Property Let participantId(ByVal value As String)
    mParticipantId = value
End Property

Public Property Get startTime() As Date
    startTime = mStartTime
End Property

Public Property Let startTime(ByVal value As Date)
    mStartTime = value
End Property

Public Property Get endTime() As Date
    endTime = mEndTime
End Property

Public Property Let endTime(ByVal value As Date)
    mEndTime = value
End Property

Public Property Get answerCollection() As Answers
    answerCollection = mAnswerCollection
End Property

Public Property Let answerCollection(ByRef value As Answers)
    mAnswerCollection = value
End Property







