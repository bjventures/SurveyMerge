Attribute VB_Name = "TestSurveyRun"
Option Explicit
Option Private Module

'@TestModule
'@Folder("SurveyMerge.Tests.Models")

Private Assert As Object
Private Fakes As Object
Private surveyRun As ModelSurveyRun

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set surveyRun = New ModelSurveyRun
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set surveyRun = Nothing
End Sub

'@TestMethod("Model")
Private Sub surveyRun_QuestionCount_WhenSetValid_ShouldSet()
    On Error GoTo TestFail
    surveyRun.questionCount = 3
   
    Assert.AreEqual CLng(3), surveyRun.questionCount

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub surveyRun_AnswerCollection_WhenSetValid_ShouldSet()
    On Error GoTo TestFail
    Dim answers As answers
    Set answers = New answers
    Dim listAnswer As ModelAnswerList
    Set listAnswer = New ModelAnswerList
    answers.Add listAnswer
    Assert.AreEqual CLng(1), answers.count
    
    surveyRun.answerCollection = answers
     
    Assert.AreEqual CLng(1), surveyRun.answerCollection.count
    
    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub





' TODO: NEED TO FINISH THIS

'Public Property Get questionCount() As Long
'    questionCount = mQuestionCount
'End Property
'
'Public Property Let questionCount(ByVal value As Long)
'    mQuestionCount = value
'End Property
'
'Public Property Get surveyName() As String
'    surveyName = mSurveyName
'End Property
'
'Public Property Let surveyName(ByVal value As String)
'    mSurveyName = value
'End Property
'
'Public Property Get participantId() As String
'    participantId = mParticipantId
'End Property
'
'Public Property Let participantId(ByVal value As String)
'    mParticipantId = value
'End Property
'
'Public Property Get startTime() As Date
'    startTime = mStartTime
'End Property
'
'Public Property Let startTime(ByVal value As Date)
'    mStartTime = value
'End Property
'
'Public Property Get endTime() As Date
'    endTime = mEndTime
'End Property
'
'Public Property Let endTime(ByVal value As Date)
'    mEndTime = value
'End Property
'
'Public Property Get answerCollection() As Answers
'    Set answerCollection = mAnswerCollection
'End Property
'
'Public Property Let answerCollection(ByRef value As Answers)
'    Set mAnswerCollection = value
'End Property
'

