Attribute VB_Name = "TestModelSurveyRun"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Models")

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
Private Sub surveyRun_QuestionCount_WhenInvalid_ShouldThrow()
    Const ExpectedError As Long = CustomError.ModelValidationError
    On Error GoTo TestFail
    surveyRun.questionCount = 0
    
Assert:
    Assert.fail "Expected error was not raised"
TestExit:
    Exit Sub
TestFail:
    If Err.number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Model")
Private Sub surveyRun_SurveyName_WhenSetValid_ShouldSet()
    On Error GoTo TestFail
    surveyRun.surveyName = "Test Name"
   
    Assert.AreEqual "Test Name", surveyRun.surveyName

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub surveyRun_ParticipantId_WhenSetValid_ShouldSet()
    On Error GoTo TestFail
    surveyRun.participantId = "Test Id"
   
    Assert.AreEqual "Test Id", surveyRun.participantId

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub surveyRun_StartTime_WhenSetValid_ShouldSet()
    On Error GoTo TestFail
    Dim todayDate As Date
    todayDate = Date
    surveyRun.startTime = todayDate
    
    Assert.AreEqual todayDate, surveyRun.startTime

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub surveyRun_EndTime_WhenSetValid_ShouldSet()
    On Error GoTo TestFail
    Dim todayDate As Date
    todayDate = Date
    surveyRun.endTime = todayDate
    
    Assert.AreEqual todayDate, surveyRun.endTime

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub surveyRun_AnswerCollection_WhenSetValid_ShouldSet()
    On Error GoTo TestFail
    Dim Answers As Answers
    Set Answers = New Answers
    Dim listAnswer As ModelAnswerList
    Set listAnswer = New ModelAnswerList
    Answers.Add listAnswer
    Assert.AreEqual CLng(1), Answers.count
    
    surveyRun.answerCollection = Answers
     
    Assert.AreEqual CLng(1), surveyRun.answerCollection.count
    
    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

