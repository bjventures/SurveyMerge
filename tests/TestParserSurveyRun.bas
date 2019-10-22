Attribute VB_Name = "TestParserSurveyRun"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Parsers")

Private Assert As Object
Private Fakes As Object
Private parser As ParserSurveyRun
Private singleSurveyRun As ModelSurveyRun
Private accessor As FileAccessor

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set accessor = New FileAccessor
    accessor.loadSurveyRunFile "answer-lines-2"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fakes = Nothing
    Set accessor = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set parser = New ParserSurveyRun
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set parser = Nothing
    Set singleSurveyRun = Nothing
End Sub

Private Function getRunLines(runNumber As Integer) As Variant
    Dim lineHeader As String
    Dim lineAnswers As String
    Dim lineTimes As String
    lineHeader = accessor.getFileRunLines(runNumber)(0)
    lineAnswers = accessor.getFileRunLines(runNumber)(1)
    lineTimes = accessor.getFileRunLines(runNumber)(2)
    getRunLines = Array(lineHeader, lineAnswers, lineTimes)
End Function

' TODO Test Fail
'@TestMethod("Parsers")
Private Sub test_ParserSurveyRun_WhenCorrectData_ShouldParseAnswers()
    On Error GoTo TestFail

    Set singleSurveyRun = parser.parse("name", "participant id", getRunLines(1))
    Dim answerCollection As Answers
    Set answerCollection = singleSurveyRun.answerCollection

    Assert.AreEqual CLng(5), answerCollection.count

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

' TODO Test Fail
'@TestMethod("Parsers")
' If there is no answer, we don't know what type it is, so use the "super" type as a placeholder.
Private Sub test_ParserSurveyRun_WhenIncorrectNumberCount_ShouldThrow()
    Const ExpectedError As Long = CustomError.SurveyRunError
    Const ExpectedDescription As String = "The question count does not match the question numbers."
    On Error GoTo Assert
    Set singleSurveyRun = parser.parse("name", "participant id", getRunLines(2))
    
    Assert.fail "Expected error was not raised"
    Exit Sub

Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Parsers")
Private Sub test_ParserSurveyRun_WhenIncorrectFileData_ShouldThrow()
    Const ExpectedError As Long = CustomError.IncorrectDataFormat
    Const ExpectedDescription As String = "There is an error in the survey run timestamp data."
    On Error GoTo Assert
    Set singleSurveyRun = parser.parse("name", "participant id", getRunLines(3))

    Assert.fail "Expected error was not raised"
    Exit Sub

Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

