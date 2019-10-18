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

'Private assertion As Boolean
'Private parser As ParserSurveyRun
'Private singleSurveyRun As ModelSurveyRun


'
'Public Function test_ParserSurveyRun_WhenCorrectData_ShouldParseAnswers() As Boolean
'
'    Set singleSurveyRun = parser.parse("name", "participant id", getAnswerLines("test-109"))
'    Dim answerCollection As Answers
'
'    Set answerCollection = singleSurveyRun.answerCollection
'    assertion = answerCollection.count = 5
'    test_ParserSurveyRun_WhenCorrectData_ShouldParseAnswers = assertion
'
'End Function
'
'Public Function test_ParserSurveyRun_WhenIncorrectNumberCount_ShouldThrow() As Boolean
'
'On Error GoTo Catch
'    Set singleSurveyRun = parser.parse("name", "participant id", getAnswerLines("test-114"))
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.SurveyRunError
'    test_ParserSurveyRun_WhenIncorrectNumberCount_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
''@Ignore UnderscoreInPublicClassModuleMember
'Public Function test_ParserSurveyRun_WhenIncorrectFileData_ShouldThrow() As Boolean
'
'On Error GoTo Catch
'    Set singleSurveyRun = parser.parse("name", "participant id", getAnswerLines("test-115"))
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.IncorrectDataFormat
'    test_ParserSurveyRun_WhenIncorrectFileData_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
'






