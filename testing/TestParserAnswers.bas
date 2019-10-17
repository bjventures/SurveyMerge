Attribute VB_Name = "TestParserAnswers"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Parsers")

Private Assert As Object
Private Fakes As Object
Private answerParser As ParserAnswers
Private lineHeader As String
Private lineAnswers As String
Private lineArray() As Variant
Private lineTimes As String
Private returnedAnswers As Answers

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set answerParser = New ParserAnswers
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set returnedAnswers = Nothing
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenLinesHaveListAnswers_ShouldReturnCorrectAnswers()
    On Error GoTo TestFail
    lineHeader = "Start Time,End Time,1,2,3,4,5,6"
    lineAnswers = "2019-10-11T08:57:50+1100,2019-10-11T08:58:26+1100,1,,,2,1,1"
    lineTimes = ",,2019-10-11T08:57:52+1100,Nil,Nil,2019-10-11T08:58:22+1100,2019-10-11T08:58:24+1100,2019-10-11T08:58:25+1100"
    lineArray = Array(lineHeader, lineAnswers, lineTimes)

    Set returnedAnswers = answerParser.parse(lineArray)
    
    Assert.AreEqual CLng(6), returnedAnswers.count
    Assert.IsTrue TypeOf returnedAnswers.item(1) Is ModelAnswerList

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenLinesHaveCheckboxAnswer_ShouldReturnCorrectAnswer()
    On Error GoTo TestFail
    lineHeader = "Start Time,End Time,1"
    lineAnswers = "2019-10-11T08:57:50+1100,2019-10-11T08:58:26+1100," & Chr$(34) & "2,4" & Chr$(34)
    lineTimes = ",,2019-10-11T08:57:52+1100,Nil,Nil,2019-10-11T08:58:22+1100"
    lineArray = Array(lineHeader, lineAnswers, lineTimes)
    Set returnedAnswers = answerParser.parse(lineArray)
    
    Assert.AreEqual CLng(1), returnedAnswers.count
    Assert.IsTrue TypeOf returnedAnswers.item(1) Is ModelAnswerCheckbox

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenLinesHaveTextAnswer_ShouldReturnCorrectAnswer()
    On Error GoTo TestFail
    lineHeader = "Start Time,End Time,1"
    lineAnswers = "2019-10-11T08:57:50+1100,2019-10-11T08:58:26+1100," & Chr$(34) & "Text answer" & Chr$(34)
    lineTimes = ",,2019-10-11T08:57:52+1100,Nil,Nil,2019-10-11T08:58:22+1100"
    lineArray = Array(lineHeader, lineAnswers, lineTimes)
    Set returnedAnswers = answerParser.parse(lineArray)
    
    Assert.AreEqual CLng(1), returnedAnswers.count
    Assert.IsTrue TypeOf returnedAnswers.item(1) Is ModelAnswerText

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenLinesHaveSliderAnswer_ShouldReturnCorrectAnswer()
    On Error GoTo TestFail
    lineHeader = "Start Time,End Time,1"
    lineAnswers = "2019-10-11T08:57:50+1100,2019-10-11T08:58:26+1100,0.25"
    lineTimes = ",,2019-10-11T08:57:52+1100,Nil,Nil,2019-10-11T08:58:22+1100"
    lineArray = Array(lineHeader, lineAnswers, lineTimes)
    Set returnedAnswers = answerParser.parse(lineArray)
    
    Assert.AreEqual CLng(1), returnedAnswers.count
    Assert.IsTrue TypeOf returnedAnswers.item(1) Is ModelAnswerSlider

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenLinesHaveInvalidNonNumericAnswer_ShouldThrow()
    Const ExpectedError As Long = CustomError.InvalidQuestionType
    On Error GoTo TestFail
    lineHeader = "Start Time,End Time,1"
    ' Invalid answer: {a1}
    lineAnswers = "2019-10-11T08:57:50+1100,2019-10-11T08:58:26+1100,a1"
    lineTimes = ",,2019-10-11T08:57:52+1100,Nil,Nil,2019-10-11T08:58:22+1100"
    lineArray = Array(lineHeader, lineAnswers, lineTimes)
    
    Set returnedAnswers = answerParser.parse(lineArray)

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

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenLinesHaveInvalidAnswerMissingQuote_ShouldThrow()
    Const ExpectedError As Long = CustomError.InvalidQuestionType
    On Error GoTo TestFail
    lineHeader = "Start Time,End Time,1"
    ' Invalid answer" {"Test}
    lineAnswers = "2019-10-11T08:57:50+1100,2019-10-11T08:58:26+1100," & Chr$(34) & "Test"
    lineTimes = ",,2019-10-11T08:57:52+1100,Nil,Nil,2019-10-11T08:58:22+1100"
    lineArray = Array(lineHeader, lineAnswers, lineTimes)
    
    Set returnedAnswers = answerParser.parse(lineArray)

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

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenLinesHaveInvalidAnswerShort_ShouldThrow()
    Const ExpectedError As Long = CustomError.InvalidQuestionType
    On Error GoTo TestFail
    lineHeader = "Start Time,End Time,1"
    ' Invalid answer" {a}
    lineAnswers = "2019-10-11T08:57:50+1100,2019-10-11T08:58:26+1100,a"
    lineTimes = ",,2019-10-11T08:57:52+1100,Nil,Nil,2019-10-11T08:58:22+1100"
    lineArray = Array(lineHeader, lineAnswers, lineTimes)
    
    Set returnedAnswers = answerParser.parse(lineArray)

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

