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
Private lineTimes As String
Private lineArray() As Variant
Private returnedAnswers As Answers
Private accessor As FileAccessor

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set answerParser = New ParserAnswers
    Set accessor = New FileAccessor
    accessor.loadSurveyRunFile "answer-lines"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fakes = Nothing
    Set answerParser = Nothing
    Set accessor = Nothing
    Set returnedAnswers = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
End Sub

'@TestCleanup
Private Sub TestCleanup()
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenValidListAnswers_ShouldReturnCorrectAnswers()
    On Error GoTo TestFail
    lineHeader = accessor.getFileRunLines(1)(0)
    lineAnswers = accessor.getFileRunLines(1)(1)
    lineTimes = accessor.getFileRunLines(1)(2)
    lineArray = Array(lineHeader, lineAnswers, lineTimes)

    Set returnedAnswers = answerParser.parse(lineArray)

    Assert.AreEqual CLng(1), returnedAnswers.count
    Assert.IsTrue TypeOf returnedAnswers.item(1) Is ModelAnswerList

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenValidCheckboxAnswers_ShouldReturnCorrectAnswer()
    On Error GoTo TestFail
    lineHeader = accessor.getFileRunLines(2)(0)
    lineAnswers = accessor.getFileRunLines(2)(1)
    lineTimes = accessor.getFileRunLines(2)(2)
    lineArray = Array(lineHeader, lineAnswers, lineTimes)

    Set returnedAnswers = answerParser.parse(lineArray)

    Assert.AreEqual CLng(3), returnedAnswers.count
    Assert.IsTrue TypeOf returnedAnswers.item(1) Is ModelAnswerCheckbox

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenValidTextAnswer_ShouldReturnCorrectAnswer()
    On Error GoTo TestFail
    lineHeader = accessor.getFileRunLines(3)(0)
    lineAnswers = accessor.getFileRunLines(3)(1)
    lineTimes = accessor.getFileRunLines(3)(2)
    lineArray = Array(lineHeader, lineAnswers, lineTimes)
    
    Set returnedAnswers = answerParser.parse(lineArray)

    Assert.AreEqual CLng(5), returnedAnswers.count
    Assert.IsTrue TypeOf returnedAnswers.item(1) Is ModelAnswerText

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenValidSliderAnswer_ShouldReturnCorrectAnswer()
    On Error GoTo TestFail
    lineHeader = accessor.getFileRunLines(4)(0)
    lineAnswers = accessor.getFileRunLines(4)(1)
    lineTimes = accessor.getFileRunLines(4)(2)
    lineArray = Array(lineHeader, lineAnswers, lineTimes)
    Set returnedAnswers = answerParser.parse(lineArray)

    Assert.AreEqual CLng(3), returnedAnswers.count
    Assert.IsTrue TypeOf returnedAnswers.item(1) Is ModelAnswerSlider

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Parsers")
' If there is no answer, we don't know what type it is, so use the "super" type as a placeholder.
Private Sub parserAnswers_Parse_WhenNoAnswer_ShouldReturnBaseAnswer()
    On Error GoTo TestFail
    lineHeader = accessor.getFileRunLines(5)(0)
    lineAnswers = accessor.getFileRunLines(5)(1)
    lineTimes = accessor.getFileRunLines(5)(2)
    lineArray = Array(lineHeader, lineAnswers, lineTimes)
    Set returnedAnswers = answerParser.parse(lineArray)

    Assert.AreEqual CLng(1), returnedAnswers.count
    Assert.IsTrue TypeOf returnedAnswers.item(1) Is ModelAnswerBase

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenHaveAllTypes_ShouldReturnAllAnswers()
    On Error GoTo TestFail
    lineHeader = accessor.getFileRunLines(6)(0)
    lineAnswers = accessor.getFileRunLines(6)(1)
    lineTimes = accessor.getFileRunLines(6)(2)
    lineArray = Array(lineHeader, lineAnswers, lineTimes)
    Set returnedAnswers = answerParser.parse(lineArray)

    Assert.AreEqual CLng(5), returnedAnswers.count
    Assert.IsTrue TypeOf returnedAnswers.item(1) Is ModelAnswerBase
    Assert.IsTrue TypeOf returnedAnswers.item(2) Is ModelAnswerList
    Assert.IsTrue TypeOf returnedAnswers.item(3) Is ModelAnswerCheckbox
    Assert.IsTrue TypeOf returnedAnswers.item(4) Is ModelAnswerText
    Assert.IsTrue TypeOf returnedAnswers.item(5) Is ModelAnswerSlider

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenWrongNumberOfTimes_ShouldThrow()
    Const ExpectedError As Long = CustomError.IncorrectDataFormat
    Const ExpectedDescription As String = "The number of questions and times does not match."
    On Error GoTo Assert
    lineHeader = accessor.getFileRunLines(7)(0)
    lineAnswers = accessor.getFileRunLines(7)(1)
    lineTimes = accessor.getFileRunLines(7)(2)
    lineArray = Array(lineHeader, lineAnswers, lineTimes)
    Set returnedAnswers = answerParser.parse(lineArray)

    Assert.fail "Expected error was not raised"
    Exit Sub
    
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenLinesHaveInvalidNonNumericAnswer_ShouldThrow()
    Const ExpectedError As Long = CustomError.InvalidQuestionType
    Const ExpectedDescription As String = "The answer text 'a1' is not a valid answer type."
    On Error GoTo Assert
    lineHeader = "Start Time,End Time,1"
    lineHeader = accessor.getFileRunLines(8)(0)
    lineAnswers = accessor.getFileRunLines(8)(1)
    lineTimes = accessor.getFileRunLines(8)(2)
    lineArray = Array(lineHeader, lineAnswers, lineTimes)
    
    Set returnedAnswers = answerParser.parse(lineArray)

    Assert.fail "Expected error was not raised"
    Exit Sub
    
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenLinesHaveInvalidAnswerMissingQuote_ShouldThrow()
    Const ExpectedError As Long = CustomError.InvalidQuestionType
    Const ExpectedDescription As String = "The answer text '""Test' is not a valid answer type."
    On Error GoTo Assert
    lineHeader = "Start Time,End Time,1"
    lineHeader = accessor.getFileRunLines(9)(0)
    lineAnswers = accessor.getFileRunLines(9)(1)
    lineTimes = accessor.getFileRunLines(9)(2)
    lineArray = Array(lineHeader, lineAnswers, lineTimes)

    Set returnedAnswers = answerParser.parse(lineArray)

    Assert.fail "Expected error was not raised"
    Exit Sub
    
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenLinesHaveInvalidAnswerShort_ShouldThrow()
    Const ExpectedError As Long = CustomError.InvalidQuestionType
    Const ExpectedDescription As String = "The answer text 'a' is not a valid answer type."

    On Error GoTo Assert
    lineHeader = "Start Time,End Time,1"
    lineHeader = accessor.getFileRunLines(10)(0)
    lineAnswers = accessor.getFileRunLines(10)(1)
    lineTimes = accessor.getFileRunLines(10)(2)
    lineArray = Array(lineHeader, lineAnswers, lineTimes)
    
    Set returnedAnswers = answerParser.parse(lineArray)

    Assert.fail "Expected error was not raised"
    Exit Sub
    
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'
' Test errors thrown in models
'

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenLinesHaveNegativeListAnswer_ShouldThrow()
    Const ExpectedError As Long = CustomError.ModelValidationError
    Const ExpectedDescription As String = "The value '-1' is not valid."

    On Error GoTo Assert
    lineHeader = "Start Time,End Time,1"
    lineHeader = accessor.getFileRunLines(11)(0)
    lineAnswers = accessor.getFileRunLines(11)(1)
    lineTimes = accessor.getFileRunLines(11)(2)
    lineArray = Array(lineHeader, lineAnswers, lineTimes)
    
    Set returnedAnswers = answerParser.parse(lineArray)

    Assert.fail "Expected error was not raised"
    Exit Sub
    
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenLinesHaveNegativeCheckboxAnswer_ShouldThrow()
    Const ExpectedError As Long = 13
    Const ExpectedDescription As String = "Type mismatch"

    On Error GoTo Assert
    lineHeader = "Start Time,End Time,1"
    lineHeader = accessor.getFileRunLines(12)(0)
    lineAnswers = accessor.getFileRunLines(12)(1)
    lineTimes = accessor.getFileRunLines(12)(2)
    lineArray = Array(lineHeader, lineAnswers, lineTimes)
    
    Set returnedAnswers = answerParser.parse(lineArray)

    Assert.fail "Expected error was not raised"
    Exit Sub
    
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenLinesHaveSliderAnswerGreaterThanOne_ShouldThrow()
    Const ExpectedError As Long = CustomError.ModelValidationError
    Const ExpectedDescription As String = "The value '-1.2' is not valid."

    On Error GoTo Assert
    lineHeader = "Start Time,End Time,1"
    lineHeader = accessor.getFileRunLines(13)(0)
    lineAnswers = accessor.getFileRunLines(13)(1)
    lineTimes = accessor.getFileRunLines(13)(2)
    lineArray = Array(lineHeader, lineAnswers, lineTimes)
    
    Set returnedAnswers = answerParser.parse(lineArray)

    Assert.fail "Expected error was not raised"
    Exit Sub
    
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenLinesHaveSliderAnswerLessThanZero_ShouldThrow()
    Const ExpectedError As Long = CustomError.ModelValidationError
    Const ExpectedDescription As String = "The value '-1.2' is not valid."

    On Error GoTo Assert
    lineHeader = "Start Time,End Time,1"
    lineHeader = accessor.getFileRunLines(13)(0)
    lineAnswers = accessor.getFileRunLines(13)(1)
    lineTimes = accessor.getFileRunLines(13)(2)
    lineArray = Array(lineHeader, lineAnswers, lineTimes)
    
    Set returnedAnswers = answerParser.parse(lineArray)

    Assert.fail "Expected error was not raised"
    Exit Sub
    
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

