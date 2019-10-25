Attribute VB_Name = "TestParserAnswers"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Parsers")

Private Assert As Object
Private Fakes As Object
Private answerParser As ParserAnswers
Private returnedAnswers As Answers
Private accessor As FileAccessor
Private saveCalcState As Long

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set answerParser = New ParserAnswers
    Set accessor = New FileAccessor
    accessor.loadSurveyRunFile "answer-lines-1"
    saveCalcState = Application.Calculation
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fakes = Nothing
    Set answerParser = Nothing
    Set accessor = Nothing
    Set returnedAnswers = Nothing
    Application.Calculation = saveCalcState
    Application.ScreenUpdating = True
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
    
    Set returnedAnswers = answerParser.parse(accessor.getFileRunLines(1))

    Assert.AreEqual CLng(1), returnedAnswers.count
    Assert.IsTrue TypeOf returnedAnswers.item(1) Is ModelAnswerList

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenValidCheckboxAnswers_ShouldReturnCorrectAnswer()
    On Error GoTo TestFail
    Set returnedAnswers = answerParser.parse(accessor.getFileRunLines(2))

    Assert.AreEqual CLng(3), returnedAnswers.count
    Assert.IsTrue TypeOf returnedAnswers.item(1) Is ModelAnswerCheckbox

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenValidTextAnswer_ShouldReturnCorrectAnswer()
    On Error GoTo TestFail
    Set returnedAnswers = answerParser.parse(accessor.getFileRunLines(3))

    Assert.AreEqual CLng(5), returnedAnswers.count
    Assert.IsTrue TypeOf returnedAnswers.item(1) Is ModelAnswerText

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenValidSliderAnswer_ShouldReturnCorrectAnswer()
    On Error GoTo TestFail
    Set returnedAnswers = answerParser.parse(accessor.getFileRunLines(4))

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
    Set returnedAnswers = answerParser.parse(accessor.getFileRunLines(5))

    Assert.AreEqual CLng(1), returnedAnswers.count
    Assert.IsTrue TypeOf returnedAnswers.item(1) Is ModelAnswerBase

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenHaveAllTypes_ShouldReturnAllAnswers()
    On Error GoTo TestFail
    Set returnedAnswers = answerParser.parse(accessor.getFileRunLines(6))

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
    Const ExpectedDescription As String = "The question count is incorrect."
    On Error GoTo Assert
    Set returnedAnswers = answerParser.parse(accessor.getFileRunLines(7))

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
    Set returnedAnswers = answerParser.parse(accessor.getFileRunLines(8))

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
    Set returnedAnswers = answerParser.parse(accessor.getFileRunLines(9))

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
    Set returnedAnswers = answerParser.parse(accessor.getFileRunLines(10))

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
    Set returnedAnswers = answerParser.parse(accessor.getFileRunLines(11))

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
    Set returnedAnswers = answerParser.parse(accessor.getFileRunLines(12))

    Assert.fail "Expected error was not raised"
    Exit Sub

Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenLinesHaveSliderAnswerGreaterThanOne_ShouldThrow()
    Const ExpectedError As Long = CustomError.ModelValidationError
    Const ExpectedDescription As String = "The value '12' is not valid."
    On Error GoTo Assert
    Set returnedAnswers = answerParser.parse(accessor.getFileRunLines(13))

    Assert.fail "Expected error was not raised"
    Exit Sub

Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parserAnswers_Parse_WhenLinesHaveSliderAnswerLessThanZero_ShouldThrow()
    Const ExpectedError As Long = CustomError.ModelValidationError
    Dim ExpectedDescription As String
    ExpectedDescription = "The value '-1" & Application.International(xlDecimalSeparator) & "2' is not valid."
    On Error GoTo Assert
    Set returnedAnswers = answerParser.parse(accessor.getFileRunLines(14))

    Assert.fail "Expected error was not raised"
    Exit Sub

Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

