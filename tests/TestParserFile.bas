Attribute VB_Name = "TestParserFile"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Parsers")

Private Assert As Object
Private Fakes As Object
Private runCount As Long
Private testFolder As String
Private parser As ParserFile
Private mockPrinter As IPrinter
Private mockTestPrinter As TestMockPrinter

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set parser = New ParserFile
    testFolder = getCurrentPath() & "tests/test-files/"
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    Set parser = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set mockTestPrinter = New TestMockPrinter
    ' Cast to specific class.
    Set mockPrinter = mockTestPrinter
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set mockTestPrinter = Nothing
    Set mockPrinter = Nothing
End Sub

'@TestMethod("Parsers")
Private Sub parse_WhenFileCorrect_ShouldReturnQuestionCount()

    runCount = parser.parse(testFolder, "test-1.csv", mockPrinter, 0)
    
    Assert.AreEqual CLng(2), runCount
End Sub

'@TestMethod("Parsers")
Private Sub parse_WhenFileCorrectStartingCountSet_ShouldReturnQuestionCount()
    runCount = parser.parse(testFolder, "test-1.csv", mockPrinter, 3)
    
    Assert.AreEqual CLng(5), runCount
End Sub

'@TestMethod("Parsers")
Private Sub parse_WhenFileCorrect_ShouldCallPrinterForEachSurveyRun()

    runCount = parser.parse(testFolder, "test-1.csv", mockPrinter, 0)
    
    Assert.AreEqual CLng(2), mockTestPrinter.validSurveyRunNumber
End Sub

'@TestMethod("Parsers")
Private Sub parse_WhenFileWhenHasSubjectId_ShouldSet()
    runCount = parser.parse(testFolder, "test-1.csv", mockPrinter, 0)
    
    Assert.AreEqual "Test ID", mockTestPrinter.surveyRun.participantId
End Sub

'@TestMethod("Parsers")
Private Sub parse_WhenFileWhenHasSurveyName_ShouldSet()
    runCount = parser.parse(testFolder, "test-1.csv", mockPrinter, 0)
    
    Assert.AreEqual "Test 1", mockTestPrinter.surveyRun.surveyName
End Sub

'@TestMethod("Parsers")
Private Sub parse_WhenHasIncorrectSurveyNameString_ShouldThrow()
    Const ExpectedError As Long = CustomError.IncorrectDataFormat
    Const ExpectedDescription As String = "There is an error in the file 'test-2.csv'. The value 'Survey Name' was not found on line 0."
    On Error GoTo Assert
    
    runCount = parser.parse(testFolder, "test-2.csv", mockPrinter, 0)
    
    Assert.fail "Expected error was not raised"
    Exit Sub
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parse_WhenHasIncorrectSubjectIdString_ShouldThrow()
    Const ExpectedError As Long = CustomError.IncorrectDataFormat
    Const ExpectedDescription As String = "There is an error in the file 'test-4.csv'. The value 'Subject ID' was not found on line 2."
    On Error GoTo Assert
    
    runCount = parser.parse(testFolder, "test-4.csv", mockPrinter, 0)
    
    Assert.fail "Expected error was not raised"
    Exit Sub
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Parsers")
Private Sub parse_WhenHasIncorrectFormatSubjectIdString_ShouldThrow()
    Const ExpectedError As Long = CustomError.IncorrectDataFormat
    Const ExpectedDescription As String = "There is an error in the file 'test-4.csv'. The value 'Subject ID' was not found on line 2."
    On Error GoTo Assert
    
    runCount = parser.parse(testFolder, "test-4.csv", mockPrinter, 0)
    
    Assert.fail "Expected error was not raised"
    Exit Sub
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

