VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestParserFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' TODO: Create tests
'@Folder("Tests.Parsers")
''
'' class module: TestParserFile
''
'Option Explicit
'Implements ITester
'
'Private assertion As Boolean
'Private parser As ParserFile
'Private runCount As Long
'Private mockPrinter As IPrinter
'Private mockTestPrinter As TestMockPrinter
'
'Private Property Get ITester_className() As String
'    ITester_className = "TestParserFile"
'End Property
'
'Private Property Get ITester_testList() As Variant
''    ITester_testList = Array( _
''        "test_ParserFile_WhenCorrect_ShouldReturnQuestionCount", _
''        "test_ParserFile_WhenCorrectStartingCountSet_ShouldReturnCount", _
''        "test_ParserFile_WhenCorrect_ShouldCallPrinterForEachSurveyRun", _
''        "test_ParserFile_WhenHasSubjectId_ShouldSet", _
''        "test_ParserFile_WhenHasSurveyName_ShouldSet", _
''        "test_ParserFile_WhenHasIncorrectSurveyNameString_ShouldThrow", _
''        "test_ParserFile_WhenHasIncorrectSubjectIdString_ShouldThrow", _
''        "test_ParserFile_WhenIncorrectFormatSubjectId_ShouldThrow" _
''    )
'    ITester_testList = Array("test_ParserFile_WhenHasSubjectId_ShouldSet")
'
'End Property
'
'Private Function ITester_runTest(ByRef methodName As String) As Boolean
'
'    If Len(methodName) > 63 Then MsgBox "The method name '" & methodName & "' is too long to run on the Mac os.", vbCritical, ProjectName
'    ITester_runTest = CallByName(Me, methodName, VbMethod)
'
'End Function
'
'Private Sub Class_Initialize()
'
'    Set parser = New ParserFile
'
'End Sub
'
'Private Sub ITester_setUp()
'    Set mockTestPrinter = New TestMockPrinter
'    Set mockPrinter = mockTestPrinter
'End Sub
'
'Private Sub ITester_breakDown()
'    Set mockTestPrinter = Nothing
'    Set mockPrinter = Nothing
'End Sub
'
'Public Function test_ParserFile_WhenCorrect_ShouldReturnQuestionCount() As Boolean
'
'    runCount = parser.parse(getTestFilePath(), getTestFileName("test-1"), mockPrinter, 0)
'    assertion = runCount = 2
'    test_ParserFile_WhenCorrect_ShouldReturnQuestionCount = assertion
'
'End Function
'
'Public Function test_ParserFile_WhenCorrectStartingCountSet_ShouldReturnCount() As Boolean
'
'    runCount = parser.parse(getTestFilePath(), getTestFileName("test-1"), mockPrinter, 3)
'    assertion = runCount = 5
'    test_ParserFile_WhenCorrectStartingCountSet_ShouldReturnCount = assertion
'
'End Function
'
'Public Function test_ParserFile_WhenCorrect_ShouldCallPrinterForEachSurveyRun() As Boolean
'
'    runCount = parser.parse(getTestFilePath(), getTestFileName("test-1"), mockPrinter, 0)
'
'    assertion = 2 = mockTestPrinter.validSurveyRunNumber
'    test_ParserFile_WhenCorrect_ShouldCallPrinterForEachSurveyRun = assertion
'
'End Function
'
'Public Function test_ParserFile_WhenHasSubjectId_ShouldSet() As Boolean
'
'    runCount = parser.parse(getTestFilePath(), getTestFileName("test-1"), mockPrinter, 0)
'
'    assertion = "aa" = mockTestPrinter.surveyRun.participantId
'
'    test_ParserFile_WhenHasSubjectId_ShouldSet = assertion
'
'End Function
'
'Public Function test_ParserFile_WhenHasSurveyName_ShouldSet() As Boolean
'
'    runCount = parser.parse(getTestFilePath(), getTestFileName("test-1"), mockPrinter, 0)
'
'    assertion = "Test 1" = mockTestPrinter.surveyRun.surveyName
'    test_ParserFile_WhenHasSurveyName_ShouldSet = assertion
'
'End Function
'
''@Ignore UnderscoreInPublicClassModuleMember
'Public Function test_ParserFile_WhenHasIncorrectSurveyNameString_ShouldThrow() As Boolean
'
'    On Error GoTo Catch
'    runCount = parser.parse(getTestFilePath(), getTestFileName("test-2"), mockPrinter, 0)
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.IncorrectDataFormat
'    test_ParserFile_WhenHasIncorrectSurveyNameString_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
'Public Function test_ParserFile_WhenHasIncorrectSubjectIdString_ShouldThrow() As Boolean
'
'    On Error GoTo Catch
'    runCount = parser.parse(getTestFilePath(), getTestFileName("test-3"), mockPrinter, 0)
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.IncorrectDataFormat
'    test_ParserFile_WhenHasIncorrectSubjectIdString_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
'Public Function test_ParserFile_WhenIncorrectFormatSubjectId_ShouldThrow() As Boolean
'
'    On Error GoTo Catch
'    runCount = parser.parse(getTestFilePath(), getTestFileName("test-4"), mockPrinter, 0)
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.IncorrectDataFormat
'    test_ParserFile_WhenIncorrectFormatSubjectId_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
'
'
'
'





