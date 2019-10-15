VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestParserSurveyRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Tests")
''
'' class module: TestParseSurveyRun
''
'Option Explicit
'Implements ITester
'
'Private assertion As Boolean
'Private parser As ParserSurveyRun
'Private singleSurveyRun As ModelSurveyRun
'
'Private Property Get ITester_className() As String
'    ITester_className = "TestParserSurveyRun"
'End Property
'
'Private Property Get ITester_testList() As Variant
'    ITester_testList = Array( _
'        "test_ParserSurveyRun_WhenCorrectData_ShouldParseMetadata", _
'        "test_ParserSurveyRun_WhenCorrectData_ShouldParseAnswers", _
'        "test_ParserSurveyRun_WhenIncorrectNumberCount_ShouldThrow", _
'        "test_ParserSurveyRun_WhenIncorrectFileData_ShouldThrow" _
'    )
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
'    Set parser = New ParserSurveyRun
'End Sub
'
'Public Sub ITester_setUp()
'End Sub
'
'Public Sub ITester_breakDown()
'End Sub
'
'Public Function test_ParserSurveyRun_WhenCorrectData_ShouldParseMetadata() As Boolean
'
'    Set singleSurveyRun = parser.parse("name", "participant id", getAnswerLines("test-109"))
'    assertion = singleSurveyRun.questionCount = 5
'    assertion = assertion And singleSurveyRun.surveyName = "name" And singleSurveyRun.participantId = "participant id"
'    assertion = assertion And singleSurveyRun.startTime = CDate("16/04/2019 3:08:07 PM")
'    assertion = assertion And singleSurveyRun.endTime = CDate("16/04/2019 3:08:14 PM")
'    test_ParserSurveyRun_WhenCorrectData_ShouldParseMetadata = assertion
'
'End Function
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
