VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' class module: TestMain
'
Option Explicit
Implements ITester

Private moduleName As String
Private assertion As Boolean
Private wsAnswers As Worksheet
Private wsAnswerTime As Worksheet

Private Property Get ITester_className() As String
    ITester_className = "TestMain"
End Property

Private Property Get ITester_testList() As Variant

    ITester_testList = Array( _
        "test_WhenMultipleFiles_ShouldMergeAllSurveyRuns", _
        "test_WhenMultipleFilesAndroidAndApple_ShouldMergeAllSurveyRuns", _
        "test_WhenSurveyRunError_ShouldPrintError" _
    )

End Property

Private Function ITester_runTest(methodName As String) As Boolean
    
    If Len(methodName) > 63 Then MsgBox "The method name '" & methodName & "' is too long to run on the Mac os.", vbCritical, ProjectName
    ITester_runTest = CallByName(Me, methodName, VbMethod)

End Function

Private Sub Class_Initialize()
End Sub
   
Private Sub ITester_setUp()
    Application.DisplayAlerts = False
    ' Ignore error if worksheet does not exist.
    On Error Resume Next
    Sheets(getWsName(WsSheet.Answers)).Delete
    Sheets(getWsName(WsSheet.AnswerTime)).Delete
    Application.DisplayAlerts = True
End Sub

Private Sub ITester_breakDown()
End Sub

Public Function test_WhenMultipleFiles_ShouldMergeAllSurveyRuns() As Boolean

    Call combineCsvFiles(getCurrentPath() & "testing/test-files/test-group-1/", False)
    Set wsAnswers = Sheets(getWsName(WsSheet.Answers))
    Set wsAnswerTime = Sheets(getWsName(WsSheet.AnswerTime))

    assertion = wsAnswers.UsedRange.Rows.count = 10 And wsAnswerTime.UsedRange.Rows.count = 10

    test_WhenMultipleFiles_ShouldMergeAllSurveyRuns = assertion

End Function

Public Function test_WhenMultipleFilesAndroidAndApple_ShouldMergeAllSurveyRuns() As Boolean

    Call combineCsvFiles(getCurrentPath() & "testing/test-files/test-group-2/", False)

    Set wsAnswers = Sheets(getWsName(WsSheet.Answers))
    Set wsAnswerTime = Sheets(getWsName(WsSheet.AnswerTime))

    assertion = wsAnswers.UsedRange.Rows.count = 9 And wsAnswerTime.UsedRange.Rows.count = 9

    test_WhenMultipleFilesAndroidAndApple_ShouldMergeAllSurveyRuns = assertion

End Function

Public Function test_WhenSurveyRunError_ShouldPrintError() As Boolean

    Call combineCsvFiles(getCurrentPath() & "testing/test-files/test-group-3/", False)

    Set wsAnswers = Sheets(getWsName(WsSheet.Answers))
    assertion = wsAnswers.Cells(3, 1) = "Error In Survey Run: The question type is not recognised."
    assertion = wsAnswers.Cells(5, 1) = "Error In Survey Run: The number of questions is inconsistent."
    assertion = wsAnswers.Cells(6, 1) = "Error In Survey Run: " & Chr(34) & "Survey Error Name" & Chr(34) & "is not a valid keyword."
    assertion = wsAnswers.Cells(9, 1) = "Error In Survey Run: The question type is not recognised."
    test_WhenSurveyRunError_ShouldPrintError = assertion

End Function

