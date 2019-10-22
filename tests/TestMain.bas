Attribute VB_Name = "TestMain"
Option Explicit
Option Private Module
'TODO: Need to do these tests

'@TestModule
'@Folder("Tests.Controllers")

Private Assert As Object
Private Fakes As Object
Private wsAnswers As Worksheet
Private wstime As Worksheet

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
    clearSpreadsheet getWsName(WsSheet.Answers)
End Sub

'@TestCleanup
Private Sub TestCleanup()
End Sub

Private Sub clearSpreadsheet(ByVal name As String)
    On Error GoTo SetupFail
    ThisWorkbook.Sheets(name).Cells.ClearContents
    
    Exit Sub
SetupFail:
    Err.Raise CustomError.SetupError, ProjectName & ".clearSpreadsheet", "The " & name & " worksheet does not exist. Please set up project correctly."
End Sub

'@TestMethod("Model")
Private Sub answerList_Value_WhenTest_Should()
    On Error GoTo TestFail
    
    Assert.Succeed
    
    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'
'Public Function test_WhenMultipleFiles_ShouldMergeAllSurveyRuns() As Boolean
'
'    combineCsvFiles getCurrentPath() & "testing/test-files/test-group-1/", False
'    Set wsAnswers = ThisWorkbook.Sheets(getWsName(WsSheet.Answers))
'    Set wstime = ThisWorkbook.Sheets(getWsName(WsSheet.Times))
'
'    assertion = wsAnswers.UsedRange.Rows.count = 10 And wstime.UsedRange.Rows.count = 10
'
'    test_WhenMultipleFiles_ShouldMergeAllSurveyRuns = assertion
'
'End Function
'
'Public Function test_WhenMultipleFilesAndroidAndApple_ShouldMergeAllSurveyRuns() As Boolean
'
'    combineCsvFiles getCurrentPath() & "testing/test-files/test-group-2/", False
'
'    Set wsAnswers = ThisWorkbook.Sheets(getWsName(WsSheet.Answers))
'    Set wstime = ThisWorkbook.Sheets(getWsName(WsSheet.Times))
'
'    assertion = wsAnswers.UsedRange.Rows.count = 9 And wstime.UsedRange.Rows.count = 9
'
'    test_WhenMultipleFilesAndroidAndApple_ShouldMergeAllSurveyRuns = assertion
'
'End Function
'
'Public Function test_WhenSurveyRunError_ShouldPrintError() As Boolean
'
'    combineCsvFiles getCurrentPath() & "testing/test-files/test-group-3/", False
'
'    Set wsAnswers = ThisWorkbook.Sheets(getWsName(WsSheet.Answers))
'    assertion = wsAnswers.Cells(3, 1) = "Error In Survey Run: The question type is not recognised."
'    assertion = wsAnswers.Cells(5, 1) = "Error In Survey Run: The number of questions is inconsistent."
'    assertion = wsAnswers.Cells(6, 1) = "Error In Survey Run: " & Chr$(34) & "Survey Error Name" & Chr$(34) & "is not a valid keyword."
'    assertion = wsAnswers.Cells(9, 1) = "Error In Survey Run: The question type is not recognised."
'    test_WhenSurveyRunError_ShouldPrintError = assertion
'
'End Function
'
'
'
'




