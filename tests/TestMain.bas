Attribute VB_Name = "TestMain"
'@TestModule
'@Folder("Tests.Controllers")

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object
Private sheets As Variant
Private wsAnswers As Worksheet
Private wsTimes As Worksheet

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set wsAnswers = ThisWorkbook.sheets(getWsName(WsSheet.Answers))
    Set wsTimes = ThisWorkbook.sheets(getWsName(WsSheet.Times))
    sheets = Array(getWsName(WsSheet.Answers), getWsName(WsSheet.Times))
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fakes = Nothing
    clearOrAddSpreadsheets (sheets)
End Sub

'@TestInitialize
Private Sub TestInitialize()
    ' In case other tests have created a UsedRange, need to reset it.
    clearOrAddSpreadsheets (sheets)
    wsAnswers.UsedRange.Clear
    wsTimes.UsedRange.Clear
 End Sub

'@TestCleanup
Private Sub TestCleanup()
End Sub

'@TestMethod("Controllers")
Private Sub combineCsvFiles_WhenMultipleFilesOldVersion_ShouldMergeAllSurveyRuns()
    On Error GoTo TestFail
    
    combineCsvFiles getCurrentPath() & TestFolder & "/test-files/test-group-1/", False
     
    Assert.AreEqual CLng(10), wsAnswers.UsedRange.rows.count
    Assert.AreEqual CLng(10), wsTimes.UsedRange.rows.count
    
    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("Controllers")
Private Sub combineCsvFiles_WhenMultipleFiles_ShouldMergeAllSurveyRuns()
    On Error GoTo TestFail
    
    combineCsvFiles getCurrentPath() & TestFolder & "/test-files/test-group-2/", False
     
    Assert.AreEqual CLng(10), wsAnswers.UsedRange.rows.count
    Assert.AreEqual CLng(10), wsTimes.UsedRange.rows.count
    
    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'@TestMethod("Controllers")
Private Sub combineCsvFiles_WhenSurveyRunError_ShouldPrintError()
    On Error GoTo TestFail
    
    combineCsvFiles getCurrentPath() & TestFolder & "/test-files/test-group-3/", False
     
    Assert.AreEqual "Error In Survey Run: The question type is not recognised.", wsAnswers.Cells(3, 1).value
    Assert.AreEqual "Error In Survey Run: Error 515: The question count is incorrect.", wsAnswers.Cells(5, 1).value
    Assert.AreEqual "Error In Survey Run: The question type is not recognised.", wsAnswers.Cells(9, 1).value
    
    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub






