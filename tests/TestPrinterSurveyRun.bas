Attribute VB_Name = "TestPrinterSurveyRun"
'@TestModule
'@Folder("Tests.Views")
Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object
Private wsAnswers As Worksheet
Private wsTimes As Worksheet
Private sheets As Variant
Private surveyRun As ModelSurveyRun
Private printer As IPrinter
Private saveCalcState As Long

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set surveyRun = getTestSurveyRun()
    Set printer = New PrinterSurveyRun
    Set wsAnswers = ThisWorkbook.sheets(getWsName(WsSheet.Answers))
    Set wsTimes = ThisWorkbook.sheets(getWsName(WsSheet.Times))
    sheets = Array(getWsName(WsSheet.Answers), getWsName(WsSheet.Times))
    saveCalcState = Application.Calculation
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fakes = Nothing
    Set surveyRun = Nothing
    Set printer = Nothing
    Set wsAnswers = Nothing
    Set wsTimes = Nothing
    clearOrAddSpreadsheets (sheets)
    Application.Calculation = saveCalcState
    Application.ScreenUpdating = True
End Sub

'@TestInitialize
Private Sub TestInitialize()
    clearOrAddSpreadsheets (sheets)
End Sub

'@TestCleanup
Private Sub TestCleanup()
End Sub

Private Function getTestSurveyRun() As ModelSurveyRun
    Dim testRun As ModelSurveyRun
    Dim answerCollection As Answers
    Dim answer As ModelAnswerBase
    
    Dim listAnswer As ModelAnswerList
    Set answer = New ModelAnswerList
    answer.isoTime = "2019-12-03T00:00:00+0000"
    Set listAnswer = answer
    listAnswer.value = 3
    
    Dim checkboxAnswer As ModelAnswerCheckbox
    Set answer = New ModelAnswerCheckbox
    answer.isoTime = "2019-12-04T00:00:00+0000"
    Set checkboxAnswer = answer
    checkboxAnswer.value = Array(7)
    
    Dim textAnswer As ModelAnswerText
    Set answer = New ModelAnswerText
    answer.isoTime = "2019-12-05T00:00:00+0000"
    Set textAnswer = answer
    textAnswer.value = "Text Answer"
    
    Dim sliderAnswer As ModelAnswerSlider
    Set answer = New ModelAnswerSlider
    answer.isoTime = "2019-12-06T00:00:00+0000"
    Set sliderAnswer = answer
    sliderAnswer.value = 0.258
    
    Set testRun = New ModelSurveyRun
    Set answerCollection = New Answers
    answerCollection.Add listAnswer
    answerCollection.Add checkboxAnswer
    answerCollection.Add textAnswer
    answerCollection.Add sliderAnswer

    testRun.surveyName = "Test Name"
    testRun.participantId = "Test ID"
    testRun.startTime = DateSerial(2019, 12, 1)
    testRun.endTime = DateSerial(2019, 12, 2)
    testRun.questionCount = 4
    testRun.answerCollection = answerCollection
    
    Set getTestSurveyRun = testRun

End Function

Private Sub runAssertionsForRow(assertRow As Long)
    Assert.AreEqual "Test Name", wsAnswers.Cells(assertRow, 1).value
    Assert.AreEqual "Test ID", wsAnswers.Cells(assertRow, 2).value
    Assert.AreEqual DateSerial(2019, 12, 1), wsAnswers.Cells(assertRow, 3).value
    Assert.AreEqual DateSerial(2019, 12, 2), wsAnswers.Cells(assertRow, 4).value
    Assert.AreEqual CDbl(3), wsAnswers.Cells(assertRow, 5).value
    Assert.AreEqual "Test Name", wsTimes.Cells(assertRow, 1).value
    Assert.AreEqual "Test ID", wsTimes.Cells(assertRow, 2).value
    Assert.AreEqual DateSerial(2019, 12, 1), wsTimes.Cells(assertRow, 3).value
    Assert.AreEqual DateSerial(2019, 12, 2), wsTimes.Cells(assertRow, 4).value
    Assert.AreEqual DateSerial(2019, 12, 3), wsTimes.Cells(assertRow, 5).value
End Sub

Private Sub runHeaderAssertions()
    Assert.AreEqual "Survey Name", wsAnswers.Cells(1, 1).value
    Assert.AreEqual "Participant ID", wsAnswers.Cells(1, 2).value
    Assert.AreEqual "Start Time", wsAnswers.Cells(1, 3).value
    Assert.AreEqual "Finish Time", wsAnswers.Cells(1, 4).value
    Assert.AreEqual CDbl(1), wsAnswers.Cells(1, 5).value
    Assert.AreEqual "Survey Name", wsTimes.Cells(1, 1).value
    Assert.AreEqual "Participant ID", wsTimes.Cells(1, 2).value
    Assert.AreEqual "Start Time", wsTimes.Cells(1, 3).value
    Assert.AreEqual "Finish Time", wsTimes.Cells(1, 4).value
    Assert.AreEqual CDbl(1), wsTimes.Cells(1, 5).value
End Sub

Private Sub runHeaderEmptyAssertions()
    Dim counter As Long
    
    For counter = 1 To 5
        Assert.AreEqual "", wsAnswers.Cells(1, counter).value
    Next counter
End Sub

'@TestMethod("Views")
Private Sub printData_WhenNotFirstSurveyRun_ShouldPrintSurveyRunNotHeader()
    On Error GoTo TestFail
    printer.printData surveyRun, 1
    
    runAssertionsForRow 3
    runHeaderEmptyAssertions
    
    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Views")
Private Sub printData_WhenFirstSurveyRun_ShouldPrintSurveyRunAndHeader()
    On Error GoTo TestFail
    printer.printData surveyRun, 0

    runAssertionsForRow 2
    runHeaderAssertions
    
    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Views")
Private Sub printData_WhenListAnswer_ShouldBeGeneralFormat()
    On Error GoTo TestFail
    printer.printData surveyRun, 0

    Assert.AreEqual "Double", TypeName(wsAnswers.Cells(2, 8).value)
    Assert.AreEqual CDbl(3), wsAnswers.Cells(2, 5).value
    Assert.AreEqual "General", wsAnswers.Cells(2, 5).NumberFormat
    
    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Views")
Private Sub printData_WhenCheckboxAnswer_ShouldBeTextFormat()
    On Error GoTo TestFail
    printer.printData surveyRun, 0

    ' Note that since there is only one number digit, Excel interprests as Double.
    ' That is why we force the format to be text and supress the error.
    Assert.AreEqual "Double", TypeName(wsAnswers.Cells(2, 8).value)
    Assert.AreEqual "7", wsAnswers.Cells(2, 6).value
    Assert.AreEqual "@", wsAnswers.Cells(2, 6).NumberFormat
    
    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Views")
Private Sub printData_WhenTextAnswer_ShouldBeGeneralFormat()
    On Error GoTo TestFail
    printer.printData surveyRun, 0

    Assert.AreEqual "String", TypeName(wsAnswers.Cells(2, 7).value)
    Assert.AreEqual "Text Answer", wsAnswers.Cells(2, 7).value
    Assert.AreEqual "General", wsAnswers.Cells(2, 7).NumberFormat
    
    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Views")
Private Sub printData_WhenSliderAnswer_ShouldBeNumberFormat()
    On Error GoTo TestFail
    printer.printData surveyRun, 0

    Assert.AreEqual "Double", TypeName(wsAnswers.Cells(2, 8).value)
    Assert.AreEqual 0.258, wsAnswers.Cells(2, 8).value
    Assert.AreEqual "General", wsAnswers.Cells(2, 8).NumberFormat
    
    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

