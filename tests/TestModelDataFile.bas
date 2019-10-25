Attribute VB_Name = "TestModelDataFile"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Models.DataFile")

Private Assert As Object
Private Fakes As Object
Private dataFile As ModelDataFile
Private accessor As FileAccessor

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set accessor = New FileAccessor
    accessor.loadSurveyRunFile "test-1"
    Set dataFile = New ModelDataFile
    dataFile.fileContents = accessor.fileText
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fakes = Nothing
    Set accessor = Nothing
    Set dataFile = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
End Sub

'@TestCleanup
Private Sub TestCleanup()
End Sub

'@TestMethod("Model")
Private Sub surveyName_WhenInitialisedWithFileText_ShouldSet()
    On Error GoTo TestFail

    Assert.AreEqual "Test 1", dataFile.surveyName

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub subjectId_WhenInitialisedWithFileText_ShouldSet()
    On Error GoTo TestFail

    Assert.AreEqual "Test ID", dataFile.subjectId

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub surveyRunCount_WhenInitialisedWithFileText_ShouldSet()
    On Error GoTo TestFail

    Assert.AreEqual CLng(2), dataFile.surveyRunCount

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub surveyRunLines_WhenInitialisedWithFileText_ShouldSet()
    On Error GoTo TestFail
    Dim dataLines As ModelDataLines
    Set dataLines = dataFile.surveyRunLines(1)

    Assert.AreEqual "Start Time,End Time,1,2,3,4,5", dataLines.header
    Assert.AreEqual "2019-04-16T15:08:07+1000,2019-04-16T15:08:14+1000,4," & Chr(34) & "4,5,6" & Chr(34) & ",0.32,," & Chr(34) & "hhii" & Chr(34), dataLines.answer
    Assert.AreEqual ",,2019-04-16T15:08:08+1000,2019-04-16T15:09:09+1000,2019-04-16T15:10:10+1000,2019-04-16T15:11:13+1000,2019-04-16T15:12:13+1000", dataLines.timeStamp

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub surveyRunLines_WhenArgumentTooHigh_ShouldThrow()
    Const ExpectedError As Long = CustomError.IncorrectDataFormat
    Const ExpectedDescription As String = "The value for 'runNumber' is not valid."
    On Error GoTo Assert
    
    ' There are only 2 survey runs in the file.
    dataFile.surveyRunLines (3)
      
    Assert.fail "Expected error was not raised"
    Exit Sub
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

