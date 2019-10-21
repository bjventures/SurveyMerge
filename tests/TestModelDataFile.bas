Attribute VB_Name = "TestModelDataFile"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Models")

Private Assert As Object
Private Fakes As Object
Private dataFile As ModelDataFile
Private FileAccessor As FileAccessor

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set FileAccessor = New FileAccessor
    FileAccessor.loadSurveyRunFile "test-1"
    Set dataFile = New ModelDataFile
    dataFile.fileContents = FileAccessor.fileText
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fakes = Nothing
    Set FileAccessor = Nothing
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
Private Sub surveyRunStrings_WhenInitialisedWithFileText_ShouldSet()
    On Error GoTo TestFail
    Dim runLines As Variant
    runLines = dataFile.surveyRunStrings(1)

    Assert.AreEqual "Start Time,End Time,1,2,3,4,5", runLines(0)
    Assert.AreEqual "Start Time,End Time,1,2,3,4,5", runLines(0)

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub surveyRunStrings_WhenArgumentTooHigh_ShouldThrow()
    Const ExpectedError As Long = CustomError.IncorrectDataFormat
    Const ExpectedDescription As String = "The value for 'runNumber' is not valid."
    On Error GoTo Assert
    
    ' There are only 2 survey runs in the file.
    dataFile.surveyRunStrings (3)
      
    Assert.fail "Expected error was not raised"
    Exit Sub
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub




