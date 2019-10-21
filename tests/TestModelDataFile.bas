Attribute VB_Name = "TestModelDataFile"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Models")

Private Assert As Object
Private Fakes As Object
Private dataFile As ModelDataFile
Private fileAccessor As fileAccessor

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set fileAccessor = New fileAccessor
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fakes = Nothing
    Set fileAccessor = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    Set dataFile = New ModelDataFile
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set dataFile = Nothing
End Sub

'@TestMethod("Model")
Private Sub metidata_WhenInitialisedWithFileText_ShouldSet()
    On Error GoTo TestFail
    fileAccessor.loadSurveyRunFile "test-1"
    dataFile.fileContents = fileAccessor.fileText

    Assert.AreEqual "Test 1", dataFile.surveyName
    Assert.AreEqual "Test ID", dataFile.subjectId

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

