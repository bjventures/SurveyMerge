Attribute VB_Name = "TestModelDataFileErrors"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Models")

Private Assert As Object
Private Fakes As Object
Private dataFile As ModelDataFile
Private accessor As FileAccessor

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set accessor = New FileAccessor
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
Private Sub fileContents_WhenMetaLabelMissing_ShouldThrow()
    Const ExpectedError As Long = CustomError.IncorrectDataFormat
    Const ExpectedDescription As String = "The value 'Survey Name' was not found."
    On Error GoTo Assert
    
    accessor.loadSurveyRunFile "test-2"
    Set dataFile = New ModelDataFile
    dataFile.fileContents = accessor.fileText

    Assert.fail "Expected error was not raised"
    Exit Sub
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Model")
Private Sub fileContents_WhenMetaDataMissing_ShouldThrow()
    Const ExpectedError As Long = CustomError.IncorrectDataFormat
    Const ExpectedDescription As String = "The value 'Survey Name' was not found."
    On Error GoTo Assert
    
    accessor.loadSurveyRunFile "test-3"
    Set dataFile = New ModelDataFile
    dataFile.fileContents = accessor.fileText

    Assert.fail "Expected error was not raised"
    Exit Sub
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub



