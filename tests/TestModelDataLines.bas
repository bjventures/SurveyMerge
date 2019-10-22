Attribute VB_Name = "TestModelDataLines"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Models.DataFile")

Private Assert As Object
Private Fakes As Object
Private dataLines As ModelDataLines

'@ModuleInitialize
Private Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set dataLines = New ModelDataLines
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fakes = Nothing
    Set dataLines = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
End Sub

'@TestCleanup
Private Sub TestCleanup()
End Sub

'@TestMethod("Model")
Private Sub header_WhenInitialisedWithText_ShouldSet()
    On Error GoTo TestFail
    dataLines.header = "Header line"
    
    Assert.AreEqual "Header line", dataLines.header

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub answer_WhenInitialisedWithText_ShouldSet()
    On Error GoTo TestFail
    dataLines.answer = "Answer line"
    
    Assert.AreEqual "Answer line", dataLines.answer

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub timeStamp_WhenInitialisedWithText_ShouldSet()
    On Error GoTo TestFail
    dataLines.timeStamp = "Timestamp line"
    
    Assert.AreEqual "Timestamp line", dataLines.timeStamp

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

