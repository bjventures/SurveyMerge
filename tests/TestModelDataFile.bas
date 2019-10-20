Attribute VB_Name = "TestModelDataFile"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private dataFile As ModelDataFile

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
    Set dataFile = New ModelDataFile
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set dataFile = Nothing
End Sub

'@TestMethod("Model")
Private Sub test_WhenSetValid_ShouldSet()
    On Error GoTo TestFail
 
    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub
