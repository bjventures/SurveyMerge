Attribute VB_Name = "TestAnswerCheckbox"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Models")

' This class implements ModelAnswerBase, which is tested separately.

Private Assert As Object
Private Fakes As Object
Private answer As ModelAnswerBase
Private checkboxAnswer As ModelAnswerCheckbox

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
    Set answer = New ModelAnswerCheckbox
    Set checkboxAnswer = answer
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set answer = Nothing
    Set checkboxAnswer = Nothing
End Sub

'@TestMethod("Model")
Private Sub answerList_Value_WhenSetValid_ShouldSet()
    On Error GoTo TestFail
    checkboxAnswer.value = Array(2, 4)

    Assert.SequenceEquals Array(CLng(2), CLng(4)), checkboxAnswer.value

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub answerCheckbox_Value_WhenContainsNegative_ShouldThrow()
    Const ExpectedError As Long = 13
    Const ExpectedDescription As String = "Type mismatch"
    On Error GoTo Assert
    checkboxAnswer.value = Array(2, -2)

    Assert.fail "Expected error was not raised"
    Exit Sub
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Model")
Private Sub answerCheckbox_Value_WhenContainsText_ShouldThrow()
    Const ExpectedError As Long = 13
    Const ExpectedDescription As String = "Type mismatch"
    On Error GoTo Assert
    checkboxAnswer.value = Array(2, "error")

    Assert.fail "Expected error was not raised"
    Exit Sub
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Model")
Private Sub answerCheckbox_Description_WhenValueSet_ShouldGetDescription()
    On Error GoTo TestFail
    checkboxAnswer.value = Array(2, 4)
    ' Recast to parent class.
    Set answer = checkboxAnswer

    Assert.AreEqual "2,4", answer.description

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

