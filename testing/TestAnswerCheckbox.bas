Attribute VB_Name = "TestAnswerCheckbox"
Option Explicit
Option Private Module

'@TestModule
'@Folder("SurveyMerge.Tests.Models")

' Note that this class implements ModelAnswerBase, which is tested separately.

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
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub answerList_Value_WhenInvalid_ShouldThrow()
    Const ExpectedError As Long = CustomError.ModelValidationError
    On Error GoTo TestFail
    checkboxAnswer.value = Array(2, "error")

Assert:
    Assert.Fail "Expected error was not raised"
TestExit:
    Exit Sub
TestFail:
    If Err.number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Model")
Private Sub answerList_Description_WhenValueSet_ShouldGetDescription()
    On Error GoTo TestFail
        checkboxAnswer.value = Array(2, 4)
    ' Recast to parent class.
    Set answer = checkboxAnswer

    Assert.AreEqual "2,4", answer.description

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub
