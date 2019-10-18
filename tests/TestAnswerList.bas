Attribute VB_Name = "TestAnswerList"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Models.AnswerModels")

' This class implements ModelAnswerBase, which is tested separately.

Private Assert As Object
Private Fakes As Object
Private answer As ModelAnswerBase
Private listAnswer As ModelAnswerList

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
    Set answer = New ModelAnswerList
    Set listAnswer = answer
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set answer = Nothing
    Set listAnswer = Nothing
End Sub

'@TestMethod("Model")
Private Sub answerList_Value_WhenSetValid_ShouldSet()
    On Error GoTo TestFail
    listAnswer.value = 3
   
    Assert.AreEqual CLng(3), listAnswer.value

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub answerList_Value_WhenZero_ShouldThrow()
    Const ExpectedError As Long = CustomError.ModelValidationError
    Const ExpectedDescription As String = "The value '0' is not valid."
    On Error GoTo Assert
    listAnswer.value = 0
    
    Assert.fail "Expected error was not raised"
    Exit Sub
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Model")
Private Sub answerList_Value_WhenNegative_ShouldThrow()
    Const ExpectedError As Long = CustomError.ModelValidationError
    Const ExpectedDescription As String = "The value '-1' is not valid."
    On Error GoTo Assert
    listAnswer.value = -1
    
    Assert.fail "Expected error was not raised"
    Exit Sub
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Model")
Private Sub answerList_Description_WhenValueSet_ShouldGetDescription()
    On Error GoTo TestFail
    listAnswer.value = 3
    ' Recast to parent class.
    Set answer = listAnswer
   
    Assert.AreEqual "3", answer.description

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

