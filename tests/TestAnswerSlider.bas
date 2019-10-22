Attribute VB_Name = "TestAnswerSlider"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Models.Answers")

' This class implements ModelAnswerBase, which is tested separately.

Private Assert As Object
Private Fakes As Object
Private answer As ModelAnswerBase
Private sliderAnswer As ModelAnswerSlider

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
    Set answer = New ModelAnswerSlider
    Set sliderAnswer = answer
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set answer = Nothing
    Set sliderAnswer = Nothing
End Sub

'@TestMethod("Model")
Private Sub answerList_Value_WhenSetValid_ShouldSet()
    On Error GoTo TestFail
    sliderAnswer.value = 0.34
   
    Assert.AreEqual CSng(0.34), sliderAnswer.value

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub answerList_Value_WhenNegative_ShouldThrow()
    Const ExpectedError As Long = CustomError.ModelValidationError
    Dim ExpectedDescription As String
    ExpectedDescription = "The value '-0" & Application.International(xlDecimalSeparator) & "34' is not valid."
    On Error GoTo Assert
    sliderAnswer.value = -0.34

    Assert.fail "Expected error was not raised"
    Exit Sub
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Model")
Private Sub answerList_Value_WhenGreaterThanOne_ShouldThrow()
    Const ExpectedError As Long = CustomError.ModelValidationError
    Const ExpectedDescription = "The value '2' is not valid."
    On Error GoTo Assert
    sliderAnswer.value = 2

    Assert.fail "Expected error was not raised"
    Exit Sub
Assert:
    Assert.AreEqual ExpectedError, Err.number
    Assert.AreEqual ExpectedDescription, Err.description
End Sub

'@TestMethod("Model")
Private Sub answerList_Description_WhenValueSet_ShouldGetDescription()
    On Error GoTo TestFail
    sliderAnswer.value = 0.34
    ' Recast to parent class.
    Set answer = sliderAnswer
    
    Assert.AreEqual "0" & Application.International(xlDecimalSeparator) & "34", answer.description

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

