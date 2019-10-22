Attribute VB_Name = "TestAnswerBase"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.Models.Answers")

Private Assert As Object
Private Fakes As Object
Private baseAnswer As ModelAnswerBase

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
    Set baseAnswer = New ModelAnswerBase
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set baseAnswer = Nothing
End Sub

'@TestMethod("Model")
Private Sub answerBase_Number_WhenSetValid_ShouldSet()
    On Error GoTo TestFail
    baseAnswer.number = 2
    
    Assert.AreEqual CLng(2), baseAnswer.number

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub answerBase_Number_WhenInvalid_ShouldThrow()
    Const ExpectedError As Long = CustomError.ModelValidationError
    On Error GoTo TestFail
    baseAnswer.number = 0

Assert:
    Assert.fail "Expected error was not raised"
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
Private Sub answerBase_Time_WhenIsoTimeNotSet_ShouldReturnMidnight()
    On Error GoTo TestFail
    Assert.AreEqual CDate(0), baseAnswer.time

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub answerBase_Time_WhenIsoTimeSet_ShouldSet()
    On Error GoTo TestFail
    baseAnswer.isoTime = "2019-04-16T15:08:07+1000"
   
    Assert.AreEqual "2019-04-16 15:08:07", Format$(baseAnswer.time, "yyyy-mm-dd hh:mm:ss")
   
    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub answerBase_IsoOffset_WhenIsoTimeSet_ShouldGetOffset()
    On Error GoTo TestFail
    baseAnswer.isoTime = "2019-04-16T15:08:07+1000"
   
    Assert.AreEqual CLng(1000), baseAnswer.isoOffset

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub answerBase_IsoTime_WhenSetInvalidTime_ShouldSetDefault()
    On Error GoTo TestFail

    Assert.AreEqual CDate(0), baseAnswer.time

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub answerBase_IsoTime_WhenSetInvalidTimeOffset_ShouldSetDefault()
    On Error GoTo TestFail
    baseAnswer.isoTime = "2019-04-16T15:08:07+1a00"

    Assert.AreEqual CLng(0), baseAnswer.isoOffset

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Model")
Private Sub answerBase_Description_WhenGet_ShouldBeNilLength()
    On Error GoTo TestFail
   
    Assert.AreEqual vbNullString, baseAnswer.description

    Exit Sub
TestFail:
    Assert.fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

