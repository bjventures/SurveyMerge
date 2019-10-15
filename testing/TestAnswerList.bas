Attribute VB_Name = "TestAnswerList"
Option Explicit
Option Private Module

'@TestModule
'@Folder("SurveyMerge.NewTests")

Private Assert As Object
Private Fakes As Object
Private answer As ModelAnswerBase
Private listAnswer As ModelAnswerList


'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
    Set answer = New ModelAnswerList
    Set listAnswer = answer
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("AnswerList")
Private Sub value_WhenSet_ShouldSet()
    On Error GoTo TestFail
    
    listAnswer.value = 2
    
    Assert.AreEqual CLng(2), listAnswer.value


    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("AnswerList")
Private Sub number_WhenSet_ShouldSet()
    On Error GoTo TestFail

    listAnswer.value = 2

    Assert.AreEqual CLng(2), listAnswer.value


    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("AnswerList")
Private Sub answerList_WhenInvalidNumber_ShouldThrow()
    Const ExpectedError As Long = CustomError.ModelValidationError
    On Error GoTo TestFail

    answer.number = 0
    
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

'@TestMethod("AnswerList")
Private Sub answerListTime_WhenSetValid_ShouldSet()
    On Error GoTo TestFail

    answer.isoTime = "2019-04-16T15:08:07+1000"

    Assert.AreEqual "2019-04-16 15:08:07", Format(answer.time, "yyyy-mm-dd hh:mm:ss")


    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("AnswerList")
Private Sub answerListTime_WhenSetValid_ShouldGetOffset()
    On Error GoTo TestFail

    answer.isoTime = "2019-04-16T15:08:07+1000"
   
    Assert.AreEqual CLng(1000), answer.isoOffset


    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("AnswerList")
Private Sub answerListIsoTime_WhenInvalidTime_ShouldThrow()
    Const ExpectedError As Long = CustomError.ModelValidationError
    On Error GoTo TestFail

    answer.isoTime = "20aa-04-16T15:08:07+1000"
    
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

'@TestMethod("AnswerList")
Private Sub answerListIsoTime_WhenInvalidTimeOffset_ShouldThrow()
    Const ExpectedError As Long = CustomError.ModelValidationError
    On Error GoTo TestFail

    answer.isoTime = "20aa-04-16T15:08:07+1a00"
    
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

'@TestMethod("AnswerList")
Private Sub answerListValue_WhenSetValid_ShouldSet()
    On Error GoTo TestFail

    listAnswer.value = 3
   
    Assert.AreEqual CLng(3), listAnswer.value


    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("AnswerList")
Private Sub answerListValue_WhenSetValid_ShouldGetDescription()
    On Error GoTo TestFail

    listAnswer.value = 3
    Set answer = listAnswer
   
    Assert.AreEqual "3", answer.description


    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("AnswerList")
Private Sub answerListValue_WhenInvalid_ShouldThrow()
    Const ExpectedError As Long = CustomError.ModelValidationError
    On Error GoTo TestFail

    listAnswer.value = 0
    
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
