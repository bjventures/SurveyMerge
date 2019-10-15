Attribute VB_Name = "TestAnswerList"
Option Explicit
Option Private Module

'@TestModule
'@Folder("SurveyMerge.NewTests")

Private Assert As Object
Private Fakes As Object
Private answer As ModelAnswerList

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
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("AnswerList")
Private Sub value_WhenSet_ShouldSet()
    On Error GoTo TestFail
    
    answer.value = 2
    
    Assert.AreEqual CLng(2), answer.value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub


'Public Function test_AnswerList_WhenSetNumber_ShouldSet() As Boolean
'
'    baseAnswer.number = 1
'
'    assertion = baseAnswer.number = 1
'    test_AnswerList_WhenSetNumber_ShouldSet = assertion
'
'End Function
'
'Public Function test_AnswerList_WhenSetInvalidNumber_ShouldThrow() As Boolean
'
'On Error GoTo Catch
'    baseAnswer.number = 0
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.ModelValidationError
'    test_AnswerList_WhenSetInvalidNumber_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
'Public Function test_AnswerList_WhenSetValidTime_ShouldSet() As Boolean
'
'    baseAnswer.isoTime = "2019-04-16T15:08:07+1000"
'    assertion = "16/04/2019 3:08:07 PM " = baseAnswer.time
'
'    test_AnswerList_WhenSetValidTime_ShouldSet = assertion
'
'End Function
'
'Public Function test_AnswerList_WhenSetValidTime_ShouldGetOffset() As Boolean
'
'    baseAnswer.isoTime = "2019-04-16T15:08:07-1000"
'    assertion = -1000 = baseAnswer.isoOffset
'
'    test_AnswerList_WhenSetValidTime_ShouldGetOffset = assertion
'
'End Function
'
'Public Function test_AnswerList_WhenSetInvalidTime_ShouldThrow() As Boolean
'
'On Error GoTo Catch
'    baseAnswer.isoTime = "20aa-04-16T15:08:07+1000"
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.ModelValidationError
'    test_AnswerList_WhenSetInvalidTime_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
'Public Function test_AnswerList_WhenSetInvalidTimeOffset_ShouldThrow() As Boolean
'
'On Error GoTo Catch
'    baseAnswer.isoTime = "20aa-04-16T15:08:07+1a00"
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.ModelValidationError
'    test_AnswerList_WhenSetInvalidTimeOffset_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
'Public Function test_AnswerList_WhenSetTruncatedTime_ShouldThrow() As Boolean
'
'On Error GoTo Catch
'    baseAnswer.isoTime = "20-04-16T15:08:07+1000"
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.ModelValidationError
'    test_AnswerList_WhenSetTruncatedTime_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
'Public Function test_AnswerList_WhenSetValue_ShouldSet() As Boolean
'
'    Dim listAnswer As ModelAnswerList
'    Set listAnswer = New ModelAnswerList
'    listAnswer.value = 3
'    assertion = listAnswer.value = 3
'
'    test_AnswerList_WhenSetValue_ShouldSet = assertion
'
'End Function
'
'Public Function test_AnswerList_WhenSetValue_ShouldGetDescription() As Boolean
'
'    Dim listAnswer As ModelAnswerList
'    Set listAnswer = New ModelAnswerList
'
'    listAnswer.value = 3
'
'    Set baseAnswer = listAnswer
'
'    assertion = baseAnswer.description = "3"
'    test_AnswerList_WhenSetValue_ShouldGetDescription = assertion
'
'End Function
'
'Public Function test_AnswerList_WhenSetIncorrectValue_ShouldThrow() As Boolean
'
'    Dim listAnswer As ModelAnswerList
'    Set listAnswer = New ModelAnswerList
'
'On Error GoTo Catch
'    listAnswer.value = -1
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.ModelValidationError
'    test_AnswerList_WhenSetIncorrectValue_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
