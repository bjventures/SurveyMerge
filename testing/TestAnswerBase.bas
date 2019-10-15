Attribute VB_Name = "TestAnswerBase"
Option Explicit
Option Private Module

'@TestModule
'@Folder("SurveyMerge.NewTests")

Private Assert As Object
Private Fakes As Object

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
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("AnswerBase")
Private Sub number_WhenSet_ShouldSet()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'Public Function test_AnswerBase_WhenSetNumber_ShouldSet() As Boolean
'
'    baseAnswer.number = 1
'
'    assertion = baseAnswer.number = 1
'    test_AnswerBase_WhenSetNumber_ShouldSet = assertion
'
'End Function
'
'Public Function test_AnswerBase_WhenNoTimeSet_ShouldReturnMidnight() As Boolean
'
'    assertion = baseAnswer.time = CDate(0)
'    test_AnswerBase_WhenNoTimeSet_ShouldReturnMidnight = assertion
'
'End Function
'
'Public Function test_AnswerBase_WhenSetInvalidNumber_ShouldThrow() As Boolean
'
'On Error GoTo Catch
'    baseAnswer.number = 0
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.ModelValidationError
'    test_AnswerBase_WhenSetInvalidNumber_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
'Public Function test_AnswerBase_WhenSetValidTime_ShouldSet() As Boolean
'
'    baseAnswer.isoTime = "2019-04-16T15:08:07+1000"
'    assertion = "16/04/2019 3:08:07 PM " = baseAnswer.time
'
'    test_AnswerBase_WhenSetValidTime_ShouldSet = assertion
'
'End Function
'
'Public Function test_AnswerBase_WhenSetValidTime_ShouldGetOffset() As Boolean
'
'    baseAnswer.isoTime = "2019-04-16T15:08:07-1000"
'    assertion = -1000 = baseAnswer.isoOffset
'
'    test_AnswerBase_WhenSetValidTime_ShouldGetOffset = assertion
'
'End Function
'
'Public Function test_AnswerBase_WhenSetInvalidTime_ShouldThrow() As Boolean
'
'On Error GoTo Catch
'    baseAnswer.isoTime = "20aa-04-16T15:08:07+1000"
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.ModelValidationError
'    test_AnswerBase_WhenSetInvalidTime_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
'Public Function test_AnswerBase_WhenSetInvalidTimeOffset_ShouldThrow() As Boolean
'
'On Error GoTo Catch
'    baseAnswer.isoTime = "20aa-04-16T15:08:07+1a00"
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.ModelValidationError
'    test_AnswerBase_WhenSetInvalidTimeOffset_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
'Public Function test_AnswerBase_WhenSetTruncatedTime_ShouldThrow() As Boolean
'
'On Error GoTo Catch
'    baseAnswer.isoTime = "20-04-16T15:08:07+1000"
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.ModelValidationError
'    test_AnswerBase_WhenSetTruncatedTime_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
'Public Function test_AnswerBase_WhenGetDescription_ShouldBeNilLengthString() As Boolean
'
'    assertion = baseAnswer.description = vbNullString
'    test_AnswerBase_WhenGetDescription_ShouldBeNilLengthString = assertion
'
'End Function
'
