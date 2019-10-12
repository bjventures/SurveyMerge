VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestModelAnswerBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' class module: TestModelAnswerBase
'
Option Explicit
Implements ITester

Private assertion As Boolean
Private baseAnswer As ModelAnswerBase

Private Property Get ITester_className() As String
    ITester_className = "TestModelAnswerBase"
End Property

Private Property Get ITester_testList() As Variant
    ITester_testList = Array( _
        "test_AnswerBase_WhenSetNumber_ShouldSet", _
        "test_AnswerBase_WhenNoTimeSet_ShouldReturnMidnight", _
        "test_AnswerBase_WhenSetInvalidNumber_ShouldThrow", _
        "test_AnswerBase_WhenSetValidTime_ShouldSet", _
        "test_AnswerBase_WhenSetValidTime_ShouldGetOffset", _
        "test_AnswerBase_WhenSetInvalidTime_ShouldThrow", _
        "test_AnswerBase_WhenSetInvalidTimeOffset_ShouldThrow", _
        "test_AnswerBase_WhenSetTruncatedTime_ShouldThrow", _
        "test_AnswerBase_WhenGetDescription_ShouldBeNilLengthString" _
    )
   
End Property

Private Function ITester_runTest(methodName As String) As Boolean
    
    If Len(methodName) > 63 Then MsgBox "The method name '" & methodName & "' is too long to run on the Mac os.", vbCritical, ProjectName
    ITester_runTest = CallByName(Me, methodName, VbMethod)

End Function

Public Sub ITester_setUp()
    Set baseAnswer = New ModelAnswerBase
End Sub

Public Sub ITester_breakDown()
    Set baseAnswer = Nothing
End Sub

Public Function test_AnswerBase_WhenSetNumber_ShouldSet() As Boolean
    
    baseAnswer.number = 1
    
    assertion = baseAnswer.number = 1
    test_AnswerBase_WhenSetNumber_ShouldSet = assertion

End Function

Public Function test_AnswerBase_WhenNoTimeSet_ShouldReturnMidnight() As Boolean
        
    assertion = baseAnswer.time = CDate(0)
    test_AnswerBase_WhenNoTimeSet_ShouldReturnMidnight = assertion

End Function

Public Function test_AnswerBase_WhenSetInvalidNumber_ShouldThrow() As Boolean

On Error GoTo Catch
    baseAnswer.number = 0
    
Finally:
    Exit Function
    
Catch:
    assertion = Err.number = CustomError.ModelValidationError
    test_AnswerBase_WhenSetInvalidNumber_ShouldThrow = assertion
    Resume Finally

End Function

Public Function test_AnswerBase_WhenSetValidTime_ShouldSet() As Boolean

    baseAnswer.isoTime = "2019-04-16T15:08:07+1000"
    assertion = "16/04/2019 3:08:07 PM " = baseAnswer.time
    
    test_AnswerBase_WhenSetValidTime_ShouldSet = assertion
    
End Function

Public Function test_AnswerBase_WhenSetValidTime_ShouldGetOffset() As Boolean

    baseAnswer.isoTime = "2019-04-16T15:08:07-1000"
    assertion = -1000 = baseAnswer.isoOffset
    
    test_AnswerBase_WhenSetValidTime_ShouldGetOffset = assertion
    
End Function

Public Function test_AnswerBase_WhenSetInvalidTime_ShouldThrow() As Boolean

On Error GoTo Catch
    baseAnswer.isoTime = "20aa-04-16T15:08:07+1000"
    
Finally:
    Exit Function
    
Catch:
    assertion = Err.number = CustomError.ModelValidationError
    test_AnswerBase_WhenSetInvalidTime_ShouldThrow = assertion
    Resume Finally
    
End Function

Public Function test_AnswerBase_WhenSetInvalidTimeOffset_ShouldThrow() As Boolean

On Error GoTo Catch
    baseAnswer.isoTime = "20aa-04-16T15:08:07+1a00"
    
Finally:
    Exit Function
    
Catch:
    assertion = Err.number = CustomError.ModelValidationError
    test_AnswerBase_WhenSetInvalidTimeOffset_ShouldThrow = assertion
    Resume Finally
    
End Function

Public Function test_AnswerBase_WhenSetTruncatedTime_ShouldThrow() As Boolean

On Error GoTo Catch
    baseAnswer.isoTime = "20-04-16T15:08:07+1000"
    
Finally:
    Exit Function
    
Catch:
    assertion = Err.number = CustomError.ModelValidationError
    test_AnswerBase_WhenSetTruncatedTime_ShouldThrow = assertion
    Resume Finally
    
End Function

Public Function test_AnswerBase_WhenGetDescription_ShouldBeNilLengthString() As Boolean

    assertion = baseAnswer.description = ""
    test_AnswerBase_WhenGetDescription_ShouldBeNilLengthString = assertion

End Function



