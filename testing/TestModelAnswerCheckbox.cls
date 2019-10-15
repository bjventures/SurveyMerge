VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestModelAnswerCheckbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Tests")
''
'' class module: TestModelAnswerCheckbox
''
'Option Explicit
'Implements ITester
'
'Private assertion As Boolean
'Private baseAnswer As ModelAnswerBase
'
'Private Property Get ITester_className() As String
'    ITester_className = "TestModelAnswerCheckbox"
'End Property
'
'Private Property Get ITester_testList() As Variant
'    ITester_testList = Array( _
'        "test_AnswerCheckbox_WhenSetNumber_ShouldSet", _
'        "test_AnswerCheckbox_WhenSetInvalidNumber_ShouldThrow", _
'        "test_AnswerCheckbox_WhenSetValidTime_ShouldSet", _
'        "test_AnswerCheckbox_WhenSetValidTime_ShouldGetOffset", _
'        "test_AnswerCheckbox_WhenSetInvalidTime_ShouldThrow", _
'        "test_AnswerCheckbox_WhenSetInvalidTimeOffset_ShouldThrow", _
'        "test_AnswerCheckbox_WhenSetTruncatedTime_ShouldThrow", _
'        "test_AnswerCheckbox_WhenSetValue_ShouldSet", _
'        "test_AnswerCheckbox_WhenSetValue_ShouldGetDescription", _
'        "test_AnswerCheckbox_WhenSetIncorrectValue_ShouldThrow" _
'    )
'End Property
'
'Private Function ITester_runTest(ByRef methodName As String) As Boolean
'
'    If Len(methodName) > 63 Then MsgBox "The method name '" & methodName & "' is too long to run on the Mac os.", vbCritical, ProjectName
'    ITester_runTest = CallByName(Me, methodName, VbMethod)
'
'End Function
'
'Public Sub ITester_setUp()
'    Set baseAnswer = New ModelAnswerCheckbox
'End Sub
'
'Public Sub ITester_breakDown()
'    Set baseAnswer = Nothing
'End Sub
'
'Public Function test_AnswerCheckbox_WhenSetNumber_ShouldSet() As Boolean
'
'    baseAnswer.number = 1
'
'    assertion = baseAnswer.number = 1
'    test_AnswerCheckbox_WhenSetNumber_ShouldSet = assertion
'
'End Function
'
'Public Function test_AnswerCheckbox_WhenSetInvalidNumber_ShouldThrow() As Boolean
'
'On Error GoTo Catch
'    baseAnswer.number = 0
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.ModelValidationError
'    test_AnswerCheckbox_WhenSetInvalidNumber_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
'Public Function test_AnswerCheckbox_WhenSetValidTime_ShouldSet() As Boolean
'
'    baseAnswer.isoTime = "2019-04-16T15:08:07+1000"
'    assertion = "16/04/2019 3:08:07 PM " = baseAnswer.time
'
'    test_AnswerCheckbox_WhenSetValidTime_ShouldSet = assertion
'
'End Function
'
'Public Function test_AnswerCheckbox_WhenSetValidTime_ShouldGetOffset() As Boolean
'
'    baseAnswer.isoTime = "2019-04-16T15:08:07-1000"
'    assertion = -1000 = baseAnswer.isoOffset
'
'    test_AnswerCheckbox_WhenSetValidTime_ShouldGetOffset = assertion
'
'End Function
'
'Public Function test_AnswerCheckbox_WhenSetInvalidTime_ShouldThrow() As Boolean
'
'On Error GoTo Catch
'    baseAnswer.isoTime = "20aa-04-16T15:08:07+1000"
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.ModelValidationError
'    test_AnswerCheckbox_WhenSetInvalidTime_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
'Public Function test_AnswerCheckbox_WhenSetInvalidTimeOffset_ShouldThrow() As Boolean
'
'On Error GoTo Catch
'    baseAnswer.isoTime = "20aa-04-16T15:08:07+1a00"
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.ModelValidationError
'    test_AnswerCheckbox_WhenSetInvalidTimeOffset_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
'Public Function test_AnswerCheckbox_WhenSetTruncatedTime_ShouldThrow() As Boolean
'
'On Error GoTo Catch
'    baseAnswer.isoTime = "20-04-16T15:08:07+1000"
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.ModelValidationError
'    test_AnswerCheckbox_WhenSetTruncatedTime_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'
'Public Function test_AnswerCheckbox_WhenSetValue_ShouldSet() As Boolean
'
'    Dim checkboxAnswer As ModelAnswerCheckbox
'    Set checkboxAnswer = baseAnswer
'    checkboxAnswer.value = Array(2, 4)
'
'    assertion = checkboxAnswer.value(0) = 2 And checkboxAnswer.value(1) = 4
'    test_AnswerCheckbox_WhenSetValue_ShouldSet = assertion
'
'End Function
'
'Public Function test_AnswerCheckbox_WhenSetValue_ShouldGetDescription() As Boolean
'
'    Dim checkboxAnswer As ModelAnswerCheckbox
'    Set checkboxAnswer = baseAnswer
'    checkboxAnswer.value = Array(1, 3, 4)
'
'    assertion = baseAnswer.description = "1,3,4"
'
'    test_AnswerCheckbox_WhenSetValue_ShouldGetDescription = assertion
'
'End Function
'
'Public Function test_AnswerCheckbox_WhenSetIncorrectValue_ShouldThrow() As Boolean
'
'    Dim checkboxAnswer As ModelAnswerCheckbox
'    Set checkboxAnswer = New ModelAnswerCheckbox
'
'On Error GoTo Catch
'    checkboxAnswer.value = Array(2, "error")
'
'Finally:
'    Exit Function
'
'Catch:
'    assertion = Err.number = CustomError.ModelValidationError
'    test_AnswerCheckbox_WhenSetIncorrectValue_ShouldThrow = assertion
'    Resume Finally
'
'End Function
'

