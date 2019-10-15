VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestParserAnswers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Tests")
'
' class module: TestParserAnswers
'
Option Explicit
Implements ITester

Private assertion As Boolean
Private parser As ParserAnswers
Private answerCollection As Answers
Private baseAnswer As ModelAnswerBase

Private Property Get ITester_className() As String
    ITester_className = "TestParserAnswers"
End Property

Private Property Get ITester_testList() As Variant
    ITester_testList = Array( _
        "test_ParserAnswers_WhenHaveListAnswerType_ShouldParse", _
        "test_ParserAnswers_WhenHaveCheckboxAnswerType_ShouldParse", _
        "test_ParserAnswers_WhenHaveSliderAnswerType_ShouldParse", _
        "test_ParserAnswers_WhenHaveTextAnswerType_ShouldParse", _
        "test_ParserAnswers_WhenHaveNilAnswer_ShouldParse", _
        "test_ParserAnswers_WhenHaveAllAnswerTypesNilAnswer_ShouldParse", _
        "test_ParserAnswers_WhenHaveDefectiveAnswer_ShouldThrow", _
        "test_ParserAnswers_WhenTextAnswerHasQuoteCommaList_ShouldParse", _
        "test_ParserAnswers_WhenTextAnswerIsEmpty_ShouldParse", _
        "test_ParserAnswers_WhenTextAnswerQuoteCommaNoAnswer_ShouldParse", _
        "test_ParserAnswers_WhenTextAnswerHasCommaQuote_ShouldParse", _
        "test_ParserAnswers_WhenTextAnswerMultiplQuoteComma_ShouldParse", _
        "test_ParserAnswers_WhenTextAnswerIsFirst_ShouldParse", _
        "test_ParserAnswers_WhenComplexTextAnswerIsLast_ShouldParse", _
        "test_ParserAnswers_WhenComplexTextAnswerThenNil_ShouldParse", _
        "test_ParserAnswers_WhenComplexTextAnswerThenList_ShouldParse", _
        "test_ParserAnswers_WhenComplexTextAnswerThenText_ShouldParse" _
    )
      
End Property

Private Function ITester_runTest(ByRef methodName As String) As Boolean
    
    If Len(methodName) > 63 Then MsgBox "The method name '" & methodName & "' is too long to run on the Mac os.", vbCritical, ProjectName
    ITester_runTest = CallByName(Me, methodName, VbMethod)

End Function

Private Sub Class_Initialize()
    Set parser = New ParserAnswers
End Sub

Public Sub ITester_setUp()
End Sub

Public Sub ITester_breakDown()
    Set answerCollection = Nothing
End Sub

Public Function test_ParserAnswers_WhenHaveListAnswerType_ShouldParse() As Boolean

    Set answerCollection = parser.parse(getAnswerLines("test-109"))
    Set baseAnswer = answerCollection.item(1)
    Dim listAnswer As New ModelAnswerList
    Set listAnswer = baseAnswer

    assertion = listAnswer.value = 8

    assertion = assertion And baseAnswer.number = 0
    assertion = assertion And baseAnswer.time = CDate("16/04/2019 3:08:08 PM")
    assertion = assertion And baseAnswer.isoOffset = 1000
    assertion = assertion And TypeName(baseAnswer) = "ModelAnswerList" And baseAnswer.description = "8"
    test_ParserAnswers_WhenHaveListAnswerType_ShouldParse = assertion

End Function

Public Function test_ParserAnswers_WhenHaveCheckboxAnswerType_ShouldParse() As Boolean
    Set answerCollection = parser.parse(getAnswerLines("test-109"))

    Set baseAnswer = answerCollection.item(2)
    Dim checkboxAnswer As ModelAnswerCheckbox
    Set checkboxAnswer = baseAnswer

    assertion = checkboxAnswer.value(0) = 4 And checkboxAnswer.value(1) = 5 And checkboxAnswer.value(2) = 6

    assertion = assertion And baseAnswer.number = 0
    assertion = assertion And baseAnswer.time = CDate("16/04/2019 3:08:09 PM")
    assertion = assertion And baseAnswer.isoOffset = 1000
    assertion = assertion And TypeName(baseAnswer) = "ModelAnswerCheckbox" And baseAnswer.description = "4,5,6"

    test_ParserAnswers_WhenHaveCheckboxAnswerType_ShouldParse = assertion

End Function

Public Function test_ParserAnswers_WhenHaveSliderAnswerType_ShouldParse() As Boolean
    Set answerCollection = parser.parse(getAnswerLines("test-109"))

    Set baseAnswer = answerCollection.item(3)
    Dim sliderAnswer As ModelAnswerSlider
    Set sliderAnswer = baseAnswer

    assertion = sliderAnswer.value = 0.32

    assertion = assertion And baseAnswer.number = 0
    assertion = assertion And baseAnswer.time = CDate("16/04/2019 3:08:10 PM")
    assertion = assertion And baseAnswer.isoOffset = 1000
    assertion = assertion And TypeName(baseAnswer) = "ModelAnswerSlider" And baseAnswer.description = "0.32"

    test_ParserAnswers_WhenHaveSliderAnswerType_ShouldParse = assertion

End Function

Public Function test_ParserAnswers_WhenHaveTextAnswerType_ShouldParse() As Boolean
    Set answerCollection = parser.parse(getAnswerLines("test-109"))

    Set baseAnswer = answerCollection.item(4)
    Dim textAnswer As ModelAnswerText
    Set textAnswer = baseAnswer

    assertion = textAnswer.value = "hh,xx" & Chr$(34) & ",b6,yyy" & Chr$(34) & ",ii"

    assertion = assertion And baseAnswer.number = 0
    assertion = assertion And baseAnswer.time = CDate("16/04/2019 3:08:13 PM")
    assertion = assertion And baseAnswer.isoOffset = 1000
    assertion = assertion And TypeName(baseAnswer) = "ModelAnswerText" And baseAnswer.description = "hh,xx" & Chr$(34) & ",b6,yyy" & Chr$(34) & ",ii"

    test_ParserAnswers_WhenHaveTextAnswerType_ShouldParse = assertion

End Function

Public Function test_ParserAnswers_WhenHaveNilAnswer_ShouldParse() As Boolean
    Set answerCollection = parser.parse(getAnswerLines("test-109"))

    Set baseAnswer = answerCollection.item(5)
    assertion = baseAnswer.number = 0
    assertion = assertion And baseAnswer.time = CDate("12:00:00 AM")
    assertion = assertion And baseAnswer.isoOffset = 0
    assertion = assertion And TypeName(baseAnswer) = "ModelAnswerBase" And baseAnswer.description = vbNullString

    test_ParserAnswers_WhenHaveNilAnswer_ShouldParse = assertion

End Function

Public Function test_ParserAnswers_WhenHaveAllAnswerTypesNilAnswer_ShouldParse() As Boolean
    Set answerCollection = parser.parse(getAnswerLines("test-109"))

    assertion = answerCollection.count = 5
    test_ParserAnswers_WhenHaveAllAnswerTypesNilAnswer_ShouldParse = assertion

End Function

Public Function test_ParserAnswers_WhenHaveDefectiveAnswer_ShouldThrow() As Boolean
    ' This question string ios missing a quote.
    On Error GoTo Catch
    Set answerCollection = parser.parse(getAnswerLines("test-110"))

Finally:
    Exit Function

Catch:
    assertion = Err.number = CustomError.InvalidQuestionType
    test_ParserAnswers_WhenHaveDefectiveAnswer_ShouldThrow = assertion
    Resume Finally

End Function

'
' Tests for text answers.
'
Public Function test_ParserAnswers_WhenTextAnswerHasQuoteCommaList_ShouldParse() As Boolean
    Set answerCollection = parser.parse(getAnswerLines("test-101"))

    Set baseAnswer = answerCollection.item(4)
    Dim textAnswer As ModelAnswerText
    Set textAnswer = baseAnswer

    assertion = textAnswer.value = "hh,xx" & Chr$(34) & ",b6,yyy" & Chr$(34) & ",ii"
    assertion = assertion And TypeName(baseAnswer) = "ModelAnswerText" And baseAnswer.description = "hh,xx" & Chr$(34) & ",b6,yyy" & Chr$(34) & ",ii"

    test_ParserAnswers_WhenTextAnswerHasQuoteCommaList_ShouldParse = assertion

End Function

Public Function test_ParserAnswers_WhenTextAnswerIsEmpty_ShouldParse() As Boolean
    Set answerCollection = parser.parse(getAnswerLines("test-116"))

    Set baseAnswer = answerCollection.item(4)
    Dim textAnswer As ModelAnswerText
    Set textAnswer = baseAnswer

    assertion = textAnswer.value = vbNullString
    assertion = assertion And TypeName(baseAnswer) = "ModelAnswerText" And baseAnswer.description = vbNullString

    test_ParserAnswers_WhenTextAnswerIsEmpty_ShouldParse = assertion

End Function

Public Function test_ParserAnswers_WhenTextAnswerQuoteCommaNoAnswer_ShouldParse() As Boolean
    Set answerCollection = parser.parse(getAnswerLines("test-108"))

    Set baseAnswer = answerCollection.item(4)
    Dim textAnswer As ModelAnswerText
    Set textAnswer = baseAnswer
    assertion = textAnswer.value = "hh,xx" & Chr$(34) & ",b6,yyy" & Chr$(34) & ",ii"
    assertion = assertion And TypeName(baseAnswer) = "ModelAnswerText" And baseAnswer.description = "hh,xx" & Chr$(34) & ",b6,yyy" & Chr$(34) & ",ii"
    test_ParserAnswers_WhenTextAnswerQuoteCommaNoAnswer_ShouldParse = assertion

End Function

Public Function test_ParserAnswers_WhenTextAnswerHasCommaQuote_ShouldParse() As Boolean
    Set answerCollection = parser.parse(getAnswerLines("test-102"))

    Set baseAnswer = answerCollection.item(4)
    Dim textAnswer As ModelAnswerText
    Set textAnswer = baseAnswer

    assertion = textAnswer.value = "hh," & Chr$(34) & "i,ii" & Chr$(34) & " ," & Chr$(34) & "2,3"
    assertion = assertion And TypeName(baseAnswer) = "ModelAnswerText" And baseAnswer.description = "hh," & Chr$(34) & "i,ii" & Chr$(34) & " ," & Chr$(34) & "2,3"

    test_ParserAnswers_WhenTextAnswerHasCommaQuote_ShouldParse = assertion

End Function

Public Function test_ParserAnswers_WhenTextAnswerMultiplQuoteComma_ShouldParse() As Boolean

    Set answerCollection = parser.parse(getAnswerLines("test-103"))

    Dim testString As String
    Set baseAnswer = answerCollection.item(4)
    Dim textAnswer As ModelAnswerText
    Set textAnswer = baseAnswer

    testString = "hh,xx" & Chr$(34) & ",k6,yy" & Chr$(34) & ",3" & Chr$(34) & ",i," & Chr$(34) & ",i"
    assertion = textAnswer.value = testString
    assertion = assertion And TypeName(baseAnswer) = "ModelAnswerText" And baseAnswer.description = testString

    test_ParserAnswers_WhenTextAnswerMultiplQuoteComma_ShouldParse = assertion

End Function

Public Function test_ParserAnswers_WhenTextAnswerIsFirst_ShouldParse() As Boolean

    Set answerCollection = parser.parse(getAnswerLines("test-104"))

    Dim testString As String
    Set baseAnswer = answerCollection.item(1)
    Dim textAnswer As ModelAnswerText
    Set textAnswer = baseAnswer

    testString = "First question"
    assertion = textAnswer.value = testString
    assertion = assertion And TypeName(baseAnswer) = "ModelAnswerText" And baseAnswer.description = testString

    test_ParserAnswers_WhenTextAnswerIsFirst_ShouldParse = assertion

End Function

Public Function test_ParserAnswers_WhenComplexTextAnswerIsLast_ShouldParse() As Boolean

    Set answerCollection = parser.parse(getAnswerLines("test-105"))

    Dim testString As String
    Set baseAnswer = answerCollection.item(4)
    Dim textAnswer As ModelAnswerText
    Set textAnswer = baseAnswer

    testString = "hh,xx" & Chr$(34) & ",x6,yy" & Chr$(34) & ",3" & Chr$(34) & ",i," & Chr$(34) & ",i"
    assertion = textAnswer.value = testString
    assertion = assertion And TypeName(baseAnswer) = "ModelAnswerText" And baseAnswer.description = testString

    test_ParserAnswers_WhenComplexTextAnswerIsLast_ShouldParse = assertion

End Function

Public Function test_ParserAnswers_WhenComplexTextAnswerThenNil_ShouldParse() As Boolean

    Set answerCollection = parser.parse(getAnswerLines("test-111"))

    Dim testString As String
    Set baseAnswer = answerCollection.item(4)
    Dim textAnswer As ModelAnswerText
    Set textAnswer = baseAnswer

    testString = "hh,xx" & Chr$(34) & ",x6,yy" & Chr$(34) & ",3" & Chr$(34) & ",i," & Chr$(34) & ",i"
    assertion = textAnswer.value = testString
    assertion = assertion And TypeName(baseAnswer) = "ModelAnswerText" And baseAnswer.description = testString

    test_ParserAnswers_WhenComplexTextAnswerThenNil_ShouldParse = assertion

End Function

Public Function test_ParserAnswers_WhenComplexTextAnswerThenList_ShouldParse() As Boolean

    Set answerCollection = parser.parse(getAnswerLines("test-112"))

    Dim testString As String
    Set baseAnswer = answerCollection.item(4)
    Dim textAnswer As ModelAnswerText
    Set textAnswer = baseAnswer

    testString = "hh,xx" & Chr$(34) & ",x6,yy" & Chr$(34) & ",3" & Chr$(34) & ",i," & Chr$(34) & ",i"
    assertion = textAnswer.value = testString
    assertion = assertion And TypeName(baseAnswer) = "ModelAnswerText" And baseAnswer.description = testString

    test_ParserAnswers_WhenComplexTextAnswerThenList_ShouldParse = assertion

End Function

Public Function test_ParserAnswers_WhenComplexTextAnswerThenText_ShouldParse() As Boolean

    Set answerCollection = parser.parse(getAnswerLines("test-113"))

    Dim testString As String
    Set baseAnswer = answerCollection.item(4)
    Dim textAnswer As ModelAnswerText
    Set textAnswer = baseAnswer

    testString = "hh,xx" & Chr$(34) & ",x6,yy" & Chr$(34) & ",3" & Chr$(34) & ",i," & Chr$(34) & ",i"
    assertion = textAnswer.value = testString
    assertion = assertion And TypeName(baseAnswer) = "ModelAnswerText" And baseAnswer.description = testString

    test_ParserAnswers_WhenComplexTextAnswerThenText_ShouldParse = assertion

End Function




