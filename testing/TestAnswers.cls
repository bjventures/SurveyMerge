VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestAnswers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' class module: TestAnswers
'
Option Explicit
Implements ITester

Private assertion As Boolean
Private answerCollection As Answers

Private Property Get ITester_className() As String
    ITester_className = "TestAnswers"
End Property

Private Property Get ITester_testList() As Variant
    ITester_testList = Array( _
        "test_Answers_WhenAddListQuestionType_ShouldAdd", _
        "test_Answers_WhenAddCheckboxQuestionType_ShouldAdd", _
        "test_Answers_WhenAddTextQuestionType_ShouldAdd", _
        "test_Answers_WhenAddSliderQuestionType_ShouldAdd", _
        "test_Answers_WhenAddMultipleQuestions_ShouldCount", _
        "test_Answers_WhenAddMultipleQuestions_ShouldGetItem", _
        "test_Answers_WhenAddMultipleQuestionsAndRemove_ShouldRemoveItem" _
    )
   
End Property

Private Function ITester_runTest(methodName As String) As Boolean
     
    If Len(methodName) > 63 Then MsgBox "The method name '" & methodName & "' is too long to run on the Mac os.", vbCritical, MsgTitle
    ITester_runTest = CallByName(Me, methodName, VbMethod)

End Function

Private Sub ITester_setUp()
    Set answerCollection = New Answers
End Sub

Private Sub ITester_breakDown()
    Set answerCollection = Nothing
End Sub

Public Function test_Answers_WhenAddListQuestionType_ShouldAdd() As Boolean
    
    Dim answer As New ModelAnswerList
    answer.value = 2

    answerCollection.Add answer
    Set answer = answerCollection.item(1)
  
    assertion = answer.value = 2
    test_Answers_WhenAddListQuestionType_ShouldAdd = assertion
    
End Function

Public Function test_Answers_WhenAddCheckboxQuestionType_ShouldAdd() As Boolean
    
    Dim answer As New ModelAnswerCheckbox
    answer.value = Array(2, 4)

    answerCollection.Add answer
    Set answer = answerCollection.item(1)
  
    assertion = answer.value(0) = 2 And answer.value(1) = 4
    test_Answers_WhenAddCheckboxQuestionType_ShouldAdd = assertion

End Function

Public Function test_Answers_WhenAddTextQuestionType_ShouldAdd() As Boolean

    Dim answer As New ModelAnswerText
    answer.value = "Test"

    answerCollection.Add answer
    Set answer = answerCollection.item(1)

    assertion = answer.value = "Test"
    test_Answers_WhenAddTextQuestionType_ShouldAdd = assertion

End Function

Public Function test_Answers_WhenAddSliderQuestionType_ShouldAdd() As Boolean

    Dim answer As New ModelAnswerSlider
    answer.value = 1.34

    answerCollection.Add answer
    Set answer = answerCollection.item(1)

    assertion = 1.34
    test_Answers_WhenAddSliderQuestionType_ShouldAdd = assertion

End Function

Public Function test_Answers_WhenAddMultipleQuestions_ShouldCount() As Boolean
    
    Dim answer As New ModelAnswerList
    answer.value = 2
    answerCollection.Add answer
    answer.value = 4
    answerCollection.Add answer
      
    assertion = answerCollection.count = 2
    test_Answers_WhenAddMultipleQuestions_ShouldCount = assertion
    
End Function

Public Function test_Answers_WhenAddMultipleQuestions_ShouldGetItem() As Boolean
    
    Dim answer As New ModelAnswerList
    answer.value = 2
    answerCollection.Add answer
    answer.value = 4
    answerCollection.Add answer
    
    Set answer = answerCollection.item(2)
  
    assertion = answer.value = 4
    test_Answers_WhenAddMultipleQuestions_ShouldGetItem = assertion
    
End Function

Public Function test_Answers_WhenAddMultipleQuestionsAndRemove_ShouldRemoveItem() As Boolean

    Dim answer1 As New ModelAnswerList
    Dim answer2 As New ModelAnswerList
    answer1.value = 2
    answerCollection.Add answer1
    answer2.value = 4
    answerCollection.Add answer2
    
    answerCollection.Remove 2
    Set answer1 = answerCollection.item(1)
    assertion = answerCollection.count = 1 And answer1.value = 2
    
    test_Answers_WhenAddMultipleQuestionsAndRemove_ShouldRemoveItem = assertion
    
End Function
