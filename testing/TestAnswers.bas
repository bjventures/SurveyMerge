Attribute VB_Name = "TestAnswers"
Option Explicit
Option Private Module

'@TestModule
'@Folder("SurveyMerge.NewTests")

Private Assert As Object
Private Fakes As Object
Private answerCollection As Answers

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
    Set answerCollection = New Answers
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Set answerCollection = Nothing
End Sub

'@TestMethod("Answers")
Private Sub add_WhenAddListQuestion_ShouldAdd()
    On Error GoTo TestFail
    
    Dim answer As ModelAnswerList
    Set answer = New ModelAnswerList
    answerCollection.Add answer
   
    Assert.AreEqual CLng(1), answerCollection.count
    Assert.IsTrue TypeOf answerCollection.item(1) Is ModelAnswerList

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Answers")
Private Sub add_WhenAddCheckboxQuestion_ShouldAdd()
    On Error GoTo TestFail
    
    Dim answer As ModelAnswerCheckbox
    Set answer = New ModelAnswerCheckbox

    answerCollection.Add answer
   
    Assert.AreEqual CLng(1), answerCollection.count
    Assert.IsTrue TypeOf answerCollection.item(1) Is ModelAnswerCheckbox

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Answers")
Private Sub add_WhenAddTextQuestion_ShouldAdd()
    On Error GoTo TestFail
    
    Dim answer As ModelAnswerText
    Set answer = New ModelAnswerText

    answerCollection.Add answer
   
    Assert.AreEqual CLng(1), answerCollection.count
    Assert.IsTrue TypeOf answerCollection.item(1) Is ModelAnswerText

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Answers")
Private Sub add_WhenSliderQuestion_ShouldAdd()
    On Error GoTo TestFail
    
    Dim answer As ModelAnswerSlider
    Set answer = New ModelAnswerSlider

    answerCollection.Add answer
   
    Assert.AreEqual CLng(1), answerCollection.count
    Assert.IsTrue TypeOf answerCollection.item(1) Is ModelAnswerSlider


    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Answers")
Private Sub add_WhenMultipleQuestions_ShouldAdd()
    On Error GoTo TestFail
    
    Dim listAnswer As ModelAnswerList
    Set listAnswer = New ModelAnswerList
    answerCollection.Add listAnswer
    Dim checkBoxAnswer As ModelAnswerCheckbox
    Set checkBoxAnswer = New ModelAnswerCheckbox
    answerCollection.Add checkBoxAnswer
    Dim sliderAnswer As ModelAnswerSlider
    Set sliderAnswer = New ModelAnswerSlider
    answerCollection.Add sliderAnswer
    Dim textAnswer As ModelAnswerText
    Set textAnswer = New ModelAnswerText
    answerCollection.Add textAnswer
    
    Assert.AreEqual CLng(4), answerCollection.count

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Answers")
Private Sub item_WhenMultipleQuestions_ShouldRetrieve()
    On Error GoTo TestFail
    
    Dim answer As ModelAnswerList
    Dim retrievedAnswer As ModelAnswerList
    
    Set answer = New ModelAnswerList
    answer.value = 2
    answerCollection.Add answer
    Set answer = New ModelAnswerList
    answer.value = 4
    answerCollection.Add answer
    
    Set retrievedAnswer = answerCollection.item(2)
   
    Assert.AreEqual CLng(2), answerCollection.count
    Assert.AreEqual CLng(4), retrievedAnswer.value

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

'@TestMethod("Answers")
Private Sub remove_WhenAddMultipleQuestionsAndRemove_ShouldRemoveItem()
    On Error GoTo TestFail

    Dim answer As ModelAnswerList
    Dim retrievedAnswer As ModelAnswerList

    Set answer = New ModelAnswerList
    answer.value = 2
    answerCollection.Add answer
    Set answer = New ModelAnswerList
    answer.value = 4
    answerCollection.Add answer

    answerCollection.Remove (2)

    Set retrievedAnswer = answerCollection.item(1)

    Assert.AreEqual CLng(1), answerCollection.count
    Assert.AreEqual CLng(2), retrievedAnswer.value

    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.description
End Sub

