Attribute VB_Name = "TestAnswers"
Option Explicit
Option Private Module

'@TestModule
'@Folder("SurveyMerge.Tests.Models")

Private Assert As Object
Private Fakes As Object
Private answerCollection As Answers

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
    Set answerCollection = New Answers
End Sub

'@TestCleanup
Private Sub TestCleanup()
    Set answerCollection = Nothing
End Sub

'@TestMethod("Model")
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

'@TestMethod("Model")
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

'@TestMethod("Model")
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

'@TestMethod("Model")
Private Sub add_WhenAddSliderQuestion_ShouldAdd()
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

'@TestMethod("Model")
Private Sub add_WhenMultipleQuestions_ShouldAdd()
    On Error GoTo TestFail
    
    Dim listAnswer As ModelAnswerList
    Set listAnswer = New ModelAnswerList
    answerCollection.Add listAnswer
    Dim checkboxAnswer As ModelAnswerCheckbox
    Set checkboxAnswer = New ModelAnswerCheckbox
    answerCollection.Add checkboxAnswer
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

'@TestMethod("Model")
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

'@TestMethod("Model")
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

