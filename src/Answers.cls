VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Answers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Custom collection of type AnswerBase
Private Answers As New Collection

Sub Add(answer As ModelAnswerBase)

    Answers.Add answer
    
End Sub

Property Get count() As Integer

    count = Answers.count

End Property

Property Get item(IndexOrName As Variant) As ModelAnswerBase

    ' Note the question number starts at 1.
    Set item = Answers(IndexOrName)

End Property

Sub Remove(IndexOrName As Variant)
  
    Answers.Remove IndexOrName

End Sub
