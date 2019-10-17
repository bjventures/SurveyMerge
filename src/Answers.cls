VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Answers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Models")
Option Explicit

' Custom collection of type AnswerBase
Private Answers As New Collection

Public Sub Add(ByRef answer As ModelAnswerBase)

    Answers.Add answer
    
End Sub

Public Property Get count() As Long

    count = Answers.count

End Property

Public Property Get item(ByRef IndexOrName As Variant) As ModelAnswerBase

    ' The question number starts at 1.
    Set item = Answers(IndexOrName)

End Property

Public Sub Remove(ByRef IndexOrName As Variant)
  
    Answers.Remove IndexOrName

End Sub

