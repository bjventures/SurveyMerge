VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelAnswerSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Models.Answers")
'
' class module: ModelAnswerSlider
' Many methods are delegated to the AnswerBase class.
'
Option Explicit
Implements ModelAnswerBase

Private baseAnswer As ModelAnswerBase
Private m_value As Single

'
' Class specific properties
'
Private Sub Class_Initialize()
    Set baseAnswer = New ModelAnswerBase
End Sub

Public Property Let value(ByVal answerValue As Single)
    If (answerValue < 0) Or (answerValue > 1) Then Err.Raise CustomError.ModelValidationError, "ModelAnswerSlider.value", "The value '" & answerValue & "' is not valid."
    m_value = answerValue
End Property

Public Property Get value() As Single
    value = m_value
End Property

'
' Delegate properties to implemented base class (AnswerBase)
'
Private Property Let ModelAnswerBase_number(ByVal number As Long)
    baseAnswer.number = number
End Property

Private Property Get ModelAnswerBase_number() As Long
    ModelAnswerBase_number = baseAnswer.number
End Property

Private Property Let ModelAnswerBase_isoTime(ByVal answerValue As String)
    baseAnswer.isoTime = answerValue
End Property

Private Property Get ModelAnswerBase_time() As Date
    ModelAnswerBase_time = baseAnswer.time
End Property

Private Property Get ModelAnswerBase_isoOffset() As Long
    ModelAnswerBase_isoOffset = baseAnswer.isoOffset
End Property

Private Property Get ModelAnswerBase_description() As String
    ModelAnswerBase_description = CStr(m_value)
End Property

