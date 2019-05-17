VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelAnswerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' class module: ModelAnswerList
' Note that many methods are delegated to the AnswerBase class.
'
Option Explicit
Implements ModelAnswerBase

Private baseAnswer As ModelAnswerBase
Private m_value As Integer

'
' Class specific properties
'
Private Sub Class_Initialize()

    Set baseAnswer = New ModelAnswerBase

End Sub

Public Property Let value(ByVal value As Integer)
    
    If value < 1 Then Call Err.Raise(CustomError.ModelValidationError)
    m_value = value
    
End Property
 
Public Property Get value() As Integer

    value = m_value
    
End Property

'
' Delegate properties to implemented base class (AnswerBase)
'
Private Property Let ModelAnswerBase_number(ByVal value As Integer)
    
    baseAnswer.number = value
    
End Property
 
Private Property Get ModelAnswerBase_number() As Integer

    ModelAnswerBase_number = baseAnswer.number
    
End Property

Private Property Let ModelAnswerBase_isoTime(ByVal value As String)
    
    baseAnswer.isoTime = value
    
End Property

Private Property Get ModelAnswerBase_time() As Date

    ModelAnswerBase_time = baseAnswer.time
    
End Property

Private Property Get ModelAnswerBase_isoOffset() As Integer

    ModelAnswerBase_isoOffset = baseAnswer.isoOffset
    
End Property

Private Property Get ModelAnswerBase_description() As String
    
    ModelAnswerBase_description = CStr(m_value)
    
End Property
