VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelAnswerCheckbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Models.AnswerModels")
'
' class module: ModelAnswerCheckbox
' Many methods are delegated to the AnswerBase class.
'
Option Explicit
Implements ModelAnswerBase

Private baseAnswer As ModelAnswerBase
Private m_value() As Long

'
' Class specific properties
'
Private Sub Class_Initialize()
    Set baseAnswer = New ModelAnswerBase
End Sub

Public Property Let value(ByVal answerValue As Variant)
    ' The value argument must be an array of integers.
    Dim bound As Long
    Dim i As Long
    bound = UBound(answerValue)
    
    ReDim m_value(bound)
    For i = 0 To bound
        If answerValue(i) < 1 Then Err.Raise CustomError.ModelValidationError, "ModelAnswerCheckbox.value", "The value '" & answerValue & "' is not valid."
        m_value(i) = answerValue(i)
    Next i
End Property

Public Property Get value() As Variant
    value = m_value
End Property

'
' Delegate properties to implemented base class (AnswerBase)
'
Private Property Let ModelAnswerBase_number(ByVal answerValue As Long)
    baseAnswer.number = answerValue
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
    Dim i As Long
    Dim bound As Long
    Dim strArray() As String
    
    bound = UBound(m_value)
    ReDim strArray(bound) As String
    For i = 0 To bound
        strArray(i) = CStr(m_value(i))
    Next i
    
    ModelAnswerBase_description = CStr(Join(strArray, ","))
End Property

