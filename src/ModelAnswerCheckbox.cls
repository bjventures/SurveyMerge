VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelAnswerCheckbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Model")
'
' class module: ModelAnswerCheckbox
' Note that many methods are delegated to the AnswerBase class.
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

Public Property Let value(ByVal value As Variant)
    ' Note that the value argument must be an array of integers.
    On Error GoTo Catch
    
    Dim bound As Long
    Dim i As Long
    
    bound = UBound(value)
    ReDim m_value(bound)
    For i = 0 To bound
        m_value(i) = value(i)
    Next i
Finally:
    Exit Property
 
Catch:
    Err.Raise CustomError.ModelValidationError
    Resume Finally

End Property

Public Property Get value() As Variant

    value = m_value
    
End Property

'
' Delegate properties to implemented base class (AnswerBase)
'
Private Property Let ModelAnswerBase_number(ByVal value As Long)
    
    baseAnswer.number = value
    
End Property

Private Property Get ModelAnswerBase_number() As Long

    ModelAnswerBase_number = baseAnswer.number
    
End Property

Private Property Let ModelAnswerBase_isoTime(ByVal value As String)
    
    baseAnswer.isoTime = value
    
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

