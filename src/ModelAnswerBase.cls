VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelAnswerBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Model")
'
' class module: ModelAnswerBase
' This is used to mimic inheritance. See: https://docs.microsoft.com/en-us/previous-versions/visualstudio/visual-basic-6/aa240846(v%3dvs.60)
'
Option Explicit

Private m_number As Long
Private m_time As Date
Private m_offset As Long

Public Property Let number(ByVal value As Long)
    
    If value < 1 Then Err.Raise CustomError.ModelValidationError
    m_number = value
    
End Property

Public Property Get number() As Long

    number = m_number
    
End Property

Public Property Let isoTime(ByVal value As String)

    On Error GoTo Catch
    m_time = DateValue(Mid$(value, 1, 10)) + TimeValue(Mid$(value, 12, 8))
    m_offset = Int(Mid$(value, 20, 5))
Finally:
    Exit Property

Catch:
    Err.Raise CustomError.ModelValidationError
    Resume Finally

End Property

Private Property Get isoTime() As String
    isoTime = "This is not used."
End Property

Public Property Get isoOffset() As Long

    isoOffset = m_offset
    
End Property

Public Property Get time() As Date

    time = m_time
    
End Property

Public Property Get description() As String
    
    description = vbNullString

End Property

