VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelDataLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Models.DataFile")
'
' class module: ModelDataLines
'
Option Explicit
Private m_Header As String
Private m_Answer As String
Private m_TimeStamp As String

Public Property Let header(ByVal line As String)
    m_Header = line
End Property

Public Property Get header() As String
    header = m_Header
End Property

Public Property Let answer(ByVal line As String)
    m_Answer = line
End Property

Public Property Get answer() As String
    answer = m_Answer
End Property

Public Property Let timeStamp(ByVal line As String)
    m_TimeStamp = line
End Property

Public Property Get timeStamp() As String
    timeStamp = m_TimeStamp
End Property


