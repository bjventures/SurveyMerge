VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestMockPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Helpers")
'
' class module: TestMockPrinter
'
Option Explicit

Implements IPrinter

Private m_ValidSurveyRunNumber As Long
Private m_SurveyRun As ModelSurveyRun
Private m_ErrorMessage As String

'@Ignore ParameterNotUsed
Public Sub IPrinter_printData(ByRef singleRun As ModelSurveyRun, ByRef lineNo As Long)
    Set m_SurveyRun = singleRun
    m_ValidSurveyRunNumber = m_ValidSurveyRunNumber + 1

End Sub

'@Ignore ParameterNotUsed
Public Sub IPrinter_printError(ByVal message As String, ByVal lineNo As Long)

    m_ErrorMessage = message

End Sub

Public Property Get validSurveyRunNumber() As Long

    validSurveyRunNumber = m_ValidSurveyRunNumber

End Property

Public Property Get surveyRun() As ModelSurveyRun
    If surveyRun Is Nothing Then Set m_SurveyRun = New ModelSurveyRun
    Set surveyRun = m_SurveyRun

End Property

Public Property Set surveyRun(ByRef value As ModelSurveyRun)
    Set m_SurveyRun = value

End Property

Public Property Get errorMessage() As String

    errorMessage = m_ErrorMessage

End Property

