VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestMockPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Tests")
'
' class module: TestMockPrinter
'
Option Explicit

Implements IPrinter

Private m_ValidSurveyRunNumber As Long
Private m_SurveyRun As ModelSurveyRun
Private m_ErrorMessage As String

Sub IPrinter_printData(ByRef surveyRun As ModelSurveyRun, ByRef lineNo As Long)

    Set m_SurveyRun = surveyRun
    m_ValidSurveyRunNumber = m_ValidSurveyRunNumber + 1

End Sub

Sub IPrinter_printError(ByRef errorMessage As String, ByRef lineNo As Long)

    m_ErrorMessage = errorMessage

End Sub

Public Property Get validSurveyRunNumber() As Long

     validSurveyRunNumber = m_ValidSurveyRunNumber

End Property

Public Property Get surveyRun() As ModelSurveyRun

     Set surveyRun = m_SurveyRun

End Property

Public Property Get errorMessage() As String

     errorMessage = m_ErrorMessage

End Property

