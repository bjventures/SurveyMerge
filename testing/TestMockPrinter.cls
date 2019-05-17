VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestMockPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' class module: TestMockPrinter
'
Option Explicit

Implements IPrinter

Private m_ValidSurveyRunNumber As Integer
Private m_SurveyRun As ModelSurveyRun
Private m_ErrorMessage As String

Sub IPrinter_printData(surveyRun As ModelSurveyRun, lineNo As Integer)

    Set m_SurveyRun = surveyRun
    m_ValidSurveyRunNumber = m_ValidSurveyRunNumber + 1

End Sub

Sub IPrinter_printError(errorMessage As String, lineNo As Integer)

    m_ErrorMessage = errorMessage

End Sub

Public Property Get validSurveyRunNumber() As Integer

     validSurveyRunNumber = m_ValidSurveyRunNumber

End Property

Public Property Get surveyRun() As ModelSurveyRun

     Set surveyRun = m_SurveyRun

End Property

Public Property Get errorMessage() As String

     errorMessage = m_ErrorMessage

End Property

