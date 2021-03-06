Attribute VB_Name = "ProjectConstants"
'@Folder("SurveyMerge.Utilities")
'
' module: Const
'
Option Explicit

Public Const ProjectName As String = "SurveyMerge"
Public Const SrcFolder As String = "src"
Public Const testFolder As String = "tests"

Enum CustomError
    IncorrectDataFormat = 515
    AnswerCountError = 516
    InvalidValue = 517
    ModelValidationError = 518
    SetupError = 519
    SurveyRunError = 520
    InvalidQuestionType = 521
    FileNotFound = 522
    DirNotFound = 522
End Enum

Enum FileCol
    keyword = 1
    metadata = 2
    answerData = 3
End Enum

' VBA does not have String Enumerations.
Enum WsSheet
    Dashboard = 1
    Answers = 2
    Times = 3
End Enum

Public Function getWsName(ByVal id As WsSheet) As String
    Select Case id
    Case WsSheet.Dashboard
        getWsName = "Dashboard"
    Case WsSheet.Answers
        getWsName = "Answers"
    Case WsSheet.Times
        getWsName = "Answer Time"
    End Select
    
End Function

