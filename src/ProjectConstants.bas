Attribute VB_Name = "ProjectConstants"
'@Folder("SurveyMerge.Utilities")
'
' module: Const
'
Option Explicit

Public Const ProjectName As String = "SurveyMerge"

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

Enum FileRow
    firstAnswer = 4
End Enum

Enum FileCol
    keyword = 1
    metadata = 2
    answerData = 3
End Enum

' VBA does not have string Enumerations.
Enum WsSheet
    Dashboard = 1
    answers = 2
    Times = 3
End Enum

Public Function getWsName(ByRef id As WsSheet) As String
    Select Case id
    Case WsSheet.Dashboard
        getWsName = "Dashboard"
    Case WsSheet.answers
        getWsName = "Answers"
    Case WsSheet.Times
        getWsName = "Answer Time"
    End Select
    
End Function

