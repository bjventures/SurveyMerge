Attribute VB_Name = "Constants"
'
' module: Constants
'
Option Explicit

Public Const MsgTitle As String = "SurveyMerge"

Enum CustomError
    UnknownKeyword = 514
    IncorrectDataFormat = 515
    AnswerCountError = 516
    InvalidValue = 517
    ModelValidationError = 518
    SetupError = 519
    SurveyRunError = 520
    InvalidQuestionType = 521
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
    Answers = 2
    AnswerTime = 3
End Enum

Function getWsName(id As WsSheet) As String
    Select Case id
        Case WsSheet.Dashboard
            getWsName = "Dashboard"
        Case WsSheet.Answers
            getWsName = "Answers"
        Case WsSheet.AnswerTime
            getWsName = "Answer Time"
    End Select
    
End Function

