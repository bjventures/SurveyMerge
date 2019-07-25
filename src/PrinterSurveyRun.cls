VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PrinterSurveyRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' class module: PrinterSurveyRun
'
Option Explicit

Implements IPrinter

Private i As Integer

Sub IPrinter_printData(surveyRun As ModelSurveyRun, lineNo As Integer)

    Dim ws As Worksheet
    Dim answer As Variant

    Set ws = Sheets(getWsName(WsSheet.Answers))
    Call printAnswerWs(ws, surveyRun, lineNo)
    
    Set ws = Sheets(getWsName(WsSheet.AnswerTime))
    Call printAnswerTimeWs(ws, surveyRun, lineNo)


End Sub

Private Sub printAnswerWs(ws As Worksheet, surveyRun As ModelSurveyRun, lineNo As Integer)
            
    Call printAnswerWsHeaderLine(ws, surveyRun, lineNo)
    
    ws.Cells(lineNo + 2, 1) = surveyRun.surveyName
    ws.Cells(lineNo + 2, 2) = surveyRun.participantId
    ws.Cells(lineNo + 2, 3) = surveyRun.startTime
    ws.Cells(lineNo + 2, 4) = surveyRun.endTime

    For i = 1 To surveyRun.questionCount
        ws.Cells(lineNo + 2, 4 + i) = surveyRun.answerCollection.item(i).description
    Next i
    ws.UsedRange.Columns(1).AutoFit
    ws.UsedRange.Columns(2).AutoFit
    ws.UsedRange.Columns(3).AutoFit
    ws.UsedRange.Columns(4).AutoFit

End Sub

Private Sub printAnswerTimeWs(ws As Worksheet, surveyRun As ModelSurveyRun, lineNo As Integer)
        
   Call printAnswerWsHeaderLine(ws, surveyRun, lineNo)
    
    ws.Cells(lineNo + 2, 1) = surveyRun.surveyName
    ws.Cells(lineNo + 2, 2) = surveyRun.participantId
    ws.Cells(lineNo + 2, 3) = surveyRun.startTime
    ws.Cells(lineNo + 2, 4) = surveyRun.endTime

    For i = 1 To surveyRun.questionCount
        ' Dates aren't set if the question is skipped by branching.
        If surveyRun.answerCollection.item(i).time <> CDate(0) Then
            ws.Cells(lineNo + 2, 4 + i) = surveyRun.answerCollection.item(i).time
        End If
    Next i
    ws.UsedRange.Columns.AutoFit

End Sub

Private Sub printAnswerWsHeaderLine(ws As Worksheet, surveyRun As ModelSurveyRun, lineNo As Integer)
    ' Print header for first survey run, do not repeat.
    If lineNo <> 0 Then Exit Sub
        
    With ws.Cells(1, 1)
        .value = "Survey Name"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    With ws.Cells(1, 2)
        .value = "Participant ID"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    With ws.Cells(1, 3)
        .value = "Start Time"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    With ws.Cells(1, 4)
        .value = "Finish Time"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    For i = 1 To surveyRun.questionCount
        With ws.Cells(1, 4 + i)
            .value = i
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Next i
    
End Sub

Sub IPrinter_printError(description As String, lineNo As Integer)

    Dim ws As Worksheet
    Set ws = Sheets(getWsName(WsSheet.Answers))
    ws.Cells(lineNo + 2, 1) = "Error In Survey Run: " & description

End Sub
