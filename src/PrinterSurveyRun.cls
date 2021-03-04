VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PrinterSurveyRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Views")
'
' class module: PrinterSurveyRun
'
Option Explicit
Implements IPrinter

Private questionCounter As Long

Public Sub IPrinter_printData(ByVal surveyRun As ModelSurveyRun, ByVal lineNo As Long)
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.sheets(getWsName(WsSheet.Answers))
    printAnswerWs ws, surveyRun, lineNo
    
    Set ws = ThisWorkbook.sheets(getWsName(WsSheet.Times))
    printTimeWs ws, surveyRun, lineNo
End Sub

Private Sub printAnswerWs(ByRef ws As Worksheet, ByRef surveyRun As ModelSurveyRun, ByVal lineNo As Long)
    Dim cell As Range
    printAnswerWsHeaderLine ws, surveyRun, lineNo
    
    ws.Cells(lineNo + 2, 1).value = surveyRun.surveyName
    ws.Cells(lineNo + 2, 2).value = surveyRun.participantId
    ws.Cells(lineNo + 2, 3).value = surveyRun.startTime
    ws.Cells(lineNo + 2, 4).value = surveyRun.endTime

    For questionCounter = 1 To surveyRun.questionCount
        Set cell = ws.Cells(lineNo + 2, 4 + questionCounter)
        ' Here need to ensure Sliders are shown as number in all regions and Checkbox questions are text with no error.
        If TypeOf surveyRun.answerCollection.item(questionCounter) Is ModelAnswerSlider Then
            cell.value = CDbl(surveyRun.answerCollection.item(questionCounter).description)
        Else
            cell.value = surveyRun.answerCollection.item(questionCounter).description
        End If
        If TypeOf surveyRun.answerCollection.item(questionCounter) Is ModelAnswerCheckbox Then
            cell.NumberFormat = "@"
            cell.Errors(xlNumberAsText).Ignore = True
        End If
    Next questionCounter
    ws.UsedRange.Columns(1).AutoFit
    ws.UsedRange.Columns(2).AutoFit
    ws.UsedRange.Columns(3).AutoFit
    ws.UsedRange.Columns(4).AutoFit
End Sub

Private Sub printTimeWs(ByRef ws As Worksheet, ByRef surveyRun As ModelSurveyRun, ByVal lineNo As Long)
    printAnswerWsHeaderLine ws, surveyRun, lineNo
    
    ws.Cells(lineNo + 2, 1).value = surveyRun.surveyName
    ws.Cells(lineNo + 2, 2).value = surveyRun.participantId
    ws.Cells(lineNo + 2, 3).value = surveyRun.startTime
    ws.Cells(lineNo + 2, 4).value = surveyRun.endTime
'    ws.Cells(lineNo + 2, 3).NumberFormat = "mm.dd.yyyy hh:mm:ss"
'Range("A2", "A50000").NumberFormat = "yyyy-mm-dd"
'Format(date_test, "mm.dd.yyyy hh:mm")
    
    
    For questionCounter = 1 To surveyRun.questionCount
        ' Dates aren't set if the question is skipped by branching.
        If surveyRun.answerCollection.item(questionCounter).time <> CDate(0) Then
            ws.Cells(lineNo + 2, 4 + questionCounter).value = surveyRun.answerCollection.item(questionCounter).time
        End If
    Next questionCounter
    ws.UsedRange.Columns.AutoFit
End Sub

Private Sub printAnswerWsHeaderLine(ByRef ws As Worksheet, ByRef surveyRun As ModelSurveyRun, ByVal lineNo As Long)
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
    
    For questionCounter = 1 To surveyRun.questionCount
        With ws.Cells(1, 4 + questionCounter)
            .value = questionCounter
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Next questionCounter
   
End Sub

Public Sub IPrinter_printError(ByVal description As String, ByVal lineNo As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets(getWsName(WsSheet.Answers))
    ws.Cells(lineNo + 2, 1).value = "Error In Survey Run: " & description
End Sub

