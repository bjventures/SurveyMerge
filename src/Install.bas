Attribute VB_Name = "Install"
'@Folder("SurveyMerge.Controller")
Option Explicit

Private Sub installEndUser()

    Dim sheetArray() As Variant
    If Not sheetExists(getWsName(WsSheet.Dashboard)) Then
        sheetArray = Array(getWsName(WsSheet.Dashboard), getWsName(WsSheet.Answers), getWsName(WsSheet.Times))
        doFirstInstall (sheetArray)
    End If

End Sub

'@Ignore ProcedureNotUsed
Private Sub installDeveloper()

    installEndUser

End Sub

Private Sub doFirstInstall(ByRef sheetArray As Variant)

    Application.ScreenUpdating = False
    createOrClearWorksheets sheetArray
    setupDashboard
    Application.ScreenUpdating = True

End Sub

Private Function sheetExists(ByVal sheetToFind As String, Optional ByRef wb As Workbook) As Boolean
    
    Dim sheet As Worksheet
    sheetExists = False
    If wb Is Nothing Then Set wb = ThisWorkbook
    For Each sheet In wb.Worksheets
        If sheetToFind = sheet.name Then
            sheetExists = True
            Exit Function
        End If
    Next sheet

End Function

Public Sub createOrClearWorksheets(ByVal sheetArray As Variant)

    Dim SheetName As Variant
    Dim sheetString As String

    For Each SheetName In sheetArray
        sheetString = CStr(SheetName)
        If sheetExists(sheetString) Then
            ThisWorkbook.Sheets(sheetString).Cells.ClearContents
        Else
            createSheet sheetString
        End If
    Next SheetName

End Sub

Private Sub setupDashboard()
    
    Dim ws As Worksheet
    Dim btnRange As Range
    Dim btn As Button
    Set ws = ThisWorkbook.Sheets(getWsName(WsSheet.Dashboard))
    ws.Activate
    
    ' Instructions
    With ws.Cells(1, 1)
        .value = "Instructions"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Font.Size = 14
    End With
    ws.Columns("A").ColumnWidth = 75
    ws.Cells(2, 1).value = getInstructions1
    ws.Cells(3, 1).value = getInstructions2
    ws.Cells(5, 1).value = getInstructions3
    
    ' Links
    With ws
        .Hyperlinks.Add Anchor:=.Range("A4"), _
        Address:="https://pielsurvey.org/contact", _
        ScreenTip:="PIEL Survey contact form", _
        TextToDisplay:="Contact PIEL Survey"
        .Hyperlinks.Add Anchor:=.Range("A6"), _
        Address:="https://github.com/bjventures/SurveyMerge", _
        ScreenTip:="Repository", _
        TextToDisplay:="Participate in open source project"
    End With
    ws.Cells(4, 1).HorizontalAlignment = xlCenter
    ws.Cells(6, 1).HorizontalAlignment = xlCenter
    ws.UsedRange.WrapText = True
    
    
    ' Button
    Set btnRange = ws.Range("A8")
    Set btn = ws.Buttons.Add(btnRange.Left + 145, btnRange.Top, 100, 25)
    With btn
        .Caption = "Combine Files"
        .name = "btnCombine"
        .Font.Bold = True
        .OnAction = "combineCsvFiles"
    End With

End Sub

Private Sub createSheet(ByVal name As String)
    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.count)).name = name
    End With
End Sub

Private Function getInstructions1() As String

    Dim returnString As String
    
    returnString = "To import the PIEL Survey data files (with '.csv' extension):" & _
                   "  1. Copy all the data files into the same folder at this Workbook." & _
                   "  2. Click on the button below." & _
                   "  3. Check the resulting imported data. Errors (if any) will be printed in the file." & vbCrLf & vbCrLf & _
                   "Each time that you click on the button, the previous imported data will be overwritten. " & _
                   "This allows you to merge the data as often as you like as new data files are received."
    getInstructions1 = returnString

End Function

Private Function getInstructions2() As String

    Dim returnString As String
    
    returnString = "Note that this is a Beta version of the data merge tool, " & _
                   "let us know of any problems by clicking on the contact link below."
    getInstructions2 = returnString

End Function

Private Function getInstructions3() As String

    Dim returnString As String
    
    returnString = "This software is an open source project. " & _
                   "Click on the link below for the repository that contains the code and license terms. " & _
                   "It would be great if your team could contribute to the project and improve it for other researchers."
    getInstructions3 = returnString

End Function




