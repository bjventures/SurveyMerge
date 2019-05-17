Attribute VB_Name = "Main"
'
' module: Main
'
Option Explicit

Private Sub install()

    Dim sheetArray() As Variant
    If Not sheetExists(getWsName(WsSheet.Dashboard)) Then
        sheetArray = Array(getWsName(WsSheet.Dashboard), getWsName(WsSheet.Answers), getWsName(WsSheet.AnswerTime))
        Call doFirstInstall(sheetArray)
    End If

End Sub

Sub combineCsvFiles(Optional currentPath As String = "", Optional showMsg As Boolean = True)
         
    On Error GoTo Catch
      
    Dim sheetArray() As Variant
    Dim success As Integer
    Dim fileArray() As String
    Dim fileName As String
    Dim i As Integer
    Dim arrayBound As Integer
    Dim parser As ParserFile
    Dim printer As PrinterSurveyRun
    Dim lineCounter As Integer
        
    If currentPath = "" Then currentPath = getCurrentPath
    fileArray = getFileList(currentPath)
    arrayBound = UBound(fileArray)
    
    If arrayBound < 1 Then
        MsgBox "No data files were found in the current directory.", vbOKOnly, MsgTitle
        Exit Sub
    End If
    
    Set parser = New ParserFile
    Set printer = New PrinterSurveyRun
    
    Application.ScreenUpdating = False
    
    sheetArray = Array(getWsName(WsSheet.Answers), getWsName(WsSheet.AnswerTime))
    Call createOrClearWorksheets(sheetArray)
    
    success = 0
    lineCounter = 0
    For i = 0 To arrayBound - 1
        lineCounter = parser.parse(currentPath, fileArray(i), printer, lineCounter)
        success = success + 1
    Next

    Application.ScreenUpdating = True
    If showMsg Then MsgBox success & " CSV files were combined.", vbOKOnly, MsgTitle

Finally:
    Exit Sub

Catch:
    If showMsg Then MsgBox "The file could not be imported. " & Err.description, vbOKOnly, MsgTitle
    Resume Finally

End Sub

Private Sub doFirstInstall(sheetArray As Variant)

    Application.ScreenUpdating = False
    Call createOrClearWorksheets(sheetArray)
    Call setupDashboard
    Application.ScreenUpdating = True

End Sub

Private Sub setupDashboard()
    
    Dim ws As Worksheet
    Dim btnRange As Range
    Dim btn As Button
    Set ws = Sheets(getWsName(WsSheet.Dashboard))
    ws.Activate
    
    ' Instructions
    With ws.Cells(1, 1)
        .value = "Instructions"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Font.Size = 14
    End With
    ws.Columns("A").ColumnWidth = 75
    ws.Cells(2, 1) = getInstructions1
    ws.Cells(3, 1) = getInstructions2
    ws.Cells(5, 1) = getInstructions3
    
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

Private Function getInstructions1() As String

    Dim returnString As String
    
    returnString = "To import the PIEL Survey data files (with '.csv' extension):"
    returnString = returnString & vbCrLf & "  1. Copy all the data files into the same folder at this Workbook."
    returnString = returnString & vbCrLf & "  2. Click on the button below."
    returnString = returnString & vbCrLf & "  3. Check the resulting imported data. Errors (if any) will be printed in the file."
    returnString = returnString & vbCrLf
    returnString = returnString & vbCrLf & "Each time that you click on the button, the previous imported data will be overwritten. "
    returnString = returnString & "This allows you to merge the data as often as you like as new data files are received."
    getInstructions1 = returnString

End Function

Private Function getInstructions2() As String

    Dim returnString As String
    
    returnString = "Note that this is a Beta version of the data merge tool, "
    returnString = returnString & "let us know of any problems by clicking on the contact link below."
    getInstructions2 = returnString

End Function

Private Function getInstructions3() As String

    Dim returnString As String
    
    returnString = "This software is an open source project. Click on the link below for the repository that contains the code and license terms. "
    returnString = returnString & "It would be great if your team could contribute to the project and improve it for other researchers."
    getInstructions3 = returnString

End Function

Private Function createSheet(name As String)
    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.count)).name = name
    End With
End Function

Private Function sheetExists(sheetToFind As String, Optional wb As Workbook) As Boolean
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

Private Function createOrClearWorksheets(sheetArray As Variant)

    Dim SheetName As Variant
    Dim sheetString As String

    For Each SheetName In sheetArray
        sheetString = CStr(SheetName)
        If sheetExists(sheetString) Then
            Sheets(sheetString).Cells.ClearContents
        Else
            createSheet (sheetString)
        End If
    Next SheetName

End Function

'
' Returns an Array of Strings containing the full path to files.
'
Private Function getFileList(currentDir As String, Optional extension As String = "csv") As String()

    Dim fileArray() As String
    Dim fileCount As Integer
    Dim fileName As String

    fileCount = 0
    ' Create an array of zero length.
    fileArray = Split("")
    fileName = Dir(currentDir & "*." & extension, vbNormal)
    Do While Len(fileName) > 0
        fileCount = fileCount + 1
        ReDim Preserve fileArray(0 To fileCount)
        fileArray(fileCount - 1) = fileName
        fileName = Dir
    Loop
    
    getFileList = fileArray
    
End Function
