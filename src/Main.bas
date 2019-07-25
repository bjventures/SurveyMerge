Attribute VB_Name = "Main"
'
' module: Main
'
Option Explicit

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
