Attribute VB_Name = "Main"
'@Folder("SurveyMerge.Controllers")
'
' module: Main
'
Option Explicit

'
' The arguments for this Sub are to enable testing.
'
Public Sub combineCsvFiles(Optional ByVal currentPath As String = vbNullString, Optional ByVal showMsg As Boolean = True)
    Dim sheetArray() As Variant
    Dim success As Long
    Dim fileArray() As String
    Dim fileCounter As Long
    Dim arrayBound As Long
    Dim parser As ParserFile
    Dim printer As PrinterSurveyRun
    Dim lineCounter As Long
    Dim isFirstFile As Boolean
        
    On Error GoTo Catch
        
    If currentPath = vbNullString Then currentPath = getCurrentPath
    fileArray = getFileList(currentPath)
    arrayBound = UBound(fileArray)
    
    If arrayBound < 1 Then Err.Raise CustomError.FileNotFound, ProjectName & ".combineCsvFiles", "No data files were found in the current directory."
    
    Set parser = New ParserFile
    Set printer = New PrinterSurveyRun
    
    Application.ScreenUpdating = False
    
    sheetArray = Array(getWsName(WsSheet.Answers), getWsName(WsSheet.Times))
    clearOrAddSpreadsheets sheetArray
    Worksheets(getWsName(WsSheet.Dashboard)).Activate
    
    success = 0
    lineCounter = 0
    For fileCounter = 0 To arrayBound - 1
        isFirstFile = IIf(fileCounter > 0, False, True)
        lineCounter = parser.parse(currentPath, fileArray(fileCounter), printer, lineCounter, isFirstFile)
        success = success + 1
    Next fileCounter

    Application.ScreenUpdating = True
    If showMsg Then MsgBox success & " CSV files were combined.", vbOKOnly, ProjectName

Finally:
    Exit Sub

Catch:
    ' Delete any imported data, if an error occured, it is unreliable.
    If Err.number <> CustomError.FileNotFound Then clearOrAddSpreadsheets sheetArray
    If showMsg Then MsgBox "The file could not be imported." & vbNewLine & Err.description, vbOKOnly, ProjectName
    Resume Finally

End Sub

'
' Returns an Array of Strings containing the full path to files.
'
Private Function getFileList(ByVal currentDir As String, Optional ByVal extension As String = "csv") As String()

    Dim fileArray() As String
    Dim fileCount As Long
    Dim fileName As String

    fileCount = 0
    ' Create an array of zero length.
    fileArray = Split(vbNullString)
    fileName = Dir(currentDir & "*." & extension, vbNormal)
    Do While Len(fileName) > 0
        fileCount = fileCount + 1
        ReDim Preserve fileArray(0 To fileCount)
        fileArray(fileCount - 1) = fileName
        fileName = Dir
    Loop
    
    getFileList = fileArray
    
End Function

