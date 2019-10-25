Attribute VB_Name = "Utilities"
'@Folder("SurveyMerge.Utilities")
'
' module: Utilities
'
Option Explicit

Public Function getCurrentPath() As String

    Dim currentPath As String
    currentPath = ActiveWorkbook.path
    If Right$(currentPath, 1) <> "/" Then currentPath = currentPath & "/"
    If Not directoryExists(currentPath) Then Err.Raise CustomError.DirNotFound, ProjectName & ".getCurrentPath", "The directory does not exist."
    getCurrentPath = currentPath

End Function

Public Function directoryExists(ByRef strDir As String) As Boolean
    
    ' Need this approach since on Mac comparing to empty string gives an incorrect result if the directory is empty.
    directoryExists = IIf(Len(Dir(strDir, vbDirectory)) = 0, False, True)

End Function

Public Function clearOrAddSpreadsheets(ByVal sheets As Variant)
    Dim result As Boolean
    Dim singleSheet As Variant
    Dim saveCalcState As Long
    Dim sheetName As String
    saveCalcState = Application.Calculation
    
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    
    For Each singleSheet In sheets
        If sheetExists(singleSheet) Then
            ThisWorkbook.sheets(singleSheet).Cells.ClearContents
        Else
            With ThisWorkbook
                .sheets.Add(After:=.sheets(.sheets.count)).name = singleSheet
            End With
        End If
    Next singleSheet

    Application.ScreenUpdating = True
    Application.Calculation = saveCalcState
End Function

Public Function sheetExists(ByVal sheetToFind As String, Optional ByRef wb As Workbook) As Boolean
    Dim Sheet As Worksheet
    sheetExists = False
    If wb Is Nothing Then Set wb = ThisWorkbook
    For Each Sheet In wb.Worksheets
        If sheetToFind = Sheet.name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function

