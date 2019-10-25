Attribute VB_Name = "Export"
'@Folder("SurveyMerge.Controllers")
Option Explicit

Private Sub exportVisualBasicCode()
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    
    Dim VBComponent As Object
    Dim count As Long
    Dim path As String
    Dim SrcDirectory As String
    Dim testingDirectory As String
    Dim extension As String
    
    If Not isFileAccessAllowed Then
        MsgBox "Failed to export files as access to the '/testing' folder has not been granted. Please reinstall and grant access.", vbOKOnly, ProjectName
        Exit Sub
    End If
        
    On Error GoTo Catch
    count = 0
    SrcDirectory = getCurrentPath & SrcFolder
    testingDirectory = getCurrentPath & testFolder
        
    If Not directoryExists(SrcDirectory) Then
        MkDir SrcDirectory
    End If
    If Not directoryExists(testingDirectory) Then
        MkDir testingDirectory
    End If
    
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
        Case Document
            ' We don't want worksheets
            GoTo NextItem
        Case ClassModule
            extension = ".cls"
        Case Form
            extension = ".frm"
        Case Module
            extension = ".bas"
        Case Else
            extension = ".txt"
        End Select
            
        ' On Error Resume Next
        ' Err.Clear
        
        Select Case InStr(VBComponent.name, "Test")
        Case 0
            path = SrcDirectory & "/" & VBComponent.name & extension
        Case Else
            path = testingDirectory & "/" & VBComponent.name & extension
        End Select
        
        VBComponent.Export path
        
        If Err.number <> 0 Then
            MsgBox "Failed to export " & VBComponent.name & " to " & path, vbCritical, ProjectName
        Else
            count = count + 1
            Debug.Print "Exported: " & VBComponent.name
        End If
NextItem:
    Next
    
Finally:
    Exit Sub
Catch:
    Dim msg As String
    If Err.number = 1004 Then
        msg = "Unable to export files. Please ensure 'Trust access to the VBA project object model' is checked"
    Else
        msg = Err.description
    End If
    MsgBox msg, vbCritical, ProjectName
    Resume Finally
    
End Sub

Private Function isFileAccessAllowed() As Boolean
    ' Grant file access is only needed on the Mac for versions later than Excel 2016 due to sandbox protection.
    #If Mac Then
        #If MAC_OFFICE_VERSION < 16 Then
            isFileAccessAllowed = True
            Exit Function
        #Else
            isFileAccessAllowed = GrantAccessToMultipleFiles(Array(getCurrentPath))
            Exit Function
        #End If
    #Else
        isFileAccessAllowed = True
    #End If
        
End Function

