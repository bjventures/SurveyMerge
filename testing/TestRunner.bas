Attribute VB_Name = "TestRunner"
'@Folder("SurveyMerge.Tests")
'
' module: TestRunner
'
Option Explicit

Private testObject As ITester

Sub runAllTests()
    
    'There will be an error if the worksheet "Dashboard" does not exist.
    If Not isProjectInstalled() Then
    MsgBox "The tests could not be run. The project is not properly installed.", vbOKOnly, ProjectName
        Exit Sub
    End If
    
    If Not isFileAccessAllowed Then
        MsgBox "The tests could not be run. Access to files has not been granted or the folder does not exist. Please reinstall and allow file access.", vbOKOnly, ProjectName
        Exit Sub
    End If
    
    If Not directoryExists(getTestFilePath) Then
        MsgBox "The tests could not be run. The folder 'testing/test-files' does not exist.", vbCritical, ProjectName
        Exit Sub
    End If

    Set testObject = New TestMain
    runTestClass
    Set testObject = New TestAnswers
    runTestClass
    Set testObject = New TestModelAnswerBase
    runTestClass
    Set testObject = New TestModelAnswerCheckbox
    runTestClass
    Set testObject = New TestModelAnswerList
    runTestClass
    Set testObject = New TestModelAnswerSlider
    runTestClass
    Set testObject = New TestModelAnswerText
    runTestClass
    Set testObject = New TestParserAnswers
    runTestClass
    Set testObject = New TestParserFile
    runTestClass
    Set testObject = New TestParserSurveyRun
    runTestClass

End Sub

Private Function isProjectInstalled() As Boolean
    
    Dim sheet As Worksheet
    isProjectInstalled = False
    For Each sheet In ThisWorkbook.Worksheets
        If sheet.name = getWsName(WsSheet.Dashboard) Then
            isProjectInstalled = True
            Exit Function
        End If
    Next sheet

End Function

Private Sub runTestClass()
        
    Dim methodName As Variant
    Dim nameLength As Long
    Dim result As Boolean
    
    On Error GoTo Catch
    
    nameLength = Len(testObject.className)
    Debug.Print String(nameLength, "=")
    Debug.Print testObject.className
    Debug.Print String(nameLength, "=")

    For Each methodName In testObject.testList
       testObject.setUp
       result = testObject.runTest(CStr(methodName))
       If result Then
           Debug.Print "Passed: " & methodName
       Else
           Debug.Print "FAILED: " & methodName
       End If
       testObject.breakDown
    Next
    
Finally:
    Exit Sub

Catch:
    MsgBox "An unexpected error occurred in the test '" & testObject.className & "." & methodName & "'." & vbCrLf & Err.description, vbCritical, ProjectName
    Resume Finally
    
End Sub

Function getTestFilePath() As String
    
    getTestFilePath = getCurrentPath & "testing/test-files/"

End Function

Function getTestFileName(ByRef fileName As String) As String
    
    getTestFileName = fileName & ".csv"

End Function


Public Function getAnswerLines(ByRef fileNameStub As String) As Variant
    ' Note that the text file should only have the 3 answer lines.

On Error GoTo Catch
    Dim inputFile As Long
    Dim fileString As String
    Dim fileName As String
    Dim lineArray As Variant
    
    fileName = getTestFilePath & getTestFileName(fileNameStub)
    inputFile = FreeFile
    Open fileName For Input As #inputFile
    fileString = Input$(LOF(inputFile), inputFile)
    Close #inputFile
    
    lineArray = Split(fileString, vbLf)
    getAnswerLines = lineArray
    
Finally:
  Exit Function
Catch:
    Dim msg As String
    If Err.number = 53 Then
        msg = "Unable to read file '" & fileNameStub & "'."
    Else
        msg = "Error no: " & Err.number & " in 'getAnswerLines'." & vbNewLine & Err.description
    End If
    MsgBox msg, vbOKOnly, ProjectName
    Close inputFile
    Resume Finally
    
End Function
