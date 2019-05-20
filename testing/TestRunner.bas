Attribute VB_Name = "TestRunner"
'
' module: TestRunner
'
Option Explicit

Private testObject As ITester

Sub runAllTests()
    
    'There will be an error if the worksheet "Dashboard" does not exist.
    If Not isProjectInstalled() Then
    MsgBox "The tests could not be run. The project is not properly installed.", vbOKOnly, MsgTitle
        Exit Sub
    End If
    
    If Not isFileAccessAllowed Then
        MsgBox "The tests could not be run. Access to files has not been granted or the folder does not exist. Please reinstall and allow file access.", vbOKOnly, MsgTitle
        Exit Sub
    End If
    
    If Not directoryExists(getTestFilePath) Then
        MsgBox "The tests could not be run. The folder 'testing/test-files' does not exist.", vbCritical, MsgTitle
        Exit Sub
    End If

    Set testObject = New TestMain
    Call runTestClass
    Set testObject = New TestAnswers
    Call runTestClass
    Set testObject = New TestModelAnswerBase
    Call runTestClass
    Set testObject = New TestModelAnswerCheckbox
    Call runTestClass
    Set testObject = New TestModelAnswerList
    Call runTestClass
    Set testObject = New TestModelAnswerSlider
    Call runTestClass
    Set testObject = New TestModelAnswerText
    Call runTestClass
    Set testObject = New TestParserAnswers
    Call runTestClass
    Set testObject = New TestParserFile
    Call runTestClass
    Set testObject = New TestParserSurveyRun
    Call runTestClass

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
    Dim nameLength As Integer
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
    MsgBox "An unexpected error occurred in the test '" & testObject.className & "." & methodName & "'." & vbCrLf & Err.description, vbCritical, MsgTitle
    Resume Finally
    
End Sub

Function getTestFilePath() As String
    
    getTestFilePath = getCurrentPath & "testing/test-files/"

End Function

Function getTestFileName(fileName As String) As String
    
    getTestFileName = fileName & ".csv"

End Function


Public Function getAnswerLines(fileNameStub As String) As Variant
    ' Note that the text file should only have the 3 answer lines.

On Error GoTo Catch
    Dim inputFile As Integer
    Dim fileString As String
    Dim fileName As String
    Dim lineArray As Variant
    
    fileName = getTestFilePath & getTestFileName(fileNameStub)
    inputFile = FreeFile
    Open fileName For Input As #inputFile
    fileString = Input(LOF(inputFile), inputFile)
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
    MsgBox msg, vbOKOnly, MsgTitle
    Close inputFile
    Resume Finally
    
End Function
