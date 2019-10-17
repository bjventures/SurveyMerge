Attribute VB_Name = "Module1"
'@Folder("Tests.Helpers")
Option Explicit

'Public Function getFileLines(fileNameStub As String) As Variant
'    ' Note that the text file should only have the 3 answer lines.
'
'On Error GoTo Catch
'    Dim inputFile As Integer
'    Dim fileString As String
'    Dim fileName As String
'    Dim lineArray As Variant
'
'    fileName = getTestFilePath & fileNameStub & ".csv"
'    inputFile = FreeFile
'    Open fileName For Input As #inputFile
'    fileString = Input(LOF(inputFile), inputFile)
'    Close #inputFile
'
'    lineArray = Split(fileString, vbLf)
'    getAnswerLines = lineArray
'
'Finally:
'  Exit Function
'Catch:
'    Dim msg As String
'    If Err.number = 53 Then
'        msg = "Unable to read file '" & fileNameStub & "'."
'    Else
'        msg = "Error no: " & Err.number & " in 'getAnswerLines'." & vbNewLine & Err.description
'    End If
'    MsgBox msg, vbOKOnly, MsgTitle
'    Close inputFile
'    Resume Finally
'
'End Function

