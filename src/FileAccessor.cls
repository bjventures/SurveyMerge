VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileAccessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests.Utilities")
Option Explicit
Private Const testFileFolder = testFolder & "/test-files/"
Private lines As Variant
Private mFileText As String

Public Property Get fileText() As String
    fileText = mFileText
End Property

Public Sub loadSurveyRunFile(fileNameStub As String)
    Dim inputFile As Integer
    Dim fileName As String
    Dim lineArray As Variant

    On Error GoTo Catch
    fileName = getCurrentPath & testFileFolder & fileNameStub & ".csv"
    inputFile = FreeFile
    Open fileName For Input As #inputFile
    mFileText = Input(LOF(inputFile), inputFile)
    ' Make sure lines are separated by vbLf
    mFileText = Replace(mFileText, vbCrLf, vbLf)
    ' Just in case there are Mac line breaks
    mFileText = Replace(mFileText, vbCr, vbNullString)
    Close #inputFile

    lines = Split(mFileText, vbLf)
Finally:
    Exit Sub
Catch:
    Dim msg As String
    If Err.number = 53 Then
        msg = "Unable to read file '" & fileNameStub & "'."
    Else
        msg = "Error no: " & Err.number & " in 'FileAccessor.loadSurveyRunFile'." & vbNewLine & Err.description
    End If
    MsgBox msg, vbOKOnly, ProjectName
    Close inputFile
    Resume Finally

End Sub

Public Function getFileRunLines(runNumber As Integer) As ModelDataLines
    ' Each run has 3 lines.
    Dim runStartIndex As Long
    Dim runLines As ModelDataLines
    Set runLines = New ModelDataLines
    runStartIndex = ((runNumber) * 4) - 3
    runLines.header = lines(runStartIndex)
    runLines.answer = lines(runStartIndex + 1)
    runLines.timeStamp = lines(runStartIndex + 2)
    Set getFileRunLines = runLines
End Function

