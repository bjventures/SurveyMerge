VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParserAnswers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SurveyMerge.Data.Parsers")
' class module: ParserAnswer
'
Option Explicit

Public Function parse(ByVal lines As Variant) As Answers
    ' lines is an array of the 3 data lines from the file.
    Dim cleanArray As Variant
    
    cleanArray = getCleanDataArray(Split(lines(1), ","))

    Set parse = getAnswerCollection(cleanArray, Split(lines(2), ","))

End Function

Private Function getAnswerCollection(ByVal answerArray As Variant, ByVal timeArray As Variant) As Answers
    ' This function returns an AnswerCollection.
    ' Note that the answers must be valid, there is no checking for validity elsewhere.
    ' It throws an error if an unknown question type is encountered.
    ' Note that the question answers are not set, this is for future random order and random questions

    Dim column As Long
    Dim item As Variant
    Dim inputString As String
    Dim cleanQuotedString As String
    Dim strLength As Long
    Dim baseAnswer As ModelAnswerBase
    Dim listAnswer As ModelAnswerList
    Dim checkboxAnswer As ModelAnswerCheckbox
    Dim textAnswer As ModelAnswerText
    Dim sliderAnswer As ModelAnswerSlider
    Dim answerCollection As Answers
    
    Set answerCollection = New Answers
    ' The starting column for the timestamp as an index.
    column = FileCol.answerData - 1
    For Each item In answerArray
        inputString = item
        strLength = Len(inputString)
        
        ' Check if it is a "null" answer
        If strLength = 0 Then
            Set baseAnswer = New ModelAnswerBase
            'baseAnswer.isoTime = timeArray(column)
            answerCollection.Add baseAnswer
            GoTo NextIteration
        End If
        
        ' Check if is text or checkbox answer.
        If strLength >= 2 Then
            If Left$(inputString, 1) = Chr$(34) And Right$(inputString, 1) = Chr$(34) Then
                cleanQuotedString = Mid$(inputString, 2, strLength - 2)
                If isIntegerArray(cleanQuotedString) Then
                    Set checkboxAnswer = New ModelAnswerCheckbox
                    Set baseAnswer = checkboxAnswer
                    baseAnswer.isoTime = timeArray(column)
                    checkboxAnswer.value = convertToIntegerArray(cleanQuotedString)
                    answerCollection.Add checkboxAnswer
                    GoTo NextIteration
                Else
                    Set textAnswer = New ModelAnswerText
                    Set baseAnswer = textAnswer
                    baseAnswer.isoTime = timeArray(column)
                    textAnswer.value = cleanQuotedString
                    answerCollection.Add textAnswer
                    GoTo NextIteration
                End If
                
                GoTo NextIteration
            End If
        End If
        
        '        ' Check if is list or float answer
        If IsNumeric(inputString) Then
            If CStr(CInt(item)) = item Then
                Set listAnswer = New ModelAnswerList
                Set baseAnswer = listAnswer
                baseAnswer.isoTime = timeArray(column)
                listAnswer.value = Int(inputString)
                answerCollection.Add listAnswer
                GoTo NextIteration
            Else
                Set sliderAnswer = New ModelAnswerSlider
                Set baseAnswer = sliderAnswer
                baseAnswer.isoTime = timeArray(column)
                sliderAnswer.value = CSng(inputString)
                answerCollection.Add sliderAnswer
                GoTo NextIteration
            End If
        End If
        ' If get here, it is an error.
        Err.Raise CustomError.InvalidQuestionType
NextIteration:
        column = column + 1
    Next
     
    Set getAnswerCollection = answerCollection

End Function

Private Function isIntegerArray(ByRef inputString As String) As Boolean
     
    If Len(inputString) < 1 Then
        isIntegerArray = False
        Exit Function
    End If
    
    Dim stringArray() As String
    Dim i As Long
    
    stringArray = Split(inputString, ",")
    For i = 0 To UBound(stringArray)
        If Not isInteger(stringArray(i)) Then
            isIntegerArray = False
            Exit Function
        End If
    Next i

    isIntegerArray = True

End Function

Private Function convertToIntegerArray(ByVal inputString As String) As Variant
 
    Dim stringArray() As String
    Dim intArray() As Long
    Dim bound As Long
    Dim i As Long
    
    stringArray = Split(inputString, ",")
    bound = UBound(stringArray)
    ReDim intArray(bound)
    
    For i = 0 To bound
        intArray(i) = CInt(stringArray(i))
    Next i
    
    convertToIntegerArray = stringArray

End Function

Private Function isInteger(ByRef item As String) As Boolean

    Dim result As Boolean
    result = False

    If IsNumeric(item) Then
        If CStr(CInt(item)) = item Then
            result = True
        End If
    End If
    isInteger = result

End Function

' The purpose of this function is to combine the aray items from a checkbox question into a single item.
' The result may be a shorter array. The size should match the number of questions.
Private Function getCleanDataArray(ByRef dataArray As Variant) As Variant

    Dim newDataArray As Variant
    Dim newCounter As Long
    Dim oldCounter As Long
    Dim endIndex As Long
        
    newDataArray = Array()
    newCounter = 0
    For oldCounter = FileCol.metadata To UBound(dataArray)
        ReDim Preserve newDataArray(newCounter)
        newDataArray(newCounter) = dataArray(oldCounter)
        If Left$(newDataArray(newCounter), 1) = Chr$(34) Then
            endIndex = getEndIndex(oldCounter, dataArray, UBound(dataArray))
            newDataArray(newCounter) = joinSubArray(dataArray, oldCounter + 1, endIndex)
            oldCounter = endIndex
        End If
        newCounter = newCounter + 1
    Next
    
    getCleanDataArray = newDataArray

End Function

'
' This recursive function is called when '"*' is found.
' It returns the index of the array for the second last item '*"' before th6e next item '"*' or the end of the array.
' It is broken by 2 edge cases: {","} and {",number,} but is better than Excel parsing for this type of file.
' TODO: Since May 2019, {"} is replaced by {'} in the PIEL App data file. In late 2010 can simplify this logic.
'
Private Function getEndIndex(ByVal counter As Long, ByVal answerArray As Variant, ByVal oldArraySize As Long) As Long
    Dim i As Long
    i = counter
    
    ' If item is possible end of text answer (i.e. has {*"}), progressively look for next one until find new question.
    ' That is the real endIndex
    Do While i < oldArraySize
        If Right$(answerArray(i), 1) = Chr$(34) Then
            If Left$(answerArray(i + 1), 1) = Chr$(34) Or IsNumeric(answerArray(i + 1)) Or Len(answerArray(i + 1)) = 0 Then
                getEndIndex = i
                Exit Function
            End If
        End If
        i = i + 1
    Loop
    ' Only get here is it is the last item and it contains {"*}
    getEndIndex = i
    
End Function

Private Function joinSubArray(ByVal mainArray As Variant, ByVal startIndex As Long, ByVal endIndex As Long, Optional ByRef separator As Variant = ",") As String
    
    Dim i As Long
    Dim arrayDim As Long
    Dim tempArray() As String
    
    arrayDim = endIndex - startIndex + 1
    ReDim tempArray(arrayDim)
    
    i = 0
    For i = 0 To arrayDim
        tempArray(i) = mainArray(startIndex + i - 1)
    Next i
    joinSubArray = Join(tempArray, separator)
    
End Function

