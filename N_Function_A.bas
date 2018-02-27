Attribute VB_Name = "N_Function_A"
Function ArrangeNumArray_merge(ByRef numArr(), Optional ByVal strJoiner As String = "-") As String
    Dim newNumArr()
    startNum = numArr(0)
    preNum = startNum
    If UBound(numArr()) = 0 Then
        ArrangeNumArray_merge = startNum
        Exit Function
    End If


    For i = 1 To UBound(numArr)

        currentNum = numArr(i)

        If currentNum - 1 > preNum Then
            If startNum = preNum Then
                reDimArrayAdd newNumArr, startNum
            Else
                reDimArrayAdd newNumArr, startNum & strJoiner & preNum
            End If
            startNum = currentNum
        End If

        If currentNum = numArr(UBound(numArr)) Then
            If currentNum - 1 > preNum Then
                reDimArrayAdd newNumArr, currentNum
            Else
                reDimArrayAdd newNumArr, startNum & strJoiner & currentNum
            End If
        End If

        preNum = currentNum
    Next
    ArrangeNumArray_merge = Join(newNumArr, ", ")
End Function

Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    'DEVELOPER: Ryan Wells (wellsr.com)
    'DESCRIPTION: Function to check if a value is in an array of values
    'INPUT: Pass the function a value to search for and an array of values of any data type.
    'OUTPUT: True if is in array, false otherwise
    Dim element As Variant
    On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
    Exit Function
IsInArrayError:
    On Error GoTo 0
    IsInArray = False
End Function
'Function IsStringInArray(stringToBeFound As String, arr As Variant) As Boolean
'  IsStringInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
'End Function

