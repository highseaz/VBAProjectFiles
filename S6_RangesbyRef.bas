Attribute VB_Name = "S6_RangesbyRef"
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

Sub reDimRanges(ByRef theRanges() As Range, ByVal StartPos As Long, ByVal endPos As Long, doc As Document)
    If SafeArrayGetDim(theRanges) = 0 Then
        index = 0
    Else
        index = UBound(theRanges) + 1
    End If

    ReDim Preserve theRanges(0 To index)
    Set theRanges(index) = doc.Range(StartPos, endPos)
    '    Debug.Print "Index: " & Index
    '    Debug.Print "Range Text: " & theRanges(Index).Text
End Sub

Sub reDimArrayAdd(ByRef theArr(), ByVal itemAdded As Variant)
    If SafeArrayGetDim(theArr) = 0 Then
        index = 0
    Else
        index = UBound(theArr) + 1
    End If
    ReDim Preserve theArr(0 To index)
    If IsObject(itemAdded) Then
        Set theArr(index) = itemAdded
    Else
        theArr(index) = itemAdded
    End If
End Sub

Sub SplitRangesWithDelimiter(ByVal oriRange As Range, ByVal delimiter As String, ByRef subRange())  'As Range
    '    RangeIndex = 0
    fContinue = True
    nEnd = 0
    finalendposition = oriRange.End

    With oriRange
        '        .StartOf WdUnits.wdStory
        Do While fContinue And nEnd <= finalendposition
            nStart = .Start
            .Find.ClearFormatting
            With .Find
                .Text = delimiter
                .Forward = True
                .Wrap = WdFindWrap.wdFindStop
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchByte = False
                .MatchAllWordForms = False
                .MatchSoundsLike = False
                .MatchWildcards = True
            End With
            If .Find.Execute Then
                nEnd = .End
            Else
                nEnd = finalendposition
                fContinue = False
            End If
            .Collapse WdCollapseDirection.wdCollapseEnd
            If nEnd <= finalendposition Then
                '            reDimRanges subRange, nStart, nEnd
                Debug.Print nStart, nEnd, finalendposition, oriRange.Document.Range(nStart, nEnd)
                reDimArrayAdd subRange, oriRange.Document.Range(nStart, nEnd)
            End If

        Loop
    End With
End Sub
