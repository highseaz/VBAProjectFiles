Attribute VB_Name = "S6_RangesbyRef"
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

Sub reDimRanges(ByRef theRanges() As Range, ByVal startPos As Long, ByVal endPos As Long)
    If SafeArrayGetDim(theRanges) = 0 Then
        index = 0
    Else
        index = UBound(theRanges) + 1
    End If

    ReDim Preserve theRanges(0 To index)
    Set theRanges(index) = ActiveDocument.Range(startPos, endPos)
    '    Debug.Print "Index: " & Index
    '    Debug.Print "Range Text: " & theRanges(Index).Text
End Sub

Sub reDimArrayAdd(ByRef theArr() As Variant, ByVal itemAdded As Variant)
    If SafeArrayGetDim(theArr) = 0 Then
        index = 0
    Else
        index = UBound(theArr) + 1
    End If
    ReDim Preserve theArr(0 To index)
    Set theArr(index) = itemAdded
End Sub

Sub SplitRangesWithDelimiter(ByRef oriRange As Range, ByVal delimiter As String, ByRef subRange() As Range) 'As Range
    '    RangeIndex = 0
    fContinue = True
    FinalEndPosition = oriRange.End

    With oriRange
        '        .StartOf WdUnits.wdStory
        Do While fContinue
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
                nEnd = FinalEndPosition
                fContinue = False
            End If

            reDimRanges subRange, nStart, nEnd
            .Collapse WdCollapseDirection.wdCollapseEnd
        Loop
    End With
End Sub
