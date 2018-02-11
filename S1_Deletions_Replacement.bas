Attribute VB_Name = "S1_Deletions_Replacement"
Sub DelBlankPara(Optional Doc As Document)

    If Doc Is Nothing Or IsMissing(Doc) Then Set Doc = ActiveDocument
    Application.ScreenUpdating = False
    For Each i In Doc.Paragraphs
        If Len(Trim(i.Range)) = 1 Then
            i.Range.Delete
            n = n + 1
        End If
    Next
    Debug.Print "Delete " & n & " blank Paragraphs"
    Application.ScreenUpdating = True
End Sub


Sub delContentinMidbracket(Optional rng As Range)
    If rng Is Nothing Or IsMissing(rng) Then Set rng = Selection
    With rng
        .Find.ClearFormatting
        .Find.Replacement.ClearFormatting
        With .Find
            .Text = "\[*\]"
            .Replacement.Text = ""
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
            .Execute Replace:=wdReplaceOne
        End With
    End With
End Sub

Sub delSpace(Optional Doc As Document)

    If Doc Is Nothing Or IsMissing(Doc) Then Set Doc = ActiveDocument

    With Doc
        trackflag = .TrackRevisions

        .TrackRevisions = False

        With .Content
            With .Find
                .Text = "([!a-zA-Z0-9_,.;:\! ])^32{1,}([!a-zA-Z])"
                .Replacement.Text = "\1\2"
                .MatchWildcards = True
                .Execute Replace:=wdReplaceAll
            End With
            With .Find
                .Text = "^32{2,}"
                .Replacement.Text = "^32"
                .MatchWildcards = True
                .Execute Replace:=wdReplaceAll
            End With
            With .Find
                .Text = " ([,.;])"
                .Replacement.Text = "\1"
                .MatchWildcards = True
                .Execute Replace:=wdReplaceAll
            End With
        End With
        .TrackRevisions = trackflag
    End With
End Sub

Sub DeleteUselessEnterinSelection()
    Dim str As String

    str = "([£»£º£¬»ò])^13([!Í¼])"
    Call DeletePatternsInSelection(str, 1, 1)

End Sub

Sub DeletePatternsInSelection(strPattern As String, posOffsetToStart As Integer, LenthofDeletion As Integer)

    Dim MyRange As Range
    Dim NumofFound As Long
    Dim numofEnd As Long

    numofEnd = Selection.End
    NumofFound = 0
    Set MyRange = Selection.Range
    With MyRange.Find
        .MatchWildcards = True
        .Execute FindText:=strPattern, Forward:=True

        While .Found And NumofFound < numofEnd
            NumofFound = MyRange.Start
            Debug.Print NumofFound
            MyRange.Find.Execute FindText:=strPattern, Forward:=True
            istart = NumofFound + posOffsetToStart
            ActiveDocument.Range(Start:=istart, End:=istart + Lenthofdel).Delete
        Wend
    End With

End Sub
Sub ReplacementWithRef()

    ActiveWindow.View.MarkupMode = wdBalloonRevisions

    Dim sFileName As String
    '
    sFileName = MYWORKPATH_CODE & "\replacement.txt"

    '    Debug.Print sFileName
    Dim result1()
    result1() = ReadFromFileByType(typeReplacement, sFileName)


    For i = 0 To UBound(result1)
        With ActiveDocument.Content.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = Trim(result1(i)(0))

            .Replacement.Text = Trim(result1(i)(1))
            .Forward = True
            .Wrap = wdFindAsk
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = Trim(result1(i)(2))
            .Execute Replace:=wdReplaceAll
        End With
        Debug.Print result1(i)(0), result1(i)(1), result1(i)(2)
    Next i

End Sub


Function StrReplaceSpecialChars(ByVal str As String, sChr As String, Optional sLen As Integer = 20)
    Dim s As String
    s = Replace(str, "'", sChr)
    s = Replace(s, "*", sChr)
    s = Replace(s, "/", sChr)
    s = Replace(s, "\", sChr)
    s = Replace(s, ":", sChr)
    s = Replace(s, "?", sChr)
    s = Replace(s, Chr(34), sChr)
    s = Replace(s, "<", sChr)
    s = Replace(s, ">", sChr)
    s = Replace(s, "|", sChr)
    s = Replace(s, " ", sChr)
    Debug.Print s

    StrReplaceSpecialChars = Left(s, sLen)

End Function
