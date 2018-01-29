Attribute VB_Name = "S1_Deletions_Replacement"
Sub DelBlankPara(Optional Doc As Document)
    If Doc Is Nothing Or IsMissing(Doc) Then Set Doc = ActiveDocument

    'É¾³ý¿Õ°×¶ÎÂä

    Application.ScreenUpdating = False
    For Each i In Doc.Paragraphs
        If Len(Trim(i.Range)) = 1 Then
            i.Range.Delete
            n = n + 1
        End If
    Next
    Debug.Print "¹²É¾³ý¿Õ°×¶ÎÂä" & n & "¸ö"
    Application.ScreenUpdating = True
End Sub


Sub delContentinMidbracket()
    ' É¾³ýÖÐÀ¨ºÅÖÐµÄÄÚÈÝ
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "\[*\]"
        .Replacement.text = ""
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceOne
    End With

End Sub

Sub delSpace(Optional Doc As Document)

    trackflag = ActiveDocument.TrackRevisions

    ActiveDocument.TrackRevisions = False

    If Doc Is Nothing Or IsMissing(Doc) Then Set Doc = ActiveDocument
    With Doc.Content
        With .Find
            .text = "([!a-zA-Z0-9_,.;:\! ])^32{1,}([!a-zA-Z])"
            .Replacement.text = "\1\2"
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
        With .Find
            .text = "^32{2,}"
            .Replacement.text = "^32"
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
        With .Find
            .text = " ([,.;])"
            .Replacement.text = "\1"
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
    End With
    ActiveDocument.TrackRevisions = trackflag
End Sub

Sub DeleteUselessEnterinSelection()
Dim str As String

 str = "([£»£º£¬])^13([!Í¼])"
 Call DeletePatternsInSelection(str, 1, 1)
 
End Sub

Sub DeletePatternsInSelection(strPattern As String, posToStart As Integer, LenthofDeletion As Integer)
   
    Dim myRange As Range
    Dim NumofFound As Long
    Dim numofEnd As Long

    numofEnd = Selection.End
    NumofFound = 0
    Set myRange = Selection.Range
    With myRange.Find
        .MatchWildcards = True
        .Execute FindText:=strPattern, Forward:=True

        While .Found And NumofFound < numofEnd
            NumofFound = myRange.Start
            Debug.Print NumofFound
            myRange.Find.Execute FindText:=strPattern, Forward:=True
            istart = NumofFound + posToStart
            ActiveDocument.Range(Start:=istart, End:=istart + Lenthofdel) = ""
        Wend
    End With

End Sub
Sub ReplacementWithRef()

    Dim sFileName As String
    '
    sFileName = ActiveDocument.Path & "\replacement.txt"  'ÀÏÎÄ¼þÃû

    '    Debug.Print sFileName
    Dim result1()
    result1() = ReadFromFileByType(typeReplacement, sFileName)


    For i = 0 To UBound(result1)
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = result1(i)(0)
            .Replacement.text = result1(i)(1)
            .Forward = True
            .Wrap = wdFindAsk
            .format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = result1(i)(2)
            .Execute Replace:=wdReplaceAll
        End With
        Debug.Print result1(i)(0), result1(i)(1), result1(i)(2)
    Next i

End Sub



