Attribute VB_Name = "S4_transferFromCNtoPCT"
Option Base 1
'从一个文档的文件导入到另一个文档中指定位置
Sub PastefromOriDocToTargetDoc(ByVal sourceDoc As Document, _
                                ByVal Startpoint As Long, _
                                ByVal Endpoint As Long, _
                                ByVal targetRangePoint As Range, _
                                Optional ByVal iFormatingtype As Integer = 2)

    Set docOriRange = sourceDoc.Range(Start:=Startpoint, End:=Endpoint)
    PCTContentFormating rng:=docOriRange, Formatingtype:=iFormatingtype
    docOriRange.Copy

    Set targetRange = targetRangePoint
    targetRange.Collapse Direction:=wdCollapseEnd
    targetRange.PasteAndFormat (wdUseDestinationStylesRecovery)
    'wdUseDestinationStylesRecovery wdFormatSurroundingFormattingWithEmphasis
    Set docOriRange = Nothing
    Set targetRange = Nothing

End Sub


Sub transferFromCNtoPCT()
    Dim docNew As Document
    Dim docOri As Document

    Set docNew = Documents.Open(TEMPLATE_Full)
    oridir = SelectedFileWithDlog
    Set docOri = Documents.Open(oridir) '  Set docOri = SelectedFileWithDlog

    Dim aStart(10), aEnd(10) As Long
    Dim tgtpnt(10) As Range
    Dim atype(10) As Integer

    ''-----------------寻找位置需要更精炼，而且不适用在页眉的情况---------------
    aStart(1) = RangeIncludingStr(PCTSplitDelimiter(0), docOri).End + 1
    If Not RangeIncludingStr(PCTSplitDelimiter(1), docOri) Is Nothing Then
        aEnd(1) = RangeIncludingStr(PCTSplitDelimiter(1), docOri).Start - 1
    Else
        aEnd(1) = RangeIncludingStr(PCTSplitDelimiter(2), docOri).Start - 1
    End If

    Set tgtpnt(1) = docNew.Bookmarks(3).Range.Next.Next

    ''---------------------------------------------------------------------
    aStart(2) = RangeIncludingStr(PCTSplitDelimiter(2), docOri).End + 1
    aEnd(2) = RangeIncludingStr(PCTSplitDelimiter(3), docOri).Start - 1
    Set tgtpnt(2) = docNew.Bookmarks(2).Range.Next.Next
    ''----------------------------------------------------------------------------
    aStart(3) = RangeIncludingStr(PCTSplitDelimiter(3), docOri).End + 1
    aEnd(3) = RangeIncludingStr(PCTSplitDelimiter(4), docOri).Start - 1
    Set tgtpnt(3) = docNew.Bookmarks(1).Range.Next
    ''format1
    ''-----------------------------------------------

    For i = 4 To 8
        aStart(i) = RangeIncludingStr(PCTSplitDelimiter(i), docOri).End
        aEnd(i) = RangeIncludingStr(PCTSplitDelimiter(i + 1), docOri).Start - 1
        Set tgtpnt(i) = RangeIncludingStr(PCTSplitDelimiter(i), docNew)

    Next i

    ''-------------------------------------------------------------------
    aStart(9) = RangeIncludingStr(PCTSplitDelimiter(9), docOri).End
    aEnd(9) = docOri.Paragraphs.last.Range.End
    Set tgtpnt(9) = docNew.Bookmarks(4).Range.Next.Next
    'format3
    ''------------------------------------------------------------------
    For i = 1 To 9

        If i = 3 Then
            atype(i) = 1
        ElseIf i = 9 Then
            atype(i) = 3
Else:
            atype(i) = 2
        End If
        '        Debug.Print i
        '        Debug.Print aStart(i)
        '        Debug.Print aEnd(i)

        PastefromOriDocToTargetDoc sourceDoc:=docOri, _
                         Startpoint:=aStart(i), _
                         Endpoint:=aEnd(i), _
                         targetRangePoint:=tgtpnt(i), _
                         iFormatingtype:=atype(i)
    Next i

    ''-------------------------------------------------------------------
    strNewFileName = oridir

    If InStr(1, strNewFileName, CONSNameFindStr) > 0 Then
        strNewFileName = Replace(strNewFileName, CONSNameFindStr, CONSNameReplaceStr)
    Else
        strNewFileName = Replace(strNewFileName, ".do", "_" & CONSNameReplaceStr & ".do")
    End If

    ''-ExtractStringByPatternFrom----------------
    ''-------------------------------------------------------------------
    docNew.SaveAs FileName:=strNewFileName
    docOri.Close SaveChanges:=wdDoNotSaveChanges

    Set docNew = Nothing
    Set docOri = Nothing

End Sub

Sub addCrossRefParagraph()
    Dim rng As Range
    Dim PasteInsitu As Object
    Set PasteInsitu = JsonReadFromConfFile("PasteInsitu")
    With ActiveDocument
        bTrack = .TrackRevisions
        For i = 1 To PasteInsitu.Count
            Set rng = RangeIncludingStr(PasteInsitu(i), ActiveDocument, True)
            Start = rng.Start

            .TrackRevisions = False
            rng.Cut
            .TrackRevisions = True
            .Range(Start, Start).Paste
        Next
        Set rng = Nothing
        .TrackRevisions = bTrack
    End With

    Set PasteInsitu = Nothing


End Sub

Sub AdjustLineSpaceOfEquationsandGraph()
    Selection.Find.ClearFormatting
    With Selection
        With .Find
            .Text = "^g"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute
        End With
        Do While .Find.Found
            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle

            .Find.Execute
        Loop

    End With
End Sub

Sub AdjustTextOfTables()

    For Each oTable In ActiveDocument.Tables
        With oTable.Range.ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphCenter
            .WidowControl = False
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            .FirstLineIndent = CentimetersToPoints(0)
            .OutlineLevel = wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .MirrorIndents = False
            .TextboxTightWrap = wdTightNone
            .AutoAdjustRightIndent = False
            .DisableLineHeightGrid = True
            .FarEastLineBreakControl = True
            .WordWrap = True
            .HangingPunctuation = True
            .HalfWidthPunctuationOnTopOfLine = False
            .AddSpaceBetweenFarEastAndAlpha = True
            .AddSpaceBetweenFarEastAndDigit = True
            .BaseLineAlignment = wdBaselineAlignAuto


        End With
    Next
End Sub

Public Sub PCTContentFormating(ByVal rng As Range, Optional ByVal Formatingtype As Integer = 2)
    '1 is title 2 is content
    '3 is drawing

    With rng
        With .ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = 24
            If Formatingtype = 3 Then .LineSpacingRule = wdLineSpaceSingle
            If Formatingtype = 1 Or 3 Then .Alignment = wdAlignParagraphCenter
            If Formatingtype = 2 Then .Alignment = wdAlignParagraphJustify
            .WidowControl = False
            .KeepWithNext = False
            .KeepTogether = False
            .PageBreakBefore = False
            .NoLineNumber = False
            .Hyphenation = True
            If Formatingtype = 1 Or 3 Then .FirstLineIndent = CentimetersToPoints(0)
            If Formatingtype = 2 Then .FirstLineIndent = CentimetersToPoints(0.35)
            .OutlineLevel = wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            If Formatingtype = 1 Or 3 Then .CharacterUnitFirstLineIndent = 0
            If Formatingtype = 2 Then .CharacterUnitFirstLineIndent = 2
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .MirrorIndents = False
            .TextboxTightWrap = wdTightNone
            .AutoAdjustRightIndent = True
            .DisableLineHeightGrid = False
            .FarEastLineBreakControl = True
            .WordWrap = True
            .HangingPunctuation = True
            .HalfWidthPunctuationOnTopOfLine = False
            .AddSpaceBetweenFarEastAndAlpha = True
            .AddSpaceBetweenFarEastAndDigit = True
            .BaseLineAlignment = wdBaselineAlignAuto
        End With
        .Font.name = "宋体"
        .Font.name = "Times New Roman"
        If Formatingtype = 1 Then .Bold = True
        If Formatingtype = 3 Then .Bold = False

    End With
End Sub


Sub acceptFormatChanges()
    With ActiveWindow.View
        .ShowFormatChanges = True
        '        .ShowRevisionsAndComments = False
        .ShowInsertionsAndDeletions = False

        ActiveDocument.AcceptAllRevisionsShown

        .ShowRevisionsAndComments = True
        .ShowFormatChanges = True
        .ShowInsertionsAndDeletions = True
    End With
End Sub
