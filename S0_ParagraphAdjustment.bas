Attribute VB_Name = "S0_ParagraphAdjustment"
Sub RemoveAutoNumbers()
    '去除所有的自动编号
    ActiveDocument.Content.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
End Sub

Sub ConvertAutoNumbersToText()
    If Selection.Type = wdSelectionIP Then
        ActiveDocument.Content.ListFormat.ConvertNumbersToText
        ActiveDocument.Content.Find.Execute FindText:="^t", replacewith:=" ", Replace:=wdReplaceAll
    Else
        Selection.Range.ListFormat.ConvertNumbersToText
        Selection.Find.Execute FindText:="^t", replacewith:=" ", Replace:=wdReplaceAll
    End If
End Sub
'采用小四，1.5倍行距，不加段号的格式。
Sub LineSpacingAndFontAdjustment4EP()
    '
'     If Rng Is Nothing Or IsMissing(Rng) Then Set Rng = Selection
    With Selection
        .WholeStory
        With .ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpace1pt5
            .FirstLineIndent = CentimetersToPoints(0)
            .OutlineLevel = wdOutlineLevelBodyText
            .CharacterUnitLeftIndent = 0
            .CharacterUnitRightIndent = 0
            .CharacterUnitFirstLineIndent = 0
            .LineUnitBefore = 0
            .LineUnitAfter = 0
            .MirrorIndents = False
            .AutoAdjustRightIndent = False
            .DisableLineHeightGrid = True
            .WordWrap = True
            '        .Alignment = wdAlignParagraphJustify
        End With
        .Font.Name = "Times New Roman"
        .Font.Size = 12
    End With
End Sub

Sub AlignParagraphCenter(Optional rng As Range)
     If rng Is Nothing Or IsMissing(rng) Then Set rng = Selection
    '
    '单倍行距，整体居中'
    With rng
        With .ParagraphFormat
            .LeftIndent = CentimetersToPoints(0)
            .RightIndent = CentimetersToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphCenter
            .WidowControl = True
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
    End With
End Sub

