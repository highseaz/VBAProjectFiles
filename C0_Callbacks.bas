Attribute VB_Name = "C0_Callbacks"
Dim x As New EventClassModule

Sub Register_Event_Handler()
 Set x.App = word.Application
End Sub


'Callback for customUI.onLoad
Sub rxiRibbonUI_onLoad(ribbon As IRibbonUI)
'VBA.SendKeys "+{F5}"
Call RightClickNewMenuAmd
Call Register_Event_Handler
End Sub


Sub but_getLabel(control As IRibbonControl, ByRef returnedVal)

    Select Case control.id

        Case "SpBnReplace__btn"
            returnedVal = "�滻" & Chr(13)
        Case "BtnReplaceRefFileEdit"
            returnedVal = "�༭�����ļ�" & Chr(13)
        Case "BtnAddClickMenu"
            returnedVal = "�Ҽ��˵�" & Chr(13)

        Case "SpBnDelete__btn"
            returnedVal = "ɾ��" & Chr(13)
        Case "BtnDeletebracket"
            returnedVal = "ɾ��[]" & Chr(13)
        Case "BtnDeleteSpace"
            returnedVal = "ɾ���ո�" & Chr(13)
        Case "BtnDeleteEnter"
            returnedVal = "ɾ���س�" & Chr(13)

        Case "SpBnPCT__btn"
            returnedVal = "����PCT�ĵ�" & Chr(13)
        Case "BtnPCTformat"
            returnedVal = "Ӧ��PCT��ʽ" & Chr(13)
        Case "BtnAdjTables"
            returnedVal = "��������ʽ" & Chr(13)
        Case "BtnAdjGraphs"
            returnedVal = "����ͼƬ��ʽ" & Chr(13)
        Case "BtnaddCrossRefParagraph"
            returnedVal = "���뽻�����ö���" & Chr(13)
 Case "BtnSplit"
            returnedVal = "���PCT�ļ�" & Chr(13)


        Case "BtnAcceptFormat"
            returnedVal = "���ܸ�ʽ�޸�" & Chr(13)

        Case "BtnEPformat"
            returnedVal = "Ӧ��EP��ʽ" & Chr(13)
        Case "BtnAddRefNum"
            returnedVal = "���븽ͼ���" & Chr(13)
        Case "BtnRmAutoNum"
            returnedVal = "ȥ���Զ����" & Chr(13)


    End Select
End Sub



Sub Button_Click(control As IRibbonControl)
    Select Case control.id
        Case "btnFileSaveAsPdfOrXps"
            Call ExportAsPDFFile

        Case "SpBnReplace__btn"
            Call ReplacementWithRef
'            Call ReplacementWithoutMatchWildcards
        Case "BtnReplacewithWildchar"
            Call ReplacementWithRef
        Case "BtnReplacewithoutWildchar"
'            Call ReplacementWithoutMatchWildcards

        Case "SpBnDelete__btn"
            Call delSpace
            '            Call delContentinMidbracket
            Call DelBlankPara

        Case "BtnDeleteSpace"
            Call delSpace
        Case "BtnDeletebracket"
            Call delContentinMidbracket
        Case "BtnDeleteEnter"
            Call DelBlankPara

        Case "BtnAcceptFormat"
            Call acceptFormatChanges

        Case "BtnEPformat"
            Call LineSpacingAndFontAdjustment4EP
            Call RemoveAutoNumbers
        Case "BtnAddRefNum"
            Call addReferenceNumber4Claimswithform
        Case "BtnRmAutoNum"
            Call ConvertAutoNumbersToText

        Case " SpBnPCT__btn"
            Call transferFromCNtoPCT
            Call AdjustTextOfTables
            Call AdjustLineSpaceOfEquationsandGraph
'            Call addCrossRefParagraph
        Case "BtnPCTformat"
            Call transferFromCNtoPCT
        Case "BtnAdjTables"
            Call AdjustTextOfTables
        Case "BtnAdjGraphs"
            Call AdjustLineSpaceOfEquationsandGraph
            Case "BtnSplit"
            Call splitall
        Case "BtnaddCrossRefParagraph"

    End Select
End Sub

'Callback for Glytemplate onAction
Sub Gallery_Click(control As IRibbonControl, id As String, index As Integer)
End Sub


