Attribute VB_Name = "C0_Callbacks"
Dim x As New Class_EventModule

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
            returnedVal = "替换" & Chr(13)
        Case "BtnReplaceRefFileEdit"
            returnedVal = "编辑配置文件" & Chr(13)
        Case "BtnAddClickMenu"
            returnedVal = "右键菜单" & Chr(13)

        Case "SpBnDelete__btn"
            returnedVal = "删除" & Chr(13)
        Case "BtnDeletebracket"
            returnedVal = "删除[]" & Chr(13)
        Case "BtnDeleteSpace"
            returnedVal = "删除空格" & Chr(13)
        Case "BtnDeleteEnter"
            returnedVal = "删除回车" & Chr(13)

        Case "SpBnPCT__btn"
            returnedVal = "制作PCT文档" & Chr(13)
        Case "BtnPCTformat"
            returnedVal = "应用PCT格式" & Chr(13)
        Case "BtnAdjTables"
            returnedVal = "调整表格格式" & Chr(13)
        Case "BtnAdjGraphs"
            returnedVal = "调整图片格式" & Chr(13)
        Case "BtnaddCrossRefParagraph"
            returnedVal = "加入交叉引用段落" & Chr(13)
        Case "BtnSplit"
            returnedVal = "拆分PCT文件" & Chr(13)


        Case "BtnAcceptFormat"
            returnedVal = "接受格式修改" & Chr(13)

        Case "BtnEPformat"
            returnedVal = "应用EP格式" & Chr(13)
        Case "BtnAddRefNum"
            returnedVal = "加入附图标记" & Chr(13)
        Case "BtnRmAutoNum"
            returnedVal = "去除自动编号" & Chr(13)


        Case "BtnSaveDocVarsFile"
            returnedVal = "保存域文件" & Chr(13)
        Case "BtnloadDocVarsFile"
            returnedVal = "加载域文件" & Chr(13)
        Case "BtnUnlinkDocVarFields"
            returnedVal = "解除域" & Chr(13)
        Case "BtnUpdateSelectDocVar"
            returnedVal = "更新域变量" & Chr(13)
        Case "BtnChangeStrToFieldWithDiag2"
            returnedVal = "转化为域" & Chr(13)

    End Select
End Sub



Sub Button_Click(control As IRibbonControl)
    Select Case control.id
        Case "btnFileSaveAsPdfOrXps"
            Call ExportAsPDFFile

        Case "SpBnReplace__btn"
            Call ReplacementWithRef
    
        Case "BtnReplacewithWildchar"
            Call ReplacementWithRef
        Case "BtnReplacewithoutWildchar"
           

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
            Call addCrossRefParagraph



        Case "BtnSaveDocVarsFile"
            Call saveDocVarsFile
        Case "BtnloadDocVarsFile"
            Call loadDocVarsFile
        Case "BtnUnlinkDocVarFields"
            Call unlinkDocVarFields
        Case "BtnUpdateSelectDocVar"
            Call updateSelectDocVar
        Case "BtnChangeStrToFieldWithDiag2"
            Call changeStrToFieldWithDiag2

    End Select
End Sub

'Callback for Glytemplate onAction
Sub Gallery_Click(control As IRibbonControl, id As String, index As Integer)
End Sub


