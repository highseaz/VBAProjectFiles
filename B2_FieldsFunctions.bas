Attribute VB_Name = "B2_FieldsFunctions"
'配置文件名默认为 word文件名-docvar.txt
'配置文件格式 key=value，#为注释 出处 https://my.oschina.net/noahxiao/blog/79465

'解除DovVariable Field，转换为普通文字
Sub unlinkDocVarFields()

    Dim varResponse As Variant

    varResponse = MsgBox("是否把文档中的DocumentVariable域替换为普通文字？", vbYesNo)

    If varResponse = vbYes Then

        Dim bTrack As Boolean
        bTrack = ActiveDocument.TrackRevisions
        ActiveDocument.TrackRevisions = False


        '遍历DocVariable域
        Dim fCount As Integer
        fCount = 0
        For Each ofld In ActiveDocument.Fields
            If ofld.Type = wdFieldDocVariable Then
                '撤消域连接
                ofld.Unlink
                '无效会被替换 Error! No document variable supplied.
                fCount = fCount + 1
            End If
        Next ofld

        ActiveDocument.TrackRevisions = bTrack
        MsgBox "完成对" & fCount & "个DocVar域替换！"

    End If

End Sub

Sub unlinkSelectedFields()
    If Selection.Fields.Count <> 0 Then
        Dim varResponse As Variant
        varResponse = MsgBox("是否把此域替换为普通文字？", vbYesNo)
        If varResponse = vbYes Then
            Dim bTrack As Boolean
            bTrack = ActiveDocument.TrackRevisions
            ActiveDocument.TrackRevisions = False

            Dim fname As String
            Dim selectFiled As Field
            Set selectFiled = Selection.Fields(1)

            If selectFiled.Type = wdFieldDocVariable Then
                If selectFiled.code = "" Then
                    MsgBox "域为空"
                    Exit Sub
                End If

                fname = getFieldName(selectFiled)

                Dim fCount As Integer
                fCount = 0
                For Each ofld In ActiveDocument.Fields
                    If ofld.Type = wdFieldDocVariable Then
                        If getFieldName(ofld) = fname Then
                            '                            oFld.result.HighlightColorIndex = wdNoHighlight
                            ofld.Unlink
                            fCount = fCount + 1
                        End If
                    End If
                Next ofld
                ActiveDocument.Variables(fname).Delete
            Else
                MsgBox "域不是DocVariable类型"
            End If

            ActiveDocument.TrackRevisions = bTrack
            MsgBox "完成对" & fCount & "个DocVar域的替换！"

        End If
    Else
        MsgBox "请选择需要更新的域！"
    End If
End Sub

'读取txt文件中的DovVariable配置
Sub loadDocVarsFile()

    Dim varResponse As Variant

    varResponse = MsgBox("是否读取载入DocVar文件中的配置，并更新所有DocVar域？", vbYesNo)

    If varResponse = vbYes Then

        Dim bTrack As Boolean
        bTrack = ActiveDocument.TrackRevisions
        ActiveDocument.TrackRevisions = False

        Dim sFileName As String
        Dim iFileNum As Integer
        Dim sBuf As String
        Dim iPos As Integer
        Dim sName As String
        Dim sValue As String

        sFileName = ActiveDocument.FullName & "-docvar.txt"

        If Len(Dir$(sFileName)) = 0 Then
            MsgBox "没有找到" & sFileName
            Exit Sub
        End If

        '读取文件
        iFileNum = FreeFile()
        Dim vCount As Integer
        vCount = 0
        Open sFileName For Input As iFileNum

        Do While Not EOF(iFileNum)
            Line Input #iFileNum, sBuf

            If InStr(1, Trim(sBuf), "#") <> 1 Then '#开头的配置认为是注释

                iPos = InStr(1, sBuf, "=") '拆分等号
                If iPos <> 0 Then
                    sName = Trim(Left(sBuf, iPos - 1)) 'key
                    sValue = Trim(Mid(sBuf, iPos + 1, Len(sBuf) - iPos)) 'value

                    If Len(sName) <> 0 Then
                        ActiveDocument.Variables(sName).Value = sValue '更新文档的Variables
                        vCount = vCount + 1
                    End If
                End If

            End If

        Loop

        Close iFileNum

        '更新全部wdFieldDocVariable域
        Dim fCount As Integer
        fCount = updateAllDocVarField()

        ActiveDocument.TrackRevisions = bTrack
        MsgBox "完成读取载入" & vCount & "个DocVar配置信息，并更新" & fCount & "个域！"

    End If

End Sub

'把光标位置所做的域修改的值更新到其它同名域
Sub updateSelectDocVar()
    If Selection.Fields.Count <> 0 Then
        Dim fname As String
        Dim fvalue As String
        Dim oldValue As String

        Dim selectField As Field
        Set selectField = Selection.Fields(1)
        fname = getFieldName(selectField)
        fvalue = getFieldValue(selectField)
        oldValue = ActiveDocument.Variables(fname).Value

        If fvalue = oldValue Then
            Dim NewValue As String
            NewValue = Trim(InputBox("输入拟修改的内容", "输入新内容"))
            ActiveDocument.Variables(fname).Value = NewValue
        Else
            Dim varResponse As Variant
            varResponse = MsgBox("是否把此域的内容更新到其它同名域？", vbYesNo)
            If varResponse = vbYes Then
                ActiveDocument.Variables(fname).Value = fvalue
            End If
        End If
        '更新全部wdFieldDocVariable域
        Dim fCount As Integer
        Dim ofld As Field
        '                fCount = updateAllDocVarField()
        For Each ofld In ActiveDocument.Fields
            If ofld.Type = wdFieldDocVariable Then
                If getFieldName(ofld) = fname Then
                    ofld.Update
                    fCount = fCount + 1
                    ofld.result.HighlightColorIndex = wdYellow
                End If
            End If
        Next
        MsgBox "完成其它[" & fname & "=" & fvalue & "]" & fCount & "个域值的更新！"
    Else
        MsgBox "请选择需要更新的域！"
    End If
End Sub

'把word中的DocVarField内容写入txt文本
Sub saveDocVarsFile()

    Dim bTrack As Boolean
    bTrack = ActiveDocument.TrackRevisions
    ActiveDocument.TrackRevisions = False

    Dim sFileName As String
    Dim sFileNameBackup As String
    Dim iFileNum As Integer
    Dim sCode As String
    Dim sPos As Integer

    sFileName = ActiveDocument.FullName & "-docvar.txt" '老文件名
    sFileNameBackup = ActiveDocument.FullName & "-docvar-" & Format(Now(), "yyyyMMddhhmmss") & ".txt" '备份文件名

    '备份原有docvar文件
    If Len(Dir$(sFileName)) <> 0 Then
        Name sFileName As sFileNameBackup
    End If

    '域修改值更新回DocumentVariables
    Dim docKey As String
    Dim docName As String
    Dim docValue As String
    Dim docOldValue As String
    Dim changeList As Collection
    Set changeList = New Collection
    Dim changeListCount As Integer

    docKey = "DOCVARIABLE"
    changeListCount = 0

    For Each ofld In ActiveDocument.Fields
        If ofld.Type = wdFieldDocVariable Then
            '从域code中提取DocVar的名字

            If Len(ofld) = 0 Then '删除无效field
                ofld.Delete
            Else
                docName = getFieldName(ofld)
                docValue = getFieldValue(ofld)

                '判断域中定义的DocVar是否存在Variables中
                On Error Resume Next
                docOldValue = ActiveDocument.Variables(docName).Value
                If Err.Number = 0 Then '存在
                    If docValue <> docOldValue Then '文档中域值与Variables中的值不相同时，说明文档中有修改

                        changeList.Add ("# 第" & ofld.code.Information(wdActiveEndPageNumber) & "页 第" & ofld.code.Information(wdFirstCharacterLineNumber) & "行 # " & docName & "=" & docValue)
                        changeListCount = changeListCount + 1
                    End If
                    Else '不存在，直接写入
                    ActiveDocument.Variables(docName) = docValue
                End If
                On Error GoTo 0
            End If
        End If

    Next ofld

    '写文件
    iFileNum = FreeFile()

    Dim vCount As Integer
    vCount = 0
    Open sFileName For Output As iFileNum

    Print #iFileNum, "# 保存时间："; Format(Now(), "yyyy年MM月dd日 hh:mm:ss")
    Print #iFileNum, ""
    For Each oVar In ActiveDocument.Variables

        Dim outline As String
        outline = oVar.name & "=" & oVar.Value
        Print #iFileNum, outline
        vCount = vCount + 1
    Next oVar

    Print #iFileNum, ""
    Print #iFileNum, "# 文档中的域值变更记录(值冲突)"
    Print #iFileNum, ""

    For Each iChange In changeList
        Print #iFileNum, iChange
    Next

    Close iFileNum

    ActiveDocument.TrackRevisions = bTrack
    MsgBox "完成对DocVar配置信息的写入，供写入" & vCount & "个DocVar，" & changeListCount & "个值冲突域！"
    Shell "Notepad.exe " & sFileName, vbNormalFocus

End Sub

'更新全部wdFieldDocVariable域，无变化不更新
Private Function updateAllDocVarField() As Integer

    Dim fCount As Integer
    fCount = 0
    For Each ofld In ActiveDocument.Fields
        If ofld.Type = wdFieldDocVariable Then
            If ActiveDocument.Variables(getFieldName(ofld)).Value <> getFieldValue(ofld) Then
                ofld.Update
                fCount = fCount + 1
                ofld.result.HighlightColorIndex = wdYellow
            End If
        End If
    Next ofld
    updateAllDocVarField = fCount
End Function

'获取DovVariable Field的name
Function getFieldName(ofld As Field) As String
    '    Debug.Print oFld.code
    If ofld.Type = wdFieldDocVariable Then
        Dim docKey As String
        docKey = "DOCVARIABLE"
        getFieldName = Trim(Mid(ofld.code, (InStr(1, ofld.code, docKey) + Len(docKey) + 1), InStr(1, ofld.code, "\*") - InStr(1, ofld.code, docKey) - Len(docKey) - 1))
    Else
        getFieldName = vbNullString
    End If
End Function

'获取DovVariable Field的Result（显示结果）
Function getFieldValue(ofld As Field) As String
    getFieldValue = Trim(ofld.result)
End Function

Sub changeStrToField(myfieldValue As String)
    Dim bTrack As Boolean
    bTrack = ActiveDocument.TrackRevisions
    ActiveDocument.TrackRevisions = False

    myfieldname = "FiledVar-" & ActiveDocument.Variables.Count + 1
    myfieldname = myfieldname & "-" & StrReplaceSpecialChars(myfieldValue, "-")
    Debug.Print myfieldname, myfieldValue

    'Application.ActiveWindow.View.ShowFieldCodes = True
    '
    With ActiveDocument
        .Variables.Add name:=myfieldname ', Value:=myfieldValue
        iPostion = WordfindPositionInText(myfieldValue)
        If iPostion = 0 Then
            Debug.Print "no such string found."
            Exit Sub
        End If

        Do While iPostion > 0 And iPostion < ActiveDocument.Range.End

            .Range(iPostion, iPostion + Len(myfieldValue)).Text = ""

            .Fields.Add Range:=.Range(iPostion, iPostion), Type:=wdFieldDocVariable, Text:=myfieldname, PreserveFormatting:=True
            iPostion = WordfindPositionInText(myfieldValue)

        Loop

        .Variables(myfieldname).Value = myfieldValue & ""
        Dim ofld As Field
        For Each ofld In .Fields
            If ofld.Type = wdFieldDocVariable Then
                If getFieldName(ofld) = myfieldname Then ofld.Update
            End If
        Next ofld
    End With
    '    ActiveDocument.Fields.Update
    Application.ActiveWindow.View.ShowFieldCodes = False
    ''save the vars
    ActiveDocument.TrackRevisions = bTrack

End Sub
Function DocVarExists(ByVal varName As String) As Boolean
    DocVarExists = False
    Dim dummy As String
    If Len(varName) > 0 Then ' it exists
        On Error Resume Next
        dummy = ActiveDocument.Variables(varName)
        If Err.Number = 0 Then DocVarExists = True

        Else: ' it doesn't exist
        Debug.Print "Variable " & varName & " doesn't exist"
    End If

End Function
Sub changeStrToFieldWithDiag2()
    Dim str As String
    str = Trim(Selection.Range.Text)
    If Selection.Fields.Count = 0 And Len(str) > 0 Then

        Dim varResponse As Variant
        varResponse = MsgBox("是否把此部分内容变更为域？" & Chr(13) & "__" & str & "__" & Chr(13) & Chr(13), vbYesNo)
        If varResponse = vbYes Then changeStrToField (str)
    ElseIf str = "" Then

        str = Trim(InputBox("input the str to be transformed", "title"))
        If str <> "" Then changeStrToField (str)
    ElseIf Selection.Fields.Count <> 0 Then
        '     todo   xxxx


    End If

End Sub
