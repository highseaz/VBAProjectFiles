Attribute VB_Name = "B2_FieldsFunctions"
'�����ļ���Ĭ��Ϊ word�ļ���-docvar.txt
'�����ļ���ʽ key=value��#Ϊע�� ���� https://my.oschina.net/noahxiao/blog/79465

'���DovVariable Field��ת��Ϊ��ͨ����
Sub unlinkDocVarFields()

    Dim varResponse As Variant

    varResponse = MsgBox("�Ƿ���ĵ��е�DocumentVariable���滻Ϊ��ͨ���֣�", vbYesNo)

    If varResponse = vbYes Then

        Dim bTrack As Boolean
        bTrack = ActiveDocument.TrackRevisions
        ActiveDocument.TrackRevisions = False


        '����DocVariable��
        Dim fCount As Integer
        fCount = 0
        For Each ofld In ActiveDocument.Fields
            If ofld.Type = wdFieldDocVariable Then
                '����������
                ofld.Unlink
                '��Ч�ᱻ�滻 Error! No document variable supplied.
                fCount = fCount + 1
            End If
        Next ofld

        ActiveDocument.TrackRevisions = bTrack
        MsgBox "��ɶ�" & fCount & "��DocVar���滻��"

    End If

End Sub

Sub unlinkSelectedFields()
    If Selection.Fields.Count <> 0 Then
        Dim varResponse As Variant
        varResponse = MsgBox("�Ƿ�Ѵ����滻Ϊ��ͨ���֣�", vbYesNo)
        If varResponse = vbYes Then
            Dim bTrack As Boolean
            bTrack = ActiveDocument.TrackRevisions
            ActiveDocument.TrackRevisions = False

            Dim fname As String
            Dim selectFiled As Field
            Set selectFiled = Selection.Fields(1)

            If selectFiled.Type = wdFieldDocVariable Then
                If selectFiled.code = "" Then
                    MsgBox "��Ϊ��"
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
                MsgBox "����DocVariable����"
            End If

            ActiveDocument.TrackRevisions = bTrack
            MsgBox "��ɶ�" & fCount & "��DocVar����滻��"

        End If
    Else
        MsgBox "��ѡ����Ҫ���µ���"
    End If
End Sub

'��ȡtxt�ļ��е�DovVariable����
Sub loadDocVarsFile()

    Dim varResponse As Variant

    varResponse = MsgBox("�Ƿ��ȡ����DocVar�ļ��е����ã�����������DocVar��", vbYesNo)

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
            MsgBox "û���ҵ�" & sFileName
            Exit Sub
        End If

        '��ȡ�ļ�
        iFileNum = FreeFile()
        Dim vCount As Integer
        vCount = 0
        Open sFileName For Input As iFileNum

        Do While Not EOF(iFileNum)
            Line Input #iFileNum, sBuf

            If InStr(1, Trim(sBuf), "#") <> 1 Then '#��ͷ��������Ϊ��ע��

                iPos = InStr(1, sBuf, "=") '��ֵȺ�
                If iPos <> 0 Then
                    sName = Trim(Left(sBuf, iPos - 1)) 'key
                    sValue = Trim(Mid(sBuf, iPos + 1, Len(sBuf) - iPos)) 'value

                    If Len(sName) <> 0 Then
                        ActiveDocument.Variables(sName).Value = sValue '�����ĵ���Variables
                        vCount = vCount + 1
                    End If
                End If

            End If

        Loop

        Close iFileNum

        '����ȫ��wdFieldDocVariable��
        Dim fCount As Integer
        fCount = updateAllDocVarField()

        ActiveDocument.TrackRevisions = bTrack
        MsgBox "��ɶ�ȡ����" & vCount & "��DocVar������Ϣ��������" & fCount & "����"

    End If

End Sub

'�ѹ��λ�����������޸ĵ�ֵ���µ�����ͬ����
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
            NewValue = Trim(InputBox("�������޸ĵ�����", "����������"))
            ActiveDocument.Variables(fname).Value = NewValue
        Else
            Dim varResponse As Variant
            varResponse = MsgBox("�Ƿ�Ѵ�������ݸ��µ�����ͬ����", vbYesNo)
            If varResponse = vbYes Then
                ActiveDocument.Variables(fname).Value = fvalue
            End If
        End If
        '����ȫ��wdFieldDocVariable��
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
        MsgBox "�������[" & fname & "=" & fvalue & "]" & fCount & "����ֵ�ĸ��£�"
    Else
        MsgBox "��ѡ����Ҫ���µ���"
    End If
End Sub

'��word�е�DocVarField����д��txt�ı�
Sub saveDocVarsFile()

    Dim bTrack As Boolean
    bTrack = ActiveDocument.TrackRevisions
    ActiveDocument.TrackRevisions = False

    Dim sFileName As String
    Dim sFileNameBackup As String
    Dim iFileNum As Integer
    Dim sCode As String
    Dim sPos As Integer

    sFileName = ActiveDocument.FullName & "-docvar.txt" '���ļ���
    sFileNameBackup = ActiveDocument.FullName & "-docvar-" & Format(Now(), "yyyyMMddhhmmss") & ".txt" '�����ļ���

    '����ԭ��docvar�ļ�
    If Len(Dir$(sFileName)) <> 0 Then
        Name sFileName As sFileNameBackup
    End If

    '���޸�ֵ���»�DocumentVariables
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
            '����code����ȡDocVar������

            If Len(ofld) = 0 Then 'ɾ����Чfield
                ofld.Delete
            Else
                docName = getFieldName(ofld)
                docValue = getFieldValue(ofld)

                '�ж����ж����DocVar�Ƿ����Variables��
                On Error Resume Next
                docOldValue = ActiveDocument.Variables(docName).Value
                If Err.Number = 0 Then '����
                    If docValue <> docOldValue Then '�ĵ�����ֵ��Variables�е�ֵ����ͬʱ��˵���ĵ������޸�

                        changeList.Add ("# ��" & ofld.code.Information(wdActiveEndPageNumber) & "ҳ ��" & ofld.code.Information(wdFirstCharacterLineNumber) & "�� # " & docName & "=" & docValue)
                        changeListCount = changeListCount + 1
                    End If
                    Else '�����ڣ�ֱ��д��
                    ActiveDocument.Variables(docName) = docValue
                End If
                On Error GoTo 0
            End If
        End If

    Next ofld

    'д�ļ�
    iFileNum = FreeFile()

    Dim vCount As Integer
    vCount = 0
    Open sFileName For Output As iFileNum

    Print #iFileNum, "# ����ʱ�䣺"; Format(Now(), "yyyy��MM��dd�� hh:mm:ss")
    Print #iFileNum, ""
    For Each oVar In ActiveDocument.Variables

        Dim outline As String
        outline = oVar.name & "=" & oVar.Value
        Print #iFileNum, outline
        vCount = vCount + 1
    Next oVar

    Print #iFileNum, ""
    Print #iFileNum, "# �ĵ��е���ֵ�����¼(ֵ��ͻ)"
    Print #iFileNum, ""

    For Each iChange In changeList
        Print #iFileNum, iChange
    Next

    Close iFileNum

    ActiveDocument.TrackRevisions = bTrack
    MsgBox "��ɶ�DocVar������Ϣ��д�룬��д��" & vCount & "��DocVar��" & changeListCount & "��ֵ��ͻ��"
    Shell "Notepad.exe " & sFileName, vbNormalFocus

End Sub

'����ȫ��wdFieldDocVariable���ޱ仯������
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

'��ȡDovVariable Field��name
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

'��ȡDovVariable Field��Result����ʾ�����
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
        varResponse = MsgBox("�Ƿ�Ѵ˲������ݱ��Ϊ��" & Chr(13) & "__" & str & "__" & Chr(13) & Chr(13), vbYesNo)
        If varResponse = vbYes Then changeStrToField (str)
    ElseIf str = "" Then

        str = Trim(InputBox("input the str to be transformed", "title"))
        If str <> "" Then changeStrToField (str)
    ElseIf Selection.Fields.Count <> 0 Then
        '     todo   xxxx


    End If

End Sub
