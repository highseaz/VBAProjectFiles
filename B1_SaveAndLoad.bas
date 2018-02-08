Attribute VB_Name = "B1_SaveAndLoad"


Function ReadFromFileByType(iEOTypes As EnumOfTypes, _
                            Optional sFileName As String, _
                            Optional strVarName As String) As Variant

    Dim iFileNum As Integer
    Dim sBuf As String
    Dim typeIndicator As String
    Dim strContent As String
    Dim val
    Dim result()
    i = 0
    iFileNum = FreeFile()

    Open sFileName For Input As iFileNum
    Do While Not EOF(iFileNum)
        Line Input #iFileNum, sBuf
        sBuf = Trim(sBuf)

        If InStr(1, sBuf, "<") = 1 Then                              'ֻ��<��ͷ���в����Դ���
            iPos1 = InStr(1, sBuf, ">")                              '���ݼ����Ų��Ϊ�������֡�
            typeIndicator = Trim(Left(sBuf, iPos1 - 1))
            typeIndicator = Right(typeIndicator, Len(typeIndicator) - 1)

            strContent = Right(sBuf, Len(sBuf) - iPos1)

            If typeIndicator = iEOTypes Then                         '������ƥ���
                val = Split(strContent, CONSTCHA1)                   '���
                '                Debug.Print val(0), val(1)
                If (Not IsMissing(strVarName)) And strVarName <> "" Then   '���ָ���˾�������ƣ��򷵻�һֵ
                    If val(0) = strVarName Then
                        ReadFromFileByType = val(1)
                        GoTo exit_func
                    Else
                        GoTo nextline
                    End If
                Else
                    '����ͷ��������е�ֵ
                    ReDim Preserve result(0 To i)
                    result(i) = val
                    i = i + 1
                End If
            End If
        End If

nextline:
    Loop
    ReadFromFileByType = result

exit_func:
    Close #iFileNum
    Erase result
End Function
'���ļ����浽�̶���Ŀ¼
Sub savetofile(ByVal iEOTypes As EnumOfTypes, ByVal strContent As String, Optional ByVal sFileName As String)

    Dim iFileNum As Integer
    iFileNum = FreeFile()

    Open sFileName For Append As iFileNum
    Print #iFileNum, "# ---------------------------------------------------"
    Print #iFileNum, "# ����ʱ�䣺"; Format(Now(), "yyyy��MM��dd�� hh:mm:ss")
    Print #iFileNum, ""
    Print #iFileNum, iEOTypes; Tab(20); strContent
    Print #iFileNum, ""
    Print #iFileNum, "# ---------------------------------------------------"
    Close iFileNum
    Shell "Notepad.exe " & sFileName, vbNormalFocus
End Sub

Sub ConvertFileEncode(InputFile As String, OutputFile As String, Optional InputFileCode As String = "UNICODE", Optional OutputFileCode As String = "UTF-8")
    'http://www.cppblog.com/sunicdavy/archive/2012/09/13/190553.html'
    Dim ReadStream As Object
    Set ReadStream = CreateObject("ADODB.Stream")

    Dim FileContent As String

    With ReadStream
        .Type = 2               'adTypeText
        .Charset = InputFileCode
        .Open
        .LoadFromFile InputFile
        FileContent = .ReadText
        .Close

    End With

    Set ReadStream = Nothing



    Dim WriteStream As Object
    Set WriteStream = CreateObject("ADODB.Stream")


    With WriteStream
        .Type = 2               'adTypeText
        .Charset = OutputFileCode
        .Open
        .WriteText FileContent
        .savetofile OutputFile, 2  'adSaveCreateOverWrite

        .Flush
        .Close

    End With

    Set WriteStream = Nothing


End Sub

