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

        If InStr(1, sBuf, "<") = 1 Then                              '只有<开头的行才予以处理’
            iPos1 = InStr(1, sBuf, ">")                              '根据尖括号拆分为两个部分’
            typeIndicator = Trim(Left(sBuf, iPos1 - 1))
            typeIndicator = Right(typeIndicator, Len(typeIndicator) - 1)

            strContent = Right(sBuf, Len(sBuf) - iPos1)

            If typeIndicator = iEOTypes Then                         '类型是匹配的
                val = Split(strContent, CONSTCHA1)                   '拆分
                '                Debug.Print val(0), val(1)
                If (Not IsMissing(strVarName)) And strVarName <> "" Then   '如果指定了具体的名称，则返回一值
                    If val(0) = strVarName Then
                        ReadFromFileByType = val(1)
                        GoTo exit_func
                    Else
                        GoTo nextline
                    End If
                Else
                    '否则就返回所有行的值
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
'将文件保存到固定的目录
Sub savetofile(ByVal iEOTypes As EnumOfTypes, ByVal strContent As String, Optional ByVal sFileName As String)
 
    Dim iFileNum As Integer
    iFileNum = FreeFile()

    Open sFileName For Append As iFileNum
    Print #iFileNum, "# ---------------------------------------------------"
    Print #iFileNum, "# 保存时间："; Format(Now(), "yyyy年MM月dd日 hh:mm:ss")
    Print #iFileNum, ""
    Print #iFileNum, iEOTypes; Tab(20); strContent
    Print #iFileNum, ""
    Print #iFileNum, "# ---------------------------------------------------"
    Close iFileNum
    Shell "Notepad.exe " & sFileName, vbNormalFocus
End Sub

