Attribute VB_Name = "S9_References_web"
Function getDictPatentInfobyPatNum(ByVal patnum As String) As Dictionary

    Dim StrFromHtml As String
    Dim url As String
    Dim XMLHTTP As Object

    url = "https://patents.google.com/patent/" & patnum & "?hl=en-us"

    Set XMLHTTP = CreateObject("MSXML2.XMLHTTP")
    XMLHTTP.Open "GET", url, False
    XMLHTTP.setRequestHeader "Content-Type", "text/xml"
    XMLHTTP.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; rv:25.0) Gecko/20100101 Firefox/25.0"
    XMLHTTP.send
    Do While Not XMLHTTP.readyState = 4
        DoEvents
    Loop
    StrFromHtml = XMLHTTP.responseText

    ''''''''' -----------save it if needed   -------------'''''''''
    '    sFileName = MYWORKPATH & "\html.text"
    '    WriteStringtoFile StrFromHtml, sFileName
    ''''''''' -----------comments end      -------------'''''''''

    Dim PatentDocInfo As New Dictionary
    Dim reobj As New Class_VBScriptRegExp
    Dim googlePatentPatten As String, googleTimePatentPatten As String

    googlePatentPatten = "<dd\sitemprop=" & Chr(34) & "(.+?)" & Chr(34) & ".*?>(.+?)<\/DD"
    If reobj.PStest(googlePatentPatten, StrFromHtml) Then
        Set infoallMatches = reobj.PSEXE
        For i = 0 To infoallMatches.Count - 1
            strKey = Trim(infoallMatches(i).submatches(0))
            infoValue = Trim(infoallMatches(i).submatches(1))
            If Not PatentDocInfo.Exists(strKey) Then
                PatentDocInfo.Add strKey, infoValue
            Else
                PatentDocInfo(strKey) = PatentDocInfo(strKey) & ", " & infoValue
            End If
'            Debug.Print strKey, ": ", Tab, PatentDocInfo(strKey)
        Next
    End If

    googleTimePatentPatten = "<DD><TIME\sitemprop=" & Chr(34) & "(.+?)" & Chr(34) & "\sdatetime=" & Chr(34) & "(\d{4}-\d{2}-\d{2})" & Chr(34) & ">"
    If reobj.PStest(googleTimePatentPatten, StrFromHtml) Then

        Set infoallMatches = reobj.PSEXE
        For i = 0 To infoallMatches.Count - 1
            strKey = Trim(infoallMatches(i).submatches(0))
            infoValue = Trim(infoallMatches(i).submatches(1))
            If Not PatentDocInfo.Exists(strKey) Then
                PatentDocInfo.Add strKey, infoValue
            Else
                PatentDocInfo(strKey) = PatentDocInfo(strKey) & ", " & infoValue
            End If
'            Debug.Print strKey, ": ", Tab, PatentDocInfo(strKey)
        Next
    End If
    
    Set getDictPatentInfobyPatNum = PatentDocInfo
    Set reobj = Nothing
    Set PatentDocInfo = Nothing
End Function

Function WriteStringtoFile(ByVal strtoPrint As String, ByVal sFileName As String) As Boolean
    WriteStringtoFile = False
    Dim iFileNum As Integer
    iFileNum = FreeFile()
    Open sFileName For Append As iFileNum
    Print #iFileNum, "# ---------------------------------------------------"
    Print #iFileNum, "# 保存时间："; Format(Now(), "yyyy年MM月dd日 hh:mm:ss")
    Print #iFileNum, ""
    Print #iFileNum, strtoPrint
    Print #iFileNum, ""
    Print #iFileNum, "# ---------------------------------------------------"
    Close iFileNum
    '        Shell "Notepad.exe " & sFileName, vbNormalFocus
    WriteStringtoFile = True
End Function

Function cleanPatentNum(ByVal refNum As String) As String
        s = Replace(refNum, " ", "")
        s = Replace(s, "*", "")
        s = Replace(s, "/", "")
        s = Replace(s, "\", "")
        s = Replace(s, ",", "")
      cleanPatentNum = s
End Function
