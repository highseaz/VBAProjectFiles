Attribute VB_Name = "F1_Find_Range_Dlog"
Function RangeIncludingStr(findwhat As String, ByVal Doc As Document, Optional ByVal spaceSenstive As Boolean = False) As Range
    '�����ض��ַ���������

    'reference��
    'https://www.experts-exchange.com/articles/1336/Using-Regular-Expressions-in-Visual-Basic-for-Applications-and-Visual-Basic-6.html
    'https://stackoverflow.com/questions/11354909/regex-word-macro-that-finds-two-words-within-a-range-of-each-other-and-then-ital

    Dim RE As Object
    Dim para As Paragraph
    Dim rang As Range
    icount = 0

    Set RE = CreateObject("VBScript.RegExp")
    If spaceSenstive Then
        RE.Pattern = findwhat
    Else
        RE.Pattern = addStrBetweenEachChr(findwhat, " *") & "([a-z0-9\- ]*)\r|\n"
    End If
    RE.IgnoreCase = True
    RE.Global = True
    For Each para In Doc.Paragraphs
        If RE.test(para.Range.text) Then
            icount = icount + 1
            
            Set rang = para.Range
        End If
    Next para
    If icount <> 1 Then Debug.Print icount & Chr(34) & findwhat & Chr(34) & " found."
    Set RangeIncludingStr = rang
    Set rang = Nothing

End Function

Function addStrBetweenEachChr(r As String, theAddedStr) As String ''�ַ�������������ַ���
    With CreateObject("vbscript.regexp")
        .Pattern = "(.)"
        .Global = True
        addStrBetweenEachChr = .Replace(r, "$1" & theAddedStr)
        addStrBetweenEachChr = Left(addStrBetweenEachChr, Len(addStrBetweenEachChr) - Len(theAddedStr))
    End With
End Function
Function WordfindPositionInText(strPatternFindWhat As String) As Long
    '����word����λ��
    Set myRange = ActiveDocument.Content
    With myRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting

        .text = strPatternFindWhat
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute
        If .Found Then
            WordfindPositionInText = myRange.Start
        Else
            WordfindPositionInText = 0
        End If
    End With
End Function

Public Function strStartWithEn(ByVal FindPattern As String, ByVal FindStr As String) As Boolean
    With CreateObject("vbscript.regexp")
        .Global = True
        .Pattern = FindPattern
        strStartWithEn = .test(FindStr)
    End With
End Function

Public Function getStrFromSelection() As String
'��ȡѡ���е��ַ���ȥ���س������˿ո�
    Dim s1 As String
    s1 = Selection
    If s1 = "" Or s1 = Chr(13) Or s1 = Chr(10) Then Exit Function
    s1 = Replace(s1, Chr(13), " ")
    s1 = Replace(s1, Chr(10), " ")
    getStrFromSelection = Trim(s1)
End Function
Public Function ExistEnterStr(ByVal str_src As String) As Boolean
'�ж��Ƿ���лس�
    If VBA.InStr(1, str_src, Chr(10), vbTextCompare) > 0 Or VBA.InStr(1, str_src, Chr(13), vbTextCompare) > 0 Then
        ExistEnterStr = True
    Else
        ExistEnterStr = False
    End If
End Function
Public Function GetClipBoardString() As String
    On Error Resume Next
    Dim MyData As New DataObject
    GetClipBoardString = ""
    MyData.GetFromClipboard
    GetClipBoardString = MyData.GetText
    Set MyData = Nothing
End Function

Public Function SelectedFileWithDlog() As String
    '�öԻ���ѡ��һ�ļ�
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        '��ѡ��
        .Filters.Clear
        '����ļ�������
        .Filters.Add "Word Files", "*.doc;*.docx"
        .Filters.Add "All Files", "*.*"
        '���������ļ�������
        If .Show = -1 Then
            'FileDialog ����� Show ������ʾ�Ի��򣬲��ҷ��� -1��������� OK���� 0��������� Cancel����
            MsgBox "��ѡ����ļ��ǣ�" & .SelectedItems(1), vbOKOnly + vbInformation, "hi"
              SelectedFileWithDlog = .SelectedItems(1)
        End If
    End With
End Function


Function ExtractStringByPatternFrom(FindPattern As String, ByVal text As String) As String
    Dim result As String
    Dim allMatches As Object
    Dim RE As Object
    Set RE = CreateObject("vbscript.regexp")
    
        RE.Pattern = FindPattern
        RE.Global = True
        RE.IgnoreCase = True
    Set allMatches = RE.Execute(text)
    
    If allMatches.Count <> 0 Then
    ' Debug.Print allMatches
        result = allMatches.Item(0)
    Else
    result = ""
    Debug.Print "ExtractString By Pattern From text failed!"
    End If
    
    ExtractStringByPatternFrom = result

End Function
