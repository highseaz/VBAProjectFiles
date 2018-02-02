Attribute VB_Name = "S2_addReferenceNumber"
Sub addReferenceNumber4str(str As String, num As String, Rangetobeamended As Range) '
    str = Trim(str)
    num = " (" & Trim(num) & ")"

    For Each fieldLoop In Rangetobeamended.Fields
        fieldLoop.Unlink
    Next fieldLoop

    rangst = Rangetobeamended.Start

    Dim strFound As String
    Dim RE As Object
    Set RE = CreateObject("vbscript.regexp")
    RE.Pattern = "(" & str & "(s|es)?)([,.;]|\s(?!\())"
    RE.Global = True
    RE.IgnoreCase = True

    If RE.test(Rangetobeamended.Text) Then
        Set allMatches = RE.Execute(Rangetobeamended.Text)
        For i = 0 To allMatches.Count - 1
            strFound = allMatches(i).submatches(0)
            istart = Rangetobeamended.Start + allMatches(i).firstindex
            iEnd = istart + Len(strFound) + Len(num) * i
            Debug.Print "找到的字符串为‘" & strFound & "’;其长度为" & Len(strFound) & ";共找到" & i & "个。"
            ActiveDocument.Range(Start:=istart, End:=iEnd).InsertAfter num
        Next
    Else
        Debug.Print "no ‘" & str & "’is found."
    End If

End Sub


Sub addReferenceNumber4Claims(str As String, num As String) '
    Dim rang As Range
    Set rang = ActiveDocument.Range(Start:=RangeIncludingStr("WHAT IS CLAIMED IS", ActiveDocument, True).End, _
                                End:=RangeIncludingStr("ABSTRACT", ActiveDocument, True).Start)
    Call addReferenceNumber4str(str, num, rang)
    Set rang = Nothing


End Sub
Sub addReferenceNumber4Claimswithform()
    Dim frm As UserForm2

    Set frm = New UserForm2
    frm.Show vbModalless

End Sub
