Attribute VB_Name = "F2_patterns"
Public Function DoesIDMatchPattern(ByVal x As String) As Boolean
    With CreateObject("vbscript.regexp")
        .Global = True
        .Pattern = CONSTCaseIDPattern
        DoesIDMatchPattern = .test(x)
    End With
End Function
