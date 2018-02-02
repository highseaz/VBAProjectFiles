Attribute VB_Name = "S3_splitbysections"
Sub splitall()
    Call InsertSectionBreak(2)
    Call splitBySections
End Sub
Sub splitBySections()
    Application.ScreenUpdating = False
    Dim myRange As Range
    Dim sourceDoc As Document
    Dim tarDoc As Document
    Dim spath As String
    Dim strBaseFilename As String
    Dim strNewFileName As String
    Dim tarDocFullName As String

    Set sourceDoc = word.ActiveDocument
    spath = sourceDoc.Path
    strBaseFilename = sourceDoc.Name

    For i = 1 To sourceDoc.Sections.Count
        Set myRange = sourceDoc.Sections(i).Range
        istart = myRange.Start
        iEnd = myRange.End - 1
        sourceDoc.Range(istart, iEnd).Copy

        Set tarDoc = word.Documents.Add(Visible:=False, Template:=TEMPLATEFullPath)
        tarDoc.Content.PasteAndFormat (wdFormatOriginalFormatting)

        strNewFileName = Replace(strBaseFilename, ".do", "_" & strNamebySectionNo(i) & ".do")
        strNewFileName = Replace(strNewFileName, ".", "_")
        strNewFileName = Replace(strNewFileName, "_doc", ".doc")

        tarDocFullName = spath & "\" & strNewFileName
        tarDoc.SaveAs2 tarDocFullName
        tarDoc.Close

        Application.PrintOut Filename:=Chr(34) & tarDocFullName & Chr(34), Background:=True
    Next

    Application.ScreenUpdating = True

    Set myRange = Nothing
    Set tarDoc = Nothing
    Set sourceDoc = Nothing


End Sub

Sub InsertSectionBreak(pagebreakerNO As Integer)

    Dim oRng As Range
    For i = 1 To pagebreakerNO
        Set oRng = word.ActiveDocument.Range

        With oRng.Find
            .ClearFormatting
            .MatchWildcards = False
            .Text = "^m"
            .Execute
        End With

        S = oRng.Start
        If S < 2 Then Exit Sub
        Debug.Print S
        Debug.Print ActiveDocument.Range(S, S + 1).Text
        ActiveDocument.Range(S, S + 1) = ""
        ActiveDocument.Range(S, S + 1).InsertBreak wdSectionBreakNextPage

    Next

End Sub

