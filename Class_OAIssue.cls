VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_OAIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private IssuesCount As Integer
Private IssuesTextSet() As Class_IssuesText
Private jsDic As Object
Private myDocPath As String
Private allRefs As New class_References

Public Property Get Count() As Integer
    Count = IssuesCount
End Property

Function IssuesTexts(ByVal i As Integer) As Class_IssuesText
    Set IssuesTexts = IssuesTextSet(i)
End Function

Public Property Get DocPath() As String
    DocPath = myDocPath
End Property


Public Property Let DocPath(ByVal DocumentFullPath As String)
    Dim myDocument As Document
    If LCase(Right(DocumentFullPath, 3)) = "pdf" Then
        myDocPath = NewfileFromPDFWithFormat(DocumentFullPath, "docx")
    ElseIf LCase(Right(DocumentFullPath, 4)) = "docx" Or LCase(Right(DocumentFullPath, 3)) = "doc" Then
        myDocPath = DocumentFullPath
    End If

    If myDocPath = vbNullString Then
        Debug.Print "Error: myDocPath is null."
        Exit Property
    End If

    Application.ScreenUpdating = False
    Set myDocument = Documents.Open(FileName:=myDocPath, Visible:=False)
    CleanDocFile myDocument
    InitializeInfo myDocument
    myDocument.Close SaveChanges:=wdDoNotSaveChanges
    Application.ScreenUpdating = True

End Property
Function Allreferences() As class_References
    Set Allreferences = allRefs
End Function
Function ReferenceItem(ByVal i As Integer) As reftype
    ReferenceItem = allRefs.refItem(i)
End Function


Private Sub InitializeInfo(myDocument As Document)

    Dim rngStart As Range
    Dim rngEnd As Range
    Dim reobj As New Class_VBScriptRegExp
    Dim para As Paragraph
    IssuesCount = 0

    Set rngStart = RangeIncludingStr("detailed action", myDocument)
    Set rngEnd = RangeIncludingStr("conclusion", myDocument)
    paraStart = myDocument.Range(0, rngStart.End).Paragraphs.Count
    paraEnd = myDocument.Range(0, rngEnd.End).Paragraphs.Count

    Dim pb As ProgressBar
    Set pb = New ProgressBar
    pb.Initialize "running", paraEnd - paraStart - 1

    For paraindex = paraStart To paraEnd
        Set para = myDocument.Paragraphs(paraindex)
        For i = 1 To jsDic.Count

            If reobj.PStest(jsDic(i)("IssuePattern"), para.Range.Text) Then
                ReDim Preserve IssuesTextSet(0 To IssuesCount)

                Dim newIssuesText As New Class_IssuesText

                refinof = FindRefInIssuesText(jsDic(i)("refParaPattern"), para.Range.Text)

                newIssuesText.clsInitialize i, para.Range.Text, refinof
                Set IssuesTextSet(IssuesCount) = newIssuesText
                IssuesCount = IssuesCount + 1
                Set newIssuesText = Nothing
                Exit For
            End If
        Next i
        pb.AddProgress 1

    Next
    Unload pb

    Set para = Nothing
    Set rngStart = Nothing
    Set rngEnd = Nothing

End Sub

Private Sub Class_Initialize()

    Set jsDic = JsonReadFromConfFile("OAPatternTypes")

End Sub
Function ENSummary() As String
    Dim Str As String
    For i = 0 To Count - 1
        Str = Str & IssuesTexts(i).ENText & vbNewLine
    Next
    ENSummary = Str

End Function
Sub InsertCNsummary(Optional ByVal rngRange As Range)
    Dim newStart, newEnd
    If rngRange Is Nothing Then
        Set rngRange = Selection.Range
    End If
    newStart = rngRange.Start
    rngRange.Text = ""
    rngRange.InsertAfter CNSummary

    newEnd = newStart + Len(CNSummary) - IssuesCount

    applyListFormat ActiveDocument.Range(newStart, newEnd)

End Sub
Function CNSummary() As String
    Dim Str As String

    For i = 0 To Count - 1
        Str = Str & IssuesTexts(i).CNText & vbNewLine
    Next

    For j = 1 To allRefs.Count

        Dim reobj As New Class_VBScriptRegExp
        If reobj.PStest("D" & j, Str, True, False) Then
            If ReferenceItem(j - 1).refNumber <> vbNullString Then
                refNum = ReferenceItem(j - 1).refNumber
                priorityDate = getDictPatentInfobyPatNum(cleanPatentNum(refNum)).Item("priorityDate")
                var = Split(priorityDate, "-")
                priorityDateCN = var(0) & "年" & var(1) & "月" & var(2) & "日"
                Str = reobj.objRegExp.Replace(Str, "D" & j & _
                                                 "(" & ReferenceItem(j - 1).refName & _
                                                 "，" & refNum & _
                                                 "，其最早优先权日为 " & priorityDateCN & _
                                                 ")")
            Else
            
            End If
        End If

        ''
        '        Debug.Print patentinfodic
    Next
    CNSummary = Str
End Function

Private Function FindRefInIssuesText(ByVal refParaPattern As String, ByVal searchTextString As String)

    If refParaPattern = vbNullString Then
        FindRefInIssuesText = vbNullString
        Exit Function
    End If

    Dim Str As String

    Dim reobj As New Class_VBScriptRegExp

    If reobj.PStest(refParaPattern, searchTextString) Then

        Set refParaallMatches = reobj.PSEXE

        refstr = Split(refParaallMatches(0), "view of")
        refdocCount = UBound(refstr) + 1

        For j = 1 To refdocCount
            singleRefString = refstr(j - 1)

            refNoPattern = "(CN|EP|JP|US)[\w\s]*?(\d[\d\/\s,-]*\d)"

            If reobj.PStest(refNoPattern, singleRefString) Then
                Set singleRefallMatches = reobj.PSEXE
                singleRefNo = singleRefallMatches(0).submatches(0) & singleRefallMatches(0).submatches(1)
            End If

            refNamePattern1 = "(?:hereinafter)\s([A-Z]\w*)"
            refNamePattern2 = "[A-Z]\w*\s(?=et al)"
            refNamePattern3 = " ([A-Z]\w*)"

            If reobj.PStest(refNamePattern1, singleRefString) Then
                Set singleRefallMatches = reobj.PSEXE
                singleRefName = singleRefallMatches(0).submatches(0)
            ElseIf reobj.PStest(refNamePattern2, singleRefString, True) Then
                Set singleRefallMatches = reobj.PSEXE
                singleRefName = singleRefallMatches(0)
            ElseIf reobj.PStest(refNamePattern3, singleRefString, True) Then
                Set singleRefallMatches = reobj.PSEXE
                singleRefName = singleRefallMatches(0)
            End If

            addrefIndex = allRefs.Add(singleRefName, singleRefNo)
            Str = Str & "D" & addrefIndex + 1 & "、"
        Next
    End If
    Set reobj = Nothing
    FindRefInIssuesText = Left(Str, Len(Str) - 1)

End Function
Sub CleanDocFile(ByRef Doc As Document)

    With Doc.Range
        .Find.ClearFormatting
        .Find.Replacement.ClearFormatting
        With .Find
            .Text = "([a-zA-Z0-9 ])[^m^13]{1,}([a-z0-9 ])"
            .Replacement.Text = "\1 \2"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchWildcards = True
        End With
        .Find.Execute Replace:=wdReplaceAll
    End With

End Sub
