VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_OAIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit

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
    '    Debug.Print myDocPath
    Set myDocument = Documents.Open(myDocPath)

    InitializeInfo myDocument
    myDocument.Close SaveChanges:=wdDoNotSaveChanges
End Property
Function Allreferences() As class_References
    Set Allreferences = allRefs
End Function
Function ReferenceItem(ByVal i As Integer) As reftype
    ReferenceItem = allRefs.refItem(i)
End Function


Private Sub InitializeInfo(myDocument As Document)
    Dim re As Object
    Dim i As Integer
    Dim para As Paragraph
    IssuesCount = 0

    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = True
    For Each para In myDocument.Paragraphs

        For i = 1 To jsDic.Count

            re.Pattern = jsDic(i)("IssuePattern") '

            If re.test(para.Range.Text) Then

                ReDim Preserve IssuesTextSet(0 To IssuesCount)

                Dim newIssuesText As New Class_IssuesText

                refinof = FindRefInIssuesText(jsDic(i)("refParaPattern"), para.Range.Text)

                newIssuesText.clsInitialize i, para.Range.Text, refinof
                Set IssuesTextSet(IssuesCount) = newIssuesText
                IssuesCount = IssuesCount + 1
                Set newIssuesText = Nothing
            End If
        Next i
    Next para
    Set para = Nothing

End Sub

Private Sub Class_Initialize()

    Set jsDic = JsonReadFromConfFile("OAPatternTypes")

End Sub
Function CNSummary() As String
    Dim str As String

    For i = 0 To Count - 1
        str = str & IssuesTexts(i).CNText & vbNewLine
    Next

    For j = 1 To allRefs.Count

        Dim reobj As New Class_VBScriptRegExp
        If reobj.PStest("D" & j, str, True, False) Then
            str = reobj.objRegExp.Replace(str, "D" & j & _
                                                 "(" & ReferenceItem(j - 1).refName & _
                                                 "-" & ReferenceItem(j - 1).refNumber & _
                                                 ")")
        End If
    Next
    CNSummary = str

End Function

Private Function FindRefInIssuesText(ByVal refParaPattern As String, ByVal searchTextString As String)
    If refParaPattern = vbNullString Then
        FindRefInIssuesText = vbNullString
        Exit Function
    End If

    Dim str As String

    Dim reobj As New Class_VBScriptRegExp
    If reobj.PStest(refParaPattern, searchTextString) Then

        Set refParaallMatches = reobj.PSEXE

        refstr = Split(refParaallMatches(0), "view of")
        refdocCount = UBound(refstr) + 1  '''count

        For j = 1 To refdocCount
            singleRefString = refstr(j - 1)

            refNoPattern = "(CN|EP|JP|US)[\w\s]*?(\d[\d\/\s,-]*\d)"

            If reobj.PStest(refNoPattern, singleRefString) Then
                Set singleRefallMatches = reobj.PSEXE
                singleRefNo = singleRefallMatches(0).submatches(0) & singleRefallMatches(0).submatches(1)
                '                Debug.Print "-D" & j & ":_" & singleRefNo
            End If

            refNamePattern1 = "(?:hereinafter)\s([A-Z]\w*)"
            refNamePattern2 = "[A-Z]\w*\s(?=et al)"
            refNamePattern3 = " ([A-Z]\w*)"

            If reobj.PStest(refNamePattern1, singleRefString) Then
                Set singleRefallMatches = reobj.PSEXE
                singleRefName = singleRefallMatches(0).submatches(0) '& singleRefallMatches(0).submatches(1)
            ElseIf reobj.PStest(refNamePattern2, singleRefString, True) Then
                Set singleRefallMatches = reobj.PSEXE
                singleRefName = singleRefallMatches(0) '.submatches(0) & singleRefallMatches(0).submatches(1)
            ElseIf reobj.PStest(refNamePattern3, singleRefString, True) Then
                Set singleRefallMatches = reobj.PSEXE
                singleRefName = singleRefallMatches(0)
            End If

            addrefIndex = allRefs.Add(singleRefName, singleRefNo)
            str = str & "D" & addrefIndex + 1 & "¡¢"
        Next
    End If
    Set reobj = Nothing
    FindRefInIssuesText = Left(str, Len(str) - 1)
End Function
