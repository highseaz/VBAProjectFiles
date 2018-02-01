VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_OAIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private IssuesCount As Integer
Private IssuesTextSet() As Class_IssuesText
Private dic As Scripting.Dictionary

Private myDocPath As String

'Private Function OAPatternAndTypes(ByVal iNo As Integer) As OAIssue
'    OAPatternAndTypes.IssueType = OAPatternTypesDic.Keys(iNo)
'    OAPatternAndTypes.IssuePattern = OAPatternTypesDic.Items(iNo)
'End Function


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
    Debug.Print myDocPath
    Set myDocument = Documents.Open(myDocPath)



    InitializeInfo myDocument
    myDocument.Close SaveChanges:=wdDoNotSaveChanges


End Property



Private Sub InitializeInfo(myDocument As Document)


    Dim RE As Object
    Dim issueTypeNo As Integer
    Dim para As Paragraph
    IssuesCount = 0

    Set RE = CreateObject("VBScript.RegExp")
    RE.IgnoreCase = True
    RE.Global = True
    For Each para In myDocument.Paragraphs

        For issueTypeNo = 0 To dic.Count - 1

            RE.Pattern = dic.Items(issueTypeNo) '

            If RE.test(para.Range.text) Then

                ReDim Preserve IssuesTextSet(0 To IssuesCount)

                Dim newIssuesText As New Class_IssuesText
                newIssuesText.ENText = para.Range.text
                newIssuesText.InfoType = dic.Keys(issueTypeNo) '

                Set IssuesTextSet(IssuesCount) = newIssuesText

                IssuesCount = IssuesCount + 1

                Set newIssuesText = Nothing
            End If
        Next issueTypeNo
    Next para
    Set para = Nothing

End Sub

Private Sub Class_Initialize()

    Set dic = OAPatternTypesDic

End Sub
