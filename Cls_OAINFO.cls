VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cls_OAINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private PCount As Integer
Private paraText() As cls_paraText
Private patterns(5) As String
Private myDocPath As String


Public Property Get count() As Integer
    count = PCount
End Property

Function Paras(ByVal i As Integer) As cls_paraText
    Set Paras = paraText(i)
End Function

Public Property Get DocPath() As String
    DocPath = myDocPath
End Property

Public Property Let DocPath(ByVal DocumentFullPath As String)
    myDocPath = DocumentFullPath

    Dim myDocument As Document
    If LCase(Right(DocumentFullPath, 3)) = "pdf" Then
        Set myDocument = Documents.Open(NewfileFromPDFWithFormat(DocumentFullPath, "docx"))
    Else
        Set myDocument = Documents.Open(DocumentFullPath)
    End If
    InitializeInfo myDocument

End Property

Private Sub Class_Initialize()

    patterns(0) = "under 35 U.S.C. 103" '''''''''''''''''''''''''''''''''''''

End Sub


Private Sub InitializeInfo(myDocument As Document)

    Dim RE As Object
    Dim para As Paragraph
    PCount = 0

    Set RE = CreateObject("VBScript.RegExp")
    RE.IgnoreCase = True
    RE.Global = True
    For Each para In myDocument.Paragraphs

        RE.Pattern = patterns(0) '''''''''''''''''''''''''''''''''''''
        
        If RE.test(para.Range.text) Then

            ReDim Preserve paraText(0 To PCount)

            Dim newparaText As New cls_paraText
            newparaText.ENText = para.Range.text
            '            newparaText.InfoType = 1 '
            '''''''''''''''''''''''''''''''''''''''''''
            Set paraText(PCount) = newparaText
            PCount = PCount + 1

            Set newparaText = Nothing
        End If


    Next para
    Set para = Nothing

End Sub

