VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_IssuesText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private strENText As String
Private strCNText As String
Private sInfoType As String ''''''''''''''''''''''''''''''''''''''''''''''''''
Private dic As Object
Private IndexinDic As Integer
Private strRefInfo As String

Public Property Get ENText() As String
    ENText = strENText
End Property
'
Public Property Get CNText() As String
    CNText = strCNText
End Property
'
Property Get InfoType() As String
    InfoType = sInfoType
End Property

'Private Property Let ENText(ByVal str As String)
'    str = Replace(str, Chr(13), vbNullString)
'    str = Replace(str, Chr(10), vbNullString)
'    strENText = Trim(str)
'End Property

'Private Property Let InfoType(ByVal NewValue As String)
'    sInfoType = Trim(NewValue)
'End Property


Sub clsInitialize(ByVal indexInPatternSet As Integer, ByVal ENText As String, ByVal RefInfo As String)
IndexinDic = indexInPatternSet

Set dic = JsonReadFromConfFile("OAPatternTypes")(indexInPatternSet)

    sInfoType = dic("IssueType")
    strENText = Replace(ENText, Chr(13), vbNullString)
    strENText = Replace(strENText, Chr(10), vbNullString)
    strENText = Trim(strENText)
    strRefInfo = RefInfo
    transformFromENtoCN

End Sub

Private Sub transformFromENtoCN()
    Dim claimIndexPatternInIssue As String
    claimIndexPatternInIssue = "(?:claim[s\s]*)(\d+[\d-, \w]*?)(?=(is)|(are)|(would))"

    Dim reobj As New Class_VBScriptRegExp
    Dim re, allMatches, temjs
    Set re = reobj.objRegExp
    re.Pattern = claimIndexPatternInIssue
    If re.Test(strENText) Then

        Set allMatches = re.Execute(strENText)
        strCNText = allMatches(0).submatches(0)

    End If
    strCNText = Replace(strCNText, "and", "ºÍ")
    strCNText = Replace(strCNText, " ", vbNullString)
    strCNText = Replace(strCNText, ",", "¡¢")
    
    If strRefInfo <> vbNullString Then
    strCNText = strCNText & dic("CNAddMiddle") & strRefInfo
    End If

    strCNText = dic("CNAddBefore") & strCNText & dic("CNAddAfter")
  
End Sub

Private Sub Class_Initialize()
 

End Sub
