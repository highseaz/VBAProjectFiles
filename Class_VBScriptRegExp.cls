VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_VBScriptRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public objRegExp As Object
Private sPattern As String
Private sText As String

Private Sub Class_Initialize()

    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = True
    objRegExp.IgnoreCase = True
End Sub

Function PStest(ByVal strPattern As String, _
                            ByVal searchTextString As String, _
                            Optional bIgnoreCase As Boolean = True, _
                            Optional bGlobal As Boolean = True) As Boolean

    objRegExp.Global = bGlobal
    objRegExp.IgnoreCase = bIgnoreCase
    sPattern = strPattern
    sText = searchTextString

    objRegExp.Pattern = sPattern
    PStest = objRegExp.test(sText)

End Function

Function PSEXE() As Object
    If sText = vbNullString Or sPattern = vbNullString Then
        Debug.Print "objRegExp is not initialized."
        Set PSEXE = Nothing
    Else
        Set PSEXE = objRegExp.Execute(sText)
    End If
End Function

Function PSReplace(ByVal strReplacement As String) As String

    If sText = vbNullString Or sPattern = vbNullString Then
        Debug.Print "objRegExp is not initialized."
        PSReplace = ""
    Else
        PSReplace = objRegExp.Replace(sText, strReplacement)
    End If
End Function

Private Sub Class_Terminate()
    Set objRegExp = Nothing
End Sub
