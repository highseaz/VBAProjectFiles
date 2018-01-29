VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Event_ClassModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents App As word.Application
Attribute App.VB_VarHelpID = -1
Private Sub App_DocumentBeforeSave(ByVal Doc As Document, SaveAsUI As Boolean, Cancel As Boolean)

    If SaveAsUI = True Then Exit Sub
    If InStr(1, Doc.Path, "C:\", vbTextCompare) >= 1 Then Exit Sub

    If InStr(1, Doc.Path, MYWORKPATH, vbTextCompare) < 1 Then
        If MsgBox("This file is now saved in " & Chr(13) & Chr(13) & Doc.Path, vbYesNo + vbCritical + vbDefaultButton2, "ȷ�ϱ���") = vbNo Then
            Cancel = True
            Exit Sub

        End If

    End If

End Sub
Private Sub App_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean)
    If InStr(1, Doc.Path, "Content.Outlook") > 0 Then

        If MsgBox("This file is now saved in " & Chr(13) & Chr(13) & Doc.Path, vbYesNo + vbCritical + vbDefaultButton2, "ȷ�Ϲر�") = vbNo Then
            Cancel = True
            Exit Sub

        End If

    End If

    If Doc.Name Like "*customer*" Then

        Call ExportModules
        Call showGitHubDesktop

    End If
End Sub



'Private Sub Document_Open()
''
''
'End Sub
'
Private Sub App_DocumentOpen(ByVal Doc As Document)
    Call ImportModules

    Call RightClickNewMenuAmd
    SendKeys "+{F5}"
End Sub