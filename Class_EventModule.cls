VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_EventModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents App As word.Application
Attribute App.VB_VarHelpID = -1
Private Sub App_DocumentBeforeSave(ByVal doc As Document, SaveAsUI As Boolean, Cancel As Boolean)

    If SaveAsUI = True Then Exit Sub
    If InStr(1, doc.Path, "C:\", vbTextCompare) >= 1 Then Exit Sub

    If InStr(1, doc.Path, MYWORKPATH, vbTextCompare) < 1 Then
        If MsgBox("This file is now saved in " & Chr(13) & Chr(13) & doc.Path, vbYesNo + vbCritical + vbDefaultButton2, "确认保存") = vbNo Then
            Cancel = True
            Exit Sub

        End If

    End If

End Sub
Private Sub App_DocumentBeforeClose(ByVal doc As Document, Cancel As Boolean)
    If InStr(1, doc.Path, "Content.Outlook") > 0 Then

        If MsgBox("This file is now saved in " & Chr(13) & Chr(13) & doc.Path, vbYesNo + vbCritical + vbDefaultButton2, "确认关闭") = vbNo Then
            Cancel = True
            Exit Sub

        End If

    End If

    If doc.name Like "*customer*" Then
        If MsgBox("ExportModules? ", vbYesNo + vbCritical + vbDefaultButton2, "ExportModules") = vbYes Then

            Call ExportModules
            Call showGitHubDesktop
        End If
    End If
End Sub

Private Sub App_DocumentOpen(ByVal doc As Document)
    Call RightClickNewMenuAmd
    SendKeys "+{F5}"
End Sub
