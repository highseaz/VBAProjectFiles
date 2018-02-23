Attribute VB_Name = "F0_FfindFolderwithKey"
Function findFolderwithKey(ByVal Key As String) As String

If Key = "" Then GoTo Exit_withNULL

Dim strFoundDir As String
    strFoundDir = Dir(MYWORKPATH_Work & "\*" & Key & "*", vbDirectory)
    If LenB(strFoundDir) > 0 Then
        'Do the rest of your code
        Debug.Print "The found path is " & strFoundDir
        findFolderwithKey = MYWORKPATH_Work & "\" & strFoundDir
        Exit Function
    End If
    
Exit_withNULL:
        findFolderwithKey = ""
        Debug.Print "Path not found."
        Exit Function
End Function

Function infoFileFullPath() As String

    If Not DocVarExists("CaseID_self") Then
'     UserFormBaseInfo.Show
          newID = Trim(InputBox("CaseID_self", " ‰»Î–¬ƒ⁄»›"))
           infoFileFullPath = findFolderwithKey(newID) & "\info.txt"
 
Else
  infoFileFullPath = findFolderwithKey(ActiveDocument.Variables("CaseID_self").Value) & "\info.txt"
   End If

End Function

Function FolderFoundOrCreated(ByVal SpecialPath As String, ByVal detailedPath As String) As String

    Dim FSO As Object
    Dim FullPath As String
    
    '    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")
    '    SpecialPath = WshShell.SpecialFolders("MyDocuments")
    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    
    FullPath = SpecialPath & detailedPath

    If FSO.FolderExists(FullPath) = False Then
        On Error Resume Next
        MkDir FullPath
        On Error GoTo 0
    End If

    If FSO.FolderExists(FullPath) = True Then
        FolderFoundOrCreated = FullPath
    Else
        FolderFoundOrCreated = "Error"
    End If
End Function
