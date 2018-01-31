Attribute VB_Name = "B0_VBAExImport"
Public Sub ExportModules()

If MsgBox("ExportModules? ", vbYesNo + vbCritical + vbDefaultButton2, "ExportModules") = vbNo Then Exit Sub

    Dim bExport As Boolean
    Dim wdSource As Document
    Dim szSourcedocument As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If

    '    On Error Resume Next
    '    Kill FolderWithVBAProjectFiles & "\*.*"
    '    On Error GoTo 0

    ''' NOTE: This doc must be opened
    szSourcedocument = ActiveDocument.Name
    Set wdSource = Application.Documents(szSourcedocument)

    If wdSource.VBProject.Protection = 1 Then
        MsgBox "The VBA in this document is protected," & _
        "not possible to export the code"
        Exit Sub
    End If

    szExportPath = FolderWithVBAProjectFiles & "\"

    For Each cmpComponent In wdSource.VBProject.VBComponents

        bExport = True
        szFileName = cmpComponent.Name

        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' Don't try to export.
                bExport = False
        End Select

        If InStr(1, szFileName, "_") > 1 And bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            ''' remove it from the project if you want
            '''wdSource.VBProject.VBComponents.Remove cmpComponent

        End If
    Next cmpComponent
    MsgBox "Export is finished"
End Sub


Public Sub ImportModules()
    Dim wdTarget As Document
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim szTargetdocument As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    '    If ActiveDocument.Name = ThisDocument.Name Then
    '        MsgBox "Select another destination document" & _
    '        "Not possible to import in this document "
    '        Exit Sub
    '    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This document must be open in Excel.
    szTargetdocument = ActiveDocument.Name
    Set wdTarget = Application.Documents(szTargetdocument)

    If wdTarget.VBProject.Protection = 1 Then
        MsgBox "The VBA in this document is protected," & _
        "not possible to Import the code"
        Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"

    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.count = 0 Then
        MsgBox "There are no files to import"
        Exit Sub
    End If

    'Delete all modules/Userforms from the Activedocument
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wdTarget.VBProject.VBComponents

    ''' Import all the code modules in the specified path
    ''' to the Activedocument.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
        If InStr(1, objFile.Name, "_") > 1 Then

            If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
                cmpComponents.Import objFile.Path
            End If
        End If

    Next objFile

    MsgBox "Import is ready"
End Sub

Function FolderWithVBAProjectFiles() As String
    '    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String

    '    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    '    SpecialPath = WshShell.SpecialFolders("MyDocuments")
    SpecialPath = MYWORKPATH_CODE
    

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If

    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
        On Error Resume Next
        MkDir SpecialPath & "VBAProjectFiles"
        On Error GoTo 0
    End If

    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
        FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
    Else
        FolderWithVBAProjectFiles = "Error"
    End If

End Function

Function DeleteVBAModulesAndUserForms()
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent

    Set VBProj = ActiveDocument.VBProject

    For Each VBComp In VBProj.VBComponents
        If VBComp.Type = vbext_ct_Document Or InStr(1, VBComp.Name, "VBA") > 1 Then
            'Thisdocument or worksheet module
            'We do nothing
        ElseIf InStr(1, VBComp.Name, "_") > 1 Then
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp
End Function


