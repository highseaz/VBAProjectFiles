Attribute VB_Name = "B3_ListProcedures"
Public Enum ProcScope
    ScopePrivate = 1
    ScopePublic = 2
    ScopeFriend = 3
    ScopeDefault = 4
End Enum

Public Enum LineSplits
    LineSplitRemove = 0
    LineSplitKeep = 1
    LineSplitConvert = 2
End Enum

Public Type ProcInfo
    ProcName As String
    ProcKind As VBIDE.vbext_ProcKind
    ProcStartLine As Long
    ProcBodyLine As Long
    ProcCountLines As Long
    ProcScope As ProcScope
    ProcDeclaration As String
End Type

Function ProcedureInfo(ProcName As String, ProcKind As VBIDE.vbext_ProcKind, _
        CodeMod As VBIDE.CodeModule) As ProcInfo

    Dim PInfo As ProcInfo
    Dim BodyLine As Long
    Dim Declaration As String
    Dim FirstLine As String


    BodyLine = CodeMod.ProcStartLine(ProcName, ProcKind)
    If BodyLine > 0 Then
        With CodeMod
            PInfo.ProcName = ProcName
            PInfo.ProcKind = ProcKind
            PInfo.ProcBodyLine = .ProcBodyLine(ProcName, ProcKind)
            PInfo.ProcCountLines = .ProcCountLines(ProcName, ProcKind)
            PInfo.ProcStartLine = .ProcStartLine(ProcName, ProcKind)

            FirstLine = .Lines(PInfo.ProcBodyLine, 1)
            If StrComp(Left(FirstLine, Len("Public")), "Public", vbBinaryCompare) = 0 Then
                PInfo.ProcScope = ScopePublic
            ElseIf StrComp(Left(FirstLine, Len("Private")), "Private", vbBinaryCompare) = 0 Then
                PInfo.ProcScope = ScopePrivate
            ElseIf StrComp(Left(FirstLine, Len("Friend")), "Friend", vbBinaryCompare) = 0 Then
                PInfo.ProcScope = ScopeFriend
            Else
                PInfo.ProcScope = ScopeDefault
            End If
            PInfo.ProcDeclaration = GetProcedureDeclaration(CodeMod, ProcName, ProcKind, LineSplitKeep)
        End With
    End If

    ProcedureInfo = PInfo

End Function


Function GetProcedureDeclaration(CodeMod As VBIDE.CodeModule, _
        ProcName As String, ProcKind As VBIDE.vbext_ProcKind, _
        Optional LineSplitBehavior As LineSplits = LineSplitRemove) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' GetProcedureDeclaration
    ' This return the procedure declaration of ProcName in CodeMod. The LineSplitBehavior
    ' determines what to do with procedure declaration that span more than one line using
    ' the "_" line continuation character. If LineSplitBehavior is LineSplitRemove, the
    ' entire procedure declaration is converted to a single line of text. If
    ' LineSplitBehavior is LineSplitKeep the "_" characters are retained and the
    ' declaration is split with vbNewLine into multiple lines. If LineSplitBehavior is
    ' LineSplitConvert, the "_" characters are removed and replaced with vbNewLine.
    ' The function returns vbNullString if the procedure could not be found.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim LineNum As Long
    Dim s As String
    Dim Declaration As String

    On Error Resume Next
    LineNum = CodeMod.ProcBodyLine(ProcName, ProcKind)
    If Err.Number <> 0 Then
        Exit Function
    End If
    s = CodeMod.Lines(LineNum, 1)
    Do While Right(s, 1) = "_"
        Select Case True
            Case LineSplitBehavior = LineSplitConvert
                s = Left(s, Len(s) - 1) & vbNewLine
            Case LineSplitBehavior = LineSplitKeep
                s = s & vbNewLine
            Case LineSplitBehavior = LineSplitRemove
                s = Left(s, Len(s) - 1) & " "
        End Select
        Declaration = Declaration & s
        LineNum = LineNum + 1
        s = CodeMod.Lines(LineNum, 1)
    Loop
    Declaration = SingleSpace(Declaration & s)
    GetProcedureDeclaration = Declaration
End Function

Private Function SingleSpace(ByVal Text As String) As String
    Dim Pos As String
    Pos = InStr(1, Text, Space(2), vbBinaryCompare)
    Do Until Pos = 0
        Text = Replace(Text, Space(2), Space(1))
        Pos = InStr(1, Text, Space(2), vbBinaryCompare)
    Loop
    SingleSpace = Text
End Function

'You can call the ProcedureInfo function using code like the following:
Sub ShowProcedureInfo()
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim compName As String
    Dim ProcName As String
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim PInfo As ProcInfo

    compName = "B0_VBAExImport"
    ProcName = "ImportModules"
    ProcKind = vbext_pk_Proc

    Set VBProj = ActiveDocument.VBProject
    Set VBComp = VBProj.VBComponents(compName)
    Set CodeMod = VBComp.CodeModule

    PInfo = ProcedureInfo(ProcName, ProcKind, CodeMod)

    Debug.Print "ProcName: " & PInfo.ProcName
    Debug.Print "ProcKind: " & CStr(PInfo.ProcKind)
    Debug.Print "ProcStartLine: " & CStr(PInfo.ProcStartLine)
    Debug.Print "ProcBodyLine: " & CStr(PInfo.ProcBodyLine)
    Debug.Print "ProcCountLines: " & CStr(PInfo.ProcCountLines)
    Debug.Print "ProcScope: " & CStr(PInfo.ProcScope)
    Debug.Print "ProcDeclaration: " & PInfo.ProcDeclaration
End Sub
Sub ListProcedures()
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim LineNum As Long
    Dim NumLines As Long
    '    Dim WS As Worksheet
    Dim rng As Range
    Dim ProcName As String
    Dim ProcKind As VBIDE.vbext_ProcKind

    Set VBProj = ActiveDocument.VBProject
    Set VBComp = VBProj.VBComponents("B0_VBAExImport")
    Set CodeMod = VBComp.CodeModule
    With CodeMod
        LineNum = .CountOfDeclarationLines + 1

        Do Until LineNum >= .CountOfLines
            ProcName = .ProcOfLine(LineNum, ProcKind)
            s = .Lines(LineNum, 1)

            Debug.Print ProcName
            Debug.Print s
            LineNum = .ProcStartLine(ProcName, ProcKind) + _
                    .ProcCountLines(ProcName, ProcKind) + 1

        Loop
    End With

End Sub

Function ProcKindString(ProcKind As VBIDE.vbext_ProcKind) As String
    Select Case ProcKind
        Case vbext_pk_Get
            ProcKindString = "Property Get"
        Case vbext_pk_Let
            ProcKindString = "Property Let"
        Case vbext_pk_Set
            ProcKindString = "Property Set"
        Case vbext_pk_Proc
            ProcKindString = "Sub Or Function"
        Case Else
            ProcKindString = "Unknown Type: " & CStr(ProcKind)
    End Select

End Function



Sub ShowProcedureInfo2()
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim compName As String
    Dim ProcName As String
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim PInfo As ProcInfo

    Dim wdTable As Table
    Dim rowNew As Row

    Set wdTable = ActiveDocument.Tables(1)


    For Each VBComp In ActiveDocument.VBProject.VBComponents
        If VBComp.Type <> vbext_ct_StdModule Then GoTo next_comp

        compName = VBComp.name
        If InStr(1, compName, "_") < 1 Then GoTo next_comp

        Set rowNew = wdTable.Rows.Add(BeforeRow:=wdTable.Rows.last)
        rowNew.Cells.Merge

        wdTable.Cell(rowNew.index, 1).Range.InsertAfter compName

        rowNew.Alignment = wdAlignRowCenter
        rowNew.Shading.BackgroundPatternColor = wdColorYellow

        Set CodeMod = VBComp.CodeModule

        With CodeMod
            LineNum = .CountOfDeclarationLines + 1
            Do Until LineNum >= .CountOfLines

                ProcName = .ProcOfLine(LineNum, ProcKind)

                LineNum = .ProcStartLine(ProcName, ProcKind) + _
                    .ProcCountLines(ProcName, ProcKind) + 1

                PInfo = ProcedureInfo(ProcName, ProcKind, CodeMod)

                Set rowNew = wdTable.Rows.Add(BeforeRow:=wdTable.Rows.last)
                wdTable.Cell(rowNew.index, 1).Range.InsertAfter PInfo.ProcDeclaration

                '                Debug.Print CompName
                '                Debug.Print "ProcName: " & PInfo.ProcName
                '                Debug.Print "ProcDeclaration: " & PInfo.ProcDeclaration
                '                Debug.Print ""
            Loop
        End With


next_comp:
    Next

    ProcKind = vbext_pk_Proc

End Sub
Sub clearCommentInVBComponents(ByVal compName As String)

    Dim N                       As Long
    Dim i                        As Long
    Dim j                        As Long
    Dim k                       As Long
    Dim l                        As Long
    Dim LineText            As String
    Dim ExitString          As String
    Dim Quotes              As Long
    Dim Q                       As Long
    Dim StartPos            As Long


    With ThisDocument.VBProject.VBComponents("Class_OAIssue").CodeModule
        For j = .CountOfLines To 1 Step -1
            LineText = Trim(.Lines(j, 1))
            If LineText = "ExitString = " & """" & "Ignore Comments In This Module" & """" Then
                Exit For
            End If
            StartPos = 1
Retry:
            N = InStr(StartPos, LineText, "'")
            Q = InStr(StartPos, LineText, """")
            Quotes = 0
            If Q < N Then
                For l = 1 To N
                    If Mid(LineText, l, 1) = """" Then
                        Quotes = Quotes + 1
                    End If
                Next l
            End If
            If Quotes / 2 = 1 Then
                StartPos = N + 1
GoTo Retry:
            Else
                Select Case N
                    Case Is = 0
                    Case Is = 1
                        .DeleteLines j, 1
                    Case Is > 1
                        .ReplaceLine j, Left(LineText, N - 1)
                        Debug.Print "line " & j & ": <" & LineText & "> Is amended."
                End Select
            End If
        Next j
    End With


    ExitString = "Ignore Comments In This Module"

End Sub
