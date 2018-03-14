VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_Claims"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private ClaimIndexPattern As String
Private ClaimsCount As Integer
Private GroupNo As Integer
Private ClaimIndexRangesArray() As Range
Private ClaimRangesArray() 'As Range
Private workRange As Range
Private delimiter As String
Private claimQuotePattern As String
Private myDocument As Document

Private claimInfoSummary As String

Private claimInforTable() As claimInfo
Private DicGroupsInfo As New Dictionary

Private pLanguage As Languagetype

Enum Languagetype
    lng_CN
    lng_EN
End Enum
Private Sub Class_Initialize()
    ClaimIndexPattern = "\d+(?=\D)"
    pLanguage = lng_CN
    '    ClassInit
End Sub

'Property Let Language(ByVal lng As Languagetype)
'    pLanguage = pLanguage
'    ClassInit
'End Property

'Property Set DocumentFile(ByVal doc As Document)
'    Set myDocument = doc
'    ClassInit
'End Property

Public Property Get Count() As Integer
    Count = ClaimsCount
End Property

Public Function getClaimIndexRanges()
    getClaimIndexRanges = ClaimIndexRangesArray
End Function


Public Sub ClassInit(Optional ByVal doc As Document, Optional ByVal lng As Languagetype = lng_CN)

    Dim js As Object
    If pLanguage = lng_CN Then
        Set js = JsonReadFromConfFile("ClaimsPattern")("CN")
    ElseIf pLanguage = lng_EN Then
        Set js = JsonReadFromConfFile("ClaimsPattern")("EN")
    Else
        Exit Sub
    End If

    workRngPattern = js("workRngPattern")
    delimiter = js("delimiter")
    claimQuotePattern = js("claimQuotePattern")

    Application.ScreenUpdating = False
    ''---___________________
    If IsMissing(doc) Or (doc Is Nothing) Then
        Dim docPath As String
        docPath = SelectedFileWithDlog
        If docPath = "" Then
            MsgBox "no doc is inputted"
            Exit Sub
        End If
        Set myDocument = Documents.Open(FileName:=docPath, Visible:=False)
    Else
        Set myDocument = doc

    End If
    ''---

    myDocument.Revisions.AcceptAll

    Set workRange = RangeIncludingStr(workRngPattern, myDocument)
    If workRange Is Nothing Then
        MsgBox "workrange is zero"
        Class_Terminate
        Exit Sub
    End If
    workRange.EndOf Unit:=wdSection, Extend:=wdExtend

    clearField

    SplitRangeWithDelimiterIntoArray workRange, delimiter, ClaimRangesArray
    FindClaimIndexRanges

    processClaimDependentInfo
    GetGrouping
    ''---

    Application.ScreenUpdating = True
End Sub

Sub clearField()
    Dim C, cdel As Integer
    C = workRange.Fields.Count
    Dim ofld As Field
    For Each ofld In workRange.Fields
        If ofld.Type = wdFieldDocVariable Then
            ofld.Unlink
            '            ofld.Delete
        End If
    Next ofld

    For i = 1 To C
        myfieldname = "_FdClaim_" & i & "_FdClaim_"
        If DocVarExists(myfieldname) Then
            ActiveDocument.Variables(myfieldname).Delete
            Debug.Print myfieldname & " deleted."
            cdel = cdel + 1
        End If
    Next
    Debug.Print cdel & " vars deleted."

End Sub
Sub makeFields()
    Dim bTrack As Boolean
    bTrack = ActiveDocument.TrackRevisions
    ActiveDocument.TrackRevisions = False
    Application.ScreenUpdating = False

    For Each rng In ClaimIndexRangesArray
        rng.Text = "_FdClaim_" & rng.Text & "_FdClaim_"
    Next

    For i = 1 To ClaimsCount
        StrToField "_FdClaim_" & i & "_FdClaim_", i
    Next
    ActiveDocument.TrackRevisions = bTrack
    Application.ScreenUpdating = True
End Sub
Private Sub StrToField(ByVal myfieldname As String, ByVal myfieldValue As String)
    With ActiveDocument
        .Variables.Add name:=myfieldname, Value:=myfieldValue
        iPostion = WordfindPositionInText(myfieldname)
        If iPostion = 0 Then
            Debug.Print "no such string found."
            Exit Sub
        End If

        Do While iPostion > 0 And iPostion < workRange.End
            .Range(iPostion, iPostion + Len(myfieldname)).Text = ""
            .Fields.Add Range:=.Range(iPostion, iPostion), Type:=wdFieldDocVariable, Text:=myfieldname, PreserveFormatting:=True
            iPostion = WordfindPositionInText(myfieldname)
        Loop


        Dim ofld As Field
        For Each ofld In .Fields
            If ofld.Type = wdFieldDocVariable Then
                If getFieldName(ofld) = myfieldname Then ofld.Update
            End If
        Next ofld
    End With
    '    ActiveDocument.Fields.Update
End Sub


Private Sub FindClaimIndexRanges()

    Dim reobj As New Class_VBScriptRegExp
    Set re = reobj.objRegExp

    ClaimsCount = 0

    For Each ClaimRange In ClaimRangesArray
        Debug.Print UBound(ClaimRangesArray)
        '        ClaimRangeStartPosition = RngStartPosition + ClaimRange.Start
        ClaimRangeStartPosition = ClaimRange.Start

        re.Pattern = ClaimIndexPattern
        If re.test(ClaimRange.Text) Then
            Set allMatches = re.Execute(ClaimRange.Text)
            ClaimIndexFound = allMatches(0)

            ClaimIndexSPos = ClaimRangeStartPosition + allMatches(0).FirstIndex
            Debug.Print "ClaimIndexFound: " & ClaimIndexFound, ClaimIndexSPos
            reDimRanges ClaimIndexRangesArray, ClaimIndexSPos, ClaimIndexSPos + Len(ClaimIndexFound), myDocument
            ClaimsCount = ClaimsCount + 1
        End If

        re.Pattern = claimQuotePattern
        If re.test(ClaimRange.Text) Then
            Set amQuoteLong = re.Execute(ClaimRange.Text)

            For i = 0 To amQuoteLong.Count - 1
                ClaimQuoteLong = amQuoteLong(i)
                Debug.Print "ClaimQuoteLong: " & ClaimQuoteLong
                ClaimQuoteLongSPos = ClaimRangeStartPosition + amQuoteLong(i).FirstIndex

                re.Pattern = ClaimIndexPattern
                If re.test(ClaimQuoteLong) Then
                    Set amClaimIndex = re.Execute(ClaimQuoteLong)

                    For j = 0 To amClaimIndex.Count - 1
                        ClaimQuoteshort = amClaimIndex(j)
                        ClaimQuoteshortSPos = ClaimQuoteLongSPos + amClaimIndex(j).FirstIndex
                        reDimRanges ClaimIndexRangesArray, ClaimQuoteshortSPos, ClaimQuoteshortSPos + Len(ClaimQuoteshort), myDocument
                        Debug.Print "----ClaimQuoteshort:" & ClaimQuoteshort & "----" & ClaimQuoteshortSPos
                    Next j
                End If

            Next i
        End If
    Next

    Debug.Print ClaimsCount & " claims found"
End Sub
Private Sub processClaimDependentInfo()
    Dim rng As Range, oneClaims
    Dim fullSM As String
    Dim reobj As New Class_VBScriptRegExp
    Dim blank As claimInfo
    Dim newclaim As claimInfo
    '    Dim ctable As New Dictionary
    ReDim Preserve claimInforTable(0 To ClaimsCount)
    i = 1
    Dim ct2 As Dictionary

    claimPattern = "(\d+)(.*?)，"
    IndClaimPattern = "一种"
    DPDClaimPattern = "权利要求(\d+)"

    For Each ClaimRange In ClaimRangesArray
        oneclaim = ClaimRange.Text
        oneclaim = Replace(oneclaim, Chr$(13), vbNullString)
        oneclaim = Replace(oneclaim, " ", vbNullString)
        Debug.Print oneclaim

        If oneclaim <> "" Then
            newclaim = blank
            With newclaim
                If reobj.PStest(claimPattern, oneclaim) Then

                    Set head = reobj.PSEXE
                    .cNo = head(0).submatches(0)

                    If reobj.PStest(IndClaimPattern, head(0).submatches(1)) Then
                        .cKind = claimKind_IND
                        .cDependentFrom = 0
                        fullSM = head(0).submatches(1)
                        .cSubjectMatter = Right(fullSM, Len(fullSM) - InStr(1, fullSM, IndClaimPattern) + 1 - Len(IndClaimPattern))
                        If reobj.PStest(DPDClaimPattern, oneclaim) Then
                            .cDependentFrom = reobj.PSEXE(0).submatches(0)
                            .cKind = claimKind_FAKEIND
                        End If

                    ElseIf reobj.PStest(DPDClaimPattern, oneclaim) Then
                        .cDependentFrom = reobj.PSEXE(0).submatches(0)
                        .cKind = claimKind_DPD
                    Else
                        GoTo nextline
                    End If
                End If
                Debug.Print .cNo, .cSubjectMatter, .cKind, .cDependentFrom
            End With

            '            ReDim Preserve claimInforTable(0 To i)
            '            claimInforTable(newclaim.cNo) = newclaim
            If i <= ClaimsCount Then claimInforTable(i) = newclaim
            i = i + 1

        End If
nextline:
    Next
    Set reobj = Nothing



End Sub

Private Sub GetGrouping()
    'base 1

    GroupNo = 0
    For i = 1 To UBound(claimInforTable)
        With claimInforTable(i)
            Debug.Print "**:" & i, .cNo, .cSubjectMatter, .cKind, .cDependentFrom
            Select Case .cKind
                Case claimKind_IND
                    GroupNo = GroupNo + 1
                    DicGroupsInfo.Add .cNo, GroupNo

                Case claimKind_DPD
                    If DicGroupsInfo.Exists(.cDependentFrom) Then
                        DicGroupsInfo.Add .cNo, DicGroupsInfo(.cDependentFrom)
                    End If
                Case claimKind_FAKEIND
                    '''find father....
                    GroupNo = GroupNo + 1
                    DicGroupsInfo.Add .cNo, GroupNo

            End Select
        End With
    Next

    For j = 1 To GroupNo
        Dim onegroup()
        Dim currentSM As String
        For i = 1 To UBound(claimInforTable)
            With claimInforTable(i)
                Key = .cNo
                If DicGroupsInfo(Key) = j Then
                    Debug.Print Key, DicGroupsInfo(Key)
                    reDimArrayAdd onegroup, Key
                    If .cSubjectMatter <> vbNullString Then currentSM = .cSubjectMatter
                End If
            End With
        Next


        ender = IIf(j = GroupNo, "。", "，")
        claimInfoSummary = claimInfoSummary & "权利要求" & _
                                            ArrangeNumArray_merge(onegroup) & _
                                            "涉及一种" & currentSM & ender


        Debug.Print claimInfoSummary
        Erase onegroup
    Next

End Sub
Sub InsertClaimSummary(Optional ByVal rngRange As Range)
    Dim newStart, newEnd
    If rngRange Is Nothing Then
        Set rngRange = Selection.Range
    End If
    newStart = rngRange.Start
    rngRange.Text = ""
    rngRange.InsertAfter claimInfoSummary
End Sub

Private Sub Class_Terminate()
    myDocument.Close SaveChanges:=wdDoNotSaveChanges
End Sub
