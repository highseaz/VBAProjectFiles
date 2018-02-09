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
Private ClaimIndexRangesArray() As Range
Private ClaimRangesArray() As Range
Private workRange As Range
Private delimiter As String
Private claimQuotePattern As String

Private pLanguage As Languagetype

Enum Languagetype
    CN
    EN
End Enum


Property Let clsInitialize(ByVal lng As Languagetype)
    pLanguage = lng
    ClassInit
End Property

Public Property Get Count() As Integer
    Count = ClaimsCount
End Property

Function getClaimIndexRanges()
    getClaimIndexRanges = ClaimIndexRangesArray
End Function
Private Sub Class_Initialize()
    ClaimIndexPattern = "\d+(?=\D)"
End Sub

Private Sub ClassInit()

    Dim js As Object
    If pLanguage = CN Then
        Set js = JsonReadFromConfFile("ClaimsPattern")("CN")
    ElseIf pLanguage = EN Then
        Set js = JsonReadFromConfFile("ClaimsPattern")("EN")
    Else
        Exit Sub
    End If

    workRngPattern = js("workRngPattern")
    delimiter = js("delimiter")
    claimQuotePattern = js("claimQuotePattern")

    Set workRange = RangeIncludingStr(workRngPattern, ActiveDocument)
    workRange.EndOf Unit:=wdSection, Extend:=wdExtend
    clearField
    SplitRangesWithDelimiter workRange, delimiter, ClaimRangesArray
    FindClaimIndexRanges
End Sub

Sub clearField()
    Dim c, cdel As Integer
    c = workRange.Fields.Count
    Dim ofld As Field
    For Each ofld In workRange.Fields
        If ofld.Type = wdFieldDocVariable Then
            ofld.Unlink
            '            ofld.Delete
        End If
    Next ofld

    For i = 1 To c
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
        .Variables.Add Name:=myfieldname, Value:=myfieldValue
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
        '        ClaimRangeStartPosition = RngStartPosition + ClaimRange.Start
        ClaimRangeStartPosition = ClaimRange.Start

        re.Pattern = ClaimIndexPattern
        If re.test(ClaimRange.Text) Then
            Set allMatches = re.Execute(ClaimRange.Text)
            ClaimIndexFound = allMatches(0)

            ClaimIndexSPos = ClaimRangeStartPosition + allMatches(0).FirstIndex
            Debug.Print "ClaimIndexFound: " & ClaimIndexFound, ClaimIndexSPos
            reDimRanges ClaimIndexRangesArray, ClaimIndexSPos, ClaimIndexSPos + Len(ClaimIndexFound)
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
                        reDimRanges ClaimIndexRangesArray, ClaimQuoteshortSPos, ClaimQuoteshortSPos + Len(ClaimQuoteshort)
                        Debug.Print "----ClaimQuoteshort:" & ClaimQuoteshort & "----" & ClaimQuoteshortSPos
                    Next j
                End If

            Next i
        End If
    Next

    Debug.Print ClaimsCount & " claims found"
End Sub


