Attribute VB_Name = "S8_TrackChangesConversion"
Sub TrackChangesConversion()

    Dim MyRevision As word.Revision
    Dim MyRevisions As word.Revisions
    Dim MyRange As word.Range
    Dim RevisionsEnabled As Boolean
    Dim selection_start As Long, selection_end As Long
    
    Set MyRevisions = word.Selection.Range.Revisions
    
    'the revision property of selection in Word appears unstable,
    'so instead manually make sure we are getting only the text within
    'the selection
    selection_start = word.Selection.Start
    selection_end = word.Selection.End
    
    If MyRevisions.Count = 0 Then
        MsgBox "Please select text that includes tracked changes.", vbExclamation, "Convert Track Changes to Amendment Formatting"
        Exit Sub
    End If
    
    RevisionsEnabled = word.ActiveDocument.TrackRevisions
    word.ActiveDocument.TrackRevisions = False
    word.Application.ScreenUpdating = False

    For Each MyRevision In MyRevisions
        
        'warn if the revision is incomplete
        Dim force_conversion As Boolean
        force_conversion = False
        If ((MyRevision.Range.End > selection_end) And (MyRevision.Range.Start >= selection_start) And (MyRevision.Range.Start < selection_end)) Or _
            (MyRevision.Range.End > selection_start) And (MyRevision.Range.Start < selection_start) And (MyRevision.Range.End <= selection_end) Then
            If MsgBox("Your text selection does not fully encompass at least one selected amendment. Do you also want to convert this partially selected region?", vbOKCancel, _
              "") <> vbCancel Then
                force_conversion = True
            End If
        End If
        
        'the revision property of selection in Word appears unstable/inaccurate,
        'so instead manually make sure we are getting only the text within
        'the selection
        If ((MyRevision.Range.Start >= selection_start) And (MyRevision.Range.End <= selection_end)) Or _
            (force_conversion = True) Then
            Set MyRange = MyRevision.Range
            
            'check if we are within a status indicator
            Dim prev_char As String, next_char As String
            prev_char = word.ActiveDocument.Range(MyRange.Start - 1, MyRange.Start)
            next_char = word.ActiveDocument.Range(MyRange.End, MyRange.End + 1)
            
            Dim range_text As String
            Dim this_is_indicator As Boolean
            this_is_indicator = False
            range_text = LCase(Trim(MyRange.Text))
                
            If ((prev_char = "(") Or next_char = ")") And ((range_text <> "original") Or _
                   (range_text = "new") Or _
                   (range_text = "original") Or _
                   (range_text = "currently amended") Or _
                   (range_text = "previously presented") Or _
                   (range_text = "cancelled") Or _
                   (range_text = "withdrawn") Or _
                   (range_text = "withdrawn - currently amended") Or _
                   (range_text = "not entered")) Then
                   this_is_indicator = True
            End If
                   
                
            If MyRevision.Type = wdRevisionDelete Then
                'check if this is not a status indicator
                If (this_is_indicator = False) Then
                        'use strikesthrough for long deletions
                        If MyRevision.Range.Characters.Count > 5 Then
                            MyRevision.Range.Font.StrikeThrough = True
                            MyRevision.Reject
                        'use double square brackets for short deletions
                        Else
                            MyRevision.Reject
                            MyRange.InsertBefore "[["
                            MyRange.InsertAfter "]]"
                            MyRange.Font.UnderLine = wdUnderlineNone
                            MyRange.Font.StrikeThrough = False
                        End If
                  Else
                    'Accept deletions of status indicators
                    MyRevision.Accept
                  End If
            ElseIf MyRevision.Type = wdRevisionInsert Then
                
                Set MyRange = MyRevision.Range
                MyRevision.Accept
                
                'do not underline status indicators
                If (this_is_indicator = False) Then
                   'use underline for insertions
                    MyRange.Font.UnderLine = wdUnderlineSingle
                End If
            End If
         End If
    Next
    
    word.ActiveDocument.TrackRevisions = RevisionsEnabled
    word.Application.ScreenUpdating = True
    
End Sub

