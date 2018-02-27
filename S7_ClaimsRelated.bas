Attribute VB_Name = "S7_ClaimsRelated"
Sub ClaimSyncAmend()
    Dim selectRange As Range
    Set selectRange = Selection.Range
    Dim myclaims As New Class_Claims
    
    
    myclaims.makeFields
    selectRange.Select
    updateSelectDocVar
End Sub

Sub insertOAsummary()
    Dim OA As New Class_OAIssue
    Dim pdffile As String
    pdffile = SelectedFileWithDlog
    If pdffile = "" Then Exit Sub
    
    OA.docPath = pdffile
    Set rng = RangeIncludingStr("\[CNsummary\]", ActiveDocument, True)
    OA.InsertCNsummary (rng)
    Debug.Print test1.ENSummary
    
    Dim claims As Class_Claims
    Set claims = New Class_Claims
    
    Set rng = RangeIncludingStr("\[ss\]", ActiveDocument, True)
    claims.InsertClaimSummary rng
End Sub

