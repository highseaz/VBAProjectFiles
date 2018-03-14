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
    Dim rng As Range
    Set rng = RangeIncludingStr("\[CNsummary\]", ActiveDocument, True)
    OA.InsertCNsummary rng
    Debug.Print OA.ENSummary
    
    Dim claims As Class_Claims
    Set claims = New Class_Claims
    claims.ClassInit , lng_CN
    
    Set rng = RangeIncludingStr("\[ss\]", ActiveDocument, True)
    claims.InsertClaimSummary rng
End Sub

