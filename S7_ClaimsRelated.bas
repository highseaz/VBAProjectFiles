Attribute VB_Name = "S7_ClaimsRelated"
Sub ClaimSyncAmend()
    Dim selectRange As Range
    Set selectRange = Selection.Range
    Dim myclaims As New Class_Claims
    Dim ClRanges() As Range
    myclaims.clsInitialize = CN
    myclaims.makeFields
    selectRange.Select
    updateSelectDocVar
End Sub

