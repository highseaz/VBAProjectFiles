Attribute VB_Name = "S1_acceptFormatChanges"
Sub acceptFormatChanges()
    With ActiveWindow.View
        .ShowFormatChanges = True
        '        .ShowRevisionsAndComments = False
        .ShowInsertionsAndDeletions = False
        
         ActiveDocument.AcceptAllRevisionsShown
         
        .ShowRevisionsAndComments = True
        .ShowFormatChanges = True
        .ShowInsertionsAndDeletions = True
    End With
End Sub

Sub ExportAsPDFFile()
'Áí´æÎªpdf
  
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        ActiveDocument.Path & Application.PathSeparator & ActiveDocument.Name & ".pdf", _
        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=True, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
    
End Sub

