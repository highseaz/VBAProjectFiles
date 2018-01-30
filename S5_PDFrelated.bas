Attribute VB_Name = "S5_PDFrelated"


Sub ExportAsPDFFile(Optional Doc As Document)

If Doc Is Nothing Or IsMissing(Doc) Then Set Doc = ActiveDocument
With Doc
    .ExportAsFixedFormat OutputFileName:= _
        .Path & Application.PathSeparator & .Name & ".pdf", _
        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=True, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
End With
End Sub

Function NewfileFromPDFWithFormat(ByVal PDFPath As String, ByVal FileExtension As String) As String

    'Saves a PDF file as another format using Adobe Professional.
    'In order to use the macro you must enable the Acrobat library from VBA editor:
    'Go to Tools -> References -> Adobe Acrobat xx.0 Type Library, where xx depends
    'on your Acrobat Professional version (i.e. 9.0 or 10.0) you have installed to your PC.
    'Alternatively you can find it Tools -> References -> Browse and check for the path
    'C:\Program Files\Adobe\Acrobat xx.0\Acrobat\acrobat.tlb
    'where xx is your Acrobat version (i.e. 9.0 or 10.0 etc.).

    Dim objAcroApp      As Acrobat.AcroApp
    Dim objAcroAVDoc    As Acrobat.AcroAVDoc
    Dim objAcroPDDoc    As Acrobat.AcroPDDoc
    Dim objJSO          As Object
    Dim boResult        As Boolean
    Dim ExportFormat    As String
    Dim NewFilePath     As String
    NewFilePath = ""

    'Check if the file exists.
    If Dir(PDFPath) = "" Then
        MsgBox "Cannot find the PDF file!" & vbCrLf & "Check the PDF path and retry.", _
                vbCritical, "File Path Error"
        Exit Function
    End If

    'Check if the input file is a PDF file.
    If LCase(Right(PDFPath, 3)) <> "pdf" Then
        MsgBox "The input file is not a PDF file!", vbCritical, "File Type Error"
        Exit Function
    End If

    'Initialize Acrobat by creating App object.
    Set objAcroApp = CreateObject("AcroExch.App")

    'Set AVDoc object.
    Set objAcroAVDoc = CreateObject("AcroExch.AVDoc")

    'Open the PDF file.
    boResult = objAcroAVDoc.Open(PDFPath, "")

    'Set the PDDoc object.
    Set objAcroPDDoc = objAcroAVDoc.GetPDDoc

    'Set the JS Object - Java Script Object.
    Set objJSO = objAcroPDDoc.GetJSObject

    'Check the type of conversion.
    Select Case LCase(FileExtension)
        Case "eps": ExportFormat = "com.adobe.acrobat.eps"
        Case "html", "htm": ExportFormat = "com.adobe.acrobat.html"
        Case "jpeg", "jpg", "jpe": ExportFormat = "com.adobe.acrobat.jpeg"
        Case "jpf", "jpx", "jp2", "j2k", "j2c", "jpc": ExportFormat = "com.adobe.acrobat.jp2k"
        Case "docx": ExportFormat = "com.adobe.acrobat.docx"
        Case "doc": ExportFormat = "com.adobe.acrobat.doc"
        Case "png": ExportFormat = "com.adobe.acrobat.png"
        Case "ps": ExportFormat = "com.adobe.acrobat.ps"
        Case "rft": ExportFormat = "com.adobe.acrobat.rft"
        Case "xlsx": ExportFormat = "com.adobe.acrobat.xlsx"
        Case "xls": ExportFormat = "com.adobe.acrobat.spreadsheet"
        Case "txt": ExportFormat = "com.adobe.acrobat.accesstext"  '"com.adobe.acrobat.plain-text"
        Case "tiff", "tif": ExportFormat = "com.adobe.acrobat.tiff"
        Case "xml": ExportFormat = "com.adobe.acrobat.xml-1-00"
        Case Else: ExportFormat = "Wrong Input"
    End Select

    'Check if the format is correct and there are no errors.
    If ExportFormat <> "Wrong Input" And Err.Number = 0 Then
        'Set the path of the new file. Note that Adobe instead of xls uses xml files.
        'That's why here the xls extension changes to xml.
        If LCase(FileExtension) <> "xls" Then
            NewFilePath = Replace(PDFPath, ".pdf", "." & LCase(FileExtension))
        Else
            NewFilePath = Replace(PDFPath, ".pdf", ".xml")
        End If

        If Dir(NewFilePath) <> "" Then
            Debug.Print "The  file:" & vbNewLine & NewFilePath & vbNewLine & "exist! "
        Else
            'Save PDF file to the new format.
            boResult = objJSO.SaveAs(NewFilePath, ExportFormat)
            'Inform the user that conversion was successfully.
            Debug.Print "The PDf file:" & vbNewLine & PDFPath & vbNewLine & vbNewLine & _
        "Was saved as: " & vbNewLine & NewFilePath, vbInformation, "Conversion finished successfully"
        End If
    Else

        'Inform the user that something went wrong.
        Debug.Print "Something went wrong!" & vbNewLine & "The conversion of the following PDF file FAILED:" & _
        vbNewLine & PDFPath, vbInformation, "Conversion failed"

    End If

    NewfileFromPDFWithFormat = NewFilePath

cleanExit:
    'Close the PDF file without saving the changes.
    boResult = objAcroAVDoc.Close(True)
    'Close the Acrobat application.
    boResult = objAcroApp.Exit
    'Release the objects.
    Set objAcroPDDoc = Nothing
    Set objAcroAVDoc = Nothing
    Set objAcroApp = Nothing
End Function

