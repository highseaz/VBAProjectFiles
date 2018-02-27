Attribute VB_Name = "S5_PDFrelated"


Sub ExportAsPDFFile(Optional doc As Document)

    If doc Is Nothing Or IsMissing(doc) Then Set doc = ActiveDocument
    With doc
        .ExportAsFixedFormat OutputFileName:= _
        .Path & Application.PathSeparator & .name & ".pdf", _
        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=True, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentWithMarkup, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
    End With
End Sub

Function NewfileFromPDFWithFormat(ByVal PDFPath As String, ByVal FileExtension As String) As String
PDFPath = LCase(PDFPath)
    'Check if the file exists.
    If Dir(PDFPath) = "" Then
        Debug.Print "File Path Error: Cannot find the PDF file!" & vbCrLf & "Check the PDF path and retry."
        Exit Function
    End If

    'Check if the input file is a PDF file.
    If LCase(Right(PDFPath, 3)) <> "pdf" Then
        Debug.Print "File Type Error: The input file is not a PDF file!"
        Exit Function
    End If
    'Saves a PDF file as another format using Adobe Professional.
    'In order to use the macro you must enable the Acrobat library from VBA editor:
    'Go to Tools -> References -> Adobe Acrobat xx.0 Type Library, where xx depends
    'on your Acrobat Professional version (i.e. 9.0 or 10.0) you have installed to your PC.
    'Alternatively you can find it Tools -> References -> Browse and check for the path
    'C:\Program Files\Adobe\Acrobat xx.0\Acrobat\acrobat.tlb
    'where xx is your Acrobat version (i.e. 9.0 or 10.0 etc.).


    Dim ExportFormat    As String
    Dim NewFilePath     As String
    NewFilePath = ""

    ''''''''''''''''---------------''''''''''''''''''''''''
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

    If LCase(FileExtension) <> "xls" Then
        NewFilePath = Replace(PDFPath, ".pdf", "." & LCase(FileExtension))
    Else
        NewFilePath = Replace(PDFPath, ".pdf", ".xml")
    End If

    If Dir(NewFilePath) <> "" Then
        Debug.Print "The  file:" & vbNewLine & NewFilePath & vbNewLine & "exist! "
        GoTo cleanExit
    End If
    ''''''''''''''''---------------''''''''''''''''''''''''



    Dim objAcroApp      As Acrobat.AcroApp
    Dim objAcroAVDoc    As Acrobat.AcroAVDoc
    Dim objAcroPDDoc    As Acrobat.AcroPDDoc
    Dim objJSO          As Object
    Dim boResult        As Boolean

    'Initialize Acrobat by creating App object.
    Set objAcroApp = CreateObject("AcroExch.App")
    Debug.Print Err.Description
    'Set AVDoc object.
    Set objAcroAVDoc = CreateObject("AcroExch.AVDoc")
    Debug.Print Err.Description
    'Open the PDF file.
    boResult = objAcroAVDoc.Open(PDFPath, "")
    Debug.Print Err.Description
    'Set the PDDoc object.
    Set objAcroPDDoc = objAcroAVDoc.GetPDDoc
    Debug.Print Err.Description
    'Set the JS Object - Java Script Object.
    Set objJSO = objAcroPDDoc.GetJSObject
    Debug.Print Err.Description
    'Check if the format is correct and there are no errors.
    If ExportFormat <> "Wrong Input" And Err.Number = 0 Then
        'Save PDF file to the new format.
        boResult = objJSO.SaveAs(NewFilePath, ExportFormat)
        'Inform the user that conversion was successfully.
        Debug.Print "The PDf file:" & vbNewLine & PDFPath & vbNewLine & vbNewLine & _
        "Was saved as: " & vbNewLine & NewFilePath

    Else

        'Inform the user that something went wrong.
        Debug.Print "The conversion of the following PDF file FAILED:" & _
        vbNewLine & PDFPath

    End If

    'Close the PDF file without saving the changes.
    boResult = objAcroAVDoc.Close(True)
    'Close the Acrobat application.
    boResult = objAcroApp.Exit
    'Release the objects.
    Set objAcroPDDoc = Nothing
    Set objAcroAVDoc = Nothing
    Set objAcroApp = Nothing

cleanExit:
    NewfileFromPDFWithFormat = NewFilePath


End Function

