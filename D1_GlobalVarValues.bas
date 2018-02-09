Attribute VB_Name = "D1_GlobalVarValues"
Sub WriteJson(ByVal jsonString As String)
    Dim confFilePath As String
    confFilePath = ThisDocument.Path & "\conf.json"
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim Fileout As Object
    Set Fileout = FSO.CreateTextFile(confFilePath, True)
    Fileout.Write jsonString
    Fileout.Close
End Sub

Function JsonReadFromConfFile() As Object
    Dim confFilePath As String
    confFilePath = ThisDocument.Path & "\conf.json"

    Dim FSO As New FileSystemObject
    Dim JsonTS As TextStream
    Dim JsonText As String

    Set JsonTS = FSO.OpenTextFile(confFilePath, ForReading)
    JsonText = JsonTS.ReadAll
    JsonTS.Close

    Set JsonReadFromConfFile = ParseJson(JsonText)

    Set JsonTS = Nothing
    Set FSO = Nothing

End Function
'''''''''''''''''''''''''chars''''''''''''''''''''''''''
Function CONSTCHA1() As String   ' which is =
    CONSTCHA1 = JsonReadFromConfFile("CHA1")
End Function
Function CONSTCHA2() As String ' which is $
    CONSTCHA2 = JsonReadFromConfFile("CHA2")
End Function
Function CONSTCaseIDPattern() As String
    CONSTCaseIDPattern = JsonReadFromConfFile("CaseIDPattern")
End Function


''''''''''''''''replaceKeywords'''''''''''''''''''''''''''
Function CONSNameFindStr() As String
    CONSNameFindStr = JsonReadFromConfFile("NameFindStr")
End Function
Function CONSNameReplaceStr() As String
    CONSNameReplaceStr = JsonReadFromConfFile("NameReplaceStr")
End Function

''''''''''''''''Paths'''''''''''''''''''''''''''

Function MYWORKPATH() As String
    '    MYWORKPATH = JsonReadFromConfFile("WORKPATH")
    MYWORKPATH = ThisDocument.Path
End Function

Function MYWORKPATH_CODE() As String
    MYWORKPATH_CODE = ThisDocument.Path & "\code"
End Function
Function TEMPLATE_Null() As String
    TEMPLATE_Null = MYWORKPATH_CODE & "\PCT-NULL.dotx"
End Function
Function TEMPLATE_Full() As String
    TEMPLATE_Full = MYWORKPATH_CODE & "\PCT-Full.docx"
End Function

Function PrintSectionNames(ByVal secNo As Integer) As String
    PrintSectionNames = JsonReadFromConfFile("NamebySectionNo")(secNo)
End Function

Function PCTSplitDelimiter(ByVal iNo As Integer) As String
    If iNo < JsonReadFromConfFile("PCTDelimiter").Count Then
        PCTSplitDelimiter = JsonReadFromConfFile("PCTDelimiter")(iNo + 1)
    Else
        PCTSplitDelimiter = ""
    End If
End Function
'Function PasteInsitu(ByVal iNo As Integer) As String
'    PasteInsitu = JsonReadFromConfFile("PasteInsitu")(iNo)
'End Function


'Function OAPatternTypesDic() As Scripting.Dictionary
'    Dim D As Object
'    Set D = CreateObject("Scripting.Dictionary")
'    D.Add "Rejection101", "under 35 U.S.C. 101"
'    D.Add "Rejection102", "under 35 U.S.C. 102"
'    D.Add "Rejection103", "under 35 U.S.C. 103"
'    D.Add "AllowableSubjectMatter", "would be allowable if rewritten"
'    Set OAPatternTypesDic = D
'    Set D = Nothing
'End Function
'


