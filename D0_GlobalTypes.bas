Attribute VB_Name = "D0_GlobalTypes"
Type typeReplacement
    strFindWhat As String
    strReplace As String
    FlagWildcard As Boolean
End Type

Type typeVarsWithNameValue
    name As String
    Value As String
End Type

Enum EnumOfTypes
    typeVarsWithNameValue = 1
    typeReplacement = 2
End Enum

Type OAIssue
    IssuePattern As String
    IssueType As String
End Type


Type reftype
    refIndex As Integer
    refName As String
    refNumber As String
'    refNumberClean As String
End Type

