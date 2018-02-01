VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_IssuesText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private strENText As String
Private strCNText As String
Private sInfoType As String ''''''''''''''''''''''''''''''''''''''''''''''''''


Public Property Get ENText() As String
    ENText = strENText
End Property
'
Public Property Get CNText() As String
    strCNText = transformFrom(strENText)
    CNText = strCNText
End Property
'
Property Get InfoType() As String
    InfoType = sInfoType
End Property

Property Let ENText(ByVal str As String)
    strENText = Trim(str)
End Property

Property Let InfoType(ByVal NewValue As String)
    sInfoType = Trim(NewValue)
End Property

Private Function transformFrom(ByVal str As String) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    transformFrom = "zzz"
End Function
