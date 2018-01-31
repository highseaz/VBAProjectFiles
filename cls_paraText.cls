VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_paraText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private strENText As String
Private strCNText As String
Private sInfoType As Integer ''''''''''''''''''''''''''''''''''''''''''''''''''


Public Property Get ENText() As String
    ENText = strENText
End Property
'
Public Property Get CNText() As String

''todo
strCNText = transformFrom(strENText)
    CNText = strCNText '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Property
'

Property Let ENText(ByVal str As String)
    strENText = str
End Property

Public Property Get InfoType() As Integer
    InfoType = sInfoType
End Property

Public Property Let InfoType(ByVal NewValue As Integer)
    sInfoType = NewValue
End Property

Function transformFrom(ByVal str As String) As String

transformFrom = "zzz"
End Function
