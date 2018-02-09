VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_VBScriptRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Public objRegExp As Object
    
   

Private Sub Class_Initialize()

 
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.IgnoreCase = True
    objRegExp.Global = True

End Sub
