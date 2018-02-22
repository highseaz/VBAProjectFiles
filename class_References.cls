VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "class_References"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private refIndex As Integer 'from 0 to count -1
Private refCollection() As reftype
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

Public Property Get Count() As Integer
    Count = refIndex
End Property
Public Function refItem(ByVal i As Integer) As reftype
    refItem = refCollection(i)
End Function
Public Function Exist(ByVal Str As String) As Integer
    Str = Trim(LCase(Str))
    Exist = -1
    If refIndex = 0 Then Exit Function
    For i = 0 To UBound(refCollection)
        If LCase(refCollection(i).refName) = Str Or LCase(refCollection(i).refNumber) = Str Then
            Exist = i
            Exit Function
        End If
    Next
End Function

Public Function Add(ByVal strName As String, ByVal strNumber As String) As Integer
    strName = Trim(strName)
    strNumber = Trim(strNumber)

    If Exist(strName) <> -1 Then
'        Debug.Print "name exist: " & strName
        Add = Exist(strName)
        Exit Function
    ElseIf Exist(strNumber) <> -1 Then
'        Debug.Print "No. exist: " & strNumber
        Add = Exist(strName)
        Exit Function
    End If
    
'    If SafeArrayGetDim(refCollection) = 0 Then
'        refIndex = 0
'    Else
'        refIndex = UBound(refCollection) + 1
'    End If
    
    Dim ref As reftype
    ref.refIndex = refIndex + 1
    ref.refName = strName
    ref.refNumber = strNumber

    ReDim Preserve refCollection(0 To refIndex)
    refCollection(refIndex) = ref
'        Debug.Print "ref " & refIndex & "add: " & refCollection(refIndex).refName & refCollection(refIndex).refNumber
    Add = refIndex
    refIndex = refIndex + 1
End Function
Private Sub Class_Initialize()
    refIndex = 0
End Sub
