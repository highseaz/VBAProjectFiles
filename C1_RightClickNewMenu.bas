Attribute VB_Name = "C1_RightClickNewMenu"
Public Sub RightClickNewMenuAmd()
    Dim Half As Byte
    Call ComReset
    On Error Resume Next
    Dim NewButton1 As CommandBarButton

    Dim comBarName
    comBarName = Array("Lists", "Text")
'    Debug.Print 1, comBarName(0), comBarName(1)
Dim i As Integer

    For i = 0 To UBound(comBarName)
        With Application.CommandBars(comBarName(i))
        
            .Reset
            Half = Int(.Controls.count / 2) '中间位置


            Set NewButton1 = .Controls.Add(Type:=msoControlButton, Before:=Half)
            With NewButton1
                .Caption = "删除回车" '命令名称
                .FaceId = 501 '命令的FaceId
                .Visible = True '可见
                .OnAction = "DeleteUselessEnterinSelection" '指定响应过程名
            End With

            Set NewButton2 = .Controls.Add(Type:=msoControlButton, Before:=Half + 1)
            With NewButton2
                .Caption = "转化为域" '命令名称
                .FaceId = 502 '命令的FaceId
                .Visible = True '可见
                .OnAction = "changeStrToFieldWithDiag2" '指定响应过程名
            End With

            Set NewButton3 = .Controls.Add(Type:=msoControlButton, Before:=Half + 2)
            With NewButton3
                .Caption = "更新同名域" '命令名称
                .FaceId = 503 '命令的FaceId
                .Visible = True '可见
                .OnAction = "updateSelectDocVar" '指定响应过程名
            End With
        End With
    Next


End Sub
Sub ComReset() '重新设置右键菜单,彻底恢复默认设置
    Application.CommandBars("Text").Reset
End Sub




