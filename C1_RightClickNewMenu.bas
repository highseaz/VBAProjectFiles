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
            Half = Int(.Controls.count / 2) '�м�λ��


            Set NewButton1 = .Controls.Add(Type:=msoControlButton, Before:=Half)
            With NewButton1
                .Caption = "ɾ���س�" '��������
                .FaceId = 501 '�����FaceId
                .Visible = True '�ɼ�
                .OnAction = "DeleteUselessEnterinSelection" 'ָ����Ӧ������
            End With

            Set NewButton2 = .Controls.Add(Type:=msoControlButton, Before:=Half + 1)
            With NewButton2
                .Caption = "ת��Ϊ��" '��������
                .FaceId = 502 '�����FaceId
                .Visible = True '�ɼ�
                .OnAction = "changeStrToFieldWithDiag2" 'ָ����Ӧ������
            End With

            Set NewButton3 = .Controls.Add(Type:=msoControlButton, Before:=Half + 2)
            With NewButton3
                .Caption = "����ͬ����" '��������
                .FaceId = 503 '�����FaceId
                .Visible = True '�ɼ�
                .OnAction = "updateSelectDocVar" 'ָ����Ӧ������
            End With
        End With
    Next


End Sub
Sub ComReset() '���������Ҽ��˵�,���׻ָ�Ĭ������
    Application.CommandBars("Text").Reset
End Sub




