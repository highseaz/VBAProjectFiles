Attribute VB_Name = "S10_outlook"
Sub sendMail()
    With ActiveDocument
        Subj = "(due:)������" & .Variables("CaseID_Client").Value & "������һ�塿" & _
                   .Variables("CaseID_Client").Value & "; " & _
                   .Variables("CaseID_Official").Value & "; " & _
                   .Variables("CaseID_self").Value & "; " & _
                    " - " & Date & " - " & CONSTCompName
    End With
    Dim objOutlook As Outlook.Application
    Dim objMail As MailItem

    Set objOutlook = New Outlook.Application '����objOutlookΪOutlookӦ�ó������
    Set objMail = objOutlook.CreateItem(olMailItem)  '����objMailΪһ���ʼ�����

ThisDocument.Content.Copy
'//need tobe amded

    With objMail
        '  .To = Recipient        '�ռ���
        .CC = CONSTMYEMAIL     '����
        .Subject = Subj        '����
        '  .body = body           '����
        .BodyFormat = olFormatRichText
        Set Editor = .GetInspector.WordEditor
        Editor.Content.Paste
        '  .Attachments.Add File  '����
        '  .send                  '����
        .Display
    End With

    Set objMail = Nothing
    Set objOutlook = Nothing

End Sub
