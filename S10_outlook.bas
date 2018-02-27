Attribute VB_Name = "S10_outlook"
Sub sendMail()
    With ActiveDocument
        Subj = "(due:)【美国" & .Variables("CaseID_Client").Value & "负责人一稿】" & _
                   .Variables("CaseID_Client").Value & "; " & _
                   .Variables("CaseID_Official").Value & "; " & _
                   .Variables("CaseID_self").Value & "; " & _
                    " - " & Date & " - " & CONSTCompName
    End With
    Dim objOutlook As Outlook.Application
    Dim objMail As MailItem

    Set objOutlook = New Outlook.Application '创建objOutlook为Outlook应用程序对象
    Set objMail = objOutlook.CreateItem(olMailItem)  '创建objMail为一个邮件对象

ThisDocument.Content.Copy
'//need tobe amded

    With objMail
        '  .To = Recipient        '收件人
        .CC = CONSTMYEMAIL     '抄送
        .Subject = Subj        '标题
        '  .body = body           '正文
        .BodyFormat = olFormatRichText
        Set Editor = .GetInspector.WordEditor
        Editor.Content.Paste
        '  .Attachments.Add File  '附件
        '  .send                  '发送
        .Display
    End With

    Set objMail = Nothing
    Set objOutlook = Nothing

End Sub
