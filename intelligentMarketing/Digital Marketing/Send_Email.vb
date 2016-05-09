Imports System.Net.Mail
Module Send_Email

    Public Function Send_Email_With_Given_Parameters(ByVal Email_Account As String, ByVal Email_Pass As String, ByVal SMTP_Server_Port As String, ByVal SMTP_Server_Host As String, ByVal SMTP_Server_Enable_SSL As Boolean, ByVal Include_Attachment As Boolean, ByVal Attachment_Path As String, ByVal Email_From_Address As String, ByVal Email_To_Address As String, ByVal Email_Message As String, ByVal Email_Subject As String, ByVal SentListBoxDisplay As ListBox, ByVal NotSentListboxDisplay As ListBox) As String

        Try
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential(Email_Account, Email_Pass)
            Smtp_Server.Port = CInt(SMTP_Server_Port)
            Smtp_Server.EnableSsl = SMTP_Server_Enable_SSL 'True
            Smtp_Server.Host = SMTP_Server_Host 'smtp.yahoo.com"




            e_mail = New MailMessage()

            'Decide on whether to include attachment
            If Include_Attachment = True Then
                Dim attachment As System.Net.Mail.Attachment
                attachment = New System.Net.Mail.Attachment(Attachment_Path) '"C:\Users\TRIUMPH\Documents\doc2.docx")

                e_mail.Attachments.Add(attachment)

            End If

            e_mail.From = New MailAddress(Email_From_Address)
            e_mail.To.Add(Email_To_Address)
            e_mail.Subject = Email_Subject
            e_mail.IsBodyHtml = False
            e_mail.Body = Email_Message
            Smtp_Server.Send(e_mail)

            SentListBoxDisplay.Items.Add(Email_To_Address & " " & Email_Subject)
            Email_Successfully_Sent += 1
            Return "Mail Sent"

        Catch error_t As Exception
            NotSentListboxDisplay.Items.Add(Email_To_Address & " " & Email_Subject)
            Email_Unsuccessfully_Sent += 1
            Return error_t.ToString
        End Try

    End Function



End Module
