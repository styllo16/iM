Public Class Login

    Public g_pass As String
    Public g_user As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        g_pass = "123"
        g_user = "Admin"
        Dim pass As String
        Dim user As String

        pass = txtPass.Text.Trim
        user = txtUserName.Text.Trim
        If g_pass = pass And g_user = user Then

            Dim s As New settings
            SMS_sender.TabControl1.TabPages.Add(s)

            Dim f As New Sms
            SMS_sender.TabControl1.TabPages.Add(f)

            SMS_sender.lblUserName.Text = user

            txtPass.Clear()
        End If

    End Sub

    Private Sub Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If My.Settings.SMS_Account_Name = "" Then

            txtUserName.Text = "Trial Version"
            My.Settings.SMS_Account_Name = "Trial Version"
            My.Settings.Save()
        Else
            txtUserName.Text = My.Settings.SMS_Account_Name

        End If
        

    End Sub
    Sub LicenseRoutine()
        If Is_Demo_Trial = True Then Exit Sub

        If Apply_License(My.Settings.LicenseKey) = True Then
            ' MessageBox.Show("License Applied")

        Else
            My.Settings.LicenseKey = InputBox("Enter your License Key", "License Key")
            My.Settings.Save()

            If Apply_License(My.Settings.LicenseKey) = True Then
                MessageBox.Show("License Applied")

            Else
                MessageBox.Show("Invalid License, Application will End. Contact support@adroitbureau.com, 0303 3931293")
                End
            End If

        End If
    End Sub

    Private Sub Button1_Click_1(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        If txtUserName.Text = "Trial Version" Then

            Is_Demo_Trial = True

            Dim p1 As New Control_Panel
            SMS_sender.TabControl1.TabPages.Add(p1)

            '  SMS_sender.grpControlPanel.Visible = True
            SMS_sender.Button2.Enabled = True
            SMS_sender.Button3.Enabled = True
            SMS_sender.Button4.Enabled = True
            SMS_sender.Button5.Enabled = True

            Me.Close()
            GoTo skip

        End If




        My.Settings.SMS_Account_Name = txtUserName.Text
        My.Settings.Save()

        LicenseRoutine()

            ToolStripLabel1.Text = "Loading Intelligent Marketing."
            


           
            Dim p As New Form1
            SMS_sender.TabControl1.TabPages.Add(p)

            Dim s As New settings
            SMS_sender.TabControl1.TabPages.Add(s)



          

            ToolStripLabel1.Text = "Log In"


        SMS_sender.lblUserName.Text = My.Settings.SMS_Account_Name

            txtPass.Clear()

            SMS_sender.Button2.Enabled = True
            SMS_sender.Button3.Enabled = True
            SMS_sender.Button4.Enabled = True
            SMS_sender.Button5.Enabled = True
        ' SMS_sender.grpControlPanel.Visible = True


skip:
    End Sub

    Private Sub Panel3_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles Panel3.Paint

    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs)

    End Sub
End Class
