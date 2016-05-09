Public Class SMS_sender

    Dim smsV, bi, liv, dva, rev As Boolean
    Public Sub Focus_Sms()
        Sms.Focus()
    End Sub

    Private Sub SMS_sender_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Application.EnableVisualStyles()

        'TabControl1.BackColor = Color.Green
        'TabControl1.BackLowColor = Color.Green
        'TabControl1.BorderColor = Color.Green
        'TabControl1.BackHighColor = Color.Green
        ''   TabControl1.ForeColor = Color.Green
        ''TabControl1.TabBackHighColor = Color.Green
        'TabControl1.TabBackLowColor = Color.Green
        'TabControl1.ControlButtonBorderColor = Color.Green
        ''TabControl1.TabBackLowColor = Color.Green
        ''TabControl1.TabBackHighColor = Color.Green

        Dim f As New Login

        TabControl1.TabPages.Add(f)

        Display_SMS_Account_Details()

        'Set Default Database Path
        Dim dir As String
        dir = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

        If My.Settings.MsACCESSFilename = "iM Dummy" Then My.Settings.MsACCESSFilename_Path = dir
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim i As String

        i = TabControl1.TabPages.SelectedIndex
        MsgBox(i.ToString)

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim S As New Pervasive_Connection
        S.Show()

    End Sub

    

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint
        'MessageBox.Show(My.Settings.iM_Authentication)
    End Sub

   

   
    Private Sub pc_Sms_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub pc_BI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'Dim sms_form As New Sms
        'SMS_sender.TabControl1.TabPages.Add(sms_form)
    End Sub

    Private Sub pc_LI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)




        'Dim sms_form As New Sms
        'SMS_sender.TabControl1.TabPages.Add(sms_form)
    End Sub

    Private Sub pc_dataValidation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


        'Dim sms_form As New Sms
        'SMS_sender.TabControl1.TabPages.Add(sms_form)

    End Sub

    Private Sub pc_Revenue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Me.Cursor = Cursors.AppStarting
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'Dim SWL_LocationIntelligence As New LI
        'SMS_sender.TabControl1.TabPages.Add(SWL_LocationIntelligence)

        'Dim SWL_Dashboard As New SWL_Customer_Dashboard
        'SMS_sender.TabControl1.TabPages.Add(SWL_Dashboard)

        If smsV = False Then

            Me.Cursor = Cursors.AppStarting
            Me.TabControl1.TabPages.Add(Sms)
            Me.Cursor = Cursors.Default

            smsV = True

        Else
            Sms.Focus()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        'Dim SWL_LocationIntelligence As New LI
        'SMS_sender.TabControl1.TabPages.Add(SWL_LocationIntelligence)
        If bi = False Then
            Me.Cursor = Cursors.AppStarting
            Me.TabControl1.TabPages.Add(SWL_Customer_Dashboard)
            Me.Cursor = Cursors.Default
            bi = True
        End If


    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If liv = False Then
            Me.Cursor = Cursors.AppStarting
            Me.TabControl1.TabPages.Add(LI)
            Me.Cursor = Cursors.Default
            'Dim SWL_Dashboard As New SWL_Customer_Dashboard
            'SMS_sender.TabControl1.TabPages.Add(SWL_Dashboard)
            liv = True
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If dva = False Then
            Me.Cursor = Cursors.AppStarting
            form_is_lunched = True
            Me.TabControl1.TabPages.Add(Locations_Data_Validation)
            Me.Cursor = Cursors.Default
            'Dim SWL_Dashboard As New SWL_Customer_Dashboard
            'SMS_sender.TabControl1.TabPages.Add(SWL_Dashboard)

            dva = True
        End If
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs)
        Dim f As New Login

        TabControl1.TabPages.Add(f)
    End Sub

    Private Sub Button6_Click_1(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        Dim dir1 As String
        dir1 = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

        Try
            System.Diagnostics.Process.Start(dir1 & "\HOW TO USE INTELLIGENT MARKETING.chm")
        Catch ex As Exception
            MessageBox.Show("Help File not found", "File not found", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ' LogError("Error", "Help File not found")
        End Try

    End Sub
End Class