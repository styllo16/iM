Public Class Control_Panel
    Dim smsV, bi, liv, dva, rev As Boolean
    Private Sub btnDigitalMarketing_Click(sender As System.Object, e As System.EventArgs) Handles btnDigitalMarketing.Click

        If smsV = False Then

            SMS_sender.Cursor = Cursors.AppStarting
            Dim sms As New Sms
            SMS_sender.TabControl1.TabPages.Add(Sms)
            SMS_sender.Cursor = Cursors.Default

            smsV = True

        Else
            SMS_sender.Focus_Sms()

        End If
    End Sub

    Private Sub Control_Panel_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnBusinessIntelligence_Click(sender As System.Object, e As System.EventArgs) Handles btnBusinessIntelligence.Click
        If bi = False Then
            SMS_sender.Cursor = Cursors.AppStarting
            SMS_sender.TabControl1.TabPages.Add(SWL_Customer_Dashboard)
            SMS_sender.Cursor = Cursors.Default
            bi = True
        End If
    End Sub

    Private Sub btnLocationIntelligence_Click(sender As System.Object, e As System.EventArgs) Handles btnLocationIntelligence.Click
        If liv = False Then
            SMS_sender.Cursor = Cursors.AppStarting
            SMS_sender.TabControl1.TabPages.Add(LI)
            SMS_sender.Cursor = Cursors.Default
            'Dim SWL_Dashboard As New SWL_Customer_Dashboard
            'SMS_sender.TabControl1.TabPages.Add(SWL_Dashboard)
            liv = True
        End If
    End Sub

    Private Sub btnRevenueCollection_Click(sender As System.Object, e As System.EventArgs) Handles btnRevenueCollection.Click

        If smsV = False Then

            SMS_sender.Cursor = Cursors.AppStarting
            Dim sms As New Sms
            SMS_sender.TabControl1.TabPages.Add(sms)
            SMS_sender.Cursor = Cursors.Default
            sms.Revenue_Collection()

            smsV = True

        Else
            SMS_sender.Focus_Sms()

        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            SMS_sender.TabControl1.TabPages.Add(settings)
            SMS_sender.Cursor = Cursors.Default
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ToolStripLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class