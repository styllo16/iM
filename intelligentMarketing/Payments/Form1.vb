Imports MPower_Payment_Integration
Imports MPowerPayments
Imports System.Math

Public Class Form1

    Dim setup = New MPowerSetup("d8928edd-ba6f-4d60-a99a-0a553a4f7da0 ", "live_public_RttZmsqGU5G4QIt3XNqXyc-AtHE", "live_private_F7JTsWGGch1KsKKS8Ku1haXZtH4", "57f72122af7db389f40e", "live")
    Dim store = New MPowerStore

    Dim Inv = New MPowerOnsiteInvoice(setup, store)
    Dim token As String
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click


        Make_Payment()
        'store.Name = "Adroit Store"


        ''   MPowerStore store = new MPowerStore {  Name = "Awesome Online Store",  Tagline = "This is my awesome tagline",  PhoneNumber = "030200001",  PostalAddress = "P. O. Box 10770 Accra North Ghana",  LogoUrl = "http://www.mylogourl.com/photo.png"};




        'Inv.AddItem("Purchase of Bulk SMS", 1, 10.99, 10.99)

        'Inv.SetTotalAmount(100.5)


        'If Inv.Create("0541235271") = True Then
        '    token = Inv.Token
        '    MessageBox.Show(token)
        '    MessageBox.Show(Inv.Status)
        '    MessageBox.Show(Inv.ResponseText)
        '    Button2.Enabled = True
        '    TextBox2.Enabled = True

        'Else


        '    MessageBox.Show(Inv.Status)
        '    MessageBox.Show(Inv.ResponseText)
        'End If




    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("http://mpowerpayments.com/")

    End Sub

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs)

        'Dim custtoken As String
        'custtoken = TextBox2.Text

        'If Inv.Charge(token, custtoken) = True Then
        '    MessageBox.Show(Inv.Status)
        '    MessageBox.Show(Inv.ResponseText)
        '    MessageBox.Show(Inv.GetReceiptUrl())
        '    MessageBox.Show(Inv.GetCustomerInfo("name"))
        '    MessageBox.Show(Inv.GetCustomerInfo("email"))
        'Else
        '    Console.WriteLine(Inv.Status)
        '    MessageBox.Show(Inv.ResponseText)
        'End If

    End Sub

    Sub Make_Payment()
        Dim token As String


        Dim setup = New MPowerSetup("d8928edd-ba6f-4d60-a99a-0a553a4f7da0 ", "live_public_RttZmsqGU5G4QIt3XNqXyc-AtHE", "live_private_F7JTsWGGch1KsKKS8Ku1haXZtH4", "57f72122af7db389f40e", "live")
        Dim store = New MPowerStore

        store.Name = "Adroit Store"

        Dim Inv = New MPowerOnsiteInvoice(setup, store)
        Dim total_amount As Double

        Try

            Inv.AddItem("Purchase of Bulk SMS", 1, CDbl(TextBox1.Text), CDbl(TextBox1.Text))
            Inv.AddTax("Tax", CDbl(TextBox1.Text) * (15 / 100))
            total_amount = CDbl((TextBox1.Text) * 1.15)
            Inv.SetTotalAmount(Math.Round(total_amount, 2))



            If Inv.Create(TextBox2.Text) = True Then
                token = Inv.Token
                '  MessageBox.Show(token)
                '  MessageBox.Show(Inv.Status)
                MessageBox.Show(Inv.ResponseText)
                '  Button2.Enabled = True
                '  TextBox2.Enabled = True

            Else


                '  MessageBox.Show(Inv.Status)
                MessageBox.Show(Inv.ResponseText)
            End If



            ' CHARGE
            Dim custtoken As String
            custtoken = InputBox("Please input the confirmation code sent to your email or phone", "Confirmation Code")

            If Inv.Charge(token, custtoken) = True Then
                ' MessageBox.Show(Inv.Status)
                MessageBox.Show(Inv.ResponseText)
                Process.Start(Inv.GetReceiptUrl())
                ' MessageBox.Show(Inv.GetCustomerInfo("name"))
                '  MessageBox.Show(Inv.GetCustomerInfo("email"))
            Else
                Console.WriteLine(Inv.Status)
                MessageBox.Show(Inv.ResponseText)
            End If



        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try




    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Make_Payment()

    End Sub
End Class
