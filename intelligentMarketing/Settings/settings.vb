Public Class settings

    Private Sub ToolStripLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel1.Click

    End Sub

    Private Sub Panel2_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub settings_Click(sender As Object, e As System.EventArgs) Handles Me.Click
        '  Make_Controls_Visible()
    End Sub

    Private Sub settings_GotFocus(sender As Object, e As System.EventArgs) Handles Me.GotFocus
        '  Make_Controls_Visible()
    End Sub

    Private Sub settings_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'Check for Valid License
        LicenseRoutine()

        Load_MySettings_Variables()
        Load_Error_Log()

        Load_MySettings_Into_Form()

        Enforce_Demo_Restrictions()
        '  Make_Controls_Visible()
    End Sub

    Sub Enforce_Demo_Restrictions()
        If Is_Demo_Trial = True Then
            rdbSQL_SERVER.Enabled = False
            rdbPervasive.Enabled = False
            GroupBox3.Enabled = False
            GroupBox5.Enabled = False
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

    Sub Load_Error_Log()
        Dim i As Integer
        Dim c As Integer

        Do
            Try
                ListBox1.Items.Add(My.Settings.Error_log.Item(i))
                i += 1
            Catch ex As Exception
                Exit Do
            End Try

        Loop


       
    End Sub
    Sub Make_Controls_Visible()
        Panel2.Visible = True
        TabControl1.Visible = True
    End Sub

    Sub Load_MySettings_Into_Form()

        'New sms gateway Anaku
        My.Settings.SMS_Account_ProviderURL = "http://121.241.242.114:8080/sendsms?username="
        SMS_ProviderURL = "http://121.241.242.114:8080/sendsms?username="
        My.Settings.Save()

        txtSMSAccountName.Text = SMS_Account_Name
        txtSMSAccountPassword.Text = SMS_Account_Password
        txtSMSType.Text = SMS_Account_Type
        txtSMSProviderURL.Text = SMS_ProviderURL
        txtSenderID.Text = SMS_SenderID

        txtEmailAccount.Text = Email_Account
        txtEmailPass.Text = Email_Pass
        txtSMTPServerPort.Text = SMTP_server_Port
        txtSMTPServerHost.Text = SMTP_server_Host
        txtSMTPServerEnableSSL.Text = SMTP_server_Enable_SSL.ToString
        txtEmail_Source.Text = Email_Source
        txtCustomerEmailField.Text = Customer_Email_Field

        txtCustomer_Table_Name.Text = Customer_Table_Name
        txtCustomer_Name_Field.Text = Customer_Field_Name
        txtCustomer_Contact_Field1.Text = Customer_Contact_Field1
        txtCustomer_Contact_Field2.Text = Customer_Contact_Field2
        txtCustomerLocation_Field.Text = Customer_Location_Field
        txtCustomerBalance_Table.Text = Customer_Balance_Table_Name
        txtCustomer_Balance_Field.Text = Customer_Balance_Field_Name
        txtCustomer_Representative_Field.Text = Customer_Rep_Field_Name
        txtCustomer_Code_Field.Text = Customer_Code_Field
        txtCustomer_Sales_Archive.Text = Customer_Sales_Archive_Field


        rdbMSACCESS.Checked = IsACCESS
        rdbSQL_SERVER.Checked = IsSQLSERVER
        rdbPervasive.Checked = IsPervasive

        txtServerName.Text = SQLSERVER_NAME
        txtServerUserID.Text = SQLUSERID
        txtServerPassword.Text = SQLPassword
        txtMSACCESS_Filename.Text = MSACESSFILE
        txtMSACCESS_File_Path.Text = MSACCESSFILE_Path
        txtSQL_Initial_Catalog.Text = SQL_Initial_Catalog
        txtPSQL_Server_Name.Text = PSQLSERVER_NAME
        txtPSQL_Database.Text = PSQLDatabase_Name
        txtPSQL_Database_UserID.Text = PSQLDatabase_UserID
        txtPSQL_Database_Password.Text = PSQLDatabase_UserPassword


        'Perculiar for Pastel
        txtContactsFieldName.Text = Contacts_Field_Name_Pastel
        txtContactsTableName.Text = Contacts_Table_Name_Pastel


        chkIntegratedSecurity.Checked = IntegratedSecurity


        ' For the location intelligence module
        txtRegion_Field.Text = Region_Field
        txtDistrict_Field.Text = District_Field

        ' For the Business intelligence Module
        txtBI_Customer_Accounts_Table.Text = BI_Customer_Accounts_Table
        txtBI_Customer_Accounts_Name.Text = BI_Customer_Accounts_Name_Field
        txtBI_Customer_Accounts_Code.Text = BI_Customer_Accounts_Code_Field


        'Public Variables for the VoiceSMS
        txtVoice_Account_Name_Field.Text = Voice_SMS_Account_Name
        txtVoice_Account_Password.Text = Voice_SMS_Account_Password
        txtVoiceProviderURL.Text = VoiceSMS_ProviderURL


        'Pubic Variables for Custom Fields
        txtCustom_Field1.Text = Custom_Field1
        txtCustom_Field2.Text = Custom_Field2
        txtCustom_Field3.Text = Custom_Field3
        txtCustom_Field4.Text = Custom_Field4
        txtCustom_Field5.Text = Custom_Field5
        txtCustom_Field6.Text = Custom_Field6
        txtCustom_Field7.Text = Custom_Field7
        txtCustom_Field8.Text = Custom_Field8
        txtCustom_Field9.Text = Custom_Field9
        txtCustom_Field10.Text = Custom_Field10

        'Google Webforms
        txt_inputURL.Text = URL_InputForms
        txt_OutputURL.Text = URL_OutForms

    End Sub

    Sub Save_MySettings()
        My.Settings.SMS_Account_Name = txtSMSAccountName.Text
        My.Settings.SMS_Account_Password = txtSMSAccountPassword.Text
        My.Settings.SMS_Account_Type = txtSMSType.Text
        My.Settings.SMS_Account_ProviderURL = txtSMSProviderURL.Text
        My.Settings.SMS_SenderID = txtSenderID.Text

        My.Settings.Email_Account_Name = txtEmailAccount.Text
        My.Settings.Email_Account_Pass = txtEmailPass.Text
        My.Settings.SMTP_server_Port = txtSMTPServerPort.Text
        My.Settings.SMTP_server_Server_Host = txtSMTPServerHost.Text
        My.Settings.SMTP_server_Enable_SSL = txtSMTPServerEnableSSL.Text
        My.Settings.Email_Source = txtEmail_Source.Text
        My.Settings.Customer_Email_Field = txtCustomerEmailField.Text

        My.Settings.Customer_Table_Name = txtCustomer_Table_Name.Text
        My.Settings.Customer_Field_Name = txtCustomer_Name_Field.Text
        My.Settings.Customer_Contact_Field1 = txtCustomer_Contact_Field1.Text
        My.Settings.Customer_Contact_Field2 = txtCustomer_Contact_Field2.Text
        My.Settings.Customer_Location_Field = txtCustomerLocation_Field.Text
        My.Settings.Customer_Balance_Table_Name = txtCustomerBalance_Table.Text
        My.Settings.Customer_Balance_FieldName = txtCustomer_Balance_Field.Text
        My.Settings.Customer_Rep_Field_Name = txtCustomer_Representative_Field.Text
        My.Settings.Customer_Code_Field = txtCustomer_Code_Field.Text
        My.Settings.Customer_Sales_Archive = txtCustomer_Sales_Archive.Text

        My.Settings.IsACCESS = rdbMSACCESS.Checked
        My.Settings.IsSQLServer = rdbSQL_SERVER.Checked
        My.Settings.IS_PERVASIVE = rdbPervasive.Checked

        My.Settings.SQLServerName = txtServerName.Text
        My.Settings.SQLServerUserID = txtServerUserID.Text
        My.Settings.SQLServerUserPassword = txtServerPassword.Text
        My.Settings.MsACCESSFilename = txtMSACCESS_Filename.Text
        My.Settings.MsACCESSFilename_Path = txtMSACCESS_File_Path.Text
        My.Settings.SQL_InitialCatalog = txtSQL_Initial_Catalog.Text
        My.Settings.PSQLSERVERNAME = txtPSQL_Server_Name.Text
        My.Settings.PSQLDatabase_Name = txtPSQL_Database.Text
        My.Settings.PSQLDatabase_UserID = txtPSQL_Database_UserID.Text
        My.Settings.PSQLDatabase_UserPassword = txtPSQL_Database_Password.Text

        My.Settings.IntegratedSecurity = chkIntegratedSecurity.Checked


        'Perculiar for Pastel
        My.Settings.Contacts_Field_Name_Pastel = txtContactsFieldName.Text
        My.Settings.Contacts_Table_Name_Pastel = txtContactsTableName.Text


        ' For the location intelligence module
        My.Settings.Region_Field = txtRegion_Field.Text
        My.Settings.District_Field = txtDistrict_Field.Text

        'For the Business Intelligence Module
        My.Settings.BI_Customer_Accounts_Table = txtBI_Customer_Accounts_Table.Text
        My.Settings.BI_Customer_Accounts_Name_Field = txtBI_Customer_Accounts_Name.Text
        My.Settings.BI_Customer_Accounts_Code_Field = txtBI_Customer_Accounts_Code.Text


        'Public Variables for the VoiceSMS
        My.Settings.Voice_SMS_Account_Name = txtVoice_Account_Name_Field.Text
        My.Settings.Voice_SMS_Account_Password = txtVoice_Account_Password.Text
        My.Settings.Voice_SMS_Provider_URL = txtVoiceProviderURL.Text


        'Pubic Variables for Custom Fields
        My.Settings.Custom_Field1 = txtCustom_Field1.Text
        My.Settings.Custom_Field2 = txtCustom_Field2.Text
        My.Settings.Custom_Field3 = txtCustom_Field3.Text
        My.Settings.Custom_Field4 = txtCustom_Field4.Text
        My.Settings.Custom_Field5 = txtCustom_Field5.Text
        My.Settings.Custom_Field6 = txtCustom_Field6.Text
        My.Settings.Custom_Field7 = txtCustom_Field7.Text
        My.Settings.Custom_Field8 = txtCustom_Field8.Text
        My.Settings.Custom_Field9 = txtCustom_Field9.Text
        My.Settings.Custom_Field10 = txtCustom_Field10.Text

        'Google Webforms
        My.Settings.URL_Input_Webform = txt_inputURL.Text
        My.Settings.URL_Output_Webform = txt_OutputURL.Text

    End Sub

    Sub ResetTextboxes()


        txtSMSAccountName.Text = ""
        txtSMSAccountPassword.Text = ""
        txtSMSType.Text = ""
        txtSMSProviderURL.Text = ""

        txtEmailAccount.Text = ""
        txtEmailPass.Text = ""
        txtSMTPServerPort.Text = ""
        txtSMTPServerHost.Text = ""
        txtSMTPServerEnableSSL.Text = ""
        txtEmail_Source.Text = ""
        txtCustomerEmailField.Text = ""
        txtCustomer_Sales_Archive.Text = ""

        txtCustomer_Table_Name.Text = ""
        txtCustomer_Name_Field.Text = ""
        txtCustomer_Contact_Field1.Text = ""
        txtCustomer_Contact_Field2.Text = ""
        txtCustomerLocation_Field.Text = ""
        txtCustomerBalance_Table.Text = ""
        txtCustomer_Balance_Field.Text = ""
        txtCustomer_Representative_Field.Text = ""



        ' For the location intelligence module
        txtRegion_Field.Text = ""
        txtDistrict_Field.Text = ""

        'Public Variables for the VoiceSMS
        txtVoice_Account_Name_Field.Text = ""
        txtVoice_Account_Password.Text = ""
        txtVoiceProviderURL.Text = ""


        'Pubic Variables for Custom Fields
        txtCustom_Field1.Text = ""
        txtCustom_Field2.Text = ""
        txtCustom_Field3.Text = ""
        txtCustom_Field4.Text = ""
        txtCustom_Field5.Text = ""
        txtCustom_Field6.Text = ""
        txtCustom_Field7.Text = ""
        txtCustom_Field8.Text = ""
        txtCustom_Field9.Text = ""
        txtCustom_Field10.Text = ""

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Save_MySettings()
        My.Settings.Save()
        Load_MySettings_Variables()
        MessageBox.Show("Settings Saved", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)




    End Sub
    Sub SAVE_SETTINGS()
        Save_MySettings()
        My.Settings.Save()
        Load_MySettings_Variables()
        MessageBox.Show("Settings Saved", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information)



    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        ResetTextboxes()
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Make_Controls_Visible()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Dim path As String
        Dim filename As String
        Dim Filename_Including_Path As String

        OpenFileDialog1.ShowDialog()
        Filename_Including_Path = OpenFileDialog1.FileName
        filename = OpenFileDialog1.SafeFileName
        path = Replace(Filename_Including_Path, filename, "")
        filename = Replace(filename, ".accdb", "")

        If filename = "OpenFileDialog1" Or filename = "" Then GoTo skip
        txtMSACCESS_File_Path.Text = path
        txtMSACCESS_Filename.Text = filename

skip:
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton2.CheckedChanged

    End Sub

    Private Sub RadioButton2_Click(sender As Object, e As System.EventArgs) Handles RadioButton2.Click
        If RadioButton2.Checked = True Then
            txtSMTPServerHost.Text = "smtp.gmail.com"
            txtSMTPServerPort.Text = "587"
            txtSMTPServerEnableSSL.Text = "True"

        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton1.CheckedChanged

    End Sub

    Private Sub RadioButton1_Click(sender As Object, e As System.EventArgs) Handles RadioButton1.Click
        If RadioButton1.Checked = True Then
            txtSMTPServerHost.Text = "smtp.mail.yahoo.com"
            txtSMTPServerPort.Text = "587"
            txtSMTPServerEnableSSL.Text = "False"

        End If
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        SAVE_SETTINGS()
    End Sub

    Private Sub GroupBox3_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox3.Enter

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case ComboBox1.Text
            Case "SAGE PASTEL EVOLUTION"

                My.Settings.Customer_Table_Name = "Client"
                My.Settings.Customer_Field_Name = "Name"
                My.Settings.Customer_Contact_Field1 = "Telephone"
                My.Settings.Customer_Contact_Field2 = "Telephone2"
                My.Settings.Customer_Location_Field = "Physical1"
                My.Settings.Customer_Balance_Table_Name = "Client"
                My.Settings.Customer_Balance_FieldName = "dcbalance"
                My.Settings.Customer_Rep_Field_Name = "RepID"
                My.Settings.Customer_Code_Field = "Dclink"




                My.Settings.IsACCESS = False
                My.Settings.IsSQLServer = True
                My.Settings.IS_PERVASIVE = False

                My.Settings.SQLServerName = InputBox("Enter SQL Server Name")
                My.Settings.SQLServerUserID = "sa"
                My.Settings.SQLServerUserPassword = ""
                'My.Settings.MsACCESSFilename = txtMSACCESS_Filename.Text
                'My.Settings.MsACCESSFilename_Path = txtMSACCESS_File_Path.Text
                My.Settings.SQL_InitialCatalog = InputBox("Enter SQL Database Name")
                '  My.Settings.PSQLSERVERNAME = PSQLSERVER_NAME
                'My.Settings.PSQLDatabase_Name = PSQLDatabase_Name
                'My.Settings.PSQLDatabase_UserID = PSQLDatabase_UserID
                'My.Settings.PSQLDatabase_UserPassword = PSQLDatabase_UserPassword

                My.Settings.IntegratedSecurity = True



                '' For the location intelligence module
                My.Settings.Region_Field = "Physical3"
                My.Settings.District_Field = "Physical2"

                ''Public Variables for the VoiceSMS
                'My.Settings.Voice_SMS_Account_Name = txtVoice_Account_Name_Field.Text
                'My.Settings.Voice_SMS_Account_Password = txtVoice_Account_Password.Text
                'My.Settings.Voice_SMS_Provider_URL = txtVoiceProviderURL.Text


                'Pubic Variables for Custom Fields
                'My.Settings.Custom_Field1 = txtCustom_Field1.Text
                'My.Settings.Custom_Field2 = txtCustom_Field2.Text
                'My.Settings.Custom_Field3 = txtCustom_Field3.Text
                'My.Settings.Custom_Field4 = txtCustom_Field4.Text
                'My.Settings.Custom_Field5 = txtCustom_Field5.Text
                'My.Settings.Custom_Field6 = txtCustom_Field6.Text
                'My.Settings.Custom_Field7 = txtCustom_Field7.Text
                'My.Settings.Custom_Field8 = txtCustom_Field8.Text
                'My.Settings.Custom_Field9 = txtCustom_Field9.Text
                'My.Settings.Custom_Field10 = txtCustom_Field10.Text

                My.Settings.Save()
                Load_MySettings_Variables()
                Load_MySettings_Into_Form()
                ' Display_Current_Database()

            Case "FOCUS"

                My.Settings.Customer_Table_Name = "Dealers"
                My.Settings.Customer_Field_Name = "Company"
                My.Settings.Customer_Contact_Field1 = "Mobile"
                My.Settings.Customer_Contact_Field2 = "Phone"
                My.Settings.Customer_Location_Field = "Address1"
                My.Settings.Customer_Balance_Table_Name = "Dealeraccts"
                My.Settings.Customer_Balance_FieldName = "SL_Balance"
                My.Settings.Customer_Rep_Field_Name = "Representative"
                My.Settings.Customer_Code_Field = "Code"

                My.Settings.BI_Customer_Accounts_Code_Field = "Code"
                My.Settings.BI_Customer_Accounts_Name_Field = "Name"
                My.Settings.BI_Customer_Accounts_Table = "Dealeraccts"
                My.Settings.Customer_Sales_Archive = "year_turn"


                My.Settings.IsACCESS = False
                My.Settings.IsSQLServer = True
                My.Settings.IS_PERVASIVE = False

                My.Settings.SQLServerName = InputBox("Enter SQL Server Name")
                My.Settings.SQLServerUserID = "sa"
                My.Settings.SQLServerUserPassword = ""
                'My.Settings.MsACCESSFilename = txtMSACCESS_Filename.Text
                'My.Settings.MsACCESSFilename_Path = txtMSACCESS_File_Path.Text
                My.Settings.SQL_InitialCatalog = InputBox("Enter SQL Database Name")
                '  My.Settings.PSQLSERVERNAME = PSQLSERVER_NAME
                'My.Settings.PSQLDatabase_Name = PSQLDatabase_Name
                'My.Settings.PSQLDatabase_UserID = PSQLDatabase_UserID
                'My.Settings.PSQLDatabase_UserPassword = PSQLDatabase_UserPassword

                My.Settings.IntegratedSecurity = True



                '' For the location intelligence module
                My.Settings.Region_Field = "area"
                My.Settings.District_Field = "county"

                ''Public Variables for the VoiceSMS
                'My.Settings.Voice_SMS_Account_Name = txtVoice_Account_Name_Field.Text
                'My.Settings.Voice_SMS_Account_Password = txtVoice_Account_Password.Text
                'My.Settings.Voice_SMS_Provider_URL = txtVoiceProviderURL.Text


                'Pubic Variables for Custom Fields
                'My.Settings.Custom_Field1 = txtCustom_Field1.Text
                'My.Settings.Custom_Field2 = txtCustom_Field2.Text
                'My.Settings.Custom_Field3 = txtCustom_Field3.Text
                'My.Settings.Custom_Field4 = txtCustom_Field4.Text
                'My.Settings.Custom_Field5 = txtCustom_Field5.Text
                'My.Settings.Custom_Field6 = txtCustom_Field6.Text
                'My.Settings.Custom_Field7 = txtCustom_Field7.Text
                'My.Settings.Custom_Field8 = txtCustom_Field8.Text
                'My.Settings.Custom_Field9 = txtCustom_Field9.Text
                'My.Settings.Custom_Field10 = txtCustom_Field10.Text

                My.Settings.Save()
                ' Display_Current_Database()
                Load_MySettings_Variables()
                Load_MySettings_Into_Form()


            Case "IM Dummy"

                Dim DocumentsDir As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

                My.Settings.Customer_Table_Name = "Contacts"
                My.Settings.Customer_Field_Name = "Company"
                My.Settings.Customer_Contact_Field1 = "Telephone1"
                My.Settings.Customer_Contact_Field2 = "Telephone2"
                My.Settings.Customer_Location_Field = "Town"
                My.Settings.Customer_Balance_Table_Name = "Accounts"
                My.Settings.Customer_Balance_FieldName = "AccountBalance"
                My.Settings.Customer_Rep_Field_Name = "SalesRep"
                My.Settings.Customer_Code_Field = "ID"

                My.Settings.BI_Customer_Accounts_Code_Field = "ID"
                My.Settings.BI_Customer_Accounts_Name_Field = "Name"
                My.Settings.BI_Customer_Accounts_Table = "Accounts"
                My.Settings.Customer_Sales_Archive = "SalesArchive"


                My.Settings.IsACCESS = True
                My.Settings.IsSQLServer = False
                My.Settings.IS_PERVASIVE = False

                '    My.Settings.SQLServerName = "SERVER\SQL"
                '  My.Settings.SQLServerUserID = "sa"
                ' My.Settings.SQLServerUserPassword = ""
                My.Settings.MsACCESSFilename = "iM Dummy"
                My.Settings.MsACCESSFilename_Path = DocumentsDir
                '    My.Settings.SQL_InitialCatalog = "fop_data"
                '  My.Settings.PSQLSERVERNAME = PSQLSERVER_NAME
                'My.Settings.PSQLDatabase_Name = PSQLDatabase_Name
                'My.Settings.PSQLDatabase_UserID = PSQLDatabase_UserID
                'My.Settings.PSQLDatabase_UserPassword = PSQLDatabase_UserPassword

                ' My.Settings.IntegratedSecurity = False



                '' For the location intelligence module
                My.Settings.Region_Field = "Region"
                My.Settings.District_Field = "District"

                ''Public Variables for the VoiceSMS
                'My.Settings.Voice_SMS_Account_Name = txtVoice_Account_Name_Field.Text
                'My.Settings.Voice_SMS_Account_Password = txtVoice_Account_Password.Text
                'My.Settings.Voice_SMS_Provider_URL = txtVoiceProviderURL.Text


                'Pubic Variables for Custom Fields
                'My.Settings.Custom_Field1 = txtCustom_Field1.Text
                'My.Settings.Custom_Field2 = txtCustom_Field2.Text
                'My.Settings.Custom_Field3 = txtCustom_Field3.Text
                'My.Settings.Custom_Field4 = txtCustom_Field4.Text
                'My.Settings.Custom_Field5 = txtCustom_Field5.Text
                'My.Settings.Custom_Field6 = txtCustom_Field6.Text
                'My.Settings.Custom_Field7 = txtCustom_Field7.Text
                'My.Settings.Custom_Field8 = txtCustom_Field8.Text
                'My.Settings.Custom_Field9 = txtCustom_Field9.Text
                'My.Settings.Custom_Field10 = txtCustom_Field10.Text

                My.Settings.Save()
                ' Display_Current_Database()
                Load_MySettings_Variables()
                Load_MySettings_Into_Form()


            Case "PASTEL PARTNER"
                Dim DocumentsDir As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

                My.Settings.Customer_Table_Name = "CusomerMaster"
                My.Settings.Customer_Field_Name = "CustomerName"
                My.Settings.Contacts_Table_Name_Pastel = "HistoryHeader"
                My.Settings.Contacts_Field_Name_Pastel = "Telephone"
                My.Settings.Customer_Contact_Field1 = "Telephone"
                My.Settings.Customer_Contact_Field2 = "Telephone"
                My.Settings.Customer_Location_Field = "Town"
                My.Settings.Customer_Balance_Table_Name = "CusomerMaster"
                My.Settings.Customer_Balance_FieldName = "Balancethis12"
                My.Settings.Customer_Rep_Field_Name = "SalesRep"
                My.Settings.Customer_Code_Field = "ID"

                My.Settings.BI_Customer_Accounts_Code_Field = "ID"
                My.Settings.BI_Customer_Accounts_Name_Field = "Name"
                My.Settings.BI_Customer_Accounts_Table = "Accounts"



                My.Settings.IsACCESS = False
                My.Settings.IsSQLServer = False
                My.Settings.IS_PERVASIVE = True

                '    My.Settings.SQLServerName = "SERVER\SQL"
                '  My.Settings.SQLServerUserID = "sa"
                ' My.Settings.SQLServerUserPassword = ""
                'My.Settings.MsACCESSFilename = "iM Dummy"
                ' My.Settings.MsACCESSFilename_Path = DocumentsDir
                '    My.Settings.SQL_InitialCatalog = "fop_data"
                '  My.Settings.PSQLSERVERNAME = PSQLSERVER_NAME
                My.Settings.PSQLDatabase_Name = InputBox("PSQLDatabase Name")
                My.Settings.PSQLDatabase_UserID = InputBox("PSQLDatabase UserID")
                My.Settings.PSQLDatabase_UserPassword = InputBox("PSQLDatabase Password")

                ' My.Settings.IntegratedSecurity = False



                '' For the location intelligence module
                My.Settings.Region_Field = "Region"
                My.Settings.District_Field = "District"

                ''Public Variables for the VoiceSMS
                'My.Settings.Voice_SMS_Account_Name = txtVoice_Account_Name_Field.Text
                'My.Settings.Voice_SMS_Account_Password = txtVoice_Account_Password.Text
                'My.Settings.Voice_SMS_Provider_URL = txtVoiceProviderURL.Text


                'Pubic Variables for Custom Fields
                'My.Settings.Custom_Field1 = txtCustom_Field1.Text
                'My.Settings.Custom_Field2 = txtCustom_Field2.Text
                'My.Settings.Custom_Field3 = txtCustom_Field3.Text
                'My.Settings.Custom_Field4 = txtCustom_Field4.Text
                'My.Settings.Custom_Field5 = txtCustom_Field5.Text
                'My.Settings.Custom_Field6 = txtCustom_Field6.Text
                'My.Settings.Custom_Field7 = txtCustom_Field7.Text
                'My.Settings.Custom_Field8 = txtCustom_Field8.Text
                'My.Settings.Custom_Field9 = txtCustom_Field9.Text
                'My.Settings.Custom_Field10 = txtCustom_Field10.Text

                My.Settings.Save()
                ' Display_Current_Database()
                Load_MySettings_Variables()
                Load_MySettings_Into_Form()


            Case "Total Accounting"

                Dim DocumentsDir As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

                My.Settings.Customer_Table_Name = "Customers"
                My.Settings.Customer_Field_Name = "[Company Name]"
                My.Settings.Customer_Contact_Field1 = "[Telephone Number]"
                My.Settings.Customer_Contact_Field2 = "fax"
                My.Settings.Customer_Location_Field = "Town"
                My.Settings.Customer_Balance_Table_Name = "GL_Accounts"
                My.Settings.Customer_Balance_FieldName = "Account"
                My.Settings.Customer_Rep_Field_Name = "Rep_ID"
                My.Settings.Customer_Code_Field = "ID"

                My.Settings.BI_Customer_Accounts_Code_Field = "ID"
                My.Settings.BI_Customer_Accounts_Name_Field = "[Company Name]"
                My.Settings.BI_Customer_Accounts_Table = "Accounts"
                ' My.Settings.Customer_Sales_Archive = "SalesArchive"


                My.Settings.IsACCESS = True
                My.Settings.IsSQLServer = False
                My.Settings.IS_PERVASIVE = False

                '    My.Settings.SQLServerName = "SERVER\SQL"
                '  My.Settings.SQLServerUserID = "sa"
                ' My.Settings.SQLServerUserPassword = ""



                Dim path As String
                Dim filename As String
                Dim Filename_Including_Path As String

                OpenFileDialog1.ShowDialog()
                Filename_Including_Path = OpenFileDialog1.FileName
                filename = OpenFileDialog1.SafeFileName
                path = Replace(Filename_Including_Path, filename, "")
                filename = Replace(filename, ".accdb", "")

                If filename = "OpenFileDialog1" Or filename = "" Then GoTo skip
                My.Settings.MsACCESSFilename_Path = path
                My.Settings.MsACCESSFilename = filename
skip:


                '    My.Settings.SQL_InitialCatalog = "fop_data"
                '  My.Settings.PSQLSERVERNAME = PSQLSERVER_NAME
                'My.Settings.PSQLDatabase_Name = PSQLDatabase_Name
                'My.Settings.PSQLDatabase_UserID = PSQLDatabase_UserID
                'My.Settings.PSQLDatabase_UserPassword = PSQLDatabase_UserPassword

                ' My.Settings.IntegratedSecurity = False


                '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
                '' For the location intelligence module
                My.Settings.Region_Field = "Region"
                My.Settings.District_Field = "District"
                '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&77

                ''Public Variables for the VoiceSMS
                'My.Settings.Voice_SMS_Account_Name = txtVoice_Account_Name_Field.Text
                'My.Settings.Voice_SMS_Account_Password = txtVoice_Account_Password.Text
                'My.Settings.Voice_SMS_Provider_URL = txtVoiceProviderURL.Text


                'Pubic Variables for Custom Fields
                'My.Settings.Custom_Field1 = txtCustom_Field1.Text
                'My.Settings.Custom_Field2 = txtCustom_Field2.Text
                'My.Settings.Custom_Field3 = txtCustom_Field3.Text
                'My.Settings.Custom_Field4 = txtCustom_Field4.Text
                'My.Settings.Custom_Field5 = txtCustom_Field5.Text
                'My.Settings.Custom_Field6 = txtCustom_Field6.Text
                'My.Settings.Custom_Field7 = txtCustom_Field7.Text
                'My.Settings.Custom_Field8 = txtCustom_Field8.Text
                'My.Settings.Custom_Field9 = txtCustom_Field9.Text
                'My.Settings.Custom_Field10 = txtCustom_Field10.Text

                My.Settings.Save()
                ' Display_Current_Database()
                Load_MySettings_Variables()
                Load_MySettings_Into_Form()


            Case "Focus Prospects"

                Dim DocumentsDir As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

                My.Settings.Customer_Table_Name = "im_data_swl"
                My.Settings.Customer_Field_Name = "School"
                My.Settings.Customer_Contact_Field1 = "Telephone"
                My.Settings.Customer_Contact_Field2 = "Telephone2"
                My.Settings.Customer_Location_Field = "Location"
                'My.Settings.Customer_Balance_Table_Name = "GL_Accounts"
                ' My.Settings.Customer_Balance_FieldName = "Account"
                My.Settings.Customer_Rep_Field_Name = "Gender"
                My.Settings.Customer_Code_Field = "ID2"

                'My.Settings.BI_Customer_Accounts_Code_Field = "ID"
                'My.Settings.BI_Customer_Accounts_Name_Field = "[Company Name]"
                'My.Settings.BI_Customer_Accounts_Table = "Accounts"
                ' My.Settings.Customer_Sales_Archive = "SalesArchive"


                My.Settings.IsACCESS = True
                My.Settings.IsSQLServer = False
                My.Settings.IS_PERVASIVE = False

                '    My.Settings.SQLServerName = "SERVER\SQL"
                '  My.Settings.SQLServerUserID = "sa"
                ' My.Settings.SQLServerUserPassword = ""



                Dim path As String
                Dim filename As String
                Dim Filename_Including_Path As String

                OpenFileDialog1.ShowDialog()
                Filename_Including_Path = OpenFileDialog1.FileName
                filename = OpenFileDialog1.SafeFileName
                path = Replace(Filename_Including_Path, filename, "")
                filename = Replace(filename, ".accdb", "")

                If filename = "OpenFileDialog1" Or filename = "" Then GoTo skip2
                My.Settings.MsACCESSFilename_Path = path
                My.Settings.MsACCESSFilename = filename
skip2:


                '    My.Settings.SQL_InitialCatalog = "fop_data"
                '  My.Settings.PSQLSERVERNAME = PSQLSERVER_NAME
                'My.Settings.PSQLDatabase_Name = PSQLDatabase_Name
                'My.Settings.PSQLDatabase_UserID = PSQLDatabase_UserID
                'My.Settings.PSQLDatabase_UserPassword = PSQLDatabase_UserPassword

                ' My.Settings.IntegratedSecurity = False


                '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
                '' For the location intelligence module
                My.Settings.Region_Field = "Region"
                My.Settings.District_Field = "District"
                '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&77

                ''Public Variables for the VoiceSMS
                'My.Settings.Voice_SMS_Account_Name = txtVoice_Account_Name_Field.Text
                'My.Settings.Voice_SMS_Account_Password = txtVoice_Account_Password.Text
                'My.Settings.Voice_SMS_Provider_URL = txtVoiceProviderURL.Text


                'Pubic Variables for Custom Fields
                'My.Settings.Custom_Field1 = txtCustom_Field1.Text
                'My.Settings.Custom_Field2 = txtCustom_Field2.Text
                'My.Settings.Custom_Field3 = txtCustom_Field3.Text
                'My.Settings.Custom_Field4 = txtCustom_Field4.Text
                'My.Settings.Custom_Field5 = txtCustom_Field5.Text
                'My.Settings.Custom_Field6 = txtCustom_Field6.Text
                'My.Settings.Custom_Field7 = txtCustom_Field7.Text
                'My.Settings.Custom_Field8 = txtCustom_Field8.Text
                'My.Settings.Custom_Field9 = txtCustom_Field9.Text
                'My.Settings.Custom_Field10 = txtCustom_Field10.Text

                My.Settings.Save()
                ' Display_Current_Database()
                Load_MySettings_Variables()
                Load_MySettings_Into_Form()







        End Select
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        SEND_Authentication_sms_to_Adroit()
        ' MessageBox.Show(My.Settings.iM_Authentication)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        My.Settings.iM_Authentication = False

    End Sub

    Private Sub rdbSQL_SERVER_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdbSQL_SERVER.CheckedChanged
        
    End Sub

    Private Sub rdbSQL_SERVER_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdbSQL_SERVER.Click
        'If rdbSQL_SERVER.Checked = True Then
        '    TabControl3.TabPages.Item(1).Select()
        '    TabControl3.TabPages.Item(1).Focus()
        'End If
    End Sub

    Private Sub rdbSQL_SERVER_ClientSizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdbSQL_SERVER.ClientSizeChanged
   
    End Sub

    Private Sub Button5_Click_1(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Dim f As New Locations_Data_Validation
        f.Show()
    End Sub
End Class