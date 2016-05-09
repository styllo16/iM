Imports System.Data
Imports Pervasive.Data.SqlClient
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Web 'add system.web as reference for the url encoding class
Imports System.Net

Module SUPREME
    'SMS Public Variables. These are always live and dynamic
    Public SMS_Account_Name As String
    Public SMS_Account_Password As String
    Public SMS_Account_Type As String
    Public SMS_ProviderURL As String
    Public SMS_SenderID As String

    'Email Public Variables
    Public Email_Account As String
    Public Email_Pass As String
    Public SMTP_server_Port As String
    Public SMTP_server_Host As String
    Public SMTP_server_Enable_SSL As Boolean
    Public Email_Source As String
    Public Customer_Email_Field As String




    ' Public Data Connection Variables
    Public Customer_Table_Name As String
    Public Customer_Field_Name As String
    Public Customer_Contact_Field1 As String
    Public Customer_Contact_Field2 As String
    Public Customer_Location_Field As String
    Public Customer_Balance_Table_Name As String
    Public Customer_Balance_Field_Name As String
    Public Customer_Rep_Field_Name As String
    Public Customer_Code_Field As String
    Public Customer_Sales_Archive_Field As String

    Public IsACCESS As Boolean
    Public IsSQLSERVER As Boolean
    Public IsPervasive As Boolean

    Public SQLSERVER_NAME As String
    Public SQLUSERID As String
    Public SQLPassword As String
    Public MSACESSFILE As String
    Public MSACCESSFILE_Path As String
    Public SQL_Initial_Catalog As String
    Public IntegratedSecurity As Boolean
    Public PSQLSERVER_NAME As String
    Public PSQLDatabase_Name As String
    Public PSQLDatabase_UserID As String
    Public PSQLDatabase_UserPassword As string

    'Perculiar for Pastel Partner
    Public Contacts_Table_Name_Pastel As String
    Public Contacts_Field_Name_Pastel As String


    ' For the location intelligence module
    Public Selected_Region_From_Attribute_Table As String
    Public Selected_Region_From_Attribute_Table_focus_specific As String
    Public Region_Field As String
    Public District_Field As String

    'For the Business Intelligence Module
    Public BI_Customer_Accounts_Table As String
    Public BI_Customer_Accounts_Name_Field As String
    Public BI_Customer_Accounts_Code_Field As String


    'Public Variables for the VoiceSMS
    Public Voice_SMS_Account_Name As String
    Public Voice_SMS_Account_Password As String
    Public Voice_Audio_URL As String
    Public VoiceSMS_ProviderURL As String

    'Pubic Variables for Custom Fields
    Public Custom_Field1 As String
    Public Custom_Field2 As String
    Public Custom_Field3 As String
    Public Custom_Field4 As String
    Public Custom_Field5 As String
    Public Custom_Field6 As String
    Public Custom_Field7 As String
    Public Custom_Field8 As String
    Public Custom_Field9 As String
    Public Custom_Field10 As String



    'Public SMS and Email Sending Status (These will have to be reset after every routine)
    Public SMS_Successufully_Sent As Integer
    Public Email_Successfully_Sent As Integer
    Public SMS_Unsuccessfully_Sent As Integer
    Public Email_Unsuccessfully_Sent As Integer

    

    'Other Global Variables
    Public Send_List_Title As String
    Const SMS_UNIT_COST_AS_APPLIED_BY_ADROIT As Double = 3.5
    Public SMS_Current_Account_Balance As String
    Public Customer_Array(5000, 5) As String
    Public Selected_District_for_Drill_Down As String
    Public Selected_District_Captial_for_Drill_Down As String
    Public InterModule_Transfer As Boolean = False
    Public Authentication_Required_Message_Shown = False

    'Google Forms
    Public URL_InputForms As String
    Public URL_OutForms As String

    'Trial Version
    Public Is_Demo_Trial As Boolean
    Public Const Contact_Number_Restriction As Integer = 50



    Sub Set_My_Settings_to_Focus_Settings()


        My.Settings.Customer_Table_Name = "Dealers"
        My.Settings.Customer_Field_Name = "Company"
        My.Settings.Customer_Contact_Field1 = "Mobile"
        My.Settings.Customer_Contact_Field2 = "Phone"
        My.Settings.Customer_Location_Field = "Address1"
        My.Settings.Customer_Balance_Table_Name = "Dealeraccts"
        My.Settings.Customer_Balance_FieldName = "SL_Balance"
        My.Settings.Customer_Rep_Field_Name = "Representative"
        My.Settings.Customer_Code_Field = "Code"

        My.Settings.IsACCESS = False
        My.Settings.IsSQLServer = True
        My.Settings.IS_PERVASIVE = False

        My.Settings.SQLServerName = "SERVER\SQL"
        My.Settings.SQLServerUserID = "sa"
        My.Settings.SQLServerUserPassword = ""
        'My.Settings.MsACCESSFilename = txtMSACCESS_Filename.Text
        'My.Settings.MsACCESSFilename_Path = txtMSACCESS_File_Path.Text
        My.Settings.SQL_InitialCatalog = "fop_data"
        '  My.Settings.PSQLSERVERNAME = PSQLSERVER_NAME
        'My.Settings.PSQLDatabase_Name = PSQLDatabase_Name
        'My.Settings.PSQLDatabase_UserID = PSQLDatabase_UserID
        'My.Settings.PSQLDatabase_UserPassword = PSQLDatabase_UserPassword

        My.Settings.IntegratedSecurity = False




        '' For the location intelligence module
        'My.Settings.Region_Field = txtRegion_Field.Text
        'My.Settings.District_Field = txtDistrict_Field.Text

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
    End Sub
    Public Function Is_Demo_Restriction_Exceeded()
        If Is_Demo_Trial = False Then
            Return False
        End If

        Dim Number_of_Contacts As Integer

        Number_of_Contacts = Get_Number_of_contacts()

        If Number_of_Contacts > Contact_Number_Restriction Then
            MessageBox.Show("Maximum Number of " & Contact_Number_Restriction & " Allowed Contacts Exceeded for Trial Version. Contact sales@adroitbureau.com", "Demo Version Exceeded", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return True
        End If
        Return False

    End Function
    Public Function Apply_License(ByVal strlicense As String) As Boolean

        Try
            Dim keydecoder As New KeyDec
            'Dim set1 As New SettingObject
            'set1.Connection = Total_Accounting_Database_Connection_String

            ' keydecoder.Users
            keydecoder.Decode(strlicense)

            If Not keydecoder.isValidKey Then

                MessageBox.Show("Your License is invalid, Contact support@adroitbureau.com, 0303 3931293")
                Return False

            End If


            If keydecoder.isExpired Then
                MessageBox.Show("Your License has Expired, Contact support@adroitbureau.com, 0303 3931293")
                Return False
            End If


            If keydecoder.isValidKey And Not keydecoder.isExpired Then
                My.Settings.SMS_Account_Name = keydecoder.UserName
                Return True
            End If


        Catch ex As Exception
            Return False
        End Try




    End Function


    Sub Set_My_Settings_to_Pastel_Evolution_Settings()


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
        '   Load_MySettings_Into_Form()
        ' Display_Current_Database()
    End Sub


    Public Sub Reset_SMS_Sending_Status_Variables()
        SMS_Successufully_Sent = 0
        SMS_Unsuccessfully_Sent = 0
    End Sub
    Public Sub Reset_Email_Sending_Status_Variables()
        Email_Successfully_Sent = 0
        Email_Unsuccessfully_Sent = 0
    End Sub
    Public Sub Blank_Customer_Array()
        Dim i As Integer
        Dim i2 As Integer

        For i = 1 To 5000
            For i2 = 1 To 5
                Customer_Array(i - 1, i2 - 1) = ""
            Next

        Next

    End Sub

    Sub Execute_Non_Query(ByVal query As String, ByVal connection_string As String)


        If IsACCESS Then
            Using connection As New OleDbConnection(connection_string)

                Dim command As OleDbCommand = connection.CreateCommand()

                command.CommandText = query

                Try

                    connection.Open()
                    command.ExecuteNonQuery()
                    connection.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

            End Using
        End If
       

        If IsSQLSERVER Then
            Using connection As New SqlConnection(connection_string)

                Dim command As SqlCommand = connection.CreateCommand()

                command.CommandText = query

                Try

                    connection.Open()
                    command.ExecuteNonQuery()
                    connection.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

            End Using
        End If

        'If IsPervasive Then

        '    Using connection As New PsqlConnection(connection_string)

        '        Dim command As PsqlCommand = connection.CreateCommand()

        '        command.CommandText = query

        '        Try

        '            connection.Open()
        '            command.ExecuteNonQuery()
        '            connection.Close()

        '        Catch ex As Exception
        '            MessageBox.Show(ex.Message)
        '        End Try

        '    End Using
        'End If

    End Sub

    Public Sub Intelligence_To_SMS(ByRef batchbox As DataGridView)
        batchbox.Rows.Clear()
        ' Sms.Button1.Text = "hid"

        Dim i As Integer
        Dim telephone As String
        Dim CompanyName As String


        For i = 1 To 5000
            If Customer_Array(i - 1, 0) = "" Then Exit For
            batchbox.Rows.Add(1)


            CompanyName = Fetch_This_Column("Company", Customer_Array(i - 1, 0), "Dealers")
            batchbox.Item(0, i - 1).Value = CompanyName

            telephone = Fetch_This_Column(Customer_Contact_Field1, Customer_Array(i - 1, 0), "Dealers")
            If telephone <> "" Then
                batchbox.Item(1, i - 1).Value = telephone

            Else
                batchbox.Item(1, i - 1).Value = Fetch_This_Column(Customer_Contact_Field2, Customer_Array(i - 1, 0), "Dealers")
            End If


            batchbox.Item(3, i - 1).Value = True

        Next
    End Sub

    Public Function Fetch_This_Column(ByVal Target As String, ByVal Code As String, ByVal Table As String) As String


        If Not IsACCESS Then

            Dim queryString As String = _
                 "SELECT " & Target & " from " & Table & " where " & Customer_Code_Field & "= " & Code & ""

            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()


                        Return dataReader(0).ToString
                    Loop
skip:
                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using
        Else

            Dim queryString As String = _
                  "SELECT " & Target & " from " & Table & " where " & Customer_Code_Field & "= " & Code & ""

            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()


                        Return dataReader(0).ToString
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using
        End If


    End Function

    Function AR_Account_Listing_Value_For_This_Customer(ByVal Id As String) As String
        Dim queryString As String = _
      "SELECT " & Customer_Balance_Field_Name & " from " & Customer_Balance_Table_Name & " where " & Customer_Code_Field & " = '" & Id & "'"


        If IsACCESS Then
            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()

                        Return dataReader(0).ToString
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    'My.Settings.Error_log.Add(ex.Message)
                End Try

                connection.Close()
            End Using
        End If

        If IsSQLSERVER Then
            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()

                        Return dataReader(0).ToString
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    'My.Settings.Error_log.Add(ex.Message)
                End Try

                connection.Close()
            End Using
        End If


        'If IsPervasive Then
        '    Using connection As New PsqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
        '        Dim command As PsqlCommand = connection.CreateCommand()
        '        command.CommandText = queryString
        '        Try
        '            connection.Open()
        '            Dim dataReader As PsqlDataReader = _
        '             command.ExecuteReader()
        '            Do While dataReader.Read()

        '                Return dataReader(0).ToString
        '            Loop

        '            dataReader.Close()

        '        Catch ex As Exception
        '            'My.Settings.Error_log.Add(ex.Message)
        '        End Try

        '        connection.Close()
        '    End Using
        'End If




    End Function

    Function Fetch_Field_Value_From_Specified_Table_Given_ID(ByVal Id As String, ByVal FieldName As String, ByVal tableName As String) As String
      


        If IsACCESS Then
            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = "SELECT " & FieldName & " from " & tableName & " where " & Customer_Code_Field & " = " & Id
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()

                        Return dataReader(0).ToString
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using
        End If

        If IsSQLSERVER Then
            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = "SELECT " & FieldName & " from " & tableName & " where " & Customer_Code_Field & " = '" & Id & "'"
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()

                        Return dataReader(0).ToString
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using
        End If


        'If IsPervasive Then
        '    Using connection As New PsqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
        '        Dim command As PsqlCommand = connection.CreateCommand()
        '        command.CommandText = "SELECT " & FieldName & " from " & tableName & " where " & Customer_Code_Field & " = '" & Id & "'"
        '        Try
        '            connection.Open()
        '            Dim dataReader As PsqlDataReader = _
        '             command.ExecuteReader()
        '            Do While dataReader.Read()

        '                Return dataReader(0).ToString
        '            Loop

        '            dataReader.Close()

        '        Catch ex As Exception
        '            MessageBox.Show("Connection failure")
        '        End Try

        '        connection.Close()
        '    End Using
        'End If




    End Function





    Public Sub Display_SMS_Account_Details()
        'lblAccountName.Text = SMS_Account_Name
        SMS_Current_Account_Balance = Check_SMS_Account_Credits(SMS_Account_Name, SMS_Account_Password)

    End Sub
    Public Sub Fill_Combo_With_Towns(ByRef Combo As ToolStripComboBox, ByVal CustomerTableName As String, ByVal TownFieldName As String)
        Dim current_town As String
        Combo.Items.Clear()

        If IsSQLSERVER Then
            Dim queryString As String = _
                 "SELECT distinct " & TownFieldName & " from " & CustomerTableName

            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        current_town = dataReader(0).ToString

                        If Trim(current_town) <> "" Then
                            Combo.Items.Add(current_town)
                        End If

                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    '   MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using

        End If


        If IsACCESS Then
            Dim queryString As String = _
                "SELECT distinct " & TownFieldName & " from " & CustomerTableName

            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        current_town = dataReader(0).ToString

                        If Trim(current_town) <> "" Then
                            Combo.Items.Add(current_town)
                        End If

                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    ' MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using


        End If


        'If IsPervasive Then

        '    Dim queryString As String = _
        '       "SELECT distinct " & TownFieldName & " from " & CustomerTableName

        '    Using connection As New psqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
        '        Dim command As PsqlCommand = connection.CreateCommand()
        '        command.CommandText = queryString
        '        Try
        '            connection.Open()
        '            Dim dataReader As PsqlDataReader = _
        '             command.ExecuteReader()
        '            Do While dataReader.Read()
        '                current_town = dataReader(0).ToString

        '                If Trim(current_town) <> "" Then
        '                    Combo.Items.Add(current_town)
        '                End If

        '            Loop

        '            dataReader.Close()

        '        Catch ex As Exception
        '            ' MessageBox.Show(ex.Message)
        '        End Try

        '        connection.Close()
        '    End Using



        'End If

    End Sub

    Public Sub Fill_Combo_With_Data(ByRef Combo As ToolStripComboBox, ByVal CustomerTableName As String, ByVal TownFieldName As String)
        Dim current_town As String
        Combo.Items.Clear()

        If IsSQLSERVER Then
            Dim queryString As String = _
                 "SELECT distinct " & TownFieldName & " from " & CustomerTableName

            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        current_town = dataReader(0).ToString

                        If Trim(current_town) <> "" Then
                            Combo.Items.Add(current_town)
                        End If

                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    '   MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using

        End If


        If IsACCESS Then
            Dim queryString As String = _
                "SELECT distinct " & TownFieldName & " from " & CustomerTableName

            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        current_town = dataReader(0).ToString

                        If Trim(current_town) <> "" Then
                            Combo.Items.Add(current_town)
                        End If

                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    ' MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using


        End If


        'If IsPervasive Then

        '    Dim queryString As String = _
        '       "SELECT distinct " & TownFieldName & " from " & CustomerTableName

        '    Using connection As New PsqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
        '        Dim command As PsqlCommand = connection.CreateCommand()
        '        command.CommandText = queryString
        '        Try
        '            connection.Open()
        '            Dim dataReader As PsqlDataReader = _
        '             command.ExecuteReader()
        '            Do While dataReader.Read()
        '                current_town = dataReader(0).ToString

        '                If Trim(current_town) <> "" Then
        '                    Combo.Items.Add(current_town)
        '                End If

        '            Loop

        '            dataReader.Close()

        '        Catch ex As Exception
        '            ' MessageBox.Show(ex.Message)
        '        End Try

        '        connection.Close()
        '    End Using



        'End If

    End Sub
    Public Sub Fill_Combo_With_Reps(ByRef Combo As ToolStripComboBox)
        Dim current_town As String
        Combo.Items.Clear()

        If IsSQLSERVER Then
            Dim queryString As String = _
                 "SELECT distinct " & Customer_Rep_Field_Name & " from " & Customer_Table_Name

            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        current_town = dataReader(0).ToString

                        If Trim(current_town) <> "" Then
                            Combo.Items.Add(current_town)
                        End If

                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    ' MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using

        Else

            Dim queryString As String = _
                 "SELECT distinct " & Customer_Rep_Field_Name & " from " & Customer_Table_Name

            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        current_town = dataReader(0).ToString

                        If Trim(current_town) <> "" Then
                            Combo.Items.Add(current_town)
                        End If

                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    '  MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using


        End If

    End Sub
    Public Sub Fill_Combo_With_Customers(ByRef Combo As ToolStripComboBox)
        Dim current_town As String
        Combo.Items.Clear()


        If IsSQLSERVER Then
            Dim queryString As String = _
                 "SELECT " & Customer_Field_Name & " from " & Customer_Table_Name

            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        Combo.Items.Add(dataReader(0).ToString)


                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    ' MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using

        Else

            Dim queryString As String = _
                 "SELECT " & Customer_Field_Name & " from " & Customer_Table_Name

            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        current_town = dataReader(0).ToString

                        If Trim(current_town) <> "" Then
                            Combo.Items.Add(current_town)
                        End If

                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    ' MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using


        End If

    End Sub
    Function MyCDBL(ByVal txt As String) As Double
        If Trim(txt) = "" Then
            Return 0

        Else
            Return CDbl(txt)
        End If


    End Function


    Public Function Check_SMS_Account_Credits(ByVal Account As String, ByVal Password As String) As String
        Dim strCreditBalance As String
        Dim strCreditBalance_Extracted As String
        Dim dblCreditBalance As Double
        Dim webClient As System.Net.WebClient = New System.Net.WebClient()

        Dim result As String



        Try
            strCreditBalance = webClient.DownloadString("http://121.241.242.114:8080/CreditCheck/checkcredits?username=" & Account & "&password=" & Password)
            strCreditBalance = Trim(strCreditBalance)
            'test = Mid("balance:23.3", 9, Len("balance:23.3") - 8)
            strCreditBalance_Extracted = Mid(strCreditBalance, 9, Len(strCreditBalance) - 8)


            If IsNumeric(strCreditBalance_Extracted) Then
                dblCreditBalance = MyCDBL(strCreditBalance_Extracted)
                result = "Balance" & (SMS_UNIT_COST_AS_APPLIED_BY_ADROIT * dblCreditBalance).ToString
            Else
                result = "Balance: ??"
            End If


        Catch ex As Exception
            'MessageBox.Show("Internet Connection is down at the moment.")

        End Try

        Return strCreditBalance


    End Function

    Function Format_Phone_number(ByVal phone_no As String) As String
        Dim strlen As Integer
        Dim strExtract As String
        Dim strFormatted As String

        strlen = Len(phone_no)

        If strlen = 10 Then
            strExtract = Mid(phone_no, 2, 9)

        ElseIf strlen = 9 Then
            strExtract = Mid(phone_no, 1, 9)

        End If

        strFormatted = "233" & strExtract

        'Remove unwanted characters
        strFormatted = Replace(strFormatted, "-", "")
        strFormatted = Replace(strFormatted, " ", "")
        strFormatted = Replace(strFormatted, ")", "")
        strFormatted = Replace(strFormatted, "(", "")


        ' MessageBox.Show(strFormatted)
        Return strFormatted
    End Function


    Public Sub SEND_SMS_Output_to_Listbox(ByVal sms As String, ByVal phone_number As String, ByVal SenderID As String, ByVal Type As Integer, ByVal Sent_Output_Listbox As ListBox, ByVal Not_Sent_Output_ListBox As ListBox, ByVal SenderName As String)

        Dim webClient As System.Net.WebClient = New System.Net.WebClient()

        ' phone_number = _Phone_number(phone_number)

        Try
            ''for other numbers in Ghana

            If SMS_Account_Name = "" Then
                SMS_Account_Name = InputBox("Please Enter SMS Account Username")
                My.Settings.SMS_Account_Name = SMS_Account_Name

                SMS_Account_Password = InputBox("Please Enter sms Account Password")
                My.Settings.SMS_Account_Password = SMS_Account_Password

                SMS_SenderID = InputBox("Please Enter Your SenderID")
                My.Settings.SMS_SenderID = SMS_SenderID

                SMS_Account_Type = InputBox("Please Enter your SMS Type")
                My.Settings.SMS_Account_Type = SMS_Account_Type


            End If


            Dim result As String = webClient.DownloadString(SMS_ProviderURL & SMS_Account_Name & "&password=" & SMS_Account_Password & "&type=" & Type.ToString & "&dlr=1&destination=" & phone_number & "&source=" & SenderID & "&message=" & sms)

            result = Mid(result, 1, 4)

            Select Case result
                Case "1701"
                    'MessageBox.Show("Message Sent", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Sent_Output_Listbox.Items.Add(SenderName & " " & phone_number)
                    SMS_Successufully_Sent += 1
                Case "1702"
                    ' MessageBox.Show("One of the parameters is blank", "SMS Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Not_Sent_Output_ListBox.Items.Add(SenderName & " " & phone_number & " " & "One of the parameters is blank")
                    SMS_Unsuccessfully_Sent += 1
                Case "1703"
                    '  MessageBox.Show("SMS Account Validation Error", "SMS Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    My.Settings.iM_Authentication = False
                    SMS_Unsuccessfully_Sent += 1
                    Not_Sent_Output_ListBox.Items.Add(SenderName & " " & phone_number & " " & "SMS Account Validation Error")

                Case "1704"
                    'MessageBox.Show("Invalid value in Type field", "SMS Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    SMS_Unsuccessfully_Sent += 1
                    Not_Sent_Output_ListBox.Items.Add(SenderName & " " & phone_number & " " & "Invalid value in Type field")

                Case "1705"
                    ' MessageBox.Show("Invalid Message", "SMS Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    SMS_Unsuccessfully_Sent += 1
                    Not_Sent_Output_ListBox.Items.Add(SenderName & " " & phone_number & " " & "Invalid Message")

                Case "1706"
                    ' MessageBox.Show("Wrong Phone Number", "SMS Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    SMS_Unsuccessfully_Sent += 1
                    Not_Sent_Output_ListBox.Items.Add(SenderName & " " & phone_number & " " & "Wrong Phone Number")

                Case "1707"
                    ' MessageBox.Show("Invalid Source", "SMS Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    SMS_Unsuccessfully_Sent += 1
                    Not_Sent_Output_ListBox.Items.Add(SenderName & " " & phone_number & " " & "Invalid Source")

                Case "1708"
                    '  MessageBox.Show("Invalid dlr field", "SMS Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    SMS_Unsuccessfully_Sent += 1
                    Not_Sent_Output_ListBox.Items.Add(SenderName & " " & phone_number & " " & "Invalid dlr field")

                Case "1709"
                    ' MessageBox.Show("User Validation Failed", "SMS Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    SMS_Unsuccessfully_Sent += 1
                    Not_Sent_Output_ListBox.Items.Add(SenderName & " " & phone_number & " " & "User Validation Failed")

                Case "1710"
                    ' MessageBox.Show("Internal Error", "SMS Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    SMS_Unsuccessfully_Sent += 1
                    Not_Sent_Output_ListBox.Items.Add(SenderName & " " & phone_number & " " & "Internal Error")

            End Select

        Catch

            MessageBox.Show("Internet Connection is down", "SMS Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            SMS_Unsuccessfully_Sent += 1
        End Try
    End Sub

    Public Sub SEND_SMS_Report_to_Adroit(ByVal sms As String, ByVal phone_number As String, ByVal SenderID As String, ByVal Type As Integer, ByVal SenderName As String)

        Dim webClient As System.Net.WebClient = New System.Net.WebClient()

        Try
            Dim result As String = webClient.DownloadString(SMS_ProviderURL & "stl-adroit" & "&password=adroit01" & "&type=0" & "&dlr=1&destination=" & phone_number & "&source=" & SenderID & "&message=" & sms)
            '  MessageBox.Show(result.ToString)
        Catch
            '    MessageBox.Show("Internet Connection is down", "SMS Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '     SMS_Unsuccessfully_Sent += 1
        End Try
    End Sub
    Public Sub SEND_Authentication_sms_to_Adroit()
        Dim error_code As String


        Dim webClient As System.Net.WebClient = New System.Net.WebClient()

        '  If InStr(SMS_Account_Name, "adro-") = 0 Then Exit Sub

        Try
            Dim result As String = webClient.DownloadString(SMS_ProviderURL & SMS_Account_Name & "&password=" & SMS_Account_Password & "&type=0" & "&dlr=1&destination=233541235271&source=iM_Auth&message=" & SMS_Account_Name & " is Authenticating Account.")
            error_code = Mid(result, 1, 4)
            If error_code <> "1701" Then GoTo Catch1
            My.Settings.iM_Authentication = True
            My.Settings.Save()
            MessageBox.Show("Authentication was Successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch

            My.Settings.iM_Authentication = False
            My.Settings.Save()
            MessageBox.Show("Authentication Failed. Please send a mail to support@adroitbureau.com", "Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Exit Sub



Catch1:
        My.Settings.iM_Authentication = False
        My.Settings.Save()
        MessageBox.Show("Authentication Failed. Please send a mail to support@adroitbureau.com", "Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)

    End Sub


    Public Function CREATE_CONNECTION_STRING(ByVal isAccess As Boolean, ByVal isSQLSERVER As Boolean, ByVal ACCESSFILENAME As String, ByVal ACCESSDIRECTORY As String, ByVal SERVERNAME As String, ByVal Initial_Catalog As String, ByVal UserID As String, ByVal PassWord As String, ByVal IntegratedSecurity As Boolean) As String
        Dim Access_ConnectionString As String
        Dim SQL_ConnectionString As String

        If isAccess = True Then
            Access_ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & ACCESSDIRECTORY & "\" & ACCESSFILENAME & ".accdb"
            Return Access_ConnectionString
            GoTo skip
        End If

        If IsPervasive = True Then
            Access_ConnectionString = "Server Name=" & PSQLSERVER_NAME & ";Database Name=" & PSQLDatabase_Name & ";User ID=" & PSQLDatabase_UserID & ";Password=" & PSQLDatabase_UserPassword & ";"
            Return Access_ConnectionString
            GoTo skip
        End If

        If isSQLSERVER = True And IntegratedSecurity = False Then
            'SQL_ConnectionString = "Data Source=" & SERVERNAME & ";Initial Catalog=" & Initial_Catalog & ";User ID=" & My.Settings.SQLServerUserID
            SQL_ConnectionString = "Server=tcp:v15dys1x20.database.windows.net,1433;Database=TotalAcct;User ID=Admin1@v15dys1x20;Password=Test1234;Trusted_Connection=False;Encrypt=True;Connection Timeout=30;"

            Return SQL_ConnectionString
            GoTo skip
        End If

        If isSQLSERVER = True And IntegratedSecurity = True Then


            SQL_ConnectionString = "Data Source=" & SERVERNAME & ";Initial Catalog=" & Initial_Catalog & ";Integrated Security=True"

            Return SQL_ConnectionString
            GoTo skip
        End If

        '        Server Name=myServerAddress;Database Name=myDataBase;User ID=myUsername;
        'Password=myPassword;

        '        Server Name=myServerAddress;Database Name=myDataBase;User ID=myUsername;
        'Password=myPassword;Pooling=False;

        ' Data Source=triumph-pc\sqlexpress;Initial Catalog=SWL9;Integrated Security=True


skip:

    End Function



    Public Sub Load_MySettings_Variables()

       

        SMS_Account_Name = My.Settings.SMS_Account_Name
        SMS_Account_Password = My.Settings.SMS_Account_Password
        SMS_Account_Type = My.Settings.SMS_Account_Type
        ' SMS_ProviderURL = My.Settings.SMS_Account_ProviderURL
        SMS_SenderID = My.Settings.SMS_SenderID

        Email_Account = My.Settings.Email_Account_Name
        Email_Pass = My.Settings.Email_Account_Pass
        SMTP_server_Port = My.Settings.SMTP_server_Port
        SMTP_server_Host = My.Settings.SMTP_server_Server_Host
        SMTP_server_Enable_SSL = My.Settings.SMTP_server_Enable_SSL
        Email_Source = My.Settings.Email_Source
        Customer_Email_Field = My.Settings.Customer_Email_Field

        Customer_Table_Name = My.Settings.Customer_Table_Name
        Customer_Field_Name = My.Settings.Customer_Field_Name
        Customer_Contact_Field1 = My.Settings.Customer_Contact_Field1
        Customer_Contact_Field2 = My.Settings.Customer_Contact_Field2
        Customer_Location_Field = My.Settings.Customer_Location_Field
        Customer_Balance_Table_Name = My.Settings.Customer_Balance_Table_Name
        Customer_Balance_Field_Name = My.Settings.Customer_Balance_FieldName
        Customer_Rep_Field_Name = My.Settings.Customer_Rep_Field_Name
        Customer_Code_Field = My.Settings.Customer_Code_Field
        Customer_Sales_Archive_Field = My.Settings.Customer_Sales_Archive

        IsACCESS = My.Settings.IsACCESS
        IsSQLSERVER = My.Settings.IsSQLServer
        IsPervasive = My.Settings.IS_PERVASIVE
        SQLSERVER_NAME = My.Settings.SQLServerName
        SQLUSERID = My.Settings.SQLServerUserID
        SQLPassword = My.Settings.SQLServerUserPassword
        MSACESSFILE = My.Settings.MsACCESSFilename
        MSACCESSFILE_Path = My.Settings.MsACCESSFilename_Path
        SQL_Initial_Catalog = My.Settings.SQL_InitialCatalog
        IntegratedSecurity = My.Settings.IntegratedSecurity
        PSQLSERVER_NAME = My.Settings.PSQLSERVERNAME
        PSQLDatabase_Name = My.Settings.PSQLDatabase_Name
        PSQLDatabase_UserID = My.Settings.PSQLDatabase_UserID
        PSQLDatabase_UserPassword = My.Settings.PSQLDatabase_UserPassword

        'Perculiar  to Pastel Partner
        Contacts_Field_Name_Pastel = My.Settings.Contacts_Field_Name_Pastel
        Contacts_Table_Name_Pastel = My.Settings.Contacts_Table_Name_Pastel


        ' For the location intelligence module
        Region_Field = My.Settings.Region_Field
        District_Field = My.Settings.District_Field

        'For the Business Intelligence Module
        BI_Customer_Accounts_Table = My.Settings.BI_Customer_Accounts_Table
        BI_Customer_Accounts_Name_Field = My.Settings.BI_Customer_Accounts_Name_Field
        BI_Customer_Accounts_Code_Field = My.Settings.BI_Customer_Accounts_Code_Field


        'Public Variables for the VoiceSMS
        Voice_SMS_Account_Name = My.Settings.Voice_SMS_Account_Name
        Voice_SMS_Account_Password = My.Settings.Voice_SMS_Account_Password
        Voice_Audio_URL = My.Settings.Voice_SMS_Audio_URL
        VoiceSMS_ProviderURL = My.Settings.Voice_SMS_Provider_URL

        'Public Variables for Custom Fields
        Custom_Field1 = My.Settings.Custom_Field1
        Custom_Field2 = My.Settings.Custom_Field2
        Custom_Field3 = My.Settings.Custom_Field3
        Custom_Field4 = My.Settings.Custom_Field4
        Custom_Field5 = My.Settings.Custom_Field5
        Custom_Field6 = My.Settings.Custom_Field6
        Custom_Field7 = My.Settings.Custom_Field7
        Custom_Field8 = My.Settings.Custom_Field8
        Custom_Field9 = My.Settings.Custom_Field9
        Custom_Field10 = My.Settings.Custom_Field10


        'GoogleForms Variables
        URL_InputForms = My.Settings.URL_Input_Webform
        URL_OutForms = My.Settings.URL_Output_Webform



    End Sub

    Public Function SWL_SCALAR_QUERY(ByVal query As String) As Double
        Dim queryString As String = query
        '   "SELECT count(Company) from Dealers where mobile = '' and phone = ''"

        If IsSQLSERVER Then
            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()


                        Return CDbl(dataReader(0))
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    '  MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using

        End If


        If IsACCESS Then
            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()


                        Return CDbl(dataReader(0))
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    'MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using


        End If


        'If IsPervasive Then
        '    Using connection As New PsqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
        '        Dim command As PsqlCommand = connection.CreateCommand()
        '        command.CommandText = queryString
        '        Try
        '            connection.Open()
        '            Dim dataReader As PsqlDataReader = _
        '             command.ExecuteReader()
        '            Do While dataReader.Read()


        '                Return CDbl(dataReader(0))
        '            Loop

        '            dataReader.Close()

        '        Catch ex As Exception
        '            'MessageBox.Show(ex.Message)
        '        End Try

        '        connection.Close()
        '    End Using


        'End If

    End Function

    Public Function Get_Number_of_contacts()
        Return SWL_SCALAR_QUERY("select count(" & Customer_Field_Name & ") from " & Customer_Table_Name)
    End Function

    Function Convert_Region_Format_From_Attribute_Table_To_Focus_Format(ByVal Target As String) As String
        Select Case Target
            Case "Ashanti"
                Return "A/R"

            Case "Central"
                Return "C/R"

            Case "Eastern"
                Return "E/R"

            Case "Greater Accra"
                Return "GAR"

            Case "Northern"
                Return "N/R"

            Case "Upper East"
                Return "UE/R"

            Case "Upper West"
                Return "UW/R"

            Case "Volta"
                Return "V/R"

            Case "Western"
                Return "W/R"

            Case "Brong Ahafo"
                Return "B/R"

            Case Else
                Return "None"
        End Select


    End Function

End Module
