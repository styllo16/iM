Imports System.Data
Imports Pervasive.Data.SqlClient
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Web 'add system.web as reference for the url encoding class
Imports System.Net
Imports System.IO
Public Class Sms

    Public Batch_Count As Integer
    Dim By_Debtors As Boolean = False
    Sub Load_Schools_Based_On_Status(ByVal status As String)
        Clear_BatchBox()

        If IsACCESS Then

            'Customized Variables
            Dim MyCustomer_Field_Name = "School"
            Dim MyCustomer_Contact_Field1 = "Telephone"
            Dim MyCustomer_Contact_Field2 = "Telephone2"
            Dim myCustomer_Email_Field = "Email"
            Dim myCustomer_Table_Name = "im_Data_SWL"
            Dim Email2 As String
            Dim Phone_Number As String

            Dim queryString As String = _
          "SELECT " & Customer_Field_Name & "," & Customer_Contact_Field1 & "," & Customer_Contact_Field2 & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & Custom_Field1 & " = '" & status & "'"


            '
            '  "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Data Directory\ESSAMUAH DATABASE.accdb"

            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        Email2 = dataReader(3).ToString
                        Phone_Number = dataReader(1).ToString

                        batchbox.Rows.Add(1)
                        batchbox.Item(0, Batch_Count).Value = dataReader(0).ToString
                        batchbox.Item(1, Batch_Count).Value = Phone_Number
                        batchbox.Item(4, Batch_Count).Value = Email2


                        'Untick Checkbox when there is no number for EMAIL
                        If batchbox.Item(4, Batch_Count).Value = "" Then
                            batchbox.Item(5, Batch_Count).Value = False
                        Else
                            batchbox.Item(5, Batch_Count).Value = True
                        End If

                        If Phone_Number <> "" Then batchbox.Item(3, Batch_Count).Value = True

                        Batch_Count += 1
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using
        End If


        Send_List_Title = batchbox.Rows.Count.ToString & " Contacts"
        Set_Text_Of_TabControl1(0, Send_List_Title)

    End Sub
    Sub Load_Schools_Based_On_Region(ByVal region As String)
        Clear_BatchBox()

        If IsACCESS Then

            'Customized Variables
            Dim MyCustomer_Field_Name = "School"
            Dim MyCustomer_Contact_Field1 = "Telephone"
            Dim MyCustomer_Contact_Field2 = "Telephone2"
            Dim myCustomer_Email_Field = "Email"
            Dim myCustomer_Table_Name = "im_Data_SWL"
            Dim Email2 As String
            Dim Phone_Number As String

            Dim queryString As String = _
          "SELECT " & Customer_Field_Name & "," & Customer_Contact_Field1 & "," & Customer_Contact_Field2 & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & Region_Field & " like '" & region & "'"


            '
            '  "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Data Directory\ESSAMUAH DATABASE.accdb"

            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        Email2 = dataReader(3).ToString
                        Phone_Number = dataReader(1).ToString

                        batchbox.Rows.Add(1)
                        batchbox.Item(0, Batch_Count).Value = dataReader(0).ToString
                        batchbox.Item(1, Batch_Count).Value = Phone_Number
                        batchbox.Item(4, Batch_Count).Value = Email2


                        'Untick Checkbox when there is no number for EMAIL
                        If batchbox.Item(4, Batch_Count).Value = "" Then
                            batchbox.Item(5, Batch_Count).Value = False
                        Else
                            batchbox.Item(5, Batch_Count).Value = True
                        End If

                        If Phone_Number <> "" Then batchbox.Item(3, Batch_Count).Value = True

                        Batch_Count += 1
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using
        End If


        Send_List_Title = batchbox.Rows.Count.ToString & " Contacts"
        Set_Text_Of_TabControl1(0, Send_List_Title)

    End Sub

    Sub Load_Schools_Based_On_Like_For_School_Name(ByVal LikeName As String)
        ' Clear_BatchBox()
        'Customized Variables
        Dim MyCustomer_Field_Name = "School"
        Dim MyCustomer_Contact_Field1 = "Telephone"
        Dim MyCustomer_Contact_Field2 = "Telephone2"
        Dim myCustomer_Email_Field = "Email"
        Dim myCustomer_Table_Name = "im_Data_SWL"
        Dim Email2 As String
        Dim Phone_Number As String


        If IsACCESS Then

           

            Dim queryString As String = _
          "SELECT " & Customer_Field_Name & "," & Customer_Contact_Field1 & "," & Customer_Contact_Field2 & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & Customer_Field_Name & " like '" & LikeName & "'"


            '
            '  "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Data Directory\ESSAMUAH DATABASE.accdb"

            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        Email2 = dataReader(3).ToString
                        Phone_Number = dataReader(1).ToString

                        batchbox.Rows.Add(1)
                        batchbox.Item(0, Batch_Count).Value = dataReader(0).ToString
                        batchbox.Item(1, Batch_Count).Value = Phone_Number
                        batchbox.Item(4, Batch_Count).Value = Email2


                        'Untick Checkbox when there is no number for EMAIL
                        If batchbox.Item(4, Batch_Count).Value = "" Then
                            batchbox.Item(5, Batch_Count).Value = False
                        Else
                            batchbox.Item(5, Batch_Count).Value = True
                        End If

                        If Phone_Number <> "" Then batchbox.Item(3, Batch_Count).Value = True

                        Batch_Count += 1
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using
        End If



        If IsSQLSERVER Then



            Dim queryString As String = _
          "SELECT " & Customer_Field_Name & "," & Customer_Contact_Field1 & "," & Customer_Contact_Field2 & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & Customer_Field_Name & " like '" & LikeName & "'"


            '
            '  "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Data Directory\ESSAMUAH DATABASE.accdb"

            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        Email2 = dataReader(3).ToString
                        Phone_Number = dataReader(1).ToString

                        batchbox.Rows.Add(1)
                        batchbox.Item(0, Batch_Count).Value = dataReader(0).ToString
                        batchbox.Item(1, Batch_Count).Value = Phone_Number
                        batchbox.Item(4, Batch_Count).Value = Email2


                        'Untick Checkbox when there is no number for EMAIL
                        If batchbox.Item(4, Batch_Count).Value = "" Then
                            batchbox.Item(5, Batch_Count).Value = False
                        Else
                            batchbox.Item(5, Batch_Count).Value = True
                        End If

                        If Phone_Number <> "" Then batchbox.Item(3, Batch_Count).Value = True

                        Batch_Count += 1
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using
        End If



        'If IsPervasive Then



        '    Dim queryString As String = _
        '  "SELECT " & Customer_Field_Name & "," & Customer_Contact_Field1 & "," & Customer_Contact_Field2 & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & Customer_Field_Name & " like '" & LikeName & "'"


        '    '
        '    '  "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Data Directory\ESSAMUAH DATABASE.accdb"

        '    Using connection As New PsqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
        '        Dim command As PsqlCommand = connection.CreateCommand()
        '        command.CommandText = queryString
        '        Try
        '            connection.Open()
        '            Dim dataReader As PsqlDataReader = _
        '             command.ExecuteReader()
        '            Do While dataReader.Read()
        '                Email2 = dataReader(3).ToString
        '                Phone_Number = dataReader(1).ToString

        '                batchbox.Rows.Add(1)
        '                batchbox.Item(0, Batch_Count).Value = dataReader(0).ToString
        '                batchbox.Item(1, Batch_Count).Value = Phone_Number
        '                batchbox.Item(4, Batch_Count).Value = Email2


        '                'Untick Checkbox when there is no number for EMAIL
        '                If batchbox.Item(4, Batch_Count).Value = "" Then
        '                    batchbox.Item(5, Batch_Count).Value = False
        '                Else
        '                    batchbox.Item(5, Batch_Count).Value = True
        '                End If

        '                If Phone_Number <> "" Then batchbox.Item(3, Batch_Count).Value = True

        '                Batch_Count += 1
        '            Loop

        '            dataReader.Close()

        '        Catch ex As Exception
        '            MessageBox.Show(ex.Message)
        '        End Try

        '        connection.Close()
        '    End Using
        'End If



        Send_List_Title = batchbox.Rows.Count.ToString & " Contacts"
        Set_Text_Of_TabControl1(0, Send_List_Title)

    End Sub

    Sub Load_Schools_Based_On_Region_and_Status(ByVal LikeName As String, ByVal status As String)
        Clear_BatchBox()

        If IsACCESS Then

            'Customized Variables
            Dim MyCustomer_Field_Name = "School"
            Dim MyCustomer_Contact_Field1 = "Telephone"
            Dim MyCustomer_Contact_Field2 = "Telephone2"
            Dim myCustomer_Email_Field = "Email"
            Dim myCustomer_Table_Name = "im_Data_SWL"
            Dim Email2 As String
            Dim Phone_Number As String

            Dim queryString As String = _
          "SELECT " & Customer_Field_Name & "," & Customer_Contact_Field1 & "," & Customer_Contact_Field2 & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & Region_Field & " like '" & LikeName & "' and " & Custom_Field1 & " = '" & status & "'"


            '
            '  "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Data Directory\ESSAMUAH DATABASE.accdb"

            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        Email2 = dataReader(3).ToString
                        Phone_Number = dataReader(1).ToString

                        batchbox.Rows.Add(1)
                        batchbox.Item(0, Batch_Count).Value = dataReader(0).ToString
                        batchbox.Item(1, Batch_Count).Value = Phone_Number
                        batchbox.Item(4, Batch_Count).Value = Email2


                        'Untick Checkbox when there is no number for EMAIL
                        If batchbox.Item(4, Batch_Count).Value = "" Then
                            batchbox.Item(5, Batch_Count).Value = False
                        Else
                            batchbox.Item(5, Batch_Count).Value = True
                        End If

                        If Phone_Number <> "" Then batchbox.Item(3, Batch_Count).Value = True

                        Batch_Count += 1
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using
        End If


        Send_List_Title = batchbox.Rows.Count.ToString & " Contacts"
        Set_Text_Of_TabControl1(0, Send_List_Title)

    End Sub

    Sub Load_Schools_Based_On_Level_and_Status(ByVal school_likeName As String, ByVal status As String)
        Clear_BatchBox()

        If IsACCESS Then

            'Customized Variables
            Dim MyCustomer_Field_Name = "School"
            Dim MyCustomer_Contact_Field1 = "Telephone"
            Dim MyCustomer_Contact_Field2 = "Telephone2"
            Dim myCustomer_Email_Field = "Email"
            Dim myCustomer_Table_Name = "im_Data_SWL"
            Dim Email2 As String
            Dim Phone_Number As String

            Dim queryString As String = _
          "SELECT " & Customer_Field_Name & "," & Customer_Contact_Field1 & "," & Customer_Contact_Field2 & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & Customer_Field_Name & " like '" & school_likeName & "' and " & Custom_Field1 & " = '" & status & "'"


            '
            '  "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Data Directory\ESSAMUAH DATABASE.accdb"

            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        Email2 = dataReader(3).ToString
                        Phone_Number = dataReader(1).ToString

                        batchbox.Rows.Add(1)
                        batchbox.Item(0, Batch_Count).Value = dataReader(0).ToString
                        batchbox.Item(1, Batch_Count).Value = Phone_Number
                        batchbox.Item(4, Batch_Count).Value = Email2


                        'Untick Checkbox when there is no number for EMAIL
                        If batchbox.Item(4, Batch_Count).Value = "" Then
                            batchbox.Item(5, Batch_Count).Value = False
                        Else
                            batchbox.Item(5, Batch_Count).Value = True
                        End If

                        If Phone_Number <> "" Then batchbox.Item(3, Batch_Count).Value = True

                        Batch_Count += 1
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using
        End If


        Send_List_Title = batchbox.Rows.Count.ToString & " Contacts"
        Set_Text_Of_TabControl1(0, Send_List_Title)

    End Sub


    Sub Display_AR_Listing()
        Clear_BatchBox()
        Dim i As Integer
        Dim Id As String
        Dim AccountBalance As Double
        Dim Accounts_Sum As Double
        Dim Telephone1 As String
        Dim Telephone2 As String
        Dim CompanyName As String
        Dim Email As String
        Dim Exception_Already_Communicated As Boolean = False

        Dim queryString As String = _
            "SELECT " & Customer_Code_Field & "," & Customer_Contact_Field1 & "," & Customer_Contact_Field2 & "," & Customer_Field_Name & "," & Customer_Email_Field & " from " & Customer_Table_Name


        If IsACCESS Then
            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        Id = dataReader(0).ToString
                        Telephone1 = dataReader(1).ToString
                        Telephone2 = dataReader(2).ToString
                        CompanyName = dataReader(3).ToString
                        Email = dataReader(4).ToString

                        AccountBalance = MyCDBL(AR_Account_Listing_Value_For_This_Customer(Id))
                        Accounts_Sum += AccountBalance

                        If AccountBalance < 0 Or AccountBalance > 0 Then
                            batchbox.Rows.Add(1)
                            batchbox.Item(0, i).Value = CompanyName

                            If Telephone1 <> "" Then
                                batchbox.Item(1, i).Value = Telephone1
                            Else
                                batchbox.Item(1, i).Value = Telephone2
                            End If

                            'Untick Checkbox when there is no number for SMS
                            If batchbox.Item(1, i).Value = "" Then
                                batchbox.Item(3, i).Value = False
                            Else
                                batchbox.Item(3, i).Value = True
                            End If

                            batchbox.Item(2, i).Value = AccountBalance
                            batchbox.Item(4, i).Value = Email

                            'Untick Checkbox when there is no number for EMAIL
                            If batchbox.Item(4, i).Value = "" Then
                                batchbox.Item(5, i).Value = False
                            Else
                                batchbox.Item(5, i).Value = True
                            End If



                            i += 1
                            Batch_Count += 1
                        End If

                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    If Exception_Already_Communicated = True Then GoTo skip_exception
                    'My.Settings.Error_log.Add(ex.Message)
                    Exception_Already_Communicated = True
skip_exception:
                End Try

                connection.Close()
            End Using


        End If

        If IsSQLSERVER Then
            Dim Email2 As String
            Dim id2 As String
            Accounts_Sum = 0

            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        id2 = dataReader(0).ToString
                        Telephone1 = dataReader(1).ToString
                        Telephone2 = dataReader(2).ToString
                        CompanyName = dataReader(3).ToString
                        Email2 = dataReader(4).ToString

                        AccountBalance = MyCDBL(AR_Account_Listing_Value_For_This_Customer(id2))
                        Accounts_Sum += AccountBalance

                        If AccountBalance < 0 Or AccountBalance > 0 Then
                            batchbox.Rows.Add(1)
                            batchbox.Item(0, i).Value = CompanyName

                            If Telephone1 <> "" Then
                                batchbox.Item(1, i).Value = Telephone1
                            Else
                                batchbox.Item(1, i).Value = Telephone2
                            End If

                            'Untick Checkbox when there is no number
                            If batchbox.Item(1, i).Value = "" Then
                                batchbox.Item(3, i).Value = False
                            Else
                                batchbox.Item(3, i).Value = True
                            End If

                            batchbox.Item(2, i).Value = AccountBalance
                            batchbox.Item(4, i).Value = Email2


                            'Untick Checkbox when there is no number for EMAIL
                            If batchbox.Item(4, i).Value = "" Then
                                batchbox.Item(5, i).Value = False
                            Else
                                batchbox.Item(5, i).Value = True
                            End If



                            i += 1
                            Batch_Count += 1
                        End If

                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    'My.Settings.Error_log.Add(ex.Message.ToString)
                End Try

                connection.Close()
            End Using

        End If

        'If IsPervasive Then
        '    Dim Email2 As String
        '    Dim id3 As String
        '    Accounts_Sum = 0

        '    Using connection As New PsqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
        '        Dim command As PsqlCommand = connection.CreateCommand()
        '        command.CommandText = "SELECT " & Customer_Code_Field & "," & Customer_Contact_Field1 & "," & Customer_Contact_Field2 & "," & Customer_Field_Name & "," & Customer_Email_Field & "," & Customer_Balance_Field_Name & " from " & Customer_Table_Name
        '        Try
        '            connection.Open()
        '            Dim dataReader As PsqlDataReader = _
        '             command.ExecuteReader()
        '            Do While dataReader.Read()
        '                id3 = dataReader(0).ToString
        '                Telephone1 = dataReader(1).ToString
        '                Telephone2 = dataReader(2).ToString
        '                CompanyName = dataReader(3).ToString
        '                Email2 = dataReader(4).ToString

        '                AccountBalance = MyCDBL(dataReader(5).ToString)  'MyCDBL(AR_Account_Listing_Value_For_This_Customer(id3)))

        '                Accounts_Sum += AccountBalance

        '                If AccountBalance < 0 Or AccountBalance > 0 Then
        '                    batchbox.Rows.Add(1)
        '                    batchbox.Item(0, i).Value = CompanyName



        '                    'Pastel
        '                    If Contacts_Table_Name_Pastel = "" Then
        '                        batchbox.Item(1, i).Value = Telephone1

        '                    Else
        '                        batchbox.Item(1, i).Value = Fetch_Field_Value_From_Specified_Table_Given_ID(id3, Contacts_Field_Name_Pastel, Contacts_Table_Name_Pastel)
        '                    End If



        '                    'Untick Checkbox when there is no number
        '                    If batchbox.Item(1, i).Value = "" Then
        '                        batchbox.Item(3, i).Value = False
        '                    Else
        '                        batchbox.Item(3, i).Value = True
        '                    End If

        '                    batchbox.Item(2, i).Value = AccountBalance
        '                    batchbox.Item(4, i).Value = Email2


        '                    'Untick Checkbox when there is no number for EMAIL
        '                    If batchbox.Item(4, i).Value = "" Then
        '                        batchbox.Item(5, i).Value = False
        '                    Else
        '                        batchbox.Item(5, i).Value = True
        '                    End If



        '                    i += 1
        '                    Batch_Count += 1
        '                End If

        '            Loop

        '            dataReader.Close()

        '        Catch ex As Exception
        '            'My.Settings.Error_log.Add(ex.Message)
        '        End Try

        '        connection.Close()
        '    End Using

        'End If

    End Sub
    Sub Display_List_by_Custome_Field(ByVal combovalue As String, ByVal customfield As String)
        Clear_BatchBox()

        Dim i As Integer
        Dim telephone As String
        Dim Companyname As String
        Dim Email As String

        ' town = InputBox("Enter Town")

        If IsACCESS Then
            Dim queryString As String = _
                  "SELECT " & Customer_Contact_Field1 & "," & Customer_Field_Name & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & customfield & "='" & combovalue & "'"

            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        telephone = Trim(dataReader(0).ToString)
                        Companyname = dataReader(1).ToString
                        Email = dataReader(2).ToString

                        batchbox.Rows.Add(1)
                        batchbox.Item(0, i).Value = Companyname
                        batchbox.Item(1, i).Value = telephone


                        If batchbox.Item(1, i).Value = "" Then
                            batchbox.Item(3, Batch_Count).Value = False
                        Else
                            batchbox.Item(3, Batch_Count).Value = True
                        End If
                        batchbox.Item(4, i).Value = Email

                        'Untick Checkbox when there is no number for EMAIL
                        If batchbox.Item(4, i).Value = "" Then
                            batchbox.Item(5, i).Value = False
                        Else
                            batchbox.Item(5, i).Value = True
                        End If



                        i += 1
                        Batch_Count += 1
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using

        End If

        If IsSQLSERVER Then
            Dim Email2 As String
            Dim queryString As String = _
                  "SELECT " & Customer_Contact_Field1 & "," & Customer_Field_Name & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & customfield & "='" & combovalue & "'"

            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        telephone = Trim(dataReader(0).ToString)
                        Companyname = dataReader(1).ToString
                        Email2 = dataReader(2).ToString

                        batchbox.Rows.Add(1)
                        batchbox.Item(0, i).Value = Companyname
                        batchbox.Item(1, i).Value = telephone

                        If batchbox.Item(1, i).Value = "" Then
                            batchbox.Item(3, Batch_Count).Value = False
                        Else
                            batchbox.Item(3, Batch_Count).Value = True
                        End If
                        batchbox.Item(4, i).Value = Email2

                        'Untick Checkbox when there is no number for EMAIL
                        If batchbox.Item(4, i).Value = "" Then
                            batchbox.Item(5, i).Value = False
                        Else
                            batchbox.Item(5, i).Value = True
                        End If


                        i += 1
                        Batch_Count += 1
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using


        End If

        'If IsPervasive Then
        '    Dim Email2 As String
        '    Dim id As String
        '    Dim queryString As String = _
        '          "SELECT " & Customer_Contact_Field1 & "," & Customer_Field_Name & "," & Customer_Email_Field & "," & Customer_Code_Field & " from " & Customer_Table_Name & " where " & customfield & "='" & combovalue & "'"

        '    Using connection As New PsqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
        '        Dim command As PsqlCommand = connection.CreateCommand()
        '        command.CommandText = queryString
        '        Try
        '            connection.Open()
        '            Dim dataReader As PsqlDataReader = _
        '             command.ExecuteReader()
        '            Do While dataReader.Read()
        '                telephone = Trim(dataReader(0).ToString)
        '                Companyname = dataReader(1).ToString
        '                Email2 = dataReader(2).ToString
        '                id = dataReader(3).ToString


        '                batchbox.Rows.Add(1)
        '                batchbox.Item(0, i).Value = Companyname


        '                'Pastel
        '                If Contacts_Table_Name_Pastel = "" Then
        '                    batchbox.Item(1, i).Value = telephone

        '                Else
        '                    batchbox.Item(1, i).Value = Fetch_Field_Value_From_Specified_Table_Given_ID(id, Contacts_Field_Name_Pastel, Contacts_Table_Name_Pastel)
        '                End If




        '                If batchbox.Item(1, i).Value = "" Then
        '                    batchbox.Item(3, Batch_Count).Value = False
        '                Else
        '                    batchbox.Item(3, Batch_Count).Value = True
        '                End If
        '                batchbox.Item(4, i).Value = Email2

        '                'Untick Checkbox when there is no number for EMAIL
        '                If batchbox.Item(4, i).Value = "" Then
        '                    batchbox.Item(5, i).Value = False
        '                Else
        '                    batchbox.Item(5, i).Value = True
        '                End If


        '                i += 1
        '                Batch_Count += 1
        '            Loop

        '            dataReader.Close()

        '        Catch ex As Exception
        '            MessageBox.Show(ex.Message)
        '        End Try

        '        connection.Close()
        '    End Using


        'End If

    End Sub

    Sub Display_List_by_District(ByVal District As String)
        Clear_BatchBox()

        Dim i As Integer
        Dim telephone As String
        Dim Companyname As String
        Dim Email As String

        ' town = InputBox("Enter Town")

        If IsACCESS Then
            Dim queryString As String = _
                  "SELECT " & Customer_Contact_Field1 & "," & Customer_Field_Name & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & District_Field & " like '" & District & "'"

            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        telephone = Trim(dataReader(0).ToString)
                        Companyname = dataReader(1).ToString
                        Email = dataReader(2).ToString

                        batchbox.Rows.Add(1)
                        batchbox.Item(0, i).Value = Companyname
                        batchbox.Item(1, i).Value = telephone


                        If batchbox.Item(1, i).Value = "" Then
                            batchbox.Item(3, Batch_Count).Value = False
                        Else
                            batchbox.Item(3, Batch_Count).Value = True
                        End If
                        batchbox.Item(4, i).Value = Email

                        'Untick Checkbox when there is no number for EMAIL
                        If batchbox.Item(4, i).Value = "" Then
                            batchbox.Item(5, i).Value = False
                        Else
                            batchbox.Item(5, i).Value = True
                        End If



                        i += 1
                        Batch_Count += 1
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using

        End If

        If IsSQLSERVER Then
            Dim Email2 As String
            Dim queryString As String = _
                "SELECT " & Customer_Contact_Field1 & "," & Customer_Field_Name & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & District_Field & " like '" & District & "'"

            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        telephone = Trim(dataReader(0).ToString)
                        Companyname = dataReader(1).ToString
                        Email2 = dataReader(2).ToString

                        batchbox.Rows.Add(1)
                        batchbox.Item(0, i).Value = Companyname
                        batchbox.Item(1, i).Value = telephone

                        If batchbox.Item(1, i).Value = "" Then
                            batchbox.Item(3, Batch_Count).Value = False
                        Else
                            batchbox.Item(3, Batch_Count).Value = True
                        End If
                        batchbox.Item(4, i).Value = Email2

                        'Untick Checkbox when there is no number for EMAIL
                        If batchbox.Item(4, i).Value = "" Then
                            batchbox.Item(5, i).Value = False
                        Else
                            batchbox.Item(5, i).Value = True
                        End If


                        i += 1
                        Batch_Count += 1
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using


        End If


        'If IsPervasive Then
        '    Dim Email2 As String
        '    Dim id As String
        '    Dim queryString As String = _
        '        "SELECT " & Customer_Contact_Field1 & "," & Customer_Field_Name & "," & Customer_Email_Field & "," & Customer_Code_Field & " from " & Customer_Table_Name & " where " & District_Field & " like '" & District & "'"

        '    Using connection As New PsqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
        '        Dim command As PsqlCommand = connection.CreateCommand()
        '        command.CommandText = queryString
        '        Try
        '            connection.Open()
        '            Dim dataReader As PsqlDataReader = _
        '             command.ExecuteReader()
        '            Do While dataReader.Read()
        '                telephone = Trim(dataReader(0).ToString)
        '                Companyname = dataReader(1).ToString
        '                Email2 = dataReader(2).ToString
        '                id = dataReader(3).ToString

        '                batchbox.Rows.Add(1)
        '                batchbox.Item(0, i).Value = Companyname


        '                'Pastel
        '                If Contacts_Table_Name_Pastel = "" Then
        '                    batchbox.Item(1, i).Value = telephone

        '                Else
        '                    batchbox.Item(1, i).Value = Fetch_Field_Value_From_Specified_Table_Given_ID(id, Contacts_Field_Name_Pastel, Contacts_Table_Name_Pastel)
        '                End If




        '                If batchbox.Item(1, i).Value = "" Then
        '                    batchbox.Item(3, Batch_Count).Value = False
        '                Else
        '                    batchbox.Item(3, Batch_Count).Value = True
        '                End If
        '                batchbox.Item(4, i).Value = Email2

        '                'Untick Checkbox when there is no number for EMAIL
        '                If batchbox.Item(4, i).Value = "" Then
        '                    batchbox.Item(5, i).Value = False
        '                Else
        '                    batchbox.Item(5, i).Value = True
        '                End If


        '                i += 1
        '                Batch_Count += 1
        '            Loop

        '            dataReader.Close()

        '        Catch ex As Exception
        '            MessageBox.Show(ex.Message)
        '        End Try

        '        connection.Close()
        '    End Using


        'End If

    End Sub


    Sub Display_List_by_Customer(ByVal Current_Customer As String)
        Clear_BatchBox()

        Dim Email As String
        Dim i As Integer
        Dim telephone As String
        Dim Companyname As String
        ' town = InputBox("Enter Town")

        If IsACCESS Then
            Dim queryString As String = _
                  "SELECT " & Customer_Contact_Field1 & "," & Customer_Field_Name & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & Customer_Field_Name & " like '" & Current_Customer & "'"

            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        telephone = Trim(dataReader(0).ToString)
                        Companyname = dataReader(1).ToString
                        Email = dataReader(2).ToString

                        batchbox.Rows.Add(1)
                        batchbox.Item(0, i).Value = Companyname
                        batchbox.Item(1, i).Value = telephone
                        batchbox.Item(3, Batch_Count).Value = True
                        batchbox.Item(4, i).Value = Email


                        'Untick Checkbox when there is no number for EMAIL
                        If batchbox.Item(4, i).Value = "" Then
                            batchbox.Item(5, i).Value = False
                        Else
                            batchbox.Item(5, i).Value = True
                        End If

                        i += 1
                        Batch_Count += 1
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using

        Else
            Dim Email2 As String
            Dim queryString As String = _
                 "SELECT " & Customer_Contact_Field1 & "," & Customer_Field_Name & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & Customer_Field_Name & " like '" & Current_Customer & "'"

            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        telephone = Trim(dataReader(0).ToString)
                        Companyname = dataReader(1).ToString
                        Email2 = dataReader(2).ToString

                        batchbox.Rows.Add(1)
                        batchbox.Item(0, i).Value = Companyname
                        batchbox.Item(1, i).Value = telephone
                        batchbox.Item(3, Batch_Count).Value = True
                        batchbox.Item(4, i).Value = Email2

                        'Untick Checkbox when there is no number for EMAIL
                        If batchbox.Item(4, i).Value = "" Then
                            batchbox.Item(5, i).Value = False
                        Else
                            batchbox.Item(5, i).Value = True
                        End If

                        i += 1
                        Batch_Count += 1
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using


        End If

    End Sub
    Sub Display_List_by_Representative(ByVal Representative As String)
        Clear_BatchBox()

        Dim Email As String
        Dim i As Integer
        Dim telephone As String
        Dim Companyname As String
        ' town = InputBox("Enter Town")

        If IsACCESS Then
            Dim queryString As String = _
                  "SELECT " & Customer_Contact_Field1 & "," & Customer_Field_Name & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & Customer_Rep_Field_Name & " like '" & Representative & "'"

            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        telephone = Trim(dataReader(0).ToString)
                        Companyname = dataReader(1).ToString
                        Email = dataReader(2).ToString

                        batchbox.Rows.Add(1)
                        batchbox.Item(0, i).Value = Companyname
                        batchbox.Item(1, i).Value = telephone


                        'Untick the check box when there is no number
                        If batchbox.Item(1, i).Value = "" Then
                            batchbox.Item(3, i).Value = False
                        Else
                            batchbox.Item(3, i).Value = True
                        End If
                        batchbox.Item(4, i).Value = Email

                        'Untick Checkbox when there is no number for EMAIL
                        If batchbox.Item(4, i).Value = "" Then
                            batchbox.Item(5, i).Value = False
                        Else
                            batchbox.Item(5, i).Value = True
                        End If


                        i += 1
                        Batch_Count += 1
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using

        Else
            Dim Email2 As String
            Dim queryString As String = _
                          "SELECT " & Customer_Contact_Field1 & "," & Customer_Field_Name & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & Customer_Rep_Field_Name & " like '" & Representative & "'"

            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        telephone = Trim(dataReader(0).ToString)
                        Companyname = dataReader(1).ToString
                        Email2 = dataReader(2).ToString

                        batchbox.Rows.Add(1)
                        batchbox.Item(0, i).Value = Companyname
                        batchbox.Item(1, i).Value = telephone

                        'Untick the check box when there is no number
                        If batchbox.Item(1, i).Value = "" Then
                            batchbox.Item(3, i).Value = False
                        Else
                            batchbox.Item(3, i).Value = True
                        End If
                        batchbox.Item(4, i).Value = Email2

                        'Untick Checkbox when there is no number for EMAIL
                        If batchbox.Item(4, i).Value = "" Then
                            batchbox.Item(5, i).Value = False
                        Else
                            batchbox.Item(5, i).Value = True
                        End If

                        i += 1
                        Batch_Count += 1
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using


        End If

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'If CheckBox1.Checked = True Then

        '    ComboBox2.Enabled = True
        '    RadioButton1.Enabled = False
        '    RadioButton2.Enabled = False
        '    RadioButton3.Enabled = False
        '    RadioButton4.Enabled = False
        '    RadioButton5.Enabled = False
        '    RadioButton6.Enabled = False
        '    RadioButton7.Enabled = False

        'Else
        '    ComboBox2.Enabled = False
        '    RadioButton1.Enabled = True
        '    RadioButton2.Enabled = True
        '    RadioButton3.Enabled = True
        '    RadioButton4.Enabled = True
        '    RadioButton5.Enabled = True
        '    RadioButton6.Enabled = True
        '    RadioButton7.Enabled = True

        'End If


    End Sub

    'Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click

    '    'If SplitContainer3.Panel1Collapsed = False Then

    '    '    SplitContainer3.Panel1Collapsed = True
    '    '    ToolStripButton2.Text = "Split pannel"

    '    'Else
    '    '    SplitContainer3.Panel1Collapsed = False
    '    '    ToolStripButton2.Text = "Show Send list "

    '    'End If

    'End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click

        If SplitContainer2.Panel2Collapsed = False Then

            SplitContainer2.Panel2Collapsed = True
            ToolStripButton3.Text = "Show template pannel"

        Else
            SplitContainer2.Panel2Collapsed = False
            ToolStripButton3.Text = "Hide template pannel"
            SplitContainer2.SplitterDistance = 1000
        End If


    End Sub
    Sub Refresh_SMS_Templates()
        ListBox1.Items.Clear()
        Dim i As Integer
        For i = 0 To My.Settings.SMS_Templates.Count - 1

            ListBox1.Items.Add(My.Settings.SMS_Templates.Item(i))
        Next
    End Sub

    Private Sub Sms_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
     
    End Sub

    Private Sub Sms_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click

    End Sub

    Private Sub Sms_Enter(sender As Object, e As System.EventArgs) Handles Me.Enter

    End Sub

   

    Private Sub Sms_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        '************************************************************
        ' Authentication Code
        If My.Settings.iM_Authentication = False Then
            ToolStrip1.Enabled = False
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            TabControl1.Enabled = False
            TabControl2.Enabled = False
        Else
            ToolStrip1.Enabled = True
            Button1.Enabled = True
            Button2.Enabled = True
            Button3.Enabled = True
            TabControl1.Enabled = True
            TabControl2.Enabled = True
        End If

        '************************************************************


        If InterModule_Transfer = False Then Exit Sub

        'Fill_Combo_With_Towns(ToolStripComboBox2, Customer_Table_Name, Customer_Location_Field)

        'Fill_Combo_With_Customers(ToolStripComboBox5)

        'Fill_Combo_With_Reps(ToolStripComboBox5)

        If Customer_Array(0, 0) <> "" Then Clear_BatchBox()
        ' Sms.Button1.Text = "hid"

        Dim i As Integer
        Dim telephone As String
        Dim CompanyName As String
        Dim email As String


        For i = 1 To 5000
            If Customer_Array(i - 1, 0) = "" Then Exit For
            batchbox.Rows.Add(1)


            CompanyName = Customer_Array(i - 1, 1) 'Fetch_This_Column(Customer_Field_Name, Customer_Array(i - 1), Customer_Table_Name)
            batchbox.Item(0, i - 1).Value = CompanyName

            telephone = Customer_Array(i - 1, 2) 'Fetch_This_Column(Customer_Contact_Field1, Customer_Array(i - 1), Customer_Table_Name)
            If telephone <> "" Then
                batchbox.Item(1, i - 1).Value = telephone

            Else
                batchbox.Item(1, i - 1).Value = Customer_Array(i - 1, 3) ' Fetch_This_Column(Customer_Contact_Field2, Customer_Array(i - 1), Customer_Table_Name)
            End If


            email = Customer_Array(i - 1, 4) ' Fetch_This_Column(Customer_Email_Field, Customer_Array(i - 1), Customer_Table_Name)
            batchbox.Item(4, i - 1).Value = email

            'Untick the check box when there is no Email
            If batchbox.Item(4, i - 1).Value = "" Then
                batchbox.Item(5, i - 1).Value = False
            Else
                batchbox.Item(5, i - 1).Value = True
            End If


            'Untick the check box when there is no Phone number
            If batchbox.Item(1, i - 1).Value = "" Then
                batchbox.Item(3, i - 1).Value = False
            Else
                batchbox.Item(3, i - 1).Value = True
            End If



        Next

        Set_Text_Of_TabControl1(0, Send_List_Title)
skip:

        InterModule_Transfer = False
    End Sub


    Private Sub Sms_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Is_Demo_Trial = True Then
            My.Settings.iM_Authentication = True
            My.Settings.Save()
        End If


        'Application.EnableVisualStyles()


        'New sms gateway Anaku
        My.Settings.SMS_Account_ProviderURL = "http://121.241.242.114:8080/sendsms?username="
        SMS_ProviderURL = "http://121.241.242.114:8080/sendsms?username="
        My.Settings.Save()


        If My.Settings.iM_Authentication = False Then
            MessageBox.Show("SMS Account Authentication is required. support@adroitbureau.com", "Intelligent Marketing", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If


        Load_MySettings_Variables()

        Refresh_SMS_Templates()
        'Fill sms template box.



        Fill_Combo_With_Towns(ToolStripComboBox3, Customer_Table_Name, Customer_Location_Field)
        'Fill_Combo_With_Towns(ToolStripComboBox2, Customer_Table_Name, "District")
        'Fill_Combo_With_Towns(ToolStripComboBox5, Customer_Table_Name, "School")

        Fill_Combo_With_Data(ToolStripComboBox5, Customer_Table_Name, Customer_Field_Name)
        Fill_Combo_With_Data(ToolStripComboBox2, Customer_Table_Name, District_Field)
        Fill_Combo_With_Data(ToolStripComboBox1, Customer_Table_Name, Custom_Field1)
        Fill_Combo_With_Data(ToolStripComboBox4, Customer_Table_Name, Custom_Field2)
        Fill_Combo_With_Data(ToolStripComboBox6, Customer_Table_Name, Custom_Field3)
        Fill_Combo_With_Data(ToolStripComboBox7, Customer_Table_Name, Custom_Field4)
        Fill_Combo_With_Data(ToolStripComboBox8, Customer_Table_Name, Custom_Field5)

        ToolStripComboBox1.Text = Custom_Field1
        ToolStripComboBox4.Text = Custom_Field2
        ToolStripComboBox6.Text = Custom_Field3
        ToolStripComboBox7.Text = Custom_Field4
        ToolStripComboBox8.Text = Custom_Field5


        'Fill_Combo_With_Customers(ToolStripComboBox5)

        'Fill_Combo_With_Reps(ToolStripComboBox1)


        SplitContainer2.Panel2Collapsed = True
        Display_Current_Database()

        Display_SMS_Account_Details()
        lblAccountName.Text = SMS_Account_Name

        lblSMSCreditBalance.Text = SMS_Current_Account_Balance      ' Display_SMS_Account_Details()




        'execute license routine
        LicenseRoutine()
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

    Sub Set_Text_Of_TabControl1(ByVal Tabindex As Integer, ByVal Text As String)
        TabControl1.TabPages(Tabindex).Text = Text

    End Sub
    Sub Clear_BatchBox()
        batchbox.Rows.Clear()
        Batch_Count = 0

        Send_List_Title = ""
        Set_Text_Of_TabControl1(0, Send_List_Title)
        reset_tabcontrol_captions()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()
        ' txtMessage.Text = ""
    End Sub

    Sub Display_Current_Database()
        If IsSQLSERVER Then
            ToolStripLabel_ConnectedDatabase.Text = SQL_Initial_Catalog
        End If

        If IsACCESS Then
            ToolStripLabel_ConnectedDatabase.Text = MSACESSFILE
        End If
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click

        If SplitContainer3.Panel1Collapsed = False Then

            SplitContainer3.Panel1Collapsed = True
            ToolStripButton1.Text = "Split pannel"

        Else
            SplitContainer3.Panel1Collapsed = False
            ToolStripButton1.Text = "Show Send list "

        End If

    End Sub

    Sub Load_All_Customers()


        Load_MySettings_Variables()
        Display_Current_Database()
        Clear_BatchBox()

        Dim Email As String

        If IsSQLSERVER Then

            Dim Mobile As String
            Dim Phone As String


            'Customized Variables
            Dim MyCustomer_Field_Name = "School"
            Dim MyCustomer_Contact_Field1 = "Telephone"
            Dim MyCustomer_Contact_Field2 = "Telephone2"
            Dim myCustomer_Email_Field = "Email"
            Dim myCustomer_Table_Name = "im_Data_SWL"



            Dim queryString As String = _
          "SELECT " & Customer_Field_Name & "," & Customer_Contact_Field1 & "," & Customer_Contact_Field2 & "," & Customer_Email_Field & " from " & Customer_Table_Name

            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()

                        batchbox.Rows.Add(1)

                        Phone = Trim(dataReader(1).ToString)
                        Mobile = Trim(dataReader(2).ToString)
                        Email = dataReader(3).ToString

                        batchbox.Item(0, Batch_Count).Value = dataReader(0).ToString
                        batchbox.Item(3, Batch_Count).Value = True

                        'Untick the check box when there is no number
                        If Phone = "" And Mobile = "" Then
                            batchbox.Item(3, Batch_Count).Value = False
                        Else
                            batchbox.Item(3, Batch_Count).Value = True
                        End If
                        batchbox.Item(4, Batch_Count).Value = Email

                        'Untick Checkbox when there is no number for EMAIL
                        If batchbox.Item(4, Batch_Count).Value = "" Then
                            batchbox.Item(5, Batch_Count).Value = False
                        Else
                            batchbox.Item(5, Batch_Count).Value = True
                        End If


                        'Utilize both phone fields
                        If Phone <> "" Then
                            batchbox.Item(1, Batch_Count).Value = Phone
                        Else
                            batchbox.Item(1, Batch_Count).Value = Mobile
                        End If


                        Batch_Count += 1
                    Loop

                    activate_checkbox(batchbox, smsCheck, emailCheck)
                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using
        End If

        '**********************************************************************
        Dim Phone_Number As String


        If IsACCESS Then

            Dim Email2 As String

            Dim queryString As String = _
          "SELECT " & Customer_Field_Name & "," & Customer_Contact_Field1 & "," & Customer_Contact_Field2 & "," & Customer_Email_Field & " from " & Customer_Table_Name


            '
            '  "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Data Directory\ESSAMUAH DATABASE.accdb"

            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        Email2 = dataReader(3).ToString
                        Phone_Number = Trim(dataReader(1).ToString)

                        batchbox.Rows.Add(1)
                        batchbox.Item(0, Batch_Count).Value = dataReader(0).ToString
                        batchbox.Item(1, Batch_Count).Value = Phone_Number
                        batchbox.Item(4, Batch_Count).Value = Email2


                        'Untick Checkbox when there is no number for EMAIL
                        If batchbox.Item(4, Batch_Count).Value = "" Then
                            batchbox.Item(5, Batch_Count).Value = False
                        Else
                            batchbox.Item(5, Batch_Count).Value = True
                        End If

                        If Phone_Number <> "" Then batchbox.Item(3, Batch_Count).Value = True

                        Batch_Count += 1
                    Loop
                    activate_checkbox(batchbox, smsCheck, emailCheck)
                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using
        End If





        'If IsPervasive Then

        '    Dim Email2 As String
        '    Dim id As String

        '    Dim queryString As String = _
        '  "SELECT " & Customer_Field_Name & "," & Customer_Contact_Field1 & "," & Customer_Contact_Field2 & "," & Customer_Email_Field & "," & Customer_Code_Field & " from " & Customer_Table_Name


        '    '
        '    '  "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Data Directory\ESSAMUAH DATABASE.accdb"

        '    Using connection As New PsqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
        '        Dim command As PsqlCommand = connection.CreateCommand()
        '        command.CommandText = queryString
        '        Try
        '            connection.Open()
        '            Dim dataReader As PsqlDataReader = _
        '             command.ExecuteReader()
        '            Do While dataReader.Read()
        '                Email2 = dataReader(3).ToString
        '                Phone_Number = Trim(dataReader(1).ToString)
        '                id = dataReader(4).ToString

        '                batchbox.Rows.Add(1)
        '                batchbox.Item(0, Batch_Count).Value = dataReader(0).ToString


        '                'Pastel
        '                If Contacts_Table_Name_Pastel = "" Then
        '                    batchbox.Item(1, Batch_Count).Value = Phone_Number

        '                Else
        '                    batchbox.Item(1, Batch_Count).Value = Fetch_Field_Value_From_Specified_Table_Given_ID(id, Contacts_Field_Name_Pastel, Contacts_Table_Name_Pastel)
        '                End If

        '                batchbox.Item(4, Batch_Count).Value = Email2


        '                'Untick Checkbox when there is no number for EMAIL
        '                If batchbox.Item(4, Batch_Count).Value = "" Then
        '                    batchbox.Item(5, Batch_Count).Value = False
        '                Else
        '                    batchbox.Item(5, Batch_Count).Value = True
        '                End If

        '                If Phone_Number <> "" Then batchbox.Item(3, Batch_Count).Value = True

        '                Batch_Count += 1
        '            Loop
        '            activate_checkbox(batchbox, smsCheck, emailCheck)
        '            dataReader.Close()

        '        Catch ex As Exception
        '            MessageBox.Show(ex.Message)
        '        End Try

        '        connection.Close()
        '    End Using
        'End If


        Send_List_Title = batchbox.Rows.Count.ToString & " Contacts"
        Set_Text_Of_TabControl1(0, Send_List_Title)

        Enforce_Demo_Restriction()
    End Sub

    Sub Enforce_Demo_Restriction()
        If Is_Demo_Restriction_Exceeded() = True Then
            Clear_BatchBox()
        End If
    End Sub







    Sub Display_Number_of_Contacts()
        Send_List_Title = batchbox.Rows.Count.ToString & " Contacts"
        Set_Text_Of_TabControl1(0, Send_List_Title)
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripComboBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripComboBox3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripDropDownButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripComboBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ToolStripComboBox5.Text = ""
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub
    Sub reset_tabcontrol_captions()
        Set_Text_Of_TabControl1(1, "Sent")
        Set_Text_Of_TabControl1(2, "Not Sent")
    End Sub
    Sub reset_send_listboxes()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()

    End Sub
    Sub Execute_Send_SMS()
        Load_MySettings_Variables()
        reset_tabcontrol_captions()


        Dim i As Integer
        Dim phone_Number As String
        Dim namex As String
        Dim AccountBalance As Double
        Dim SMS_Message As String

        SMS_Message = txtmessage.Text



        For i = 0 To Batch_Count - 1
            SMS_Message = txtmessage.Text

            phone_Number = Trim(batchbox.Item(1, i).Value)
            namex = batchbox.Item(0, i).Value
            AccountBalance = MyCDBL(batchbox.Item(2, i).Value)
            phone_Number = Format_Phone_number(phone_Number)
            If phone_Number = "" Then GoTo skip
            If phone_Number = "null" Then GoTo skip
            If batchbox.Item(3, i).Value = False Then GoTo skip


            'Place Account Balance in SMS if required
            If AccountBalance < 30 And By_Debtors = True Then GoTo skip
            SMS_Message = Replace(SMS_Message, "#debtvalue", AccountBalance.ToString)
            SMS_Message = Replace(SMS_Message, "#name", namex.ToString)
            '  MessageBox.Show(SMS_Message)


            SEND_SMS_Output_to_Listbox(SMS_Message, phone_Number, SMS_SenderID, 0, ListBox2, ListBox3, namex)
skip:
            ToolStripProgressBar1.Value = (CInt((i / Batch_Count) * 100))
        Next

        ' Send the messages to Managers Also
        ' SEND_SMS_Output_to_Listbox(SMS_Message, "233541235271", SMS_SenderID, 0, ListBox1, ListBox2, "Kojo Sam-Woode")
        'SEND_SMS_Output_to_Listbox(txtMessage.Text, "233233200204", "Essamuah", 0, ListBox1, ListBox2, "Frank Acheampong")

        SEND_SMS_Report_to_Adroit(Adroit_SMS_Report, "233541235271", "iM Report", 0, "iM")

        MessageBox.Show("SMS Routine finished", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ToolStripProgressBar1.Value = 0

        Display_SMS_Account_Details()
        lblAccountName.Text = SMS_Account_Name

        lblSMSCreditBalance.Text = SMS_Current_Account_Balance      ' Display_SMS_Account_Details()

        'Display Send Status in Tab
        Set_Text_Of_TabControl1(1, "Sent (" & SMS_Successufully_Sent & ")")
        Set_Text_Of_TabControl1(2, "Not Sent (" & SMS_Unsuccessfully_Sent & ")")

    End Sub
    Function Adroit_SMS_Report() As String
        Dim Report As String
        Report = SMS_Account_Name & " Sent: " & SMS_Successufully_Sent & ", " & Check_SMS_Account_Credits(SMS_Account_Name, SMS_Account_Password) & ", " & "Not Sent: " & SMS_Unsuccessfully_Sent
        Return Report
    End Function


    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'Fill_Combo_With_Towns(ToolStripComboBox2, Customer_Table_Name, "District")

        'Fill_Combo_With_Customers(ToolStripComboBox5)

        'Fill_Combo_With_Reps(ToolStripComboBox5)
        'Timer1.Enabled = False

    End Sub

    Private Sub ToolStripDropDownButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton1.Click

    End Sub

    Private Sub ToolStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

    End Sub

    Private Sub ToolStripMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        By_Debtors = False
        Load_All_Customers()
        Send_List_Title = "All Contacts"
        Set_Text_Of_TabControl1(0, Send_List_Title)

    End Sub

    Private Sub ToolStripComboBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox2.Click

    End Sub

    Private Sub ToolStripComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox2.SelectedIndexChanged

        'By_Debtors = False

        'Display_List_by_Town(ToolStripComboBox2.Text)
        'Send_List_Title = ToolStripDropDownButton1.Text & ": " & ToolStripComboBox2.Text
        'Set_Text_Of_TabControl1(0, Send_List_Title)


        Display_List_by_District(ToolStripComboBox2.Text)
        Enforce_Demo_Restriction()
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim test As String

        'MessageBox.Show(InStr("balance:23.3", ":").ToString)


        MessageBox.Show(test)

    End Sub

    Private Sub ToolStripComboBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripComboBox5_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        'By_Debtors = False
        'Display_List_by_Customer(ToolStripComboBox5.Text)
        'Send_List_Title = ToolStripComboBox5.Text
        'Set_Text_Of_TabControl1(0, Send_List_Title)

        Load_Schools_Based_On_Like_For_School_Name(ToolStripComboBox5.Text)
        Enforce_Demo_Restriction()
    End Sub



    'Private Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ToolStripComboBox5.SelectedIndexChanged
    '    'By_Debtors = False
    '    'Display_List_by_Representative(ToolStripComboBox5.Text)

    '    'Send_List_Title = ToolStripDropDownButton5.Text & ": " & ToolStripComboBox5.Text
    '    'Set_Text_Of_TabControl1(0, Send_List_Title)
    'End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Sub Execute_Test_SMS()
        Load_MySettings_Variables()
        ListBox4.Items.Clear()

        Dim i As Integer
        Dim phone_Number As String
        Dim namex As String
        Dim AccountBalance As Double
        Dim SMS_Message As String

        SMS_Message = txtmessage.Text



        For i = 0 To Batch_Count - 1
            SMS_Message = txtmessage.Text

            phone_Number = batchbox.Item(1, i).Value
            namex = batchbox.Item(0, i).Value
            AccountBalance = MyCDBL(batchbox.Item(2, i).Value)
            phone_Number = Format_Phone_number(phone_Number)
            If phone_Number = "" Then GoTo skip
            If phone_Number = "null" Then GoTo skip
            If batchbox.Item(3, i).Value = False Then GoTo skip


            'Place Account Balance in SMS if required
            If AccountBalance < 30 And By_Debtors = True Then GoTo skip
            SMS_Message = Replace(SMS_Message, "#debtvalue", AccountBalance.ToString)
            SMS_Message = Replace(SMS_Message, "#name", namex.ToString)
            '  MessageBox.Show(SMS_Message)


            ListBox4.Items.Add(SMS_Message & " " & phone_Number & " " & SMS_SenderID & " " & namex)
skip:
            ToolStripProgressBar1.Value = (CInt((i / Batch_Count) * 100))
        Next

        ' Send the messages to Managers Also
        'ListBox2.Items.Add(SMS_Message & " 233541235271 " & SMS_SenderID, 0, ListBox1, ListBox2, " Kojo Sam-Woode")
        'SEND_SMS_Output_to_Listbox(txtMessage.Text, "233233200204", "Essamuah", 0, ListBox1, ListBox2, "Frank Acheampong")

        MessageBox.Show("SMS Routine finished", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ToolStripProgressBar1.Value = 0

        Display_SMS_Account_Details()
    End Sub

    Private Sub AllDebtorsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        By_Debtors = True
        Display_AR_Listing()

        Send_List_Title = "Debtors List"
        Set_Text_Of_TabControl1(0, Send_List_Title)
    End Sub

    Private Sub ListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.Click
        txtmessage.Text = ListBox1.SelectedItem
        txtEmail.Text = ListBox1.SelectedItem

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        My.Settings.SMS_Templates.Add(txtmessage.Text)
        Refresh_SMS_Templates()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        'Fill_Combo_With_Towns(ToolStripComboBox2, Customer_Table_Name, "District")

        'Fill_Combo_With_Customers(ToolStripComboBox5)

        'Fill_Combo_With_Reps(ToolStripComboBox5)
    End Sub

    Private Sub batchbox_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles batchbox.CellContentClick

    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '  Intelligence_To_SMS()

        Clear_BatchBox()
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



    Private Sub Sms_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown


    End Sub


    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Reset_SMS_Sending_Status_Variables()
        Execute_Send_SMS()
    End Sub


    Private Sub chkAttachment_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAttachment.CheckedChanged
        If chkAttachment.Checked = True Then
            btnBrowse.Enabled = True
            txtAttachment.Enabled = True
        Else
            btnBrowse.Enabled = False
            txtAttachment.Enabled = False
        End If
    End Sub

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        OpenFileDialog1.ShowDialog()
        txtAttachment.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Reset_Email_Sending_Status_Variables()
        Load_MySettings_Variables()


        Dim i As Integer
        Dim EmailAddress As String
        Dim namex As String
        Dim AccountBalance As Double
        Dim Email_Message As String
        Dim Email_report As String

        Email_Message = txtEmail.Text



        For i = 0 To Batch_Count - 1
            Email_Message = txtEmail.Text

            EmailAddress = batchbox.Item(4, i).Value
            namex = batchbox.Item(0, i).Value
            AccountBalance = MyCDBL(batchbox.Item(2, i).Value)
            If EmailAddress = "" Then GoTo skip
            If EmailAddress = "null" Then GoTo skip
            If batchbox.Item(5, i).Value = False Then GoTo skip


            'Place Account Balance in SMS if required
            If AccountBalance < 30 And By_Debtors = True Then GoTo skip
            Email_Message = Replace(Email_Message, "#debtvalue", AccountBalance.ToString)
            Email_Message = Replace(Email_Message, "#name", namex.ToString)

            '  MessageBox.Show(SMS_Message)


            Email_report = Send_Email_With_Given_Parameters(Email_Account, Email_Pass, SMTP_server_Port, SMTP_server_Host, SMTP_server_Enable_SSL, chkAttachment.Checked, txtAttachment.Text, Email_Source, EmailAddress, txtEmail.Text, txtEmail_Subject.Text, ListBox2, ListBox3)
            'SEND_SMS_Output_to_Listbox(SMS_Message, phone_Number, SMS_SenderID, 0, ListBox2, ListBox3, namex)
skip:
            ToolStripProgressBar1.Value = (CInt((i / Batch_Count) * 100))
        Next

        ' Send the messages to Managers Also
        ' SEND_SMS_Output_to_Listbox(SMS_Message, "233541235271", SMS_SenderID, 0, ListBox1, ListBox2, "Kojo Sam-Woode")
        'SEND_SMS_Output_to_Listbox(txtMessage.Text, "233233200204", "Essamuah", 0, ListBox1, ListBox2, "Frank Acheampong")

        MessageBox.Show("Email Routine finished", "Email", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ToolStripProgressBar1.Value = 0

        'Display Send Status in Tab
        Set_Text_Of_TabControl1(1, "Sent (" & Email_Successfully_Sent & ")")
        Set_Text_Of_TabControl1(2, "Not Sent (" & Email_Unsuccessfully_Sent & ")")

    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Execute_Test_SMS()
    End Sub

    Private Sub ToolStripButton2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Display_SMS_Account_Details()

        If SMS_Current_Account_Balance = "" Then
            MessageBox.Show("SMS Account Details Not Entered")
        Else
            MessageBox.Show(SMS_Current_Account_Balance)
        End If



    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        My.Settings.SMS_Templates.Add(txtmessage.Text)
        My.Settings.Save()
        Refresh_SMS_Templates()

    End Sub

    Private Sub Button4_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        My.Settings.SMS_Templates.Add(txtEmail.Text)
        My.Settings.Save()
        Refresh_SMS_Templates()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Set_Text_Of_TabControl1(0, "Info")
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        System.Diagnostics.Process.Start(MSACCESSFILE_Path & "\" & MSACESSFILE & ".accdb")

    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        Dim path As String
        Dim filename As String
        Dim filename_Without_Extension As String
        Dim Filename_Including_Path As String
        Dim i As Integer

        Clear_BatchBox()

        OpenFileDialog1.ShowDialog()
        Filename_Including_Path = OpenFileDialog1.FileName
        filename = OpenFileDialog1.SafeFileName
        path = Replace(Filename_Including_Path, filename, "")

        If filename = "OpenFileDialog1" Or filename = "" Then GoTo skip

        Dim stream As FileStream = File.Open(Filename_Including_Path, FileMode.Open, FileAccess.Read)

        Dim excelReader As Excel.IExcelDataReader = Excel.ExcelReaderFactory.CreateOpenXmlReader(stream)
        '...IExcelDataReader
        '3. DataSet - The result of each spreadsheet will be created in the result.Tables
        Dim result As DataSet = excelReader.AsDataSet()
        '...
        '4. DataSet - Create column names from first row
        excelReader.IsFirstRowAsColumnNames = True
        Dim result1 As DataSet = excelReader.AsDataSet()

        '5. Data Reader methods

        Dim Contact_Name As String
        Dim Contact_Phone_Number As String
        Dim Amount_Owing As Double
        Dim Email As String


        While excelReader.Read()
            batchbox.Rows.Add(1)

            Try
                Contact_Name = excelReader(0).ToString
            Catch ex As Exception
                Contact_Name = ""
            End Try


            Try
                Contact_Phone_Number = excelReader(1).ToString
                If Contact_Phone_Number <> "" Or Contact_Phone_Number <> "null" Then batchbox.Item(3, i).Value = True
            Catch ex As Exception
                Contact_Phone_Number = ""
            End Try


            'Try
            '    Amount_Owing = excelReader(2).ToString
            'Catch ex As Exception
            '    Amount_Owing = 0
            'End Try

            'Try
            '    Email = excelReader(3).ToString
            'Catch ex As Exception
            '    Email = ""
            'End Try




            batchbox.Item(0, i).Value = Contact_Name
            batchbox.Item(1, i).Value = Contact_Phone_Number
            'batchbox.Item(2, i).Value = Amount_Owing
            'batchbox.Item(4, i).Value = Email

            'If excelReader(1) = Nothing Then
            '    ListBox1.Items.Add("")

            'Else
            '    ListBox1.Items.Add(excelReader(1).ToString)
            'End If

            '  MessageBox.Show(excelReader(1).ToString)
            i += 1
        End While

        Send_List_Title = i & "Contacts (" & filename & ")"
        Set_Text_Of_TabControl1(0, Send_List_Title)

skip:
    End Sub

    Private Sub ToolStripDropDownButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton3.Click

    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim f As New settings
        f.Show()

    End Sub

    Private Sub AllSchoolsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllSchoolsToolStripMenuItem.Click
        Load_All_Customers()
    End Sub

    Private Sub PrivateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Load_Schools_Based_On_Status("Private")
    End Sub

    Private Sub PublicToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Load_Schools_Based_On_Status("Public")
    End Sub

    Private Sub SeniorHighToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SeniorHighToolStripMenuItem.Click

    End Sub

    Private Sub UniversityToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UniversityToolStripMenuItem.Click

    End Sub

    Private Sub JuniorHighToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles JuniorHighToolStripMenuItem.Click

    End Sub

    Private Sub PrimaryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrimaryToolStripMenuItem.Click

    End Sub

    Private Sub KindergartenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KindergartenToolStripMenuItem.Click

    End Sub

    Private Sub NurseryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NurseryToolStripMenuItem.Click

    End Sub

    Private Sub GreaterAccraToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GreaterAccraToolStripMenuItem.Click

    End Sub

    Private Sub CentralRegionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CentralRegionToolStripMenuItem.Click

    End Sub

    Private Sub WesternRegionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WesternRegionToolStripMenuItem.Click

    End Sub

    Private Sub EasternRegionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EasternRegionToolStripMenuItem.Click

    End Sub

    Private Sub AshantiRegionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AshantiRegionToolStripMenuItem.Click

    End Sub

    Private Sub VoltaRegionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VoltaRegionToolStripMenuItem.Click

    End Sub

    Private Sub BrongAhafoRegionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BrongAhafoRegionToolStripMenuItem.Click

    End Sub

    Private Sub NorthernRegionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NorthernRegionToolStripMenuItem.Click

    End Sub

    Private Sub UpperEastRegionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpperEastRegionToolStripMenuItem.Click

    End Sub

    Private Sub UpperWestRegionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpperWestRegionToolStripMenuItem.Click

    End Sub

    Private Sub ToolStripMenuItem1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click

    End Sub

    Private Sub IslamicSchoolsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IslamicSchoolsToolStripMenuItem.Click
        Load_Schools_Based_On_Like_For_School_Name("%Islamic%")
    End Sub

    Private Sub MethodistToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MethodistToolStripMenuItem.Click
        Load_Schools_Based_On_Like_For_School_Name("%Meth%")
    End Sub

    Private Sub PresbyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PresbyToolStripMenuItem.Click
        Load_Schools_Based_On_Like_For_School_Name("%Presb%")
    End Sub

    Private Sub CatholicToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CatholicToolStripMenuItem.Click
        Load_Schools_Based_On_Like_For_School_Name("%Cath%")
    End Sub

    Private Sub SpecialSchoolsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SpecialSchoolsToolStripMenuItem.Click
        Load_Schools_Based_On_Like_For_School_Name("%Special%")
    End Sub

    Private Sub ToolStripComboBox3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox3.Click

    End Sub

    Private Sub ToolStripComboBox3_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox3.SelectedIndexChanged
        ''Display_List_by_Town(ToolStripComboBox3.Text)
        Display_List_by_Custome_Field(ToolStripComboBox3.Text, Customer_Location_Field)
        Enforce_Demo_Restriction()
    End Sub

    Private Sub AllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AllToolStripMenuItem.Click
        Load_Schools_Based_On_Region("%Accra%")
    End Sub

    Private Sub PrivateSchoolsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrivateSchoolsToolStripMenuItem.Click
        Load_Schools_Based_On_Region_and_Status("%Accra%", "Private")
    End Sub

    Private Sub PublicSchoolsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PublicSchoolsToolStripMenuItem.Click
        Load_Schools_Based_On_Region_and_Status("%Accra%", "Public")
    End Sub

    Private Sub ToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem3.Click
        Load_Schools_Based_On_Region_and_Status("%Central%", "Private")
    End Sub

    Private Sub ToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem4.Click
        Load_Schools_Based_On_Region_and_Status("%Central%", "Public")
    End Sub

    Private Sub ToolStripMenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem9.Click
        Load_Schools_Based_On_Region_and_Status("%Eastern%", "Private")
    End Sub

    Private Sub ToolStripMenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem10.Click
        Load_Schools_Based_On_Region_and_Status("%Eastern%", "Public")
    End Sub

    Private Sub ToolStripMenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem6.Click
        Load_Schools_Based_On_Region_and_Status("%Western%", "Private")
    End Sub

    Private Sub ToolStripMenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem7.Click
        Load_Schools_Based_On_Region_and_Status("%Western%", "Public")
    End Sub

    Private Sub ToolStripMenuItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem12.Click
        Load_Schools_Based_On_Region_and_Status("%Ashanti%", "Private")
    End Sub

    Private Sub ToolStripMenuItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem13.Click
        Load_Schools_Based_On_Region_and_Status("%Ashanti%", "Public")
    End Sub

    Private Sub ToolStripMenuItem15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem15.Click
        Load_Schools_Based_On_Region_and_Status("%Volta%", "Private")
    End Sub

    Private Sub ToolStripMenuItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem16.Click
        Load_Schools_Based_On_Region_and_Status("%Volta%", "Public")
    End Sub

    Private Sub ToolStripMenuItem18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem18.Click
        Load_Schools_Based_On_Region_and_Status("%Brong%", "Private")
    End Sub

    Private Sub ToolStripMenuItem19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem19.Click
        Load_Schools_Based_On_Region_and_Status("%Brong%", "Public")
    End Sub

    Private Sub ToolStripMenuItem21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem21.Click
        Load_Schools_Based_On_Region_and_Status("%Northern%", "Private")
    End Sub

    Private Sub ToolStripMenuItem22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem22.Click
        Load_Schools_Based_On_Region_and_Status("%Northern%", "Public")
    End Sub

    Private Sub ToolStripMenuItem24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem24.Click
        Load_Schools_Based_On_Region_and_Status("%Upper East%", "Private")
    End Sub

    Private Sub ToolStripMenuItem25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem25.Click
        Load_Schools_Based_On_Region_and_Status("%Upper East%", "Public")
    End Sub

    Private Sub ToolStripMenuItem27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem27.Click
        Load_Schools_Based_On_Region_and_Status("%Upper West%", "Private")
    End Sub

    Private Sub ToolStripMenuItem28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem28.Click
        Load_Schools_Based_On_Region_and_Status("%Upper West%", "Public")
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        Load_Schools_Based_On_Region("%Central%")
    End Sub

    Private Sub ToolStripMenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem8.Click
        Load_Schools_Based_On_Region("%Eastern%")
    End Sub

    Private Sub ToolStripMenuItem5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem5.Click
        Load_Schools_Based_On_Region("%Western%")
    End Sub

    Private Sub ToolStripMenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem11.Click
        Load_Schools_Based_On_Region("%Ashanti%")
    End Sub

    Private Sub ToolStripMenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem14.Click
        Load_Schools_Based_On_Region("%Volta%")
    End Sub

    Private Sub ToolStripMenuItem17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem17.Click
        Load_Schools_Based_On_Region("%Brong%")
    End Sub

    Private Sub ToolStripMenuItem20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem20.Click
        Load_Schools_Based_On_Region("%Northern%")
    End Sub

    Private Sub ToolStripMenuItem23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem23.Click
        Load_Schools_Based_On_Region("%Upper East%")
    End Sub

    Private Sub ToolStripMenuItem26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem26.Click
        Load_Schools_Based_On_Region("%Upper West%")
    End Sub

    Private Sub ToolStripComboBox5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox5.Click

    End Sub

    Private Sub ToolStripComboBox5_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox5.SelectedIndexChanged
        Load_Schools_Based_On_Like_For_School_Name(ToolStripComboBox5.Text)
        Enforce_Demo_Restriction()
    End Sub

    Private Sub ToolStripMenuItem30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem30.Click
        Load_Schools_Based_On_Level_and_Status("%University%", "Private")
        Enforce_Demo_Restriction()
    End Sub

    Private Sub ToolStripMenuItem29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem29.Click
        Load_Schools_Based_On_Like_For_School_Name("%University%")
        Enforce_Demo_Restriction()
    End Sub

    Private Sub ToolStripMenuItem32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem32.Click
        Load_Schools_Based_On_Like_For_School_Name("%Senior%")
        Enforce_Demo_Restriction()
    End Sub

    Private Sub ToolStripMenuItem35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem35.Click
        Load_Schools_Based_On_Like_For_School_Name("%Junior%")
        Enforce_Demo_Restriction()
    End Sub

    Private Sub ToolStripMenuItem38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem38.Click
        Load_Schools_Based_On_Like_For_School_Name("%Primary%")
    End Sub

    Private Sub ToolStripMenuItem41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem41.Click
        Load_Schools_Based_On_Like_For_School_Name("%Kindergarten%")
        Enforce_Demo_Restriction()
    End Sub



    Private Sub ToolStripMenuItem44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem44.Click
        Load_Schools_Based_On_Like_For_School_Name("%Nursery%")
        Enforce_Demo_Restriction()
    End Sub

    Private Sub ToolStripMenuItem31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem31.Click
        Load_Schools_Based_On_Level_and_Status("%University%", "Public")
    End Sub

    Private Sub ToolStripMenuItem33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem33.Click
        Load_Schools_Based_On_Level_and_Status("%Senior%", "Private")
    End Sub

    Private Sub ToolStripMenuItem34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem34.Click
        Load_Schools_Based_On_Level_and_Status("%Senior%", "Public")
    End Sub

    Private Sub ToolStripMenuItem36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem36.Click
        Load_Schools_Based_On_Level_and_Status("%Junior%", "Private")
    End Sub

    Private Sub ToolStripMenuItem37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem37.Click
        Load_Schools_Based_On_Level_and_Status("%Junior%", "Public")
    End Sub

    Private Sub ToolStripMenuItem39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem39.Click
        Load_Schools_Based_On_Level_and_Status("%Primary%", "Private")
    End Sub

    Private Sub ToolStripMenuItem40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem40.Click
        Load_Schools_Based_On_Level_and_Status("%Primary%", "Public")
    End Sub

    Private Sub ToolStripMenuItem42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem42.Click
        Load_Schools_Based_On_Level_and_Status("%Kindergarten%", "Private")
    End Sub

    Private Sub ToolStripMenuItem43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem43.Click
        Load_Schools_Based_On_Level_and_Status("%Kindergarten%", "Public")
    End Sub

    Private Sub ToolStripMenuItem45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem45.Click
        Load_Schools_Based_On_Level_and_Status("%Nursery%", "Private")
    End Sub

    Private Sub ToolStripMenuItem46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem46.Click
        Load_Schools_Based_On_Level_and_Status("%Nursery%", "Public")
    End Sub

    Private Sub ToolStripDropDownButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton5.Click

    End Sub

    Private Sub ToolStripDropDownButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton6.Click

    End Sub

    Private Sub ToolStripDropDownButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton7.Click

    End Sub

    Private Sub ToolStripMenuItem47_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem47.Click
        Revenue_Collection()

    End Sub

    Public Sub Revenue_Collection()
        By_Debtors = True
        Display_AR_Listing()

        txtmessage.Text = "dear #name your account is overdue to the tune of GHS #debtvalue"

        Send_List_Title = "Debtors List"
        Set_Text_Of_TabControl1(0, Send_List_Title)

        Enforce_Demo_Restriction()
    End Sub

    Private Sub ToolStripComboBox1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.Click

    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        Display_List_by_Custome_Field(ToolStripComboBox1.Text, Custom_Field1)
        Enforce_Demo_Restriction()
    End Sub

    Private Sub ToolStripComboBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox4.Click

    End Sub

    Private Sub ToolStripComboBox4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox4.SelectedIndexChanged
        Display_List_by_Custome_Field(ToolStripComboBox4.Text, Custom_Field2)
        Enforce_Demo_Restriction()
    End Sub

    Private Sub ToolStripComboBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox6.Click

    End Sub

    Private Sub ToolStripComboBox6_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox6.SelectedIndexChanged
        Display_List_by_Custome_Field(ToolStripComboBox6.Text, Custom_Field3)
        Enforce_Demo_Restriction()
    End Sub

    Private Sub ToolStripComboBox7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox7.Click

    End Sub

    Private Sub ToolStripComboBox7_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox7.SelectedIndexChanged
        Display_List_by_Custome_Field(ToolStripComboBox7.Text, Custom_Field4)
        Enforce_Demo_Restriction()
    End Sub

    Private Sub ToolStripComboBox8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripComboBox8.Click

    End Sub

    Private Sub ToolStripComboBox8_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStripComboBox8.SelectedIndexChanged
        Display_List_by_Custome_Field(ToolStripComboBox8.Text, Custom_Field5)
        Enforce_Demo_Restriction()
    End Sub

    Private Sub ToolStripDropDownButton2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton2.Click

    End Sub

    Private Sub ToolStripButton6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Clear_BatchBox()

        Load_MySettings_Variables()





        Fill_Combo_With_Towns(ToolStripComboBox3, Customer_Table_Name, Customer_Location_Field)
        'Fill_Combo_With_Towns(ToolStripComboBox2, Customer_Table_Name, "District")
        'Fill_Combo_With_Towns(ToolStripComboBox5, Customer_Table_Name, "School")

        Fill_Combo_With_Data(ToolStripComboBox5, Customer_Table_Name, Customer_Field_Name)
        Fill_Combo_With_Data(ToolStripComboBox2, Customer_Table_Name, District_Field)
        Fill_Combo_With_Data(ToolStripComboBox1, Customer_Table_Name, Custom_Field1)
        Fill_Combo_With_Data(ToolStripComboBox4, Customer_Table_Name, Custom_Field2)
        Fill_Combo_With_Data(ToolStripComboBox6, Customer_Table_Name, Custom_Field3)
        Fill_Combo_With_Data(ToolStripComboBox7, Customer_Table_Name, Custom_Field4)
        Fill_Combo_With_Data(ToolStripComboBox8, Customer_Table_Name, Custom_Field5)
    End Sub

    Private Sub TabPage5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage5.Click

    End Sub

    Private Sub Panel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub ToolStripMenuItem48_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem48.Click
        Dim fs As New Number_Formatter
        fs.Show()


    End Sub

    Private Sub OutputFormToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OutputFormToolStripMenuItem.Click
        Dim fs As New Pervasive_Connection
        fs.Show()

    End Sub

    Private Sub ToolStrip2_ItemClicked(sender As System.Object, e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStrip2.ItemClicked

    End Sub

    Private Sub Button5_Click_2(sender As System.Object, e As System.EventArgs)
        Dim f As New Locations_Data_Validation
        f.Show()


    End Sub



    Private Sub smsCheck_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles smsCheck.Click
        'tick or untick the checkboxes 
        check_checkbox(smsCheck, 3, 1, batchbox)
    End Sub

    Private Sub emailCheck_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles emailCheck.Click
        'tick or untick the checkboxes 
        check_checkbox(emailCheck, 5, 4, batchbox)
    End Sub

    Private Sub ToolStripDropDownButton8_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripDropDownButton8.Click
        'If Is_Demo_Trial = True Then
        '    MessageBox.Show()
        'End If


    End Sub
End Class