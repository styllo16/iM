Imports System.Data
Imports Pervasive.Data.SqlClient
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Web 'add system.web as reference for the url encoding class
Imports System.Net

Public Class Regional_Drill_Down

    Private Sub Regional_Drill_Down_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Load_Regional_Drill_Down_Data_Into_Drill_Down_Datagrid(Selected_Region_From_Attribute_Table)
    End Sub
    Sub Load_Regional_Drill_Down_Data_Into_Drill_Down_Datagrid(ByVal Region As String)
        Dim i As Integer
        Blank_Customer_Array()
        Dim current_code As String


        Dim queryString As String = _
           "SELECT " & Customer_Code_Field & ", " & Customer_Field_Name & ", " & Customer_Contact_Field1 & ", " & District_Field & ", " & Customer_Location_Field & " from " & Customer_Table_Name & " where " & Region_Field & " like '%" & Region & "%' or " & Region_Field & " like '%" & Selected_Region_From_Attribute_Table_focus_specific & "%' "


        If IsSQLSERVER Then
            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        current_code = dataReader(0).ToString
                        dg_Regional_Drill_Down.Rows.Add(1)

                        dg_Regional_Drill_Down.Item(0, i).Value = i + 1
                        dg_Regional_Drill_Down.Item(1, i).Value = current_code
                        dg_Regional_Drill_Down.Item(2, i).Value = dataReader(1).ToString
                        dg_Regional_Drill_Down.Item(3, i).Value = dataReader(2).ToString
                        dg_Regional_Drill_Down.Item(4, i).Value = dataReader(3).ToString
                        'dg_Regional_Drill_Down.Item(5, i).Value = dataReader(4).ToString
                        'dg_Regional_Drill_Down.Item(6, i).Value = dataReader(5).ToString
                        'dg_Regional_Drill_Down.Item(7, i).Value = dataReader(6).ToString
                        'dg_Regional_Drill_Down.Item(8, i).Value = dataReader(7).ToString

                        Customer_Array(i, 0) = current_code
                        i += 1
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
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
                        current_code = dataReader(0).ToString
                        dg_Regional_Drill_Down.Rows.Add(1)

                        dg_Regional_Drill_Down.Item(0, i).Value = i + 1
                        dg_Regional_Drill_Down.Item(1, i).Value = current_code
                        dg_Regional_Drill_Down.Item(2, i).Value = dataReader(1).ToString
                        dg_Regional_Drill_Down.Item(3, i).Value = dataReader(2).ToString
                        dg_Regional_Drill_Down.Item(4, i).Value = dataReader(3).ToString
                        'dg_Regional_Drill_Down.Item(5, i).Value = dataReader(4).ToString
                        'dg_Regional_Drill_Down.Item(6, i).Value = dataReader(5).ToString
                        'dg_Regional_Drill_Down.Item(7, i).Value = dataReader(6).ToString
                        'dg_Regional_Drill_Down.Item(8, i).Value = dataReader(7).ToString


                        Customer_Array(i, 0) = current_code
                        i += 1
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
        '        command.CommandText = queryString
        '        Try
        '            connection.Open()
        '            Dim dataReader As PsqlDataReader = _
        '             command.ExecuteReader()
        '            Do While dataReader.Read()
        '                current_code = dataReader(0).ToString
        '                dg_Regional_Drill_Down.Rows.Add(1)

        '                dg_Regional_Drill_Down.Item(0, i).Value = i + 1
        '                dg_Regional_Drill_Down.Item(1, i).Value = current_code
        '                dg_Regional_Drill_Down.Item(2, i).Value = dataReader(1).ToString
        '                dg_Regional_Drill_Down.Item(3, i).Value = dataReader(2).ToString
        '                dg_Regional_Drill_Down.Item(4, i).Value = dataReader(3).ToString
        '                dg_Regional_Drill_Down.Item(5, i).Value = dataReader(4).ToString
        '                'dg_Regional_Drill_Down.Item(6, i).Value = dataReader(5).ToString
        '                'dg_Regional_Drill_Down.Item(7, i).Value = dataReader(6).ToString
        '                'dg_Regional_Drill_Down.Item(8, i).Value = dataReader(7).ToString

        '                Customer_Array(i, 0) = current_code
        '                i += 1
        '            Loop

        '            dataReader.Close()

        '        Catch ex As Exception
        '            MessageBox.Show(ex.Message)
        '        End Try

        '        connection.Close()
        '    End Using
        'End If



    End Sub

   



    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        Select_Digital_Marketing()
        'Intelligence_To_SMS()
        Me.Close()

        'Intelligence_To_SMS(Sms.batchbox)


    End Sub

    Sub Select_Digital_Marketing()
        SMS_sender.TabControl1.TabPages.Item(0).Select()
        SMS_sender.TabControl1.TabPages.Item(0).Focus()
    End Sub

    Private Sub dg_Regional_Drill_Down_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dg_Regional_Drill_Down.CellContentClick

    End Sub
End Class