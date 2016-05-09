Imports System.Data
Imports Pervasive.Data.SqlClient
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Web 'add system.web as reference for the url encoding class
Imports System.Net
Imports Microsoft.Reporting.WinForms


Public Class SWL_Customer_Dashboard

    Dim TOP As Integer = 10

    Private Sub SWL_Customer_Dashboard_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        '************************************************************
        ' Authentication Code
        If My.Settings.iM_Authentication = False Then
            ToolStrip1.Enabled = False
            ReportViewer1.Enabled = False
        Else
            ToolStrip1.Enabled = True
            ReportViewer1.Enabled = True
        End If

        '************************************************************
    End Sub
    Private Sub SWL_Customer_Dashboard_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'Intelligent_Marketing_SchemaDataSet.Basic_Graph' table. You can move, or remove it, as needed.
        '   Me.Basic_GraphTableAdapter.Fill(Me.Intelligent_Marketing_SchemaDataSet.Basic_Graph)

        '  Me.ReportViewer1.RefreshReport()


        'Me.ReportViewer1.RefreshReport()
        'TOP = CInt(ToolStripComboBox1.Text)
        'Load_DataSet_With_Top_Customers(Customer_Sales_Archive_Field)
        'Send_List_Title = "All time Best Customers"

        Load_MySettings_Variables()

        If Is_Demo_Restriction_Exceeded() = True Then
            MessageBox.Show("Business Intelligence is Disabled. Contact sales@adroitbureau.com", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Enabled = False
        End If


    End Sub
    Sub Blank_Customer_Array()
        Dim i As Integer
        Dim i2 As Integer


        For i = 1 To 5000
            For i2 = 1 To 5
                Customer_Array(i - 1, i2 - 1) = ""
            Next
        Next

    End Sub


    Function SWL_Fetch_Number_of_Customers_in_Database() As Integer
        Dim queryString As String = _
            "SELECT count(Company) from Dealers"

        Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
            Dim command As SqlCommand = connection.CreateCommand()
            command.CommandText = queryString
            Try
                connection.Open()
                Dim dataReader As SqlDataReader = _
                 command.ExecuteReader()
                Do While dataReader.Read()


                    Return CInt(dataReader(0))
                Loop

                dataReader.Close()

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            connection.Close()
        End Using




    End Function


    Function SWL_Fetch_Number_of_Customers_General(ByVal query As String) As Integer
        Dim queryString As String = query
        '   "SELECT count(Company) from Dealers where mobile = '' and phone = ''"

        Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
            Dim command As SqlCommand = connection.CreateCommand()
            command.CommandText = queryString
            Try
                connection.Open()
                Dim dataReader As SqlDataReader = _
                 command.ExecuteReader()
                Do While dataReader.Read()


                    Return CInt(dataReader(0))
                Loop

                dataReader.Close()

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            connection.Close()
        End Using




    End Function


    Sub Load_DataSet_With_Top_Customers(ByVal Time_Frame As String)
        InterModule_Transfer = True
        Blank_Customer_Array()

        Intelligent_Marketing_SchemaDataSet.Clear()
        Dim i As Integer
        Dim CURRENTVALUE As Double
        Dim CurrentCode As String
        Dim TOP_SUM As Double

        Dim queryString As String = _
        "SELECT " & BI_Customer_Accounts_Name_Field & " , " & Time_Frame & ", " & BI_Customer_Accounts_Code_Field & " from " & BI_Customer_Accounts_Table & " order by " & Customer_Sales_Archive_Field & " desc"

        If IsSQLSERVER Then

            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        CURRENTVALUE = CInt(dataReader(1))
                        CurrentCode = dataReader(2).ToString

                        Dim newrow As DataRow = Intelligent_Marketing_SchemaDataSet.Basic_Graph.NewRow
                        i += 1
                        If i > TOP Then GoTo skip
                        newrow("Category") = dataReader(0).ToString
                        newrow("Data") = CURRENTVALUE.ToString
                        newrow("List") = i.ToString

                        Customer_Array(i - 1, 0) = CurrentCode
                        Customer_Array(i - 1, 1) = Fetch_Field_Value_From_Specified_Table_Given_ID(CurrentCode, Customer_Field_Name, Customer_Table_Name)
                        Customer_Array(i - 1, 2) = Fetch_Field_Value_From_Specified_Table_Given_ID(CurrentCode, Customer_Contact_Field1, Customer_Table_Name)
                        Customer_Array(i - 1, 3) = Fetch_Field_Value_From_Specified_Table_Given_ID(CurrentCode, Customer_Contact_Field2, Customer_Table_Name)
                        Customer_Array(i - 1, 4) = Fetch_Field_Value_From_Specified_Table_Given_ID(CurrentCode, Customer_Email_Field, Customer_Table_Name)



                        TOP_SUM += CURRENTVALUE

                        Intelligent_Marketing_SchemaDataSet.Basic_Graph.Rows.Add(newrow)



                    Loop
skip:
                    Dim newrow2 As DataRow = Intelligent_Marketing_SchemaDataSet.Basic_Graph.NewRow

                    newrow2("Category") = "Others"
                    newrow2("Data") = (SWL_SCALAR_QUERY("select sum(" & Time_Frame & ") from dealeraccts where " & Time_Frame & " > 0 and accttype = 'SL'")) - TOP_SUM
                    newrow2("List") = i.ToString

                    Intelligent_Marketing_SchemaDataSet.Basic_Graph.Rows.Add(newrow2)



                    dataReader.Close()

                Catch ex As Exception
                    'My.Settings.Error_log.Add(ex.Message)
                End Try

                connection.Close()
            End Using

            ReportViewer1.RefreshReport()

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
                        CURRENTVALUE = CInt(dataReader(1))
                        CurrentCode = dataReader(2).ToString

                        Dim newrow As DataRow = Intelligent_Marketing_SchemaDataSet.Basic_Graph.NewRow
                        i += 1
                        If i > TOP Then GoTo skip2
                        newrow("Category") = dataReader(0).ToString
                        newrow("Data") = CURRENTVALUE.ToString
                        newrow("List") = i.ToString

                        Customer_Array(i - 1, 0) = CurrentCode
                        Customer_Array(i - 1, 1) = Fetch_Field_Value_From_Specified_Table_Given_ID(CurrentCode, Customer_Field_Name, Customer_Table_Name)
                        Customer_Array(i - 1, 2) = Fetch_Field_Value_From_Specified_Table_Given_ID(CurrentCode, Customer_Contact_Field1, Customer_Table_Name)
                        Customer_Array(i - 1, 3) = Fetch_Field_Value_From_Specified_Table_Given_ID(CurrentCode, Customer_Contact_Field2, Customer_Table_Name)
                        Customer_Array(i - 1, 4) = Fetch_Field_Value_From_Specified_Table_Given_ID(CurrentCode, Customer_Email_Field, Customer_Table_Name)

                        TOP_SUM += CURRENTVALUE

                        Intelligent_Marketing_SchemaDataSet.Basic_Graph.Rows.Add(newrow)



                    Loop
skip2:
                    Dim newrow2 As DataRow = Intelligent_Marketing_SchemaDataSet.Basic_Graph.NewRow

                    newrow2("Category") = "Others"
                    newrow2("Data") = (SWL_SCALAR_QUERY("select sum(" & Time_Frame & ") from dealeraccts where " & Time_Frame & " > 0 and accttype = 'SL'")) - TOP_SUM
                    newrow2("List") = i.ToString

                    Intelligent_Marketing_SchemaDataSet.Basic_Graph.Rows.Add(newrow2)



                    dataReader.Close()

                Catch ex As Exception
                    '  MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using

            ReportViewer1.RefreshReport()

        End If

        '        If IsPervasive Then

        '            Using connection As New PsqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
        '                Dim command As PsqlCommand = connection.CreateCommand()
        '                command.CommandText = queryString
        '                Try
        '                    connection.Open()
        '                    Dim dataReader As PsqlDataReader = _
        '                     command.ExecuteReader()
        '                    Do While dataReader.Read()
        '                        CURRENTVALUE = CInt(dataReader(1))
        '                        CurrentCode = dataReader(2).ToString

        '                        Dim newrow As DataRow = Intelligent_Marketing_SchemaDataSet.Basic_Graph.NewRow
        '                        i += 1
        '                        If i > TOP Then GoTo skip3
        '                        newrow("Category") = dataReader(0).ToString
        '                        newrow("Data") = CURRENTVALUE.ToString
        '                        newrow("List") = i.ToString

        '                        Customer_Array(i - 1, 0) = CurrentCode
        '                        Customer_Array(i - 1, 1) = Fetch_Field_Value_From_Specified_Table_Given_ID(CurrentCode, Customer_Field_Name, Customer_Table_Name)
        '                        Customer_Array(i - 1, 2) = Fetch_Field_Value_From_Specified_Table_Given_ID(CurrentCode, Customer_Contact_Field1, Customer_Table_Name)
        '                        Customer_Array(i - 1, 3) = Fetch_Field_Value_From_Specified_Table_Given_ID(CurrentCode, Customer_Contact_Field2, Customer_Table_Name)
        '                        Customer_Array(i - 1, 4) = Fetch_Field_Value_From_Specified_Table_Given_ID(CurrentCode, Customer_Email_Field, Customer_Table_Name)

        '                        TOP_SUM += CURRENTVALUE

        '                        Intelligent_Marketing_SchemaDataSet.Basic_Graph.Rows.Add(newrow)



        '                    Loop
        'skip3:
        '                    Dim newrow2 As DataRow = Intelligent_Marketing_SchemaDataSet.Basic_Graph.NewRow

        '                    newrow2("Category") = "Others"
        '                    newrow2("Data") = (SWL_SCALAR_QUERY("select sum(" & Time_Frame & ") from dealeraccts where " & Time_Frame & " > 0 and accttype = 'SL'")) - TOP_SUM
        '                    newrow2("List") = i.ToString

        '                    Intelligent_Marketing_SchemaDataSet.Basic_Graph.Rows.Add(newrow2)



        '                    dataReader.Close()

        '                Catch ex As Exception
        '                    '  MessageBox.Show(ex.Message)
        '                End Try

        '                connection.Close()
        '            End Using

        '            ReportViewer1.RefreshReport()

        '        End If

    End Sub


    Sub Load_DataSet_With_Custom_Top_Customers(ByVal account_type As String, ByVal PreferredColumn As String)
        Blank_Customer_Array()

        Intelligent_Marketing_SchemaDataSet.Clear()
        Dim i As Integer
        Dim CURRENTVALUE As Double
        Dim CurrentCode As String
        Dim TOP_SUM As Double

        Dim queryString As String = _
              "SELECT name,life_turn,code from dealeraccts where accttype = 'SL' order by life_turn desc"


        If Not IsACCESS Then
            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        CURRENTVALUE = CInt(dataReader(1))
                        CurrentCode = dataReader(2).ToString

                        Dim newrow As DataRow = Intelligent_Marketing_SchemaDataSet.Basic_Graph.NewRow

                        If Fetch_This_Column(PreferredColumn, CurrentCode, "dealers") <> account_type Then GoTo NextLoop
                        i += 1
                        If i > TOP Then GoTo skip
                        newrow("Category") = dataReader(0).ToString
                        newrow("Data") = CURRENTVALUE.ToString
                        newrow("List") = i.ToString

                        TOP_SUM += CURRENTVALUE
                        Customer_Array(i - 1, 0) = CurrentCode

                        Intelligent_Marketing_SchemaDataSet.Basic_Graph.Rows.Add(newrow)


NextLoop:
                    Loop
skip:
                    If i < TOP + 1 Then GoTo skip2
                    Dim newrow2 As DataRow = Intelligent_Marketing_SchemaDataSet.Basic_Graph.NewRow

                    newrow2("Category") = "Other " & account_type
                    newrow2("Data") = (SWL_SCALAR_QUERY("select sum(dealeraccts.life_turn) from dealeraccts, dealers where dealers.code = dealeraccts.code and  dealeraccts.life_turn > 0 and dealeraccts.accttype = 'SL' and dealers.type_code = '" & account_type & "'")) - TOP_SUM
                    newrow2("List") = i.ToString

                    Intelligent_Marketing_SchemaDataSet.Basic_Graph.Rows.Add(newrow2)

skip2:

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using

            ReportViewer1.RefreshReport()
        Else
            Using connection As New OleDbConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As OleDbCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As OleDbDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        CURRENTVALUE = CInt(dataReader(1))
                        CurrentCode = dataReader(2).ToString

                        Dim newrow As DataRow = Intelligent_Marketing_SchemaDataSet.Basic_Graph.NewRow

                        If Fetch_This_Column(PreferredColumn, CurrentCode, "dealers") <> account_type Then GoTo NextLoop2
                        i += 1
                        If i > TOP Then GoTo skip3
                        newrow("Category") = dataReader(0).ToString
                        newrow("Data") = CURRENTVALUE.ToString
                        newrow("List") = i.ToString

                        TOP_SUM += CURRENTVALUE
                        Customer_Array(i - 1, 0) = CurrentCode

                        Intelligent_Marketing_SchemaDataSet.Basic_Graph.Rows.Add(newrow)


NextLoop2:
                    Loop
skip3:
                    If i < TOP + 1 Then GoTo skip4
                    Dim newrow2 As DataRow = Intelligent_Marketing_SchemaDataSet.Basic_Graph.NewRow

                    newrow2("Category") = "Other " & account_type
                    newrow2("Data") = (SWL_SCALAR_QUERY("select sum(dealeraccts.life_turn) from dealeraccts, dealers where dealers.code = dealeraccts.code and  dealeraccts.life_turn > 0 and dealeraccts.accttype = 'SL' ")) - TOP_SUM
                    newrow2("List") = i.ToString

                    Intelligent_Marketing_SchemaDataSet.Basic_Graph.Rows.Add(newrow2)

skip4:

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using

            ReportViewer1.RefreshReport()
        End If




    End Sub





    Private Sub ToolStripLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TOP = CInt(ToolStripComboBox1.Text)
        '  Load_DataSet_With_Top_Customers()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        TOP = CInt(ToolStripComboBox1.Text)
        Load_DataSet_With_Top_Customers(Customer_Sales_Archive_Field)
        Send_List_Title = "All time Best Customers"

        'With ReportViewer1.LocalReport
        '    .ReportEmbeddedResource = "intelligentMarketing.BusinessIntelligence_Report1.rdlc"
        '    .ReportPath = ""
        '    .DataSources.Clear()
        '    .DataSources.Add(New ReportDataSource("DataSet1", Basic_GraphBindingSource))
        '    '  ReportViewer1.RefreshReport()
        'End With


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MessageBox.Show(SWL_Fetch_Number_of_Customers_in_Database.ToString)
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        TOP = CInt(ToolStripComboBox1.Text)
        Load_DataSet_With_Top_Customers("Year_turn")

        Send_List_Title = "Top " & ToolStripComboBox1.Text & " Customers of the Year"
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        TOP = CInt(ToolStripComboBox1.Text)
        Load_DataSet_With_Top_Customers("Period_Turn")

        Send_List_Title = "Top " & ToolStripComboBox1.Text & " Customers of the Period"
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TOP = CInt(ToolStripComboBox1.Text)
        Load_DataSet_With_Custom_Top_Customers("DISTRIBUTORS", "Type_Code")
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        Select_Digital_Marketing()
        ' Intelligence_To_SMS()
    End Sub

    Sub Select_Digital_Marketing()
        SMS_sender.TabControl1.TabPages.Item(0).Select()

    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub ToolStrip2_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStrip2.ItemClicked

    End Sub
End Class