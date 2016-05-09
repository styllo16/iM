Imports System.Data
Imports Pervasive.Data.SqlClient
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Web 'add system.web as reference for the url encoding class
Imports System.Net
Imports System.IO
Imports DotSpatial.Data

Public Class Locations_Data_Validation
    Dim TownsfeatureSet, EditFeatureSet As FeatureSet
    Dim DocumentsDir As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
    Dim Batch_count As Integer
    Dim Current_Batch_Box_RowIndex As Integer


    Private Sub Locations_Data_Validation_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        statusLabel1.Text = ""







    End Sub

    Function Fetch_Customer_Code(ByVal RowIndex As Integer) As String

        Return Batch_Box.Item(0, RowIndex).Value


    End Function

    Sub Clear_BatchBox()
        Batch_Box.Rows.Clear()
        Batch_count = 0
    End Sub
    Function Fetch_Number_of_Customers() As Integer
        Return SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name)

    End Function

    Sub Load_Unvalidated_Locations()
        statusLabel1.Text = "Loading unvalidated data..."
        Load_MySettings_Variables()
        Clear_BatchBox()

        Dim sCustomerCode As String
        Dim sCustomerName As String
        Dim sCustomerTown As String
        Dim sCustomerDistrict As String
        Dim sCustomerRegion As String
        Dim sCustomerContact As String
        Dim Number_of_Customers As Integer




        Number_of_Customers = Fetch_Number_of_Customers()

        Dim queryString As String = _
                "SELECT " & Customer_Code_Field & "," & Customer_Field_Name & "," & Customer_Contact_Field1 & "," & Customer_Location_Field & "," & District_Field & "," & Region_Field & " from " & Customer_Table_Name

        If IsSQLSERVER Then

            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        sCustomerCode = dataReader(0).ToString
                        sCustomerName = dataReader(1).ToString
                        sCustomerContact = dataReader(2).ToString
                        sCustomerTown = dataReader(3).ToString
                        sCustomerDistrict = dataReader(4).ToString
                        sCustomerRegion = dataReader(5).ToString

                        If Is_Location_Valid(sCustomerTown, sCustomerDistrict, sCustomerRegion) Then GoTo skipLoop

                        Batch_Box.Rows.Add(1)


                        Batch_Box.Item(0, Batch_count).Value = sCustomerCode
                        Batch_Box.Item(1, Batch_count).Value = sCustomerName
                        Batch_Box.Item(2, Batch_count).Value = sCustomerTown
                        Batch_Box.Item(3, Batch_count).Value = sCustomerDistrict
                        Batch_Box.Item(4, Batch_count).Value = sCustomerRegion
                        Batch_Box.Item(5, Batch_count).Value = sCustomerContact


                        ProgressBar1.Value = (Batch_count / Number_of_Customers) * 100

                        Batch_count += 1
skipLoop:
                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using
        End If

        '**********************************************************************



        If IsACCESS Then


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

                        sCustomerCode = dataReader(0).ToString
                        sCustomerName = dataReader(1).ToString
                        sCustomerContact = dataReader(2).ToString
                        sCustomerTown = dataReader(3).ToString
                        sCustomerDistrict = dataReader(4).ToString
                        sCustomerRegion = dataReader(5).ToString

                        If Is_Location_Valid(sCustomerTown, sCustomerDistrict, sCustomerRegion) Then GoTo skipLoop2

                        Batch_Box.Rows.Add(1)


                        Batch_Box.Item(0, Batch_count).Value = sCustomerCode
                        Batch_Box.Item(1, Batch_count).Value = sCustomerName
                        Batch_Box.Item(2, Batch_count).Value = sCustomerTown
                        Batch_Box.Item(3, Batch_count).Value = sCustomerDistrict
                        Batch_Box.Item(4, Batch_count).Value = sCustomerRegion
                        Batch_Box.Item(5, Batch_count).Value = sCustomerContact

                        ProgressBar1.Value = (Batch_count / Number_of_Customers) * 100
                        Batch_count += 1
skipLoop2:






                    Loop

                    dataReader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

                connection.Close()
            End Using
        End If


        '        If IsPervasive Then

        '            '
        '            '  "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\Data Directory\ESSAMUAH DATABASE.accdb"

        '            Using connection As New PsqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
        '                Dim command As PsqlCommand = connection.CreateCommand()
        '                command.CommandText = queryString
        '                Try
        '                    connection.Open()
        '                    Dim dataReader As PsqlDataReader = _
        '                     command.ExecuteReader()
        '                    Do While dataReader.Read()
        '                        sCustomerCode = dataReader(0).ToString
        '                        sCustomerName = dataReader(1).ToString
        '                        sCustomerContact = dataReader(2).ToString
        '                        sCustomerTown = dataReader(3).ToString
        '                        sCustomerDistrict = dataReader(4).ToString
        '                        sCustomerRegion = dataReader(5).ToString

        '                        If Is_Location_Valid(sCustomerTown, sCustomerDistrict, sCustomerRegion) Then GoTo skipLoop3

        '                        Batch_Box.Rows.Add(1)


        '                        Batch_Box.Item(0, Batch_count).Value = sCustomerCode
        '                        Batch_Box.Item(1, Batch_count).Value = sCustomerName
        '                        Batch_Box.Item(2, Batch_count).Value = sCustomerTown
        '                        Batch_Box.Item(3, Batch_count).Value = sCustomerDistrict
        '                        Batch_Box.Item(4, Batch_count).Value = sCustomerRegion
        '                        Batch_Box.Item(5, Batch_count).Value = sCustomerContact

        '                        ProgressBar1.Value = (Batch_count / Number_of_Customers) * 100
        '                        Batch_count += 1
        'skipLoop3:
        '                    Loop

        '                    dataReader.Close()

        '                Catch ex As Exception
        '                    MessageBox.Show(ex.Message)
        '                End Try

        '                connection.Close()
        '            End Using
        '        End If


      
        ProgressBar1.Value = 0
        statusLabel1.Text = "Ready"
    End Sub


    Function IM_Dummy_Connection_String() As String
        Return CREATE_CONNECTION_STRING(True, False, "IM Dummy", DocumentsDir, "Null", "Null", "Null", "Null", False)
    End Function

    Function Is_Location_Valid(ByVal sTown As String, ByVal sDistrict As String, ByVal sRegion As String) As Boolean
        Dim Validated_Town As String
        Dim Validated_District As String
        Dim Validated_Region As String

        sRegion = Convert_Region_Names_From_Full_to_Short_Hand(sRegion)


        Dim queryString As String = _
              "SELECT TOWN_NAME, DISTRICT, REGION FROM LOCS WHERE TOWN_NAME = '" & sTown & "' AND DISTRICT = '" & sDistrict & "' AND REGION = '" & sRegion & "'"



        Using connection As New OleDbConnection(IM_Dummy_Connection_String)
            Dim command As OleDbCommand = connection.CreateCommand()
            command.CommandText = queryString
            Try
                connection.Open()
                Dim dataReader As OleDbDataReader = _
                 command.ExecuteReader()
                Do While dataReader.Read()
                    Validated_Town = dataReader(0).ToString
                    Validated_District = dataReader(1).ToString
                    Validated_Region = dataReader(2).ToString

                    If Validated_District <> "" And Validated_Town <> "" And Validated_Region <> "" Then

                        Return True
                    End If

                Loop

                dataReader.Close()

            Catch ex As Exception
                ' MessageBox.Show(ex.Message)
                Return False
            End Try

            connection.Close()

            Return False
        End Using






    End Function

    Public Sub Fill_Combo_With_Data_Locs_Validator(ByRef Combo As ComboBox, ByVal CustomerTableName As String, ByVal TownFieldName As String)
        Dim current_town As String
        Combo.Items.Clear()

      

        Dim I As Integer

        Dim queryString As String = _
                "SELECT distinct " & TownFieldName & " from " & CustomerTableName



        Using connection As New OleDbConnection(IM_Dummy_Connection_String)
            Dim command As OleDbCommand = connection.CreateCommand()
            command.CommandText = queryString
            Try
                connection.Open()
                Dim dataReader As OleDbDataReader = _
                 command.ExecuteReader()
                Do While dataReader.Read()
                    current_town = dataReader(0).ToString

                    If Trim(current_town) <> "" Then
                        Combo.Items.Add(Trim(current_town))
                    End If

                    ProgressBar1.Value = (I / 18300) * 100

                    I += 1
                Loop

                dataReader.Close()

            Catch ex As Exception
                ' MessageBox.Show(ex.Message)
            End Try

            connection.Close()
        End Using






        ProgressBar1.Value = 0


    End Sub

    Public Sub Fill_listbox_With_Data_Locs_Validator_for_District_Suggestion(ByVal Combo As ListBox, ByVal CustomerTableName As String, ByVal DistrictField_to_Display As String, ByVal Town_in_question As String)
        Dim current_town As String
        Combo.Items.Clear()
        Dim I As Integer


        Dim queryString As String = _
                "SELECT distinct " & DistrictField_to_Display & " from " & CustomerTableName & " Where TOWN_NAME = '" & Town_in_question & "'"



        Using connection As New OleDbConnection(IM_Dummy_Connection_String)
            Dim command As OleDbCommand = connection.CreateCommand()
            command.CommandText = queryString
            Try
                connection.Open()
                Dim dataReader As OleDbDataReader = _
                 command.ExecuteReader()
                Do While dataReader.Read()
                    current_town = dataReader(0).ToString

                    If Trim(current_town) <> "" Then
                        Combo.Items.Add(Trim(current_town))
                    End If

                    ProgressBar1.Value = (I / 18300) * 100

                    I += 1
                Loop

                dataReader.Close()

            Catch ex As Exception
                ' MessageBox.Show(ex.Message)
            End Try

            connection.Close()
        End Using



        ProgressBar1.Value = 0


    End Sub

    Public Sub Fill_listbox_With_Data_Locs_Validator_for_Region_Suggestion(ByVal Combo As ListBox, ByVal CustomerTableName As String, ByVal DistrictField_to_Display As String, ByVal Town_in_question As String)
        Dim current_town As String
        Combo.Items.Clear()
        Dim I As Integer

        Dim queryString As String = _
                "SELECT distinct " & DistrictField_to_Display & " from " & CustomerTableName & " Where TOWN_NAME = '" & Town_in_question & "'"

       


        Using connection As New OleDbConnection(IM_Dummy_Connection_String)
            Dim command As OleDbCommand = connection.CreateCommand()
            command.CommandText = queryString
            Try
                connection.Open()
                Dim dataReader As OleDbDataReader = _
                 command.ExecuteReader()
                Do While dataReader.Read()
                    current_town = dataReader(0).ToString

                    If Trim(current_town) <> "" Then
                        Combo.Items.Add(Convert_Region_Names_From_Short_to_Long_Hand(current_town))
                    End If

                    ProgressBar1.Value = (I / 18300) * 100

                    I += 1
                Loop

                dataReader.Close()

            Catch ex As Exception
                'MessageBox.Show(ex.Message)
            End Try

            connection.Close()
        End Using








        ProgressBar1.Value = 0


    End Sub

    Function Convert_Region_Names_From_Full_to_Short_Hand(ByVal Longhand As String)
        Select Case Longhand
            Case "Greater Accra"
                Return "GAR"

            Case "Central"
                Return "CR"

            Case "Western"
                Return "WR"

            Case "Eastern"
                Return "ER"

            Case "Volta"
                Return "VR"

            Case "Northern"
                Return "NR"

            Case "Brong Ahafo"
                Return "BAR"

            Case "Upper East"
                Return "UER"

            Case "Upper West"
                Return "UWR"

            Case "Ashanti"
                Return "AR"

        End Select


    End Function

    Function Convert_Region_Names_From_Short_to_Long_Hand(ByVal Short1 As String)
        Select Case Short1
            Case "GAR"
                Return "Greater Accra"

            Case "CR"
                Return "Central"

            Case "WR"
                Return "Western"

            Case "ER"
                Return "Eastern"

            Case "VR"
                Return "Volta"

            Case "NR"
                Return "Northern"

            Case "BAR"
                Return "Brong Ahafo"

            Case "UER"
                Return "Upper East"

            Case "UWR"
                Return "Upper West"

            Case "AR"
                Return "Ashanti"

        End Select


    End Function

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub Batch_Box_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Batch_Box.CellContentClick

    End Sub

    Private Sub Batch_Box_Click(sender As Object, e As System.EventArgs) Handles Batch_Box.Click

        Try

            Current_Batch_Box_RowIndex = Batch_Box.SelectedCells(0).RowIndex

            txtTown.Text = Batch_Box.Item(2, Current_Batch_Box_RowIndex).Value
            txtDistrict.Text = Batch_Box.Item(3, Current_Batch_Box_RowIndex).Value
            txtRegion.Text = Batch_Box.Item(4, Current_Batch_Box_RowIndex).Value

        Catch ex As Exception
            Exit Sub
        End Try


    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        Try

      
        Me.Cursor = Cursors.WaitCursor
        statusLabel1.Text = "Loading combo box..."

        Fill_Combo_With_Data_Locs_Validator(cbo_Town, "locs", "TOWN_NAME")
        Fill_Combo_With_Data_Locs_Validator(cbo_District, "locs", "DISTRICT")
        statusLabel1.Text = "Loading unvalidated location..."
        Load_Unvalidated_Locations()

        Me.Cursor = Cursors.Default
            statusLabel1.Text = "Ready"

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click

        Try

        
        Current_Batch_Box_RowIndex = Batch_Box.SelectedCells(0).RowIndex

        If Is_Location_Valid(cbo_Town.Text, cbo_District.Text, cbo_Region.Text) Then

            Update_This_Row_In_The_Database(Fetch_Customer_Code(Current_Batch_Box_RowIndex), cbo_Town.Text, cbo_District.Text, cbo_Region.Text)
            Delete_This_Row(Current_Batch_Box_RowIndex)
        Else

            MessageBox.Show("Invalid Location")
        End If


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs)
        Batch_Box.Rows.RemoveAt(0)

    End Sub

    Sub Delete_This_Row(ByVal i As Integer)
        Batch_Box.Rows.RemoveAt(i)
    End Sub

    Sub Update_This_Row_In_The_Database(ByVal ID As String, ByVal sTown As String, ByVal sDistrict As String, ByVal sRegion As String)

        Execute_Non_Query("Update " & Customer_Table_Name & " set " & Customer_Location_Field & " = '" & sTown & "' ," & District_Field & " = '" & sDistrict & "' ," & Region_Field & " = '" & sRegion & "' Where " & Customer_Code_Field & " = '" & ID & "'", CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))

    End Sub

    Private Sub Button7_Click(sender As System.Object, e As System.EventArgs)
        Current_Batch_Box_RowIndex = Batch_Box.SelectedCells(0).RowIndex







    End Sub

    Private Sub cbo_Town_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbo_Town.SelectedIndexChanged


        Fill_listbox_With_Data_Locs_Validator_for_District_Suggestion(ListBox1, "locs", "District", cbo_Town.Text)
    End Sub

    Private Sub cbo_Town_TextChanged(sender As Object, e As System.EventArgs) Handles cbo_Town.TextChanged
        Fill_listbox_With_Data_Locs_Validator_for_District_Suggestion(ListBox1, "locs", "DISTRICT", cbo_Town.Text)

    End Sub

    Private Sub cbo_District_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbo_District.SelectedIndexChanged

    End Sub

    Private Sub cbo_District_TextChanged(sender As Object, e As System.EventArgs) Handles cbo_District.TextChanged
        If cbo_District.Text = "" Then
            ListBox1.Items.Clear()
            Exit Sub
        End If


        Fill_listbox_With_Data_Locs_Validator_for_Region_Suggestion(ListBox1, "locs", "Region", cbo_District.Text)
    End Sub

    Private Sub Button3_Click_1(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Current_Batch_Box_RowIndex = Batch_Box.SelectedCells(0).RowIndex

        If Is_Location_Valid(cbo_Town.Text, cbo_District.Text, cbo_Region.Text) Then

          
        Else

            MessageBox.Show("Invalid Location")
        End If

    End Sub

    Private Sub mnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnEdit.Click
        Dim f As New validation
        f.ShowDialog()

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click

        Try
            m_LI.Town = cbo_Town.SelectedItem
            m_LI.Region = cbo_Region.SelectedItem
            m_LI.District = cbo_District.SelectedItem

            Dim f As New frmEditTownName
            f.ShowDialog()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cbo_Region_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbo_Region.SelectedIndexChanged

    End Sub

    Private Sub btnSync_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSync.Click

        Me.Cursor = Cursors.WaitCursor
        Delete_All_Records_in_Locs_Table()

        Dim cb_townName As String = "\map\Towns\"
        Dim arrayRegion(10) As String
        Dim i As Integer
        arrayRegion(0) = "GAR"
        arrayRegion(1) = "ER"
        arrayRegion(2) = "AR"
        arrayRegion(3) = "WR"
        arrayRegion(4) = "VR"
        arrayRegion(5) = "CR"
        arrayRegion(6) = "NR"
        arrayRegion(7) = "UER"
        arrayRegion(8) = "UWR"
        arrayRegion(9) = "BAR"

        For x = 0 To 9
            'Accra
            Dim cb_TownshapeFilename As String = arrayRegion(x)

            '*********
            'status text
            ProgressBar1.Value = (x + 1 / 10) * 100
            statusLabel1.Text = "Synchronizing " & cb_TownshapeFilename & "..."

            '**********************
            TownsfeatureSet = FeatureSet.Open(DocumentsDir & cb_townName & cb_TownshapeFilename & ".shp")
            ' this code opens the shapefile and stores the information
            For Each row As DataRow In TownsfeatureSet.DataTable.Rows
                'i += 1
                m_LI.Town = row("TOWN").ToString
                m_LI.Town = Replace(m_LI.Town, "'", "")


                m_LI.Region = cb_TownshapeFilename.ToString
                m_LI.Region = Replace(m_LI.Region, "'", "")

                m_LI.District = row("DISTRICT").ToString
                m_LI.District = Replace(m_LI.District, "'", "")

                'saving distric code and reading it once
                'If district_code = 1 Then
                m_LI.DistrictCode = row("DIST_CODE")
                '    district_code = 0
                'End If


                Insert_Row_Into_Locs_Table(m_LI.Town, m_LI.District, m_LI.Region)


            Next



        Next

        ProgressBar1.Value = 0

        Me.Cursor = Cursors.Default
        statusLabel1.Text = "Ready"
    End Sub

    Sub Insert_Row_Into_Locs_Table(ByVal sTown As String, ByVal sDistrict As String, ByVal sRegion As String)
        Execute_Non_Query("Insert into locs (TOWN_NAME, Region, District) values ('" & sTown & "','" & sRegion & "','" & sDistrict & "')", IM_Dummy_Connection_String)

    End Sub

    Sub Delete_All_Records_in_Locs_Table()
        Execute_Non_Query("Delete from locs", IM_Dummy_Connection_String)
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub
End Class