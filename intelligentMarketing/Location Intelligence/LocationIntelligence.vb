'Required namespaces
Imports DotSpatial.Controls
Imports DotSpatial.Symbology
Imports Pervasive.Data.SqlClient
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports DotSpatial.Data
Imports DotSpatial.Serialization
Imports DotSpatial.Controls.BruTileLayer
Imports BruTile
Imports System.Drawing
Imports DotSpatial.Projections
Imports DotSpatial.Topology
Imports System.ComponentModel.Composition

'Imports Microsoft.Reporting.WinForms



Public Class LI

    

    Dim MyLayer, JHS, Primary, hsSales, District,
        hsDebtors, hsCustomers, pSales, pDebtors, pCustomers As IMapFeatureLayer
    Dim mapFeature_occupation, mapFeature_employment As IMapFeatureLayer
    Dim mapSatellite As BruTileLayer
    Dim maps As Integer
    Dim idx As Integer
    Dim layerIndex, spotSynchIndex, districtSynchIndex, regionSynchIndex As Integer
    Dim selectedIdx As Integer
    Dim baseMap, info, firstclick, sa_sales, sa_customers, sa_debtors As Boolean
    Dim baseMapType As String
    Dim DocumentsDir As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
    Dim color1 As System.Drawing.Color = Color.LightYellow
    Dim color2 As System.Drawing.Color = Color.DarkRed
    Dim By_Debtors As Boolean = False


    Public Batch_Count As Integer
   

    Sub DISPLAY_REPORT_TITLE()
        Dim Datasource As String

        If IsSQLSERVER Then
            Datasource = SQL_Initial_Catalog
        End If

        If IsACCESS Then
            Datasource = MSACESSFILE
        End If

        If IsPervasive Then
            Datasource = PSQLDatabase_Name
        End If

        Datasource = UCase(Datasource)

        ToolStripLabel1.Text = "GEOSPATIAL DISTRIBUTION OF " & UCase(cbAnalysis.Text) & " IN GHANA.  (DATASOURCE: " & Datasource & ")"
    End Sub



    Private Sub LI_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.GotFocus
        '************************************************************
        ' Authentication Code
        If My.Settings.iM_Authentication = False Then
            Map1.Enabled = False
            ToolStrip2.Enabled = False
            dgvAttributeTable.Enabled = False
            Legend1.Enabled = False

            'If Authentication_Required_Message_Shown = False Then
            '    MessageBox.Show("Authentication Required.")
            '    Authentication_Required_Message_Shown = True
            'End If

        Else
            Map1.Enabled = True
            ToolStrip2.Enabled = True
            dgvAttributeTable.Enabled = True
            Legend1.Enabled = True
        End If

        '************************************************************
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load



        Load_MySettings_Variables()
        cbAnalysis.Visible = False
        mainContainer.Panel2Collapsed = True
        SplitContainer2.Panel2Collapsed = True
        Legend1.ClearSelection()

        ' activate menu items
        aSales.Enabled = True
        aDebtors.Enabled = True

        If Is_Demo_Restriction_Exceeded() = True Then
            MessageBox.Show("Location Intelligence is Disabled. Contact sales@adroitbureau.com", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Enabled = False
        End If

    End Sub

    Private Sub btAttribute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btAttribute.Click


    End Sub

    Private Sub btnZoomIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZoomIn.Click

        Map1.ZoomIn()

    End Sub

    Private Sub btnZoomOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZoomOut.Click

        Map1.ZoomToMaxExtent()

    End Sub



#Region "Functions"

    ''''''''''''''''''''''''''''''FUNCTIONS'''''''''''''''''''''''''''''''''''''''''

    Private Sub LabelsDisplay()

        'Check the number of layers from MapControl
        If Map1.Layers.Count > 0 Then
            'Declare a MapPolygonLayer
            Dim stateLayer As MapPolygonLayer
            'TypeCast the first layer from MapControl to MapPolygonLayer.
            'Layers are 0 based, therefore 0 is going to grab the first layer from the MapControl
            stateLayer = CType(Map1.Layers(0), MapPolygonLayer)

            'Check whether stateLayer is polygon layer or not
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'add StateName as labels on the stateLayer
                '[STATE_NAME] is an attribute from the given example US States shape file.
                Map1.AddLabels(stateLayer, "[Region]", New Font("Tahoma", 8.0), Color.Black)

            End If
        Else
            MessageBox.Show("Please add a layer to the map.")
        End If


    End Sub


    Private Sub symbolization(ByVal fieldName As String, ByVal idxLayer As Integer, ByVal analysis As String)

        ' Try
        'check the number of layers from map control
        If Map1.Layers.Count > 0 Then
            'Delacre a MapPolygonLayer
            Dim stateLayer As MapPolygonLayer
            'Type cast the FirstLayer of MapControl to MapPolygonLayer
            stateLayer = CType(Map1.Layers(idxLayer), MapPolygonLayer)
            'Check the MapPolygonLayer ( Make sure that it has a polygon layer)
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else


                'Create a new PolygonScheme
                Dim scheme As New PolygonScheme

                'set colour

                Select Case analysis

                    Case "Penrollment"
                        scheme.EditorSettings.StartColor = Color.LightYellow
                        scheme.EditorSettings.EndColor = Color.Red
                        scheme.EditorSettings.UseGradient = True

                    Case "Jenrollment"
                        scheme.EditorSettings.StartColor = Color.LightYellow
                        scheme.EditorSettings.EndColor = Color.Red
                        scheme.EditorSettings.UseGradient = True

                    Case Else
                        scheme.EditorSettings.StartColor = Color.LightYellow
                        scheme.EditorSettings.EndColor = Color.DarkRed
                        scheme.EditorSettings.UseGradient = True
                End Select

                'Set the ClassificationType for the PolygonScheme via EditotSettings
                scheme.EditorSettings.ClassificationType = ClassificationType.Quantities
                scheme.EditorSettings.NumBreaks = 8
                scheme.EditorSettings.IntervalMethod = IntervalMethod.Quantile
                scheme.EditorSettings.IntervalSnapMethod = IntervalSnapMethod.DataValue
                'Set the UniqueValue field name
                'Here STATE_NAME would be the Unique value field
                scheme.EditorSettings.FieldName = fieldName
                'create categories on the scheme based on the attributes table and field name
                'In this case field name is STATE_NAME
                scheme.CreateCategories(stateLayer.DataSet.DataTable)
                'Set the scheme to stateLayer's symbology


                stateLayer.Symbology = scheme
            End If
        Else
            MessageBox.Show("Please add a layer to the map.")
        End If

        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
    End Sub

    ''' <summary>
    ''' This sub uses to fill the shapefile attribute in the datagrid
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ShapefileAttribute()



        Dim selectedlayerType As String

        'find index and type of selected layer
        Try

            Dim dt As DataTable

            selectedlayerType = Map1.Layers.SelectedLayer.GetType.ToString

            For x = 0 To Map1.Layers.Count - 1


                If Map1.Layers(x).IsSelected = True Then
                    selectedIdx = x

                End If

            Next

            Select Case selectedlayerType

                Case "DotSpatial.Controls.MapPolygonLayer"

                    Dim stateLayer As MapPolygonLayer

                    stateLayer = CType(Map1.Layers(selectedIdx), MapPolygonLayer)

                    'Get the shapefile's attribute table to our datatable dt
                    dt = stateLayer.DataSet.DataTable

                    'Set the datagridview datasource from datatable dt
                    dgvAttributeTable.DataSource = dt

                    DisplayAttributeLabel(stateLayer.LegendText)


                Case "DotSpatial.Controls.MapPointLayer"

                    Dim stateLayer As MapPointLayer

                    stateLayer = CType(Map1.Layers(selectedIdx), MapPointLayer)

                    'Get the shapefile's attribute table to our datatable dt
                    dt = stateLayer.DataSet.DataTable

                    'Set the datagridview datasource from datatable dt
                    dgvAttributeTable.DataSource = dt

                    DisplayAttributeLabel(stateLayer.LegendText)

            End Select
            '*****************************************************************************


        Catch ex As Exception

        End Try


    End Sub

    Private Sub ShapefileAttribute(ByVal layerIndex As Integer)

        Dim selectedlayerType As String

        'find index and type of selected layer
        Try



            Dim dt As DataTable
            'find type of layer selected, if polygon or point
            selectedlayerType = Map1.Layers.SelectedLayer.GetType.ToString
            'loop trough the legend

            selectedIdx = Map1.Layers.IndexOf(Map1.Layers.SelectedLayer)



            Select Case selectedlayerType

                Case "DotSpatial.Controls.MapPolygonLayer"

                    Dim stateLayer As MapPolygonLayer

                    stateLayer = CType(Map1.Layers(selectedIdx), MapPolygonLayer)

                    'Get the shapefile's attribute table to our datatable dt
                    dt = stateLayer.DataSet.DataTable

                    'Set the datagridview datasource from datatable dt
                    dgvAttributeTable.DataSource = dt

                    DisplayAttributeLabel(stateLayer.LegendText)


                Case "DotSpatial.Controls.MapPointLayer"

                    Dim stateLayer As MapPointLayer

                    stateLayer = CType(Map1.Layers(selectedIdx), MapPointLayer)

                    'Get the shapefile's attribute table to our datatable dt
                    dt = stateLayer.DataSet.DataTable

                    'Set the datagridview datasource from datatable dt
                    dgvAttributeTable.DataSource = dt

                    DisplayAttributeLabel(stateLayer.LegendText)

            End Select

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    ' ''' <summary>
    ' ''' Function to increase the index number of the layers added
    ' ''' anytime the primary or jhs enrollment is selected, then the index is counted
    ' ''' </summary>
    ' ''' <remarks></remarks>
    'Private Sub index()

    '    layerIndex = Map1.Layers.Count - 1

    'End Sub

    ' ''' <summary>
    ' ''' Function to decrease the index number of the layers added
    ' ''' </summary>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Private Function removeIndex()
    '    If idx <> 0 Then
    '        idx = idx - 1

    '    End If
    '    Return idx
    'End Function


    ''' <summary>
    ''' sub that displays name of shape file attribute table selected
    ''' in the label lblAttributeName
    ''' </summary>
    ''' <param name="attributeName"></param>
    ''' <remarks></remarks>
    Private Sub DisplayAttributeLabel(ByVal attributeName As String)

        lblAttributeName.Text = attributeName

    End Sub

#End Region

    Private Sub dgvAttributeTable_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAttributeTable.CellClick

    End Sub


    'Private Sub cbAnalysis_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAnalysis.SelectedIndexChanged
    '    DISPLAY_REPORT_TITLE()

    '    Dim analysis As String
    '    analysis = cbAnalysis.SelectedItem
    '    Select Case analysis

    '        Case "Sales"

    '            symbolization("Sales", 0, "Reional")

    '        Case "Customers"

    '            symbolization("Customers", 0, "Reional")

    '        Case "Debtors"

    '            symbolization("Debtors", 0, "Reional")

    '    End Select

    'End Sub

    Private Sub dgvAttributeTable_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAttributeTable.Click
        InterModule_Transfer = True


        For Each row As DataGridViewRow In dgvAttributeTable.SelectedRows

            If Map1.Layers.SelectedLayer.GetType.ToString = "DotSpatial.Controls.MapPolygonLayer" Then
                Dim stateLayer As MapPolygonLayer
                stateLayer = CType(Map1.Layers(selectedIdx), MapPolygonLayer)

                If stateLayer Is Nothing Then
                    MessageBox.Show("The layer is not a polygon layer.")
                Else
                    stateLayer.SelectByAttribute("[Region] =" + "'" + row.Cells("Region").Value + "'")

                End If

            ElseIf Map1.Layers.SelectedLayer.GetType.ToString = "DotSpatial.Controls.MapPointLayer" Then
                Dim stateLayer As MapPointLayer
                stateLayer = CType(Map1.Layers(selectedIdx), MapPointLayer)

                If stateLayer Is Nothing Then
                    MessageBox.Show("The layer is not a polygon layer.")
                Else
                    stateLayer.SelectByAttribute("[TOWN] =" + "'" + row.Cells("TOWN").Value + "'")

                End If

            End If

        Next


        'Dim s As Integer
        's.ToString("###,###")



        Dim region As String
        Selected_Region_From_Attribute_Table = dgvAttributeTable.Item(0, dgvAttributeTable.SelectedCells(0).RowIndex).Value.ToString

        Selected_Region_From_Attribute_Table_focus_specific = Convert_Region_Format_From_Attribute_Table_To_Focus_Format(Selected_Region_From_Attribute_Table)

        Dim i As Integer
        Blank_Customer_Array()
        Dim current_code As String
        Dim CustomerName As String
        Dim CustomerContact1 As String
        Dim CustomerContact2 As String
        Dim CustomerEmail As String


        Dim queryString As String = _
           "SELECT " & Customer_Code_Field & ", " & Customer_Field_Name & ", " & Customer_Contact_Field1 & ", " & Customer_Contact_Field2 & ", " & Customer_Email_Field & " from " & Customer_Table_Name & " where " & Region_Field & "= '" & Selected_Region_From_Attribute_Table & "'"


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
                        CustomerName = dataReader(1).ToString
                        CustomerContact1 = dataReader(2).ToString
                        CustomerContact2 = dataReader(3).ToString
                        CustomerEmail = dataReader(4).ToString


                        Customer_Array(i, 0) = current_code
                        Customer_Array(i, 1) = CustomerName
                        Customer_Array(i, 2) = CustomerContact1
                        Customer_Array(i, 3) = CustomerContact2
                        Customer_Array(i, 4) = CustomerEmail

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
                        CustomerName = dataReader(1).ToString
                        CustomerContact1 = dataReader(2).ToString
                        CustomerContact2 = dataReader(3).ToString
                        CustomerEmail = dataReader(4).ToString


                        Customer_Array(i, 0) = current_code
                        Customer_Array(i, 1) = CustomerName
                        Customer_Array(i, 2) = CustomerContact1
                        Customer_Array(i, 3) = CustomerContact2
                        Customer_Array(i, 4) = CustomerEmail
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
        '                CustomerName = dataReader(1).ToString
        '                CustomerContact1 = dataReader(2).ToString
        '                CustomerContact2 = dataReader(3).ToString
        '                CustomerEmail = dataReader(4).ToString


        '                Customer_Array(i, 0) = current_code
        '                Customer_Array(i, 1) = CustomerName
        '                Customer_Array(i, 2) = CustomerContact1
        '                Customer_Array(i, 3) = CustomerContact2
        '                Customer_Array(i, 4) = CustomerEmail
        '                i += 1
        '            Loop

        '            dataReader.Close()

        '        Catch ex As Exception
        '            MessageBox.Show(ex.Message)
        '        End Try

        '        connection.Close()
        '    End Using



        'End If







        Send_List_Title = Selected_Region_From_Attribute_Table & " Regional Contacts"

    End Sub

    Private Sub dgvAttributeTable_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAttributeTable.DoubleClick
        Drill_Down_to_Region()
    End Sub

    Private Sub dgvAttributeTable_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgvAttributeTable.SelectionChanged

        'For Each row As DataGridViewRow In dgvAttributeTable.SelectedRows
        '    Dim stateLayer As MapPolygonLayer
        '    stateLayer = CType(Map1.Layers(selectedIdx), MapPolygonLayer)
        '    If stateLayer Is Nothing Then
        '        MessageBox.Show("The layer is not a polygon layer.")
        '    Else
        '        stateLayer.SelectByAttribute("[Region] =" + "'" + row.Cells("Region").Value + "'")

        '    End If
        'Next


    End Sub
#Region "District capital code (Point)"

    ''' <summary>
    ''' Add attribute of point file
    ''' </summary>
    ''' <param name="layerIndex"></param>
    ''' <remarks></remarks>

    'Private Sub ShapefileAttributePoint(ByVal layerIndex As Integer)

    '    selectedIdx = layerIndex
    '    Dim selectedLayer As String

    '    'Declare a datatable
    '    Dim dt As DataTable
    '    If Map1.Layers.Count > 0 Then
    '        Dim stateLayer As MapPointLayer
    '        'this loop iterates all the items and compares to the desired layer
    '        '  For x = 0 To Map1.Layers.Count - 1
    '        stateLayer = CType(Map1.Layers(layerIndex), MapPointLayer)
    '        stateLayer.LegendText = "District capital"
    '        selectedLayer = stateLayer.LegendText
    '        DisplayAttributeLabel(selectedLayer)



    '        If stateLayer Is Nothing Then
    '            MessageBox.Show("The layer is not a point layer.")
    '        Else
    '            'Get the shapefile's attribute table to our datatable dt
    '            dt = stateLayer.DataSet.DataTable

    '            'Set the datagridview datasource from datatable dt
    '            dgvAttributeTable.DataSource = dt




    '        End If
    '    Else
    '        MessageBox.Show("Please add a layer to the map.")
    '    End If
    'End Sub

    ''' <summary>
    ''' Display labels of point features (name of districts capital)
    ''' </summary>
    ''' <param name="indx"></param>
    ''' <remarks></remarks>
    'Private Sub LabelsDisplayPoint(ByVal indx As Integer)

    '    Try
    '        'Check the number of layers from MapControl
    '        If Map1.Layers.Count > 0 Then

    '            'Declare a MapPolygonLayer
    '            Dim stateLayer As MapPointLayer
    '            'TypeCast the first layer from MapControl to MapPolygonLayer.
    '            'Layers are 0 based, therefore 0 is going to grab the first layer from the MapControl
    '            stateLayer = CType(Map1.Layers(indx), MapPointLayer)

    '            'Check whether stateLayer is polygon layer or not
    '            If stateLayer Is Nothing Then
    '                MessageBox.Show("The layer is not a point layer.")
    '            Else
    '                'add StateName as labels on the stateLayer
    '                '[STATE_NAME] is an attribute from the given example US States shape file.
    '                Map1.AddLabels(stateLayer, "[Name]", New Font("Tahoma", 6.0), Color.Black)

    '            End If
    '        Else
    '            MessageBox.Show("Please add a layer to the map.")
    '        End If

    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try

    'End Sub


#End Region



#Region " Loading shape file with focus sub procedure"

    Sub Load_Shape_File_With_Focus_Data()
        Dim i As Integer
        Dim Sales_data(10) As Double
        Dim Customer_Count(10) As Integer
        Dim Debtors_Data(10) As Integer

        ' Load Sales Data into Shape File
        'Try
        '    Sales_data(0) = SWL_SCALAR_QUERY("Select sum(dealeraccts.life_turn) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'A/R' and dealeraccts.life_turn > 0")
        '    Sales_data(1) = SWL_SCALAR_QUERY("Select sum(dealeraccts.life_turn) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'C/R' and dealeraccts.life_turn > 0")
        '    Sales_data(2) = SWL_SCALAR_QUERY("Select sum(dealeraccts.life_turn) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'E/R'  and dealeraccts.life_turn > 0")
        '    Sales_data(3) = SWL_SCALAR_QUERY("Select sum(dealeraccts.life_turn) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'GAR'  and dealeraccts.life_turn > 0")
        '    Sales_data(4) = SWL_SCALAR_QUERY("Select sum(dealeraccts.life_turn) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'N/R' and dealeraccts.life_turn > 0")
        '    Sales_data(5) = SWL_SCALAR_QUERY("Select sum(dealeraccts.life_turn) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'UE/R' and dealeraccts.life_turn > 0")
        '    Sales_data(6) = SWL_SCALAR_QUERY("Select sum(dealeraccts.life_turn) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'UW/R' and dealeraccts.life_turn > 0")
        '    Sales_data(7) = SWL_SCALAR_QUERY("Select sum(dealeraccts.life_turn) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'V/R' and dealeraccts.life_turn > 0")
        '    Sales_data(8) = SWL_SCALAR_QUERY("Select sum(dealeraccts.life_turn) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'W/R' and dealeraccts.life_turn > 0")
        '    Sales_data(9) = SWL_SCALAR_QUERY("Select sum(dealeraccts.life_turn) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'B/R' and dealeraccts.life_turn > 0")

        'Catch ex As Exception

        'End Try


        'Load Customer Distribution into Shape File

        Try
            Customer_Count(0) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%'")
            Customer_Count(1) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%'")
            Customer_Count(2) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%'")
            Customer_Count(3) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%'")
            Customer_Count(4) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%'")
            Customer_Count(5) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%'")
            Customer_Count(6) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper West%'")
            Customer_Count(7) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%'")
            Customer_Count(8) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%'")
            Customer_Count(9) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%'")

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


        'Try
        '    Debtors_Data(0) = SWL_SCALAR_QUERY("Select sum(dealeraccts.sl_balance) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'A/R' and dealeraccts.sl_balance > 0")
        '    Debtors_Data(1) = SWL_SCALAR_QUERY("Select sum(dealeraccts.sl_balance) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'C/R' and dealeraccts.sl_balance > 0")
        '    Debtors_Data(2) = SWL_SCALAR_QUERY("Select sum(dealeraccts.sl_balance) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'E/R'  and dealeraccts.sl_balance > 0")
        '    Debtors_Data(3) = SWL_SCALAR_QUERY("Select sum(dealeraccts.sl_balance) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'GAR'  and dealeraccts.sl_balance > 0")
        '    Debtors_Data(4) = SWL_SCALAR_QUERY("Select sum(dealeraccts.sl_balance) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'N/R' and dealeraccts.sl_balance > 0")
        '    Debtors_Data(5) = SWL_SCALAR_QUERY("Select sum(dealeraccts.sl_balance) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'UE/R' and dealeraccts.sl_balance > 0")
        '    Debtors_Data(6) = SWL_SCALAR_QUERY("Select sum(dealeraccts.sl_balance) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'UW/R' and dealeraccts.sl_balance > 0")
        '    Debtors_Data(7) = SWL_SCALAR_QUERY("Select sum(dealeraccts.sl_balance) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'V/R' and dealeraccts.sl_balance > 0")
        '    Debtors_Data(8) = SWL_SCALAR_QUERY("Select sum(dealeraccts.sl_balance) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'W/R' and dealeraccts.sl_balance > 0")
        '    Debtors_Data(9) = SWL_SCALAR_QUERY("Select sum(dealeraccts.sl_balance) from dealeraccts, dealers where dealers.code = dealeraccts.code and dealers.area = 'B/R' and dealeraccts.sl_balance > 0")

        'Catch ex As Exception

        'End Try

        'Load Debt Distribution in Graph




        'Declare a datatable
        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPolygonLayer
            stateLayer = TryCast(Map1.Layers(regionSynchIndex), MapPolygonLayer)
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                For Each row As DataRow In dt.Rows
                    'Dim males As Double = row("Sales")
                    'males = males * 100
                    'row("Debtors") = males

                    row("Sales") = Sales_data(i)
                    row("Customers") = Customer_Count(i)
                    row("Debtors") = Debtors_Data(i)

                    i += 1
                Next
                'Set the datagridview datasource from datatable dt
                dgvAttributeTable.DataSource = dt

                'saving to the shapefile
                stateLayer.DataSet.Save()
                stateLayer.UnSelectAll()

            End If
        Else
            MessageBox.Show("Please add a layer to the map.")
        End If




    End Sub

#End Region
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click

        Dim frm As New DotSpatial.Controls.LayoutForm
        frm.MapControl = Map1
        frm.Show()


    End Sub

    Private Sub Form1_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        If baseMap = False Then
            Map1.ZoomToMaxExtent()
        Else
            If Map1.Layers.Count > 0 Then
                ZoomToLayer(Map1, Map1.Layers.Count - 1)
            End If
        End If
    End Sub

#Region "Drill Down sub procedure"

    ''' <summary>
    ''' This is the code that calls the drill down for each region selected
    ''' </summary>
    ''' <remarks></remarks>

    Sub Drill_Down_to_Region()

        Selected_Region_From_Attribute_Table = dgvAttributeTable.Item(0, dgvAttributeTable.SelectedCells(0).RowIndex).Value.ToString

        Select Case Selected_Region_From_Attribute_Table

            Case "Ashanti"
                Dim f As New Regional_Drill_Down
                f.Text = Selected_Region_From_Attribute_Table & " Regional Drill-Down"
                f.ShowDialog()

                'This will Fetch the drill down information from the database.

            Case "Central"
                Dim f As New Regional_Drill_Down
                f.Text = Selected_Region_From_Attribute_Table & " Regional Drill-Down"
                f.ShowDialog()


            Case "Eastern"
                Dim f As New Regional_Drill_Down
                f.Text = Selected_Region_From_Attribute_Table & " Regional Drill-Down"
                f.ShowDialog()


            Case "Greater Accra"
                Dim f As New Regional_Drill_Down
                f.Text = Selected_Region_From_Attribute_Table & " Regional Drill-Down"
                f.ShowDialog()


            Case "Northern"
                Dim f As New Regional_Drill_Down
                f.Text = Selected_Region_From_Attribute_Table & " Regional Drill-Down"
                f.ShowDialog()


            Case "Upper East"
                Dim f As New Regional_Drill_Down
                f.Text = Selected_Region_From_Attribute_Table & " Regional Drill-Down"
                f.ShowDialog()


            Case "Upper West"
                Dim f As New Regional_Drill_Down
                f.Text = Selected_Region_From_Attribute_Table & " Regional Drill-Down"
                f.ShowDialog()


            Case "Volta"
                Dim f As New Regional_Drill_Down
                f.Text = Selected_Region_From_Attribute_Table & " Regional Drill-Down"
                f.ShowDialog()


            Case "Western"
                Dim f As New Regional_Drill_Down
                f.Text = Selected_Region_From_Attribute_Table & " Regional Drill-Down"
                f.ShowDialog()


            Case "Brong Ahafo"
                Dim f As New Regional_Drill_Down
                f.Text = Selected_Region_From_Attribute_Table & " Regional Drill-Down"
                f.ShowDialog()


        End Select


    End Sub

#End Region

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Select_Digital_Marketing()
    End Sub
    Sub Select_Digital_Marketing()
        SMS_sender.TabControl1.TabPages.Item(0).Select()
        SMS_sender.TabControl1.TabPages.Item(0).Focus()
    End Sub

    Private Sub Legend1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Legend1.Click

        ShapefileAttribute()

    End Sub

#Region "JHS and Primary School Enrollment codes"
    ''' <summary>
    ''' Codes to add the JHS and Primary school enrollment data
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub mPrimary_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mPrimary.Click


        If mPrimary.CheckState = CheckState.Checked Then
            Dim i As Integer
            'AddLayer() method is used to add a shape file in the MapControl
            Primary = Map1.AddLayer(DocumentsDir & "\map\PrimaryEnrollment.shp")
            Primary.Projection = KnownCoordinateSystems.Projected.World.WebMercator
            Primary.LegendText = "Primary pupils enrollment"
            ' index
            i = Map1.Layers.IndexOf(Primary)
            Map1.Layers(i).IsSelected = True
            If i <> 0 Then
                Map1.Layers(i - 1).IsSelected = False
            End If
            symbolization("Total_from", i, "Penrollment")
            ShapefileAttribute(i)

        Else
            Map1.Layers.Remove(Primary)

        End If

    End Sub

    Private Sub mJHS_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mJHS.Click

        If mJHS.CheckState = CheckState.Checked Then
            Dim j As Integer
            'AddLayer() method is used to add a shape file in the MapControl
            JHS = Map1.AddLayer(DocumentsDir & "\map\JHS.shp")
            JHS.Projection = KnownCoordinateSystems.Projected.World.WebMercator
            JHS.LegendText = "JHS enrollment"

            ' index
            j = Map1.Layers.IndexOf(JHS)

            Map1.Layers(j).IsSelected = True
            If j <> 0 Then
                Map1.Layers(j - 1).IsSelected = False
            End If
            symbolization("Private_fr", j, "Jenrollment")
            ShapefileAttribute(j)


        Else


            Map1.Layers.Remove(JHS)



        End If

    End Sub
#End Region
#Region "Distric Analysis drill down code"

    Private Sub Map1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Map1.Click


        'If info = False Then Exit Sub
        'If firstclick = False Then
        '    firstclick = True
        '    Exit Sub
        'End If

        'Try

        '    Dim region As String

        '    Dim fLayer As IFeatureLayer
        '    '
        '    'Here code for load fLayer
        '    If Map1.Layers.Count = 0 Then
        '        Exit Sub
        '    End If
        '    If baseMap = True And selectedIdx = 0 Then
        '        MsgBox("Base map layer is selected" & vbCrLf & "Ensure that the right layer is selected")

        '    Else

        '        'If Map1.Layers.SelectedLayer.GetType.ToString = "DotSpatial.Controls.MapPolygonLayer" Then
        '        fLayer = Map1.Layers(selectedIdx)



        '        '*****************************************REGION SELECTION
        '        '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        '        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        '        If Not IsNothing(fLayer) Then
        '            Dim result As List(Of IFeature) = New List(Of IFeature)()
        '            fLayer.Selection.SelectionState = True
        '            result = fLayer.Selection.ToFeatureList
        '            If result.Count = 0 Then
        '                firstclick = True
        '                Exit Sub
        '            End If

        '            For Each feature As IFeature In result
        '                If Not IsDBNull(feature.ShapeIndex) Then
        '                    'MsgBox(fLayer.DataSet.ShapeIndices.IndexOf(feature.DataRow.Item(2))) '

        '                    region = feature.DataRow.Item(0).ToString
        '                    '  MsgBox(region.ToString)

        '                    ' Set up the delays for the ToolTip.
        '                    ToolTip1.AutoPopDelay = 5000
        '                    ToolTip1.InitialDelay = 1000
        '                    ToolTip1.ReshowDelay = 500
        '                    ToolTip1.BackColor = Color.LightYellow
        '                    ' Force the ToolTip text to be displayed whether or not the form is active.
        '                    ToolTip1.ShowAlways = True

        '                    ' Set up the ToolTip text for the Button and Checkbox.
        '                    ToolTip1.SetToolTip(Me.Map1, region)
        '                Else
        '                    MsgBox("No Region has been selected")
        '                End If
        '            Next
        '        End If

        '        'Map1.Cursor = Cursors.AppStarting
        '        'Dim f As New DistrictAnalysis
        '        'f.lblTitle.Text = region & " District level analysis"
        '        'f.p_RegionSelected = region


        '        'f.Show()
        '        'f.WindowState = FormWindowState.Normal

        '        Map1.Cursor = Cursors.Default
        '        'Else
        '        MsgBox("Layer selected in the ledend is not a polygon")
        '    End If


        '    'End If


        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
    End Sub
    ''' <summary>
    ''' This calls the district analysis form when the map is doubleclicked
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>

    Private Sub Map1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Map1.DoubleClick
        'Try

        '    Dim region As String

        '    Dim fLayer As IFeatureLayer
        '    '
        '    'Here code for load fLayer

        '    If baseMap = True And selectedIdx = 0 Then
        '        MsgBox("Base map layer is selected" & vbCrLf & "Ensure that the right layer is selected")

        '    Else

        '        'If Map1.Layers.SelectedLayer.GetType.ToString = "DotSpatial.Controls.MapPolygonLayer" Then
        '        fLayer = Map1.Layers(selectedIdx)


        '        '*****************************************REGION SELECTION
        '        '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        '        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        '        If Not IsNothing(fLayer) Then
        '            Dim result As List(Of IFeature) = New List(Of IFeature)()
        '            result = fLayer.Selection.ToFeatureList()
        '            For Each feature As IFeature In result
        '                If Not IsDBNull(feature.ShapeIndex) Then
        '                    'MsgBox(fLayer.DataSet.ShapeIndices.IndexOf(feature.DataRow.Item(2))) '

        '                    region = feature.DataRow.Item(0).ToString
        '                    '  MsgBox(region.ToString)
        '                Else
        '                    MsgBox("No Region has been selected")
        '                End If
        '            Next
        '        End If

        '        Map1.Cursor = Cursors.AppStarting
        '        Dim f As New DistrictAnalysis
        '        f.lblTitle.Text = region & " District level analysis"
        '        f.p_RegionSelected = region


        '        f.Show()
        '        f.WindowState = FormWindowState.Normal

        '        Map1.Cursor = Cursors.Default
        '        'Else
        '        MsgBox("Layer selected in the ledend is not a polygon")
        '    End If


        '    'End If


        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try


        '************************************************************************

    End Sub

#End Region

    Private Sub DistrictCapitalsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dCapital.Click
        Try


            Dim j As Integer
            If dCapital.CheckState = CheckState.Checked Then

                'AddLayer() method is used to add a shape file in the MapControl
                District = Map1.AddLayer("\map\Districts\Towns.shp")
                District.LegendText = "District capital"
                'run function to set index
                j = Map1.Layers.IndexOf(District)


                Map1.Layers(j).IsSelected = True
                If j <> 0 Then
                    Map1.Layers(j - 1).IsSelected = False
                End If
                ShapefileAttribute(j)

                'now we'll get a reference to the PolygonLayer's Symbolizer and apply styling changes
                District.Symbolizer = New PointSymbolizer(Color.Black, DotSpatial.Symbology.PointShape.Star, 8)
            Else


                Map1.Layers.Remove(District)



            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub



    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Select_Digital_Marketing()
    End Sub





    'Private Sub btnRoads_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRoads.Click

    '    ' Dim path As String
    '    ' path = System.Environment.GetFolderPath(Environment.SpecialFolder.MyComputer)

    '    If btnRoads.Checked = True Then
    '        Try
    '            Map1.AddLayer("C:\map\Road\Road.shp").LegendText = "Roads"

    '            Dim j As Integer
    '            j = index()
    '            TownIdx = j


    '            ' ShapefileAttributePoint(TownIdx)
    '            ' LabelsDisplayPoint(TownIdx)

    '            'Check the number of layers from MapControl
    '            If Map1.Layers.Count > 0 Then
    '                'Declare a MapPolygonLayer
    '                Dim stateLayer As MapLineLayer
    '                'TypeCast the first layer from MapControl to MapPolygonLayer.
    '                'Layers are 0 based, therefore 0 is going to grab the first layer from the MapControl
    '                stateLayer = CType(Map1.Layers(j), MapLineLayer)

    '                'Check whether stateLayer is polygon layer or not
    '                If stateLayer Is Nothing Then
    '                    MessageBox.Show("The layer is not a LineLayer layer.")
    '                Else
    '                    'add StateName as labels on the stateLayer
    '                    '[STATE_NAME] is an attribute from the given example US States shape file.
    '                    'Map1.AddLabels(stateLayer, "[SEGMENT]", New Font("Tahoma", 8.0), Color.Black)

    '                    Dim road As LineSymbolizer
    '                    road = New LineSymbolizer(Color.Black, 0.2)
    '                    road.ScaleMode = ScaleMode.Geographic
    '                    'road.SetFillColor(Color.Green )
    '                    stateLayer.Symbolizer = road
    '                    Map1.FunctionMode = FunctionMode.ZoomIn

    '                End If
    '            End If



    '        Catch Ex As Exception
    '            MsgBox(Ex.Message)
    '        End Try

    '    ElseIf btnRoads.Checked = False Then


    '        Map1.Layers.Remove(Map1.Layers(TownIdx))
    '        primaryIdx = removeIndex()

    '    End If
    'End Sub

    Private Sub btnPan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPan.Click

        Map1.FunctionMode = FunctionMode.Pan

    End Sub

    Private Sub btnPointer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPointer.Click

        Map1.FunctionMode = FunctionMode.None

    End Sub

    Private Sub showLabel(ByVal labelIndex As Integer)

        Dim selectedlayerType As String
        Try
            'find index and type of selected layer
            'find type of layer selected, if polygon or point
            selectedlayerType = Map1.Layers(labelIndex).GetType.ToString





            Dim field As String
            With Map1.Layers(labelIndex)
                If .LegendText = "Sales analysis" Then

                    field = "Sales"
                ElseIf .LegendText = "Customers analysis" Then
                    field = "Customers"
                ElseIf .LegendText = "Debtors analysis" Then
                    field = "Debtors"

                End If

            End With


            'this loop iterates all the items and compares to the desired layer
            ' For x = 0 To Map1.Layers.Count - 1

            ' selIndx = Map1.Layers(x).LegendText

            'If Map1.Layers(x).IsSelected = True Then
            ' selectedIdx = x



            Dim labelLayer As IMapLabelLayer
            labelLayer = New MapLabelLayer

            Dim cat As ILabelCategory
            cat = labelLayer.Symbology.Categories(0)
            '    If ShowLabelToolStripMenuItem.Checked = True Then

            If field = "Sales" Or field = "Debtors" Then
                cat.Expression = field & vbNewLine & "Ghc. [" &
                                    field & "]"
            ElseIf field = "Customers" Then
                cat.Expression = field & vbNewLine & "[" &
                    field & "]"
            End If

            With cat.Symbolizer
                .BackColorEnabled = True
                .BackColor = Color.FromArgb(128, Color.LightBlue)
                If baseMap = True Then
                    .BorderVisible = True
                    .BorderColor = Color.Black
                Else
                    .BorderVisible = False
                End If

                .FontStyle = FontStyle.Bold
                .FontColor = Color.Black
                .Orientation = ContentAlignment.MiddleCenter
                .Alignment = StringAlignment.Center
            End With



            Select Case selectedlayerType

                Case "DotSpatial.Controls.MapPolygonLayer"
                    If Map1.Layers.Count > 0 Then
                        Dim stateLayer As MapPolygonLayer
                        stateLayer = CType(Map1.Layers(labelIndex), MapPolygonLayer)
                        stateLayer.ShowLabels = True
                        stateLayer.LabelLayer = labelLayer
                    Else
                        MsgBox("Add a layer to the map")
                    End If

                Case "DotSpatial.Controls.MapPointLayer"
                    If Map1.Layers.Count > 0 Then
                        Dim stateLayer As MapPointLayer
                        stateLayer = CType(Map1.Layers(labelIndex), MapPointLayer)
                        stateLayer.ShowLabels = True
                        stateLayer.LabelLayer = labelLayer
                    Else
                        MsgBox("Add a layer to the map")
                    End If
            End Select


            '    Else
            'cat.Expression = "[Region]"
            'cat.Symbolizer.Orientation = ContentAlignment.MiddleCenter
            'cat.Symbolizer.Alignment = StringAlignment.Center
            'cat.Symbolizer.FontSize = 8



            '   End If


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub RemoveLabels(ByVal labelIndex As Integer)
        Dim selectedlayerType As String
        Try


            'find index and type of selected layer
            'find type of layer selected, if polygon or point
            selectedlayerType = Map1.Layers(labelIndex).GetType.ToString

            Select Case selectedlayerType

                Case "DotSpatial.Controls.MapPolygonLayer"
                    If Map1.Layers.Count > 0 Then
                        Dim stateLayer As MapPolygonLayer
                        stateLayer = CType(Map1.Layers(labelIndex), MapPolygonLayer)
                        stateLayer.ShowLabels = False
                    Else
                        MsgBox("There are no layers on the map")
                    End If

                Case "DotSpatial.Controls.MapPointLayer"
                    If Map1.Layers.Count > 0 Then
                        Dim stateLayer As MapPointLayer
                        stateLayer = CType(Map1.Layers(labelIndex), MapPointLayer)
                        stateLayer.ShowLabels = False
                    Else
                        MsgBox("There are no layers on the map")
                    End If
            End Select
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub




#Region "Codes for online map"

#Region "Adding base map code"

    ''' <summary>
    ''' Adds base map on click
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub mBaseMap_ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mBaseMap.ButtonClick


        If baseMap = False Then
            SMS_sender.Cursor = Cursors.WaitCursor
            AddBaseMap("Google map", True)
            baseMap = True
            SMS_sender.Cursor = Cursors.Default

            mGmap.Checked = True
            mBRmap.Checked = False
            mBSmap.Checked = False
            mOsm.Checked = False

            baseMap = True

        Else

            SMS_sender.Cursor = Cursors.WaitCursor
            AddBaseMap("Google map", False)
            SMS_sender.Cursor = Cursors.Default
            mGmap.Checked = False
            mBRmap.Checked = False
            mBSmap.Checked = False
            mOsm.Checked = False
            baseMap = False

        End If

    End Sub
    Private Sub GoogleMapToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mGmap.Click
        If mGmap.Checked = True Then

            'If baseMap = True Then
            '    Map1.Layers.Remove(mapSatellite)
            'End If
            SMS_sender.Cursor = Cursors.WaitCursor
            baseMapType = "Google map"
            AddBaseMap(baseMapType, True)
            baseMap = True
            SMS_sender.Cursor = Cursors.Default

            mGmap.Checked = True
            mBRmap.Checked = False
            mBSmap.Checked = False
            mOsm.Checked = False

            baseMap = True



        Else
            SMS_sender.Cursor = Cursors.WaitCursor
            AddBaseMap(baseMapType, False)
            SMS_sender.Cursor = Cursors.Default
            mGmap.Checked = False
            baseMap = False

        End If



    End Sub
    Private Sub BingRoadMapToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mBRmap.Click

        If mBRmap.Checked = True Then


            'If baseMap = True Then
            '    Map1.Layers.Remove(mapSatellite)
            'End If
            SMS_sender.Cursor = Cursors.WaitCursor
            baseMapType = "Bing road map"
            AddBaseMap(baseMapType, True)

            SMS_sender.Cursor = Cursors.Default

            mBRmap.Checked = True
            mGmap.Checked = False
            mBSmap.Checked = False
            mOsm.Checked = False

            baseMap = True

        Else
            SMS_sender.Cursor = Cursors.WaitCursor
            AddBaseMap(baseMapType, False)
            SMS_sender.Cursor = Cursors.Default

            mBRmap.Checked = False
            baseMap = False


        End If

    End Sub

    Private Sub BingSatelliteMapToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mBSmap.Click

        If mBSmap.Checked = True Then


            'If baseMap = True Then
            '    Map1.Layers.Remove(mapSatellite)
            'End If
            SMS_sender.Cursor = Cursors.WaitCursor
            baseMapType = "Bing satellite"
            AddBaseMap(baseMapType, True)

            SMS_sender.Cursor = Cursors.Default

            mBSmap.Checked = True
            mBRmap.Checked = False
            mGmap.Checked = False
            mOsm.Checked = False

            baseMap = True
        Else
            SMS_sender.Cursor = Cursors.WaitCursor
            AddBaseMap(baseMapType, False)
            SMS_sender.Cursor = Cursors.Default
            mBSmap.Checked = False
            baseMap = False

        End If





    End Sub

    Private Sub OpenStreetMapToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mOsm.Click

        If mOsm.Checked = True Then


            'If baseMap = True Then
            '    Map1.Layers.Remove(mapSatellite)
            'End If
            SMS_sender.Cursor = Cursors.WaitCursor
            baseMapType = "Open street"
            AddBaseMap(baseMapType, True)

            SMS_sender.Cursor = Cursors.Default

            mBRmap.Checked = False
            mBSmap.Checked = False
            mGmap.Checked = False
            mOsm.Checked = True

            baseMap = True

        Else
            SMS_sender.Cursor = Cursors.WaitCursor
            AddBaseMap(baseMapType, False)
            SMS_sender.Cursor = Cursors.Default
            mOsm.Checked = False

            baseMap = False

        End If

    End Sub

#End Region

    ''' <summary>
    ''' This  a sub that adds  on-line base map
    ''' </summary>
    ''' <param name="maptype">the type of map to display</param>
    ''' <param name="Add">whether to add a map or remove</param>
    ''' <remarks></remarks>
    Private Sub AddBaseMap(ByVal maptype As String, ByVal Add As Boolean)


        Select Case maptype

            Case "Bing road map"

                'End If
                'mapSatellite.Reproject(KnownCoordinateSystems.Projected.UtmWgs1984.WGS1984UTMZone30N)

                If Add = True Then
                    If baseMap = True Then
                        Map1.Layers.Remove(mapSatellite)
                    End If
                    'this code reproject the map


                    'this create the map
                    'change projection of map to webmercator
                    'changes the legend text
                    'insert the map at index layer 0

                    mapSatellite = BruTileLayer.CreateBingRoadsLayer
                    Map1.ProjectionModeReproject = ActionMode.Always
                    mapSatellite.Projection = KnownCoordinateSystems.Projected.World.WebMercator
                    mapSatellite.LegendText = "Base map"
                    Map1.Layers.Insert(0, mapSatellite) ' Map1.Layers.Add(mapSatellite)

                Else
                    Map1.Layers.Remove(mapSatellite)
                End If

            Case "Bing satellite"


                If Add = True Then
                    If baseMap = True Then
                        Map1.Layers.Remove(mapSatellite)
                    End If
                    'this code reproject the map
                    Map1.ProjectionModeReproject = ActionMode.Always

                    'this create the map
                    'change projection of map to webmercator
                    'changes the legend text
                    'insert the map at index layer 0
                    mapSatellite = BruTileLayer.CreateBingAerialLayer
                    Map1.Projection = KnownCoordinateSystems.Projected.World.WebMercator
                    mapSatellite.LegendText = "Base map"
                    Map1.Layers.Insert(0, mapSatellite)
                Else
                    Map1.Layers.Remove(mapSatellite)
                End If


            Case "Google map"


                If Add = True Then
                    If baseMap = True Then
                        Map1.Layers.Remove(mapSatellite)
                    End If
                    'this code reproject the map
                    Map1.ProjectionModeReproject = ActionMode.Always

                    'this create the map
                    'change projection of map to webmercator
                    'changes the legend text
                    'insert the map at index layer 0
                    mapSatellite = BruTileLayer.CreateGoogleMapLayer
                    Map1.Projection = KnownCoordinateSystems.Projected.World.WebMercator
                    mapSatellite.LegendText = "Base map"
                    Map1.Layers.Insert(0, mapSatellite)
                Else
                    Map1.Layers.Remove(mapSatellite)
                End If

            Case "Open street"


                If Add = True Then
                    If baseMap = True Then
                        Map1.Layers.Remove(mapSatellite)
                    End If
                    'this code reproject the map
                    Map1.ProjectionModeReproject = ActionMode.Always

                    'this create the map
                    'change projection of map to webmercator
                    'changes the legend text
                    'insert the map at index layer 0
                    mapSatellite = BruTileLayer.CreateOsmLayer
                    Map1.Projection = KnownCoordinateSystems.Projected.World.WebMercator
                    Map1.Layers.Insert(0, mapSatellite)
                    mapSatellite.LegendText = "Base map"

                Else
                    Map1.Layers.Remove(mapSatellite)
                End If

        End Select
    End Sub

    Public Sub Load_map()



        Load_MySettings_Variables()

        mainContainer.Panel2Collapsed = True

        '                                 ADDING MAP
        '****************************************************************************
        'AddLayer() method is used to add a shape file in the MapControl
        MyLayer = Map1.AddLayer("C:\Users\Proto\Desktop\iM Standard\map\Regional.shp") '.LegendText = "Regional analysis"

        'test
        ' Map1.AddLayer().LegendText = "Regional analysis"
        ' change projection
        MyLayer.Reproject(KnownCoordinateSystems.Projected.World.WebMercator)
        Map1.ProjectionModeReproject = ActionMode.Always


        'zooms map to extent
        '  Map1.ZoomToMaxExtent()

        ' cbAnalysis.SelectedIndex = 1

        'add analysis into the combo box
        cbAttribute.Items.Insert(0, "Regional analysis table")
        '********************************************************************************************

        ' LabelsDisplay()

        Load_Shape_File_With_Focus_Data()

        '  symbolization("Customers", 0, "Regional")

        ' This is the code that connects to the focus database and loads the information



        '*******************************************************

        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPolygonLayer
            stateLayer = TryCast(Map1.Layers(1), MapPolygonLayer)
            'If stateLayer Is Nothing Then
            '    MessageBox.Show("The layer is not a polygon layer.")

            'End If

            stateLayer.UnSelectAll()

        End If

        '**********************************************************



        lblAnalysis2.Visible = False
        cbAttribute.Visible = False


        ' Setting the map in selection mode
        Map1.FunctionMode = FunctionMode.Select

        DISPLAY_REPORT_TITLE()

    End Sub
    Private Sub SalesMapToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalesMapToolStripMenuItem.Click

        ' Load_map()


    End Sub

    ''' <summary>
    ''' Applies styling and labeling to a MapPolygonLayer.
    ''' </summary>
    ''' <param name="Map">Dotspatial Map</param>
    ''' <param name="PolygonLayer">MapPolygonLayer. IMapPolygonLayer</param>
    ''' <param name="LabelColumnName">Column name to use for labeling. String.</param>
    ''' <param name="LabelColor">Label color. System.Drawing.Color</param>
    ''' <param name="LabelFont"></param>
    ''' <param name="LayerFillColor">Fill color. System.Drawing.Color</param>
    ''' <param name="LayerOutlineColor">Outline color. System.Drawing.Color</param>
    ''' <param name="Opacity">Transparency of the fill color. 0=transparent, 1=opaque. Single</param>
    ''' <remarks>This is a utility function for quickly applying styles and labels to polygon layers.</remarks>
    Private Sub ApplyPolygonLayerStyle(ByVal Map As Map, ByVal PolygonLayer As IMapPolygonLayer, ByVal LabelColumnName As String, ByVal startColor As System.Drawing.Color, ByVal endColor As System.Drawing.Color, ByVal Opacity As Single)




        '**************************************************************************
        ''                      Labels
        'Dim labelLayer As IMapLabelLayer = New MapLabelLayer()
        'Dim category As ILabelCategory = labelLayer.Symbology.Categories(0)


        'category.Expression = "[Region] GHS: [" & LabelColumnName & "]"

        'category.Symbolizer.BackColorEnabled = True
        'category.Symbolizer.BackColor = Color.FromArgb(128, Color.LightBlue)
        'category.Symbolizer.BorderVisible = False
        'category.Symbolizer.FontStyle = FontStyle.Bold
        'category.Symbolizer.FontColor = Color.Black
        'category.Symbolizer.Orientation = ContentAlignment.MiddleCenter
        'category.Symbolizer.Alignment = StringAlignment.Center
        'PolygonLayer.ShowLabels = True
        'PolygonLayer.LabelLayer = labelLayer


        '***********************************************************************************
        'now we'll get a reference to the PolygonLayer's Symbolizer and apply styling changes
        'Dim MySymbolizer As PolygonSymbolizer = PolygonLayer.Symbolizer
        'With MySymbolizer
        '    '.SetFillColor(LayerFillColor.ToTransparent(Opacity))
        '    .SetOutline(Color.Gray, 1)
        'End With

        '*******************************************
        'Create a new PolygonScheme
        Dim scheme As New PolygonScheme

        'set colour


        'set transparency if there is base map
        If baseMap = True Then
            scheme.EditorSettings.StartColor = (startColor.ToTransparent(0.3))
            scheme.EditorSettings.EndColor = (endColor.ToTransparent(0.8))
            scheme.EditorSettings.UseGradient = False

        Else

            scheme.EditorSettings.StartColor = (startColor)
            scheme.EditorSettings.EndColor = (endColor)
            scheme.EditorSettings.UseGradient = True

        End If
        'Set the ClassificationType for the PolygonScheme via EditotSettings
        scheme.EditorSettings.ClassificationType = ClassificationType.Quantities

        scheme.EditorSettings.IntervalMethod = IntervalMethod.Quantile
        'scheme.EditorSettings.IntervalSnapMethod = IntervalSnapMethod.DataValue
        scheme.EditorSettings.IntervalRoundingDigits = 0
        scheme.EditorSettings.NumBreaks = 6

        'Set the UniqueValue field name
        'Here STATE_NAME would be the Unique value field
        scheme.EditorSettings.FieldName = LabelColumnName
        'create categories on the scheme based on the attributes table and field name
        'In this case field name is STATE_NAME
        scheme.CreateCategories(PolygonLayer.DataSet.DataTable)
        'Set the scheme to stateLayer's symbology
        PolygonLayer.Symbolizer.SetOutline(Color.LightGray, 1)


        PolygonLayer.Symbology = scheme
        ' PolygonLayer.Symbolizer.SetFillColor(Color.Transparent)

        '**********************
        'show the labels
        '    PolygonLayer.ShowLabels = True


        'we want all the labels shown, even if they collide
        '   PolygonLayer.LabelLayer.Symbolizer.PreventCollisions = False


        'refresh the map
        Map.Refresh()
    End Sub
    'Private Sub Map_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Map1.MouseUp
    '    If e.Button <> Windows.Forms.MouseButtons.Right Then
    '        Dim ClickCoordinate As Coordinate = Map1.PixelToProj(e.Location)
    '        MsgBox(ClickCoordinate.Y & " " & ClickCoordinate.X)
    '    End If
    'End Sub


#Region "codes for point symbology"





    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="layerindex"></param>
    ''' <remarks></remarks>
    Private Sub ApplyPointLayerStyle(ByVal layerindex As Integer, ByVal fieldName As String)

        Try

            If Map1.Layers.Count > 0 Then
                'Delacre a MapPolygonLayer
                Dim stateLayer As MapPointLayer
                'Type cast the FirstLayer of MapControl to MapPolygonLayer

                stateLayer = CType(Map1.Layers(layerindex), MapPointLayer)
                'Check the MapPolygonLayer ( Make sure that it has a polygon layer)
                If stateLayer Is Nothing Then
                    MessageBox.Show("The layer is not a point layer.")
                Else
                    'Create a new PolygonScheme
                    Dim scheme As New PointScheme


                    scheme.EditorSettings.UseGradient = False

                    scheme.EditorSettings.HueSatLight = False
                    scheme.EditorSettings.RampColors = True


                    Select Case fieldName
                        Case "Customers"
                            scheme.EditorSettings.UseColorRange = True
                            scheme.EditorSettings.StartColor = (Color.Blue.ToTransparent(0.5))
                            scheme.EditorSettings.EndColor = (Color.Blue.ToTransparent(0.8))
                        Case "Sales"
                            scheme.EditorSettings.UseColorRange = True
                            scheme.EditorSettings.StartColor = (Color.GreenYellow.ToTransparent(0.5))
                            scheme.EditorSettings.EndColor = (Color.GreenYellow.ToTransparent(0.8))
                        Case "Debtors"
                            scheme.EditorSettings.UseColorRange = True
                            scheme.EditorSettings.StartColor = (Color.Red.ToTransparent(0.5))
                            scheme.EditorSettings.EndColor = (Color.Red.ToTransparent(0.8))
                    End Select


                    'Set the ClassificationType for the PolygonScheme via EditotSettings
                    scheme.EditorSettings.ClassificationType = ClassificationType.Quantities
                    scheme.EditorSettings.IntervalMethod = IntervalMethod.Quantile
                    scheme.EditorSettings.NumBreaks = 5



                    scheme.EditorSettings.UseSizeRange = True

                    scheme.EditorSettings.StartSize = 1
                    scheme.EditorSettings.EndSize = 30




                    scheme.EditorSettings.FieldName = fieldName
                    'create categories on the scheme based on the attributes table and field name
                    'In this case field name is STATE_NAME
                    scheme.CreateCategories(stateLayer.DataSet.DataTable)
                    'Set the scheme to stateLayer's symbology
                    stateLayer.Symbology = scheme
                    MyLayer = CType(stateLayer, IMapFeatureLayer)
                End If
            Else
                MessageBox.Show("Please add a layer to the map.")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        ' Map1.Refresh()
    End Sub


#End Region
#End Region

#Region "Codes for hotspot analysis"

    Sub Load_Shape_File_Point_For_District_Capitals(ByVal index As Integer)


        'Reading from shapefile.
        Dim i As Integer
        Dim towns As String
        Dim district As String
        Dim region As String
        Dim Rowcount As Integer

        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPointLayer
            stateLayer = TryCast(Map1.Layers(index), MapPointLayer) 'change index
            If stateLayer Is Nothing Then
                MessageBox.Show("Select spot anaylisis layer in the legend.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                Rowcount = dt.Rows.Count


                For Each row As DataRow In dt.Rows
                    'Dim males As Double = row("Sales")
                    'males = males * 100
                    'row("Debtors") = males

                    district = Trim(row("DISTRICT"))

                    Dim intCount As String
                    intCount = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & District_Field & " = '" & district & "'")
                    ' row("Sales") =
                    row("Customers") = intCount
                    'row("Debtors") = Debtors_Data(i)

                    ToolStripProgressBar1.Value = (i / Rowcount) * 100

                    i += 1

                Next


                '   "Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & Region_Field & "= 'A/R'"
                'Set the datagridview datasource from datatable dt
                dgvAttributeTable.DataSource = dt

                'saving to the shapefile
                stateLayer.DataSet.Save()
                stateLayer.UnSelectAll()

                Map1.Refresh()
            End If
        Else
            MessageBox.Show("Please add a layer to the map.")
        End If


        ToolStripProgressBar1.Value = 0
    End Sub

    Sub Load_Shape_File_Point_With_Focus_Data(ByVal index As Integer)
        Dim i As Integer
        Dim Sales_data(10) As Double
        Dim Customer_Count(10) As Integer
        Dim Debtors_Data(10) As Integer

        '******************************************************************************************
        'Reading from shapefile.
        Dim towns As String
        Dim district As String
        Dim region As String


        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPointLayer
            stateLayer = TryCast(Map1.Layers(index), MapPointLayer) 'change index
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                For Each row As DataRow In dt.Rows
                    'Dim males As Double = row("Sales")
                    'males = males * 100
                    'row("Debtors") = males

                    towns = row("Name")
                    district = row("District")
                    region = row("Region")

                    Dim intCount As String
                    intCount = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & District_Field & " = '" & district & "' and " & Customer_Location_Field & " = '" & towns & "'")
                    ' row("Sales") =
                    row("Customers") = intCount
                    'row("Debtors") = Debtors_Data(i)
                    i += 1

                Next


                '   "Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & Region_Field & "= 'A/R'"
                'Set the datagridview datasource from datatable dt
                dgvAttributeTable.DataSource = dt

                'saving to the shapefile
                stateLayer.DataSet.Save()
                stateLayer.UnSelectAll()
            End If
        Else
            MessageBox.Show("Please add a layer to the map.")
        End If

        '***********************************************************************************


    End Sub

#End Region









    Private Sub aSales_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles aSales.Click

        Legend1.ClearSelection()
        Try
            If aSales.Checked = True Then

                ' Map1.Projection = KnownCoordinateSystems.Projected.World.WebMercator
                'set reprojection mode always true so that the reprojection message will not appear
                Map1.ProjectionModeReproject = ActionMode.Always
                pSales = Map1.AddLayer(DocumentsDir & "\map\Ghana.shp")
                pSales.LegendText = "Sales analysis"


                'reproject the map


                'apply symbology
                ApplyPolygonLayerStyle(Me.Map1, pSales, "Sales", color1, color2, 0.3)
                'determine index
                layerIndex = Map1.Layers.IndexOf(pSales)
                'passing index to the synchronization index
                regionSynchIndex = layerIndex
                selectedIdx = layerIndex
                'zoom to layer
                ZoomToLayer(Map1, layerIndex)

                Map1.Layers(layerIndex).IsSelected = True
                'If layerIndex <> 0 Then
                '    Map1.Layers(layerIndex - 1).IsSelected = False
                'End If

                'activate synchronization menu
                If mRegionalSynch.Enabled = False Then
                    mRegionalSynch.Enabled = True
                End If
            Else
                Map1.Layers.Remove(pSales)
                Region_Synch_Deactivation()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub aCustomers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles aCustomers.Click
        Legend1.ClearSelection()


        Try
            If aCustomers.Checked = True Then
                'set reprojection mode always true so that the reprojection message will not appear
                Map1.ProjectionModeReproject = ActionMode.Always
                pCustomers = Map1.AddLayer(DocumentsDir & "\map\Ghana.shp")
                pCustomers.LegendText = "Customers analysis"



                Dim MyFont As Font
                MyFont = New Font("Tahoma", 8.0)

                'reproject the map
                'pCustomers.Projection = KnownCoordinateSystems.Projected.UtmWgs1984.WGS1984ComplexUTMZone30N

                '8888888888888888888888888888888888888888888888
                ' Load_Shape_File_With_Focus_Data()
                '*************************************************


                'apply symbology
                ApplyPolygonLayerStyle(Me.Map1, pCustomers, "Customers", color1, color2, 0.3)
                'determine index
                layerIndex = Map1.Layers.IndexOf(pCustomers)
                regionSynchIndex = layerIndex
                selectedIdx = layerIndex
                'zoom to layer
                ZoomToLayer(Map1, layerIndex)
                Map1.Layers(layerIndex).IsSelected = True




                'If layerIndex <> 0 Then
                '    Map1.Layers(layerIndex - 1).IsSelected = False
                'End If
                'activate synchronization menu
                If mRegionalSynch.Enabled = False Then
                    mRegionalSynch.Enabled = True
                End If


            Else
                Map1.Layers.Remove(pCustomers)
                Region_Synch_Deactivation()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try



    End Sub

    Private Sub aDebtors_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles aDebtors.Click
        Legend1.ClearSelection()
        Try
            If aDebtors.Checked = True Then
                'set reprojection mode always true so that the reprojection message will not appear
                Map1.ProjectionModeReproject = ActionMode.Always
                pDebtors = Map1.AddLayer(DocumentsDir & "\map\Ghana.shp")
                pDebtors.LegendText = "Debtors analysis"

                Dim MyFont As Font
                MyFont = New Font("Tahoma", 8.0)

                'reproject the map
                '   pDebtors.Projection = KnownCoordinateSystems.Projected.zo




                'apply symbology
                ApplyPolygonLayerStyle(Me.Map1, pDebtors, "Debtors", color1, color2, 0.3)
                'determine index
                layerIndex = Map1.Layers.IndexOf(pDebtors)
                regionSynchIndex = layerIndex
                selectedIdx = layerIndex
                'zoom to layer
                ZoomToLayer(Map1, layerIndex)

                Map1.Layers(layerIndex).IsSelected = True
                'If layerIndex <> 0 Then
                '    Map1.Layers(layerIndex - 1).IsSelected = False
                'End If
                'activate synchronization menu
                If mRegionalSynch.Enabled = False Then
                    mRegionalSynch.Enabled = True
                End If
            Else
                Map1.Layers.Remove(pDebtors)
                Region_Synch_Deactivation()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub sCustomers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sCustomers.Click
        Legend1.ClearSelection()
        Try
            If sCustomers.Checked = True Then
                SMS_sender.Cursor = Cursors.WaitCursor

                hsCustomers = Map1.AddLayer(DocumentsDir & "\map\Districts\Districts.shp")
                hsCustomers.LegendText = "Customers spot analysis"

                'determine index of added layer
                layerIndex = Map1.Layers.IndexOf(hsCustomers)
                ' pass index for synchronization
                spotSynchIndex = layerIndex
                selectedIdx = layerIndex
                'apply symbology
                ApplyPointLayerStyle(layerIndex, "Customers")

                'activate synchronization menu
                If mSpotSynch.Enabled = False Then
                    mSpotSynch.Enabled = True
                End If
                SMS_sender.Cursor = Cursors.Default
                Map1.Layers(layerIndex).IsSelected = True
                'If layerIndex <> 0 Then
                '    Map1.Layers(layerIndex - 1).IsSelected = False
                'End If
                ZoomToLayer(Map1, layerIndex)

                sa_customers = True
            Else
                Map1.Layers.Remove(hsCustomers)
                'synchronization deactivation
                hotSpot_Synch_Deactivation()
                sa_customers = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub sDebtors_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sDebtors.Click

        Legend1.ClearSelection()
        Try


            If sDebtors.Checked = True Then

                SMS_sender.Cursor = Cursors.WaitCursor

                hsDebtors = Map1.AddLayer(DocumentsDir & "\map\Districts\Districts.shp")
                hsDebtors.LegendText = "Debtors spot analysis"

                'determine index of added layer
                layerIndex = Map1.Layers.IndexOf(hsDebtors)
                ' pass index for synchronization
                spotSynchIndex = layerIndex
                selectedIdx = layerIndex
                'apply symbology
                ApplyPointLayerStyle(layerIndex, "Debtors")
                'activate synchronization menu
                If mSpotSynch.Enabled = False Then
                    mSpotSynch.Enabled = True
                End If
                ZoomToLayer(Map1, layerIndex)
                SMS_sender.Cursor = Cursors.Default

                Map1.Layers(layerIndex).IsSelected = True
                'If layerIndex <> 0 Then
                '    Map1.Layers(layerIndex - 1).IsSelected = False
                'End If
                sa_debtors = True
            Else
                Map1.Layers.Remove(hsDebtors)
                'synchronization deactivation
                hotSpot_Synch_Deactivation()
                sa_debtors = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub sSales_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sSales.Click

        Legend1.ClearSelection()
        Try
            If sSales.Checked = True Then
                SMS_sender.Cursor = Cursors.WaitCursor

                hsSales = Map1.AddLayer(DocumentsDir & "\map\Districts\Districts.shp")
                hsSales.LegendText = "Sales spot analysis"
                Dim s As String = hsSales.ProjectionString
                'determine index of added layer
                layerIndex = Map1.Layers.IndexOf(hsSales)
                ' pass index for synchronization
                spotSynchIndex = layerIndex
                selectedIdx = layerIndex
                'apply symbology
                ApplyPointLayerStyle(layerIndex, "Sales")

                'activate synchronization menu
                If mSpotSynch.Enabled = False Then
                    mSpotSynch.Enabled = True
                End If
                ZoomToLayer(Map1, layerIndex)
                SMS_sender.Cursor = Cursors.Default
                Map1.Layers(layerIndex).IsSelected = True
                'If layerIndex <> 0 Then
                '    Map1.Layers(layerIndex - 1).IsSelected = False
                'End If
                'set sales anyalisis to be true
                'this will be used when synchronizing
                sa_sales = True
            Else
                Map1.Layers.Remove(hsSales)
                'synchronization deactivation
                hotSpot_Synch_Deactivation()
                sa_sales = False

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub RegionalDataToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mRegionalSynch.Click
        SMS_sender.Cursor = Cursors.WaitCursor
        Load_Shape_File_With_Focus_Data()

        If Not IsNothing(pSales) Then
            ApplyPolygonLayerStyle(Me.Map1, pSales, "Sales", color1, color2, 0.3)
        End If
        If Not IsNothing(pCustomers) Then
            ApplyPolygonLayerStyle(Me.Map1, pCustomers, "Customers", color1, color2, 0.3)
        End If
        If Not IsNothing(pDebtors) Then
            ApplyPolygonLayerStyle(Me.Map1, pDebtors, "Debtors", color1, color2, 0.3)
        End If
        Map1.Refresh()
        SMS_sender.Cursor = Cursors.Default
    End Sub

    Private Sub SpotDataToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mSpotSynch.Click

        SMS_sender.Cursor = Cursors.WaitCursor
        Load_Shape_File_Point_For_District_Capitals(spotSynchIndex)

        'apply symbology
        If Not IsNothing(hsCustomers) Then
            ApplyPointLayerStyle(spotSynchIndex, "Customers")
        End If
        If Not IsNothing(hsSales) Then
            ApplyPointLayerStyle(layerIndex, "Sales")
        End If
        If Not IsNothing(hsDebtors) Then
            ApplyPointLayerStyle(layerIndex, "Debtors")
        End If

        SMS_sender.Cursor = Cursors.Default

    End Sub

    Private Sub Region_Synch_Deactivation()

        'this code deactivate  sychronization for hotspot analysis
        If aSales.CheckState = CheckState.Unchecked And
            aCustomers.CheckState = CheckState.Unchecked And
            aDebtors.CheckState = CheckState.Unchecked Then

            mRegionalSynch.Enabled = False
        End If

    End Sub
    Private Sub hotSpot_Synch_Deactivation()

        'this code deactivate  sychronization for hotspot analysis
        If sSales.CheckState = CheckState.Unchecked And
            sCustomers.CheckState = CheckState.Unchecked And
            sDebtors.CheckState = CheckState.Unchecked Then

            mSpotSynch.Enabled = False
        End If

    End Sub

    Private Sub ShowLabelsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowLabelsToolStripMenuItem.Click
        Dim i As Integer = 0
        Dim a As Integer = Map1.Layers.IndexOf(Map1.Layers.SelectedLayer)
        If baseMap = True Then
            i = 1
        End If
        For x = i To Map1.Layers.Count - 1
            If x = a Then
                showLabel(x)
            End If
        Next

    End Sub

    Private Sub RemoveLabelsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveLabelsToolStripMenuItem.Click
        Dim i As Integer = 0
        If baseMap = True Then
            i = 1
        End If
        For x = i To Map1.Layers.Count - 1
            If Map1.Layers(x).IsSelected = True Then
                RemoveLabels(x)
            End If
        Next

    End Sub


    Private Sub mDistrictSynch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mDistrictSynch.Click

    End Sub

    Private Sub ZoomInToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZoomInToolStripMenuItem.Click
        Map1.FunctionMode = FunctionMode.ZoomIn

    End Sub

    Private Sub ZoomOutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZoomOutToolStripMenuItem.Click

        Map1.FunctionMode = FunctionMode.ZoomOut

    End Sub

    Private Sub SelectToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectToolStripMenuItem.Click
        Map1.FunctionMode = FunctionMode.Select
    End Sub

    Private Sub ZoomToNextToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZoomToNextToolStripMenuItem.Click
        Map1.ZoomToNext()
    End Sub

    Private Sub ZoomToPreviousToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZoomToPreviousToolStripMenuItem.Click
        Map1.ZoomToPrevious()
    End Sub

    Private Sub ZoomToExtentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZoomToExtentToolStripMenuItem.Click
        Map1.ZoomToMaxExtent()
    End Sub

    Private Sub PanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PanToolStripMenuItem.Click
        Map1.FunctionMode = FunctionMode.Pan
    End Sub

    Private Sub InfoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InfoToolStripMenuItem.Click
        Map1.FunctionMode = FunctionMode.Info
    End Sub

    Private Sub PointerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PointerToolStripMenuItem.Click
        Map1.FunctionMode = FunctionMode.None
    End Sub

    Private Sub UnselectToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UnselectToolStripMenuItem.Click
        Map1.ClearSelection()
    End Sub


    Private Sub InfoPanelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InfoPanelToolStripMenuItem.Click
        If mainContainer.Panel2Collapsed = False Then

            mainContainer.Panel2Collapsed = True




            '****************************************************
        Else
            mainContainer.Panel2Collapsed = False



            'Dim num As Integer = Map1.Layers.Count - 1
            'ShapefileAttribute(num)




        End If


        'If Map1.Layers.Count = 1 Then
        '    '********************************************
        '    Dim stateLayer As MapPolygonLayer
        '    stateLayer = CType(Map1.Layers(1), MapPolygonLayer)
        '    stateLayer.UnSelectAll()
        If baseMap = False Then
            Map1.ZoomToMaxExtent()
        End If


        '    If stateLayer Is Nothing Then
        '        MessageBox.Show("The layer is not a polygon layer.")


        '    End If
        'End If

        ' Map1.ZoomToMaxExtent()
    End Sub

    Private Sub PanelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PanelToolStripMenuItem.Click
        If mapContainer.Panel1Collapsed = True Then
            mapContainer.Panel1Collapsed = False

        Else
            mapContainer.Panel1Collapsed = True
        End If





    End Sub

    Private Sub ToolStrip3_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ToolStrip3.DoubleClick

        mainContainer.Panel2Collapsed = True

    End Sub

    Private Sub ToolStrip3_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStrip3.ItemClicked

    End Sub

    Private Sub ToolStripButton3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        mapContainer.Panel1Collapsed = True
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        mainContainer.Panel2Collapsed = True
    End Sub

    Private Sub ZoomToLayerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZoomToLayerToolStripMenuItem.Click
        ZoomToLayer(Map1, Map1.Layers.IndexOf(Map1.Layers.SelectedLayer))

    End Sub


    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPointLayer
            stateLayer = TryCast(Map1.Layers(0), MapPointLayer)
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable
                'Add new column

                'Dim column As New DataColumn("Customers", Type.GetType(("System.Int32")))
                'Dim s As New DataColumn("Sales", Type.GetType(("System.Int32")))
                'Dim d As New DataColumn("Debtors", Type.GetType(("System.Int32")))
                Dim a As Integer = 2
                Dim b As Integer = 5
                Dim c As Integer = 3
                'dt.Columns.Add(column)
                'dt.Columns.Add(s)
                'dt.Columns.Add(d)
                'calculate values
                For Each row As DataRow In dt.Rows
                    a = a + 2
                    b = b + 3
                    c = c + 4

                    row("CUSTOMERS") = a
                    row("SALES") = b
                    row("DEBTORS") = c

                Next
                'Set the datagridview datasource from datatable dt
                dgvAttributeTable.DataSource = dt
                stateLayer.DataSet.Save()

            End If
        Else
            MessageBox.Show("Please add a layer to the map.")
        End If
    End Sub



    Private Sub GreaterAccraToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GreaterAccraToolStripMenuItem.Click
        Dim Region As String = GreaterAccraToolStripMenuItem.Text

        Open_district_analysis(Region)
    End Sub
    Private Sub Open_district_analysis(ByVal region As String)
        Map1.Cursor = Cursors.AppStarting
        Dim f As New DistrictAnalysis
        f.lblTitle.Text = region & " District level analysis"
        If baseMap = True Then
            f.DistrictbaseMap = True
            f.typeOfBaseMap = baseMapType
        End If
        f.p_RegionSelected = region
        Map1.Cursor = Cursors.Default
        f.Show()



    End Sub
    Private Sub WesternToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WesternToolStripMenuItem.Click
        Dim Region As String = WesternToolStripMenuItem.Text
        Open_district_analysis(Region)
    End Sub

    Private Sub CentralToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CentralToolStripMenuItem.Click
        Dim Region As String = CentralToolStripMenuItem.Text
        Open_district_analysis(Region)
    End Sub

    Private Sub AshantiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AshantiToolStripMenuItem.Click
        Dim Region As String = AshantiToolStripMenuItem.Text
        Open_district_analysis(Region)
    End Sub

    Private Sub VoltaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VoltaToolStripMenuItem.Click
        Dim Region As String = VoltaToolStripMenuItem.Text
        Open_district_analysis(Region)
    End Sub

    Private Sub EasternToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EasternToolStripMenuItem.Click
        Dim Region As String = EasternToolStripMenuItem.Text
        Open_district_analysis(Region)
    End Sub

    Private Sub BrongAhafoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BrongAhafoToolStripMenuItem.Click
        Dim Region As String = BrongAhafoToolStripMenuItem.Text
        Open_district_analysis(Region)
    End Sub

    Private Sub NorthernToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NorthernToolStripMenuItem.Click
        Dim Region As String = NorthernToolStripMenuItem.Text
        Open_district_analysis(Region)
    End Sub

    Private Sub UpperEastToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpperEastToolStripMenuItem.Click
        Dim Region As String = UpperEastToolStripMenuItem.Text
        Open_district_analysis(Region)
    End Sub

    Private Sub UpperWestToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpperWestToolStripMenuItem.Click
        Dim Region As String = UpperWestToolStripMenuItem.Text
        Open_district_analysis(Region)
    End Sub

    'Private Sub Map1_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles Map1.MouseHover
    '    If info = False Then Exit Sub

    '    Try

    '        Dim region As String

    '        Dim fLayer As IFeatureLayer
    '        '
    '        'Here code for load fLayer
    '        If Map1.Layers.Count = 0 Then
    '            Exit Sub
    '        End If
    '        If baseMap = True And selectedIdx = 0 Then
    '            MsgBox("Base map layer is selected" & vbCrLf & "Ensure that the right layer is selected")

    '        Else

    '            'If Map1.Layers.SelectedLayer.GetType.ToString = "DotSpatial.Controls.MapPolygonLayer" Then
    '            fLayer = Map1.Layers(selectedIdx)


    '            '*****************************************REGION SELECTION
    '            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '            '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    '            If Not IsNothing(fLayer) Then
    '                Dim result As List(Of IFeature) = New List(Of IFeature)()
    '                fLayer.Selection.SelectionState = True
    '                result = fLayer.Selection.ToFeatureList
    '                If result.Count = 0 Then
    '                    Exit Sub
    '                End If

    '                For Each feature As IFeature In result
    '                    If Not IsDBNull(feature.ShapeIndex) Then
    '                        'MsgBox(fLayer.DataSet.ShapeIndices.IndexOf(feature.DataRow.Item(2))) '

    '                        region = feature.DataRow.Item(0).ToString
    '                        '  MsgBox(region.ToString)

    '                        ' Set up the delays for the ToolTip.
    '                        ToolTip1.AutoPopDelay = 5000
    '                        ToolTip1.InitialDelay = 1000
    '                        ToolTip1.ReshowDelay = 500
    '                        ' Force the ToolTip text to be displayed whether or not the form is active.
    '                        ToolTip1.ShowAlways = True

    '                        ' Set up the ToolTip text for the Button and Checkbox.
    '                        ToolTip1.SetToolTip(Me.Map1, region)
    '                    Else
    '                        MsgBox("No Region has been selected")
    '                    End If
    '                Next
    '            End If

    '            Map1.Cursor = Cursors.AppStarting
    '            'Dim f As New DistrictAnalysis
    '            'f.lblTitle.Text = region & " District level analysis"
    '            'f.p_RegionSelected = region


    '            'f.Show()
    '            'f.WindowState = FormWindowState.Normal

    '            'Map1.Cursor = Cursors.Default
    '            ''Else
    '            'MsgBox("Layer selected in the ledend is not a polygon")
    '        End If


    '        'End If


    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub



    Private Sub Map1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Map1.Load

    End Sub

    Private Sub ToolStripButton7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        If info = True Then
            info = False
            Map1.FunctionMode = FunctionMode.None
            ToolTip1.RemoveAll()
            ToolTip1.ShowAlways = False
            ToolStripButton7.BackColor = Color.Red
            '    SplitContainer2.Panel2Collapsed = True
        Else

            Map1.FunctionMode = FunctionMode.Select
            info = True
            ToolStripButton7.BackColor = Color.DarkRed
            SplitContainer2.Panel2Collapsed = False

        End If


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



    Private Sub Map1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Map1.MouseDoubleClick
        If info = False Then Exit Sub

        Try

            Dim value1, value2 As String

            Dim fLayer As IFeatureLayer
            '
            'Here code for load fLayer
            If Map1.Layers.Count = 0 Then
                Exit Sub
            End If
            If baseMap = True And selectedIdx = 0 Then
                MsgBox("Base map layer is selected" & vbCrLf & "Ensure that the right layer is selected")

            Else

                'If Map1.Layers.SelectedLayer.GetType.ToString = "DotSpatial.Controls.MapPolygonLayer" Then
                fLayer = Map1.Layers(selectedIdx)



                '*****************************************REGION SELECTION
                '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
                '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
                If Not IsNothing(fLayer) Then
                    Dim result As List(Of IFeature) = New List(Of IFeature)()
                    fLayer.Selection.SelectionState = True
                    result = fLayer.Selection.ToFeatureList
                    If result.Count = 0 Then
                        firstclick = True
                        Exit Sub
                    End If

                    For Each feature As IFeature In result
                        If Not IsDBNull(feature.ShapeIndex) Then
                            'MsgBox(fLayer.DataSet.ShapeIndices.IndexOf(feature.DataRow.Item(2))) '

                            If Map1.Layers.SelectedLayer.GetType.ToString = "DotSpatial.Controls.MapPolygonLayer" Then
                                value1 = feature.DataRow.Item("Region").ToString

                                'Display Customers
                                Display_Clients_in_Region_Polygon(value1)

                                'had to put s - 0 because of some bugs
                                lblName.Text = value1
                                If Not IsNothing(pSales) Then
                                    If pSales.IsSelected = True Then
                                        value2 = ""
                                        value2 = value1 & vbCrLf &
                                                 "Sales: GHS " & feature.DataRow.Item("SALES").ToString
                                    End If
                                End If



                                If Not IsNothing(pDebtors) Then
                                    If pDebtors.IsSelected = True Then
                                        value2 = ""
                                        value2 = value1 & vbCrLf &
                                                 "Debt: GHS " & feature.DataRow.Item("DEBTORS").ToString
                                    End If
                                End If

                                If Not IsNothing(pCustomers) Then

                                    If pCustomers.IsSelected = True Then
                                        value2 = ""
                                        value2 = value1 & vbCrLf &
                                                 "Customers: " & feature.DataRow.Item("CUSTOMERS").ToString
                                    End If

                                End If
                            End If

                            If Map1.Layers.SelectedLayer.GetType.ToString = "DotSpatial.Controls.MapPointLayer" Then

                                value1 = feature.DataRow.Item("DISTRICT").ToString

                                'Execute Fetch Customers
                                Display_Clients_in_District_Spot(value1)

                                lblName.Text = value1
                                If Not IsNothing(hsSales) Then
                                    If hsSales.IsSelected = True Then
                                        value2 = ""
                                        value2 = value1 & vbCrLf &
                                                 "Sales: GHS " & feature.DataRow.Item("SALES").ToString

                                    End If
                                End If


                                If Not IsNothing(hsDebtors) Then
                                    If hsDebtors.IsSelected = True Then
                                        value2 = ""
                                        value2 = value1 & vbCrLf &
                                                 "Debt: GHS " & feature.DataRow.Item("DEBTORS").ToString
                                    End If
                                End If


                                If Not IsNothing(hsCustomers) Then
                                    If hsCustomers.IsSelected = True Then
                                        value2 = ""
                                        value2 = value1 & vbCrLf &
                                                 "Customers: " & feature.DataRow.Item("CUSTOMERS").ToString
                                    End If
                                End If
                            End If


                            '  MsgBox(region.ToString)

                            ' Set up the delays for the ToolTip.
                            ToolTip1.AutoPopDelay = 5000
                            ToolTip1.InitialDelay = 1000
                            ToolTip1.ReshowDelay = 500
                            ToolTip1.BackColor = Color.LightYellow
                            ' Force the ToolTip text to be displayed whether or not the form is active.
                            ' ToolTip1.ShowAlways = True

                            ' Set up the ToolTip text for the Button and Checkbox.
                            ToolTip1.SetToolTip(Me.Map1, value2)
                        Else
                            MsgBox("No Region has been selected")
                        End If
                    Next
                End If

                'Map1.Cursor = Cursors.AppStarting
                'Dim f As New DistrictAnalysis
                'f.lblTitle.Text = region & " District level analysis"
                'f.p_RegionSelected = region


                'f.Show()
                'f.WindowState = FormWindowState.Normal

                'Map1.Cursor = Cursors.Default
                'Else
                '   MsgBox("Layer selected in the ledend is not a polygon")
            End If


            'End If


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        Map1.ClearSelection()
    End Sub

    Private Sub ToolStripButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton8.Click
        SplitContainer2.Panel2Collapsed = True
    End Sub

    Sub Display_Clients_in_Region_Polygon(ByVal RegionName As String)
        Load_MySettings_Variables()

        Clear_BatchBox()

        Dim Email As String

        Dim queryString As String = _
        "SELECT " & Customer_Field_Name & "," & Customer_Contact_Field1 & "," & Customer_Contact_Field2 & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & Region_Field & " = '" & RegionName & "'"


        If IsSQLSERVER Then

            Dim Mobile As String
            Dim Phone As String


            'Customized Variables
            Dim MyCustomer_Field_Name = "School"
            Dim MyCustomer_Contact_Field1 = "Telephone"
            Dim MyCustomer_Contact_Field2 = "Telephone2"
            Dim myCustomer_Email_Field = "Email"
            Dim myCustomer_Table_Name = "im_Data_SWL"




            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()

                        batchbox.Rows.Add(1)

                        Phone = dataReader(1).ToString
                        Mobile = dataReader(2).ToString
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


        'If IsPervasive Then

        '    Dim Email2 As String
        '    Dim id As String



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
        '                ' id = dataReader(4).ToString

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

        '            dataReader.Close()

        '        Catch ex As Exception
        '            MessageBox.Show(ex.Message)
        '        End Try

        '        connection.Close()
        '    End Using
        'End If




    End Sub


    Sub Display_Clients_in_District_Spot(ByVal DistrictName As String)
        Load_MySettings_Variables()

        Clear_BatchBox()

        Dim Email As String

        Dim queryString As String = _
   "SELECT " & Customer_Field_Name & "," & Customer_Contact_Field1 & "," & Customer_Contact_Field2 & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & District_Field & " = '" & DistrictName & "'"


        If IsSQLSERVER Then

            Dim Mobile As String
            Dim Phone As String


            'Customized Variables
            Dim MyCustomer_Field_Name = "School"
            Dim MyCustomer_Contact_Field1 = "Telephone"
            Dim MyCustomer_Contact_Field2 = "Telephone2"
            Dim myCustomer_Email_Field = "Email"
            Dim myCustomer_Table_Name = "im_Data_SWL"




            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()

                        batchbox.Rows.Add(1)

                        Phone = dataReader(1).ToString
                        Mobile = dataReader(2).ToString
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


        'If IsPervasive Then

        '    Dim Email2 As String
        '    Dim id As String



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
        '                ' id = dataReader(4).ToString

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

        '            dataReader.Close()

        '        Catch ex As Exception
        '            MessageBox.Show(ex.Message)
        '        End Try

        '        connection.Close()
        '    End Using
        'End If




    End Sub

    Sub Clear_BatchBox()
        batchbox.Rows.Clear()
        Batch_Count = 0

        Send_List_Title = ""

        ' txtMessage.Text = ""
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Reset_SMS_Sending_Status_Variables()
        Execute_Send_SMS()
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
        'lblAccountName.Text = SMS_Account_Name

        ' lblSMSCreditBalance.Text = SMS_Current_Account_Balance      ' Display_SMS_Account_Details()

        'Display Send Status in Tab
        Set_Text_Of_TabControl1(2, "Sent (" & SMS_Successufully_Sent & ")")
        Set_Text_Of_TabControl1(3, "Not Sent (" & SMS_Unsuccessfully_Sent & ")")

    End Sub

    Function Adroit_SMS_Report() As String
        Dim Report As String
        Report = SMS_Account_Name & " Sent: " & SMS_Successufully_Sent & ", " & Check_SMS_Account_Credits(SMS_Account_Name, SMS_Account_Password) & ", " & "Not Sent: " & SMS_Unsuccessfully_Sent
        Return Report
    End Function

    Sub Set_Text_Of_TabControl1(ByVal Tabindex As Integer, ByVal Text As String)
        TabControl1.TabPages(Tabindex).Text = Text

    End Sub
    Sub reset_tabcontrol_captions()
        Set_Text_Of_TabControl1(2, "Sent")
        Set_Text_Of_TabControl1(3, "Not Sent")
    End Sub
    Sub reset_send_listboxes()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()

    End Sub

    Private Sub batchbox_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles batchbox.CellContentClick

    End Sub

    Private Sub LocalMapToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LocalMapToolStripMenuItem.Click
        layerOpacity = 1
        AddLocalMap(Map1, Legend1, "Ghana", MapType.regional, LayerType.polygon)

    End Sub

    Private Sub EmploymentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menu_stats_employment.Click
        Legend1.ClearSelection()


        Try
            If menu_stats_employment.Checked = True Then
                'set reprojection mode always true so that the reprojection message will not appear
                Map1.ProjectionModeReproject = ActionMode.Always
                mapFeature_employment = Map1.AddLayer(DocumentsDir & "\map\stats\employment_data.shp")
                mapFeature_employment.LegendText = "Employment data"



                Dim MyFont As Font
                MyFont = New Font("Tahoma", 8.0)

                'reproject the map
                'pCustomers.Projection = KnownCoordinateSystems.Projected.UtmWgs1984.WGS1984ComplexUTMZone30N

                '8888888888888888888888888888888888888888888888
                ' Load_Shape_File_With_Focus_Data()
                '*************************************************


                'apply symbology
                ApplyPolygonLayerStyle(Me.Map1, mapFeature_employment, "all_employ", color1, color2, 0.3)
                'determine index
                layerIndex = Map1.Layers.IndexOf(mapFeature_employment)
                regionSynchIndex = layerIndex
                selectedIdx = layerIndex
                'zoom to layer
                ZoomToLayer(Map1, layerIndex)
                Map1.Layers(layerIndex).IsSelected = True




                'If layerIndex <> 0 Then
                '    Map1.Layers(layerIndex - 1).IsSelected = False
                'End If
                'activate synchronization menu
                If mRegionalSynch.Enabled = False Then
                    mRegionalSynch.Enabled = True
                End If


            Else
                Map1.Layers.Remove(mapFeature_employment)
                Region_Synch_Deactivation()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Private Sub OccupationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menu_stats_occupation.Click

        Legend1.ClearSelection()


        Try
            If menu_stats_occupation.Checked = True Then
                'set reprojection mode always true so that the reprojection message will not appear
                Map1.ProjectionModeReproject = ActionMode.Always
                mapFeature_occupation = Map1.AddLayer(DocumentsDir & "\map\stats\reg_occupation_data.shp")
                mapFeature_occupation.LegendText = "Occupation data"



                Dim MyFont As Font
                MyFont = New Font("Tahoma", 8.0)

                'reproject the map
                'pCustomers.Projection = KnownCoordinateSystems.Projected.UtmWgs1984.WGS1984ComplexUTMZone30N

                '8888888888888888888888888888888888888888888888
                ' Load_Shape_File_With_Focus_Data()
                '*************************************************


                'apply symbology
                ApplyPolygonLayerStyle(Me.Map1, mapFeature_occupation, "Managers", Color.LightYellow, Color.LightGreen, 0.3)
                'determine index
                layerIndex = Map1.Layers.IndexOf(mapFeature_occupation)
                regionSynchIndex = layerIndex
                selectedIdx = layerIndex
                'zoom to layer
                ZoomToLayer(Map1, layerIndex)
                Map1.Layers(layerIndex).IsSelected = True




                'If layerIndex <> 0 Then
                '    Map1.Layers(layerIndex - 1).IsSelected = False
                'End If
                'activate synchronization menu
                If mRegionalSynch.Enabled = False Then
                    mRegionalSynch.Enabled = True
                End If


            Else
                Map1.Layers.Remove(mapFeature_occupation)
                Region_Synch_Deactivation()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
End Class


