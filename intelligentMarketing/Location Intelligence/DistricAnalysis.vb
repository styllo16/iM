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


Public Class DistrictAnalysis
    Dim MyLayer, JHS, Primary, hsSales, District,
        hsDebtors, hsCustomers, pSales, pDebtors, pCustomers As IMapFeatureLayer
    Public p_RegionSelected, typeOfBaseMap As String
    Dim mapSatellite As BruTileLayer
    Dim maps As Integer
    Private idx As Integer
    Dim layerIndex, spotSynchIndex, districtSynchIndex, regionSynchIndex As Integer
    Private selectedIdx As Integer
    Public info, firstclick, DistrictbaseMap, sa_sales, sa_customers, sa_debtors As Boolean
    Dim reg As String
    Dim DocumentsDir As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

    Dim Rows_In_Backend_Database As Integer
    Dim Rows_In_Shapefile As Integer

    Dim Batch_Count As Integer

    Dim By_Debtors As Boolean = False


    Private Sub DistrictAnalysis_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized
        Legend1.ClearSelection()
        SplitContainer2.Panel2Collapsed = True
        'select regional file for spot analysis
        Select Case p_RegionSelected
            Case "Greater Accra"
                reg = "GAR"
            Case "Central"
                reg = "CR"
            Case "Ashanti"
                reg = "AR"
            Case "Volta"
                reg = "VR"
            Case "Western"
                reg = "WR"
            Case "Northern"
                reg = "NR"
            Case "Eastern"
                reg = "ER"
            Case "Upper East"
                reg = "UER"
            Case "Upper West"
                reg = "UWR"
            Case "Brong Ahafo"
                reg = "BAR"


        End Select


        Try
            'show on-line map if there is on-line map active in the main window
            If DistrictbaseMap = True Then
                SMS_sender.Cursor = Cursors.WaitCursor
                AddBaseMap(typeOfBaseMap, True)
                SMS_sender.Cursor = Cursors.Default
            End If

            Map1.ProjectionModeReproject = ActionMode.Always
            pCustomers = Map1.AddLayer(DocumentsDir & "\map\Districts\" & p_RegionSelected & ".shp")
            pCustomers.LegendText = "District analysis"


            'reproject the map



            'determine index
            layerIndex = Map1.Layers.IndexOf(pCustomers)
            'passing index to the synchronization index
            districtSynchIndex = layerIndex
            'zoom to layer
            ZoomToLayer(Map1, layerIndex)

            'apply symbology
            ApplyPolygonLayerStyle(Me.Map1, pCustomers, "Customers", Color.ForestGreen, Color.DarkSeaGreen, 0.3)

            Map1.Layers(layerIndex).IsSelected = True


            ''activate synchronization menu
            'If mRegionalSynch.Enabled = False Then
            '    mRegionalSynch.Enabled = True
            'End If
          

            aCustomers.Checked = True


            'AddLayer() method is used to add a shape file in the MapControl
            '  Map1.AddLayer(DocumentsDir & "\map\Districts\" & p_RegionSelected & ".shp").LegendText = p_RegionSelected & " Region"

            mainContainer.Panel2Collapsed = True

            'zooms map to extent
            '    Map1.ZoomToMaxExtent()

            ' cbAnalysis.SelectedIndex = 0

            'add analysis into the combo box
            ' cbAttribute.Items.Insert(0, "Regional analysis table")

            'display names of regions on map
            ' LabelsDisplay()

            'populate_selected_regions()

            ' symbolization("Customers", 0, "District analysis")




            ' This is the code that connects to the focus database and loads the information
            '   Load_Shape_File_With_Focus_Data()


            '*******************************************************
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
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

    Private Sub AddBaseMap(ByVal maptype As String, ByVal Add As Boolean)


        Select Case maptype

            Case "Bing road map"

                'End If
                'mapSatellite.Reproject(KnownCoordinateSystems.Projected.UtmWgs1984.WGS1984UTMZone30N)

                If Add = True Then
                    If DistrictbaseMap = True Then
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
                    If DistrictbaseMap = True Then
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
                    If DistrictbaseMap = True Then
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
                    If DistrictbaseMap = True Then
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

    Private Sub Region_Synch_Deactivation()

        'this code deactivate  sychronization for hotspot analysis
        If aSales.CheckState = CheckState.Unchecked And
            aCustomers.CheckState = CheckState.Unchecked And
            aDebtors.CheckState = CheckState.Unchecked Then

            ' mRegionalSynch.Enabled = False
        End If

    End Sub


    Private Sub btnZoomIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZoomIn.Click

        Map1.ZoomIn()

    End Sub

    Private Sub btnZoomOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZoomOut.Click

        Map1.ZoomOut()

    End Sub

    Private Sub btnZoomE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZoomE.Click

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
                Map1.AddLabels(stateLayer, "[DISTRICT]", New Font("Tahoma", 8.0), Color.Black)

            End If
        Else
            MessageBox.Show("Please add a layer to the map.")
        End If


    End Sub


    Private Sub symbolization(ByVal fieldName As String, ByVal idxLayer As Integer, ByVal analysis As String)


        Dim comboItemsNumber As Integer

        Try
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

                    


                            scheme.EditorSettings.StartColor = Color.LightYellow
                            scheme.EditorSettings.EndColor = Color.Red
                            scheme.EditorSettings.UseGradient = True

                            comboItemsNumber = cbAttribute.Items.Count
                            cbAttribute.Items.Insert(comboItemsNumber, "Central Region analysis table")
                            cbAttribute.SelectedIndex = comboItemsNumber

          
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

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub




    Private Sub ShapefileAttribute(ByVal layerIndex As Integer)

        selectedIdx = layerIndex

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Map1.Layers(selectedIdx).LegendText = "District capital" Then
            ShapefileAttributePoint(selectedIdx)
            Exit Sub
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim selectedLayer As String

        'Declare a datatable
        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPolygonLayer
            'this loop iterates all the items and compares to the desired layer
            '  For x = 0 To Map1.Layers.Count - 1
            stateLayer = CType(Map1.Layers(layerIndex), MapPolygonLayer)
            selectedLayer = stateLayer.LegendText
            DisplayAttributeLabel(selectedLayer)



            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                'Set the datagridview datasource from datatable dt
                dgvAttributeTable.DataSource = dt




            End If
        Else
            MessageBox.Show("Please add a layer to the map.")
        End If
    End Sub

    ''' <summary>
    ''' Add attribute of point file
    ''' </summary>
    ''' <param name="layerIndex"></param>
    ''' <remarks></remarks>

    Private Sub ShapefileAttributePoint(ByVal layerIndex As Integer)

        selectedIdx = layerIndex
        Dim selectedLayer As String

        'Declare a datatable
        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPointLayer
            'this loop iterates all the items and compares to the desired layer
            '  For x = 0 To Map1.Layers.Count - 1
            stateLayer = CType(Map1.Layers(layerIndex), MapPointLayer)
            stateLayer.LegendText = "District capital"
            selectedLayer = stateLayer.LegendText
            DisplayAttributeLabel(selectedLayer)



            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a point layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                'Set the datagridview datasource from datatable dt
                dgvAttributeTable.DataSource = dt




            End If
        Else
            MessageBox.Show("Please add a layer to the map.")
        End If
    End Sub

    ''' <summary>
    ''' Function to increase the index number of the layers added
    ''' anytime the primary or jhs enrollment is selected, then the index is counted
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    'Private Function index() As Integer

    '    idx = idx + 1
    '    Return idx
    'End Function

    ''' <summary>
    ''' Function to decrease the index number of the layers added
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function removeIndex()
        If idx <> 0 Then
            idx = idx - 1

        End If
        Return idx
    End Function


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


    'Private Sub cbAnalysis_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbAnalysis.SelectedIndexChanged

    '    Dim analysis As String
    '    analysis = cbAnalysis.SelectedItem
    '    Select Case analysis

    '        Case "SALES"

    '            symbolization("Sales", 0, "Distric analysis")

    '        Case "CUSTOMERS"

    '            symbolization("Customers", 0, "District analysis")

    '        Case "DEBTORS"

    '            symbolization("Debtors", 0, "District analysis")

    '    End Select

    'End Sub


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
        
        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' select district base on selection NEW CODE
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvAttributeTable_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAttributeTable.CellClick

        'If Map1.Layers(selectedIdx).LegendText = "District capital" Then


        '    For Each row As DataGridViewRow In dgvAttributeTable.SelectedRows
        '        Dim stateLayer As MapPointLayer
        '        stateLayer = CType(Map1.Layers(selectedIdx), MapPointLayer)
        '        If stateLayer Is Nothing Then
        '            MessageBox.Show("The layer is not a point layer.")
        '        Else
        '            stateLayer.SelectByAttribute("[Name] =" + "'" + row.Cells("Name").Value + "'")

        '        End If
        '    Next

        '    Exit Sub

        'End If


        'For Each row As DataGridViewRow In dgvAttributeTable.SelectedRows
        '    Dim stateLayer As MapPolygonLayer
        '    stateLayer = CType(Map1.Layers(selectedIdx), MapPolygonLayer)
        '    If stateLayer Is Nothing Then
        '        MessageBox.Show("The layer is not a polygon layer.")
        '    Else
        '        stateLayer.SelectByAttribute("[DISTRICT] =" + "'" + row.Cells("DISTRICT").Value + "'")

        '    End If
        'Next
        For Each row As DataGridViewRow In dgvAttributeTable.SelectedRows

            If Map1.Layers.SelectedLayer.GetType.ToString = "DotSpatial.Controls.MapPolygonLayer" Then
                Dim stateLayer As MapPolygonLayer
                stateLayer = CType(Map1.Layers(selectedIdx), MapPolygonLayer)

                If stateLayer Is Nothing Then
                    MessageBox.Show("The layer is not a polygon layer.")
                Else
                    stateLayer.SelectByAttribute("[DISTRICT] =" + "'" + row.Cells("DISTRICT").Value + "'")

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
    End Sub

    Private Sub DistrictAnalysis_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.SizeChanged

        
            If Map1.Layers.Count > 0 Then
                ZoomToLayer(Map1, Map1.Layers.Count - 1)
            End If

    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click

        Dim frm As New DotSpatial.Controls.LayoutForm
        frm.MapControl = Map1
        frm.Show()


    End Sub



    Private Sub Legend1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Legend1.Click



        ShapefileAttribute()

    End Sub


#Region "Codes fro loading data into shapefile for the various districts"

    Sub Load_Shape_File_With_Focus_Data()
        Dim i As Integer
        Dim Sales_data(10) As Double
        Dim Customer_Count(10) As Integer
        Dim Debtors_Data(10) As Integer

        '' Load Sales Data into Shape File
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


        ''Load Customer Distribution into Shape File

        'Try
        '    Customer_Count(0) = SWL_SCALAR_QUERY("Select count(dealers.code) from dealers, dealeraccts where dealers.code = dealeraccts.code and dealers.area = 'AR' and dealeraccts.accttype = 'SL'")
        '    Customer_Count(1) = SWL_SCALAR_QUERY("Select count(dealers.code) from dealers, dealeraccts where dealers.code = dealeraccts.code and dealers.area = 'CR' and dealeraccts.accttype = 'SL'")
        '    Customer_Count(2) = SWL_SCALAR_QUERY("Select count(dealers.code) from dealers, dealeraccts where dealers.code = dealeraccts.code and dealers.area = 'ER' and dealeraccts.accttype = 'SL'")
        '    Customer_Count(3) = SWL_SCALAR_QUERY("Select count(dealers.code) from dealers, dealeraccts where dealers.code = dealeraccts.code and dealers.area = 'GAR' and dealeraccts.accttype = 'SL'")
        '    Customer_Count(4) = SWL_SCALAR_QUERY("Select count(dealers.code) from dealers, dealeraccts where dealers.code = dealeraccts.code and dealers.area = 'NR' and dealeraccts.accttype = 'SL'")
        '    Customer_Count(5) = SWL_SCALAR_QUERY("Select count(dealers.code) from dealers, dealeraccts where dealers.code = dealeraccts.code and dealers.area = 'UER' and dealeraccts.accttype = 'SL'")
        '    Customer_Count(6) = SWL_SCALAR_QUERY("Select count(dealers.code) from dealers, dealeraccts where dealers.code = dealeraccts.code and dealers.area = 'UWR' and dealeraccts.accttype = 'SL'")
        '    Customer_Count(7) = SWL_SCALAR_QUERY("Select count(dealers.code) from dealers, dealeraccts where dealers.code = dealeraccts.code and dealers.area = 'VR' and dealeraccts.accttype = 'SL'")
        '    Customer_Count(8) = SWL_SCALAR_QUERY("Select count(dealers.code) from dealers, dealeraccts where dealers.code = dealeraccts.code and dealers.area = 'WR' and dealeraccts.accttype = 'SL'")
        '    Customer_Count(9) = SWL_SCALAR_QUERY("Select count(dealers.code) from dealers, dealeraccts where dealers.code = dealeraccts.code and dealers.area = 'BAR' and dealeraccts.accttype = 'SL'")

        'Catch ex As Exception

        'End Try


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
            stateLayer = TryCast(Map1.Layers(0), MapPolygonLayer)
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                For Each row As DataRow In dt.Rows
                    'Dim males As Double = row("Sales")
                    'males = males * 100
                    'row("Debtors") = males

                    'row("Sales") = Sales_data(i)
                    'row("Customers") = Customer_Count(i)
                    'row("Debtors") = Debtors_Data(i)
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

    ''' <summary>
    ''' Codes that load information into various regions
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
#Region "Population of regions"

#Region "Populate selected regions"

    Private Sub populate_selected_regions()

        Select Case p_RegionSelected

            Case "Upper East"

                upper_east()

            Case "Upper West"

                upper_west()

            Case "Northern"

                northern()

            Case "Brong Ahafo"

                brong_ahafo()

            Case "Ashanti"

                ashanti()

            Case "Volta"

                volta()

            Case "Eastern"

                eastern()

            Case "Western"

                western()

            Case "Central"

                central()

            Case "Greater Accra"

                greater_accra()

        End Select

    End Sub




#End Region

#Region "Upper East"

    Private Sub upper_east()

        Dim i As Integer
        Dim Sales_data(10) As Double
        Dim Customer_Count(10) As Integer
        Dim Debtors_Data(10) As Integer




        'Load Customer Distribution into Shape File

        Try
            Customer_Count(0) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Bawku West%'")
            Customer_Count(1) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Bongo%'")
            Customer_Count(2) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Builsa%'")
            Customer_Count(3) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Bawku Municipal%'")
            Customer_Count(4) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Bolgatanga%'")
            Customer_Count(5) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Talensi%'")
            Customer_Count(6) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Kasena Nankana East%'")
            Customer_Count(7) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Kasena Nankana West%'")
            Customer_Count(8) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Garu%'")


        Catch ex As Exception

        End Try

        'Try
        '    Customer_Count(0) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Bawku West%'")
        '    Customer_Count(1) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Bongo%'")
        '    Customer_Count(2) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Builsa%'")
        '    Customer_Count(3) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Bawku Municipal%'")
        '    Customer_Count(4) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Bolgatanga%'")
        '    Customer_Count(5) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Talensi%'")
        '    Customer_Count(6) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Kasena Nankana East%'")
        '    Customer_Count(7) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Kasena Nankana West%'")
        '    Customer_Count(8) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Garu%'")


        'Catch ex As Exception

        'End Try

        'Try
        '    Sales_data(0) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Bawku West%'")
        '    Sales_data(1) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Bongo%'")
        '    Sales_data(2) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Builsa%'")
        '    Sales_data(3) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Bawku Municipal%'")
        '    Sales_data(4) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Bolgatanga%'")
        '    Sales_data(5) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Talensi%'")
        '    Sales_data(6) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Kasena Nankana East%'")
        '    Sales_data(7) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Kasena Nankana West%'")
        '    Sales_data(8) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Upper East%' and " & District_Field & " like '%Garu%'")


        'Catch ex As Exception

        'End Try


        'Declare a datatable
        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPolygonLayer
            stateLayer = TryCast(Map1.Layers(0), MapPolygonLayer)
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                For Each row As DataRow In dt.Rows
                    'Dim males As Double = row("Sales")
                    'males = males * 100
                    'row("Debtors") = males


                    row("Customers") = Customer_Count(i)

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

#Region "Upper West"

    Private Sub upper_west()

        Dim i As Integer
        Dim Sales_data(10) As Double
        Dim Customer_Count(10) As Integer
        Dim Debtors_Data(10) As Integer
        Dim Distric_In_Shape_File As String

        'Declare a datatable
        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPolygonLayer
            stateLayer = TryCast(Map1.Layers(0), MapPolygonLayer)
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                For Each row As DataRow In dt.Rows
                    'Dim males As Double = row("Sales")
                    'males = males * 100
                    'row("Debtors") = males

                    Distric_In_Shape_File = row("DISTRICT")
                    row("Customers") = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & District_Field & " = '" & Distric_In_Shape_File & "'")

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

#Region "Northern"

    Private Sub northern()

        Dim i As Integer
        Dim Sales_data(10) As Double
        Dim Customer_Count(20) As Integer
        Dim Debtors_Data(10) As Integer




        'Load Customer Distribution into Shape File

        Try
            Customer_Count(0) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Mamprusi West%'")
            Customer_Count(1) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Savelugu%'")
            Customer_Count(2) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Tolon%'")
            Customer_Count(3) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Yendi%'")
            Customer_Count(4) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Tamale%'")
            Customer_Count(5) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Mamprusi East%'")
            Customer_Count(6) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Central Gonja%'")
            Customer_Count(7) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%West Gonja%'")
            Customer_Count(8) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Bole%'")
            Customer_Count(9) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Sawla%'")
            Customer_Count(10) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Nanumba North%'")
            Customer_Count(11) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Karaga%'")
            Customer_Count(12) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Gushiegu%'")
            Customer_Count(13) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Kpandai%'")
            Customer_Count(14) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%East Gonja%'")
            Customer_Count(15) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Zabzugu%'")
            Customer_Count(16) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Bunkpurugu%'")
            Customer_Count(17) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Nanumba South %'")
            Customer_Count(18) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Chereponi%'")
            Customer_Count(19) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Northern%' and " & District_Field & " like '%Saboba%'")







        Catch ex As Exception

        End Try

        'Declare a datatable
        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPolygonLayer
            stateLayer = TryCast(Map1.Layers(0), MapPolygonLayer)
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                For Each row As DataRow In dt.Rows
                    'Dim males As Double = row("Sales")
                    'males = males * 100
                    'row("Debtors") = males


                    row("Customers") = Customer_Count(i)

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

#Region "Brong Ahafo"

    Private Sub brong_ahafo()

        Dim i As Integer
        Dim Sales_data(10) As Double
        Dim Customer_Count(22) As Integer
        Dim Debtors_Data(10) As Integer




        'Load Customer Distribution into Shape File

        Try
            Customer_Count(0) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Asutifi%'")
            Customer_Count(1) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Berekum%'")
            Customer_Count(2) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Techiman%'")
            Customer_Count(3) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Kintampo North%'")
            Customer_Count(4) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Kintampo South%'")
            Customer_Count(5) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Wenchi%'")
            Customer_Count(6) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Tain%'")
            Customer_Count(7) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Jaman South%'")
            Customer_Count(8) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Jaman North%'")
            Customer_Count(9) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Asunafo North%'")
            Customer_Count(10) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Asunafo South%'")
            Customer_Count(11) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Tano North%'")
            Customer_Count(12) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Tano South%'")
            Customer_Count(13) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Pru%'")
            Customer_Count(14) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Atebubu%'")
            Customer_Count(15) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Dormaa East%'")
            Customer_Count(16) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Dormaa Municipal%'")
            Customer_Count(17) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Nkoranza South %'")
            Customer_Count(18) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Nkoranza North%'")
            Customer_Count(19) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Sunyani%'")
            Customer_Count(20) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Sunyani West%'")
            Customer_Count(21) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Brong%' and " & District_Field & " like '%Sene%'")







        Catch ex As Exception

        End Try

        'Declare a datatable
        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPolygonLayer
            stateLayer = TryCast(Map1.Layers(0), MapPolygonLayer)
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                For Each row As DataRow In dt.Rows
                    'Dim males As Double = row("Sales")
                    'males = males * 100
                    'row("Debtors") = males


                    row("Customers") = Customer_Count(i)

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

#Region "Volta"

    Private Sub volta()

        Dim i As Integer
        Dim Sales_data(10) As Double
        Dim Customer_Count(18) As Integer
        Dim Debtors_Data(10) As Integer




        'Load Customer Distribution into Shape File

        Try
            Customer_Count(0) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%Keta%'")
            Customer_Count(1) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%North Tongu%'")
            Customer_Count(2) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%Krachi East%'")
            Customer_Count(3) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%Krachi West%'")
            Customer_Count(4) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%South Dayi%'")
            Customer_Count(5) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%North Dayi%'")
            Customer_Count(6) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%South Tongu%'")
            Customer_Count(7) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%Biakoye%'")
            Customer_Count(8) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%Akatsi%'")
            Customer_Count(9) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%Hohoe%'")
            Customer_Count(10) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%Kadjebi%'")
            Customer_Count(11) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%Ho%'")
            Customer_Count(12) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%Adaklu Anyigbe%'")
            Customer_Count(13) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%Jasikan%'")
            Customer_Count(14) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%Ketu North%'")
            Customer_Count(15) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%Ketu South%'")
            Customer_Count(16) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%Nkwanta North%'")
            Customer_Count(17) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Volta%' and " & District_Field & " like '%Nkwanta South %'")







        Catch ex As Exception

        End Try

        'Declare a datatable
        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPolygonLayer
            stateLayer = TryCast(Map1.Layers(0), MapPolygonLayer)
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                For Each row As DataRow In dt.Rows
                    'Dim males As Double = row("Sales")
                    'males = males * 100
                    'row("Debtors") = males


                    row("Customers") = Customer_Count(i)

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

#Region "Ashanti"

    Private Sub ashanti()

        Dim i As Integer
        Dim Sales_data(10) As Double
        Dim Customer_Count(27) As Integer
        Dim Debtors_Data(10) As Integer




        'Load Customer Distribution into Shape File

        Try
            Customer_Count(0) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Ahafo Ano North%'")
            Customer_Count(1) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Ahafo Ano South%'")
            Customer_Count(2) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Amansie West%'")
            Customer_Count(3) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Asante Akim North%'")
            Customer_Count(4) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Asante Akim South%'")
            Customer_Count(5) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Ejura Sekyeredumase%'")
            Customer_Count(6) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Kumasi%'")
            Customer_Count(7) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Atwima Mponua%'")
            Customer_Count(8) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Atwima Nwabiagya%'")
            Customer_Count(9) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Amansi Central%'")
            Customer_Count(10) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Obuasi%'")
            Customer_Count(11) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Adansi North%'")
            Customer_Count(12) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Adansi South%'")
            Customer_Count(13) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Bosome Freho%'")
            Customer_Count(14) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Afiyga Sekyere%'")
            Customer_Count(15) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Sekyere Central%'")
            Customer_Count(16) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Mampong%'")
            Customer_Count(17) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Offinso North%'")
            Customer_Count(18) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Offinso Municipal%'")
            Customer_Count(19) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Ejisu Juaben Municipal%'")
            Customer_Count(20) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Sekyere East%'")
            Customer_Count(21) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Bekwai Municipal%'")
            Customer_Count(22) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Kwabre East%'")
            Customer_Count(23) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Afigya Kwabre%'")
            Customer_Count(24) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Atwima Kwanwoma%'")
            Customer_Count(25) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Bosumtwi%'")
            Customer_Count(26) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & District_Field & " like '%Sekyere Afram Plains%'")







        Catch ex As Exception

        End Try

        'Declare a datatable
        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPolygonLayer
            stateLayer = TryCast(Map1.Layers(0), MapPolygonLayer)
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                For Each row As DataRow In dt.Rows
                    'Dim males As Double = row("Sales")
                    'males = males * 100
                    'row("Debtors") = males


                    row("Customers") = Customer_Count(i)

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

#Region "Eastern"

    Private Sub eastern()

        Dim i As Integer
        Dim Sales_data(10) As Double
        Dim Customer_Count(21) As Integer
        Dim Debtors_Data(10) As Integer




        'Load Customer Distribution into Shape File

        Try
            Customer_Count(0) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Akwapem North%'")
            Customer_Count(1) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Akwapem South%'")
            Customer_Count(2) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Asuogyaman%'")
            Customer_Count(3) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Fanteakwa%'")
            Customer_Count(4) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Kwaebibirem%'")
            Customer_Count(5) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%New Juabeng%'")
            Customer_Count(6) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Suhum%'")
            Customer_Count(7) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%West Akim%'")
            Customer_Count(8) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Yilo Krobo%'")
            Customer_Count(9) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Kwahu West%'")
            Customer_Count(10) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Atiwa%'")
            Customer_Count(11) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%East Akim%'")
            Customer_Count(12) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Birim South%'")
            Customer_Count(13) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Birim Municipal%'")
            Customer_Count(14) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Kwahu East%'")
            Customer_Count(15) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Kwahu South%'")
            Customer_Count(16) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Lower Manya%'")
            Customer_Count(17) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Upper Manya%'")
            Customer_Count(18) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Birim North%'")
            Customer_Count(19) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Akyem Mansa%'")
            Customer_Count(20) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Eastern%' and " & District_Field & " like '%Kwahu North%'")






        Catch ex As Exception

        End Try

        'Declare a datatable
        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPolygonLayer
            stateLayer = TryCast(Map1.Layers(0), MapPolygonLayer)
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                For Each row As DataRow In dt.Rows
                    'Dim males As Double = row("Sales")
                    'males = males * 100
                    'row("Debtors") = males


                    row("Customers") = Customer_Count(i)

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

#Region "Western"

    Private Sub western()

        Dim i As Integer
        Dim Sales_data(10) As Double
        Dim Customer_Count(17) As Integer
        Dim Debtors_Data(10) As Integer




        'Load Customer Distribution into Shape File

        Try
            Customer_Count(0) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%' and " & District_Field & " like '%Ahanta West%'")
            Customer_Count(1) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%' and " & District_Field & " like '%Aowin%'")
            Customer_Count(2) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%' and " & District_Field & " like '%Bekwai%'")
            Customer_Count(3) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%' and " & District_Field & " like '%Jomoro%'")
            Customer_Count(4) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%' and " & District_Field & " like '%Bia%'")
            Customer_Count(5) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%' and " & District_Field & " like '%Juabeso%'")
            Customer_Count(6) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%' and " & District_Field & " like '%Wassa Amenfi East%'")
            Customer_Count(7) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%' and " & District_Field & " like '%Wassa Amenfi West%'")
            Customer_Count(8) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%' and " & District_Field & " like '%Prestea%'")
            Customer_Count(9) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%' and " & District_Field & " like '%Tarkwa%'")
            Customer_Count(10) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%' and " & District_Field & " like '%Sefwi Wiawso%'")
            Customer_Count(11) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%' and " & District_Field & " like '%Sefwi Akontombra%'")
            Customer_Count(12) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%' and " & District_Field & " like '%Shama%'")
            Customer_Count(13) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%' and " & District_Field & " like '%Takoradi%'")
            Customer_Count(14) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%' and " & District_Field & " like '%Mpohor%'")
            Customer_Count(15) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%' and " & District_Field & " like '%Nzema East%'")
            Customer_Count(16) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Western%' and " & District_Field & " like '%Ellembelle%'")






        Catch ex As Exception

        End Try

        'Declare a datatable
        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPolygonLayer
            stateLayer = TryCast(Map1.Layers(0), MapPolygonLayer)
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                For Each row As DataRow In dt.Rows
                    'Dim males As Double = row("Sales")
                    'males = males * 100
                    'row("Debtors") = males


                    row("Customers") = Customer_Count(i)

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

#Region "Central"

    Private Sub central()

        Dim i As Integer
        Dim Sales_data(10) As Double
        Dim Customer_Count(17) As Integer
        Dim Debtors_Data(10) As Integer




        'Load Customer Distribution into Shape File

        Try
            Customer_Count(0) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%' and " & District_Field & " like '%Ajumako%'")
            Customer_Count(1) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%' and " & District_Field & " like '%Abura%'")
            Customer_Count(2) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%' and " & District_Field & " like '%Asikuma%'")
            Customer_Count(3) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%' and " & District_Field & " like '%Cape Coast%'")
            Customer_Count(4) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%' and " & District_Field & " like '%Komenda%'")
            Customer_Count(5) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%' and " & District_Field & " like '%Mfantsiman%'")
            Customer_Count(6) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%' and " & District_Field & " like '%Lower Denkyira%'")
            Customer_Count(7) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%' and " & District_Field & " like '%Assin North%'")
            Customer_Count(8) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%' and " & District_Field & " like '%Assin South%'")
            Customer_Count(9) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%' and " & District_Field & " like '%Upper Denkyira West%'")
            Customer_Count(10) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%' and " & District_Field & " like '%Upper Denkyira East%'")
            Customer_Count(11) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%' and " & District_Field & " like '%Ewutu Senya%'")
            Customer_Count(12) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%' and " & District_Field & " like '%Effutu Municipal%'")
            Customer_Count(13) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%' and " & District_Field & " like '%Agona East%'")
            Customer_Count(14) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%' and " & District_Field & " like '%Agona West%'")
            Customer_Count(15) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%' and " & District_Field & " like '%Gomoa West%'")
            Customer_Count(16) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Central%' and " & District_Field & " like '%Gomoa East%'")






        Catch ex As Exception

        End Try

        'Declare a datatable
        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPolygonLayer
            stateLayer = TryCast(Map1.Layers(0), MapPolygonLayer)
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                For Each row As DataRow In dt.Rows
                    'Dim males As Double = row("Sales")
                    'males = males * 100
                    'row("Debtors") = males


                    row("Customers") = Customer_Count(i)

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

#Region "Greater Accra"

    Private Sub greater_accra()


        Dim i As Integer
        Dim Sales_data(10) As Double
        Dim Customer_Count(10) As Integer
        Dim Debtors_Data(10) As Integer




        'Load Customer Distribution into Shape File

        Try
            Customer_Count(0) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Dangbe East%'")
            Customer_Count(1) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%'  and " & District_Field & " like '%Dangbe West%'")
            Customer_Count(2) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Ga East%'")
            Customer_Count(3) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Ga West%'")
            Customer_Count(4) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Weija%'")
            Customer_Count(5) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Adenta%'")
            Customer_Count(6) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Tema%'")
            Customer_Count(7) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Ashaiman%'")
            Customer_Count(8) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Ledzokuku%'")
            Customer_Count(9) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Accra%'")



        Catch ex As Exception

        End Try


        'Try
        '    Customer_Count(0) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Dangbe East%'")
        '    Customer_Count(1) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Dangbe West%'")
        '    Customer_Count(2) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Ga East%'")
        '    Customer_Count(3) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Ga West%'")
        '    Customer_Count(4) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Weija%'")
        '    Customer_Count(5) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Adenta%'")
        '    Customer_Count(6) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Tema%'")
        '    Customer_Count(7) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Ashaiman%'")
        '    Customer_Count(8) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Ledzokuku%'")
        '    Customer_Count(9) = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Accra%'")



        'Catch ex As Exception

        'End Try

        'Try
        '    Sales_data(0) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Balance_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Dangbe East%'")
        '    Sales_data(1) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Balance_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Dangbe West%'")
        '    Sales_data(2) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Balance_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Ga East%'")
        '    Sales_data(3) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Balance_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Ga West%'")
        '    Sales_data(4) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Balance_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Weija%'")
        '    Sales_data(5) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Balance_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Adenta%'")
        '    Sales_data(6) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Balance_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Tema%'")
        '    Sales_data(7) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Balance_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Ashaiman%'")
        '    Sales_data(8) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Balance_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Ledzokuku%'")
        '    Sales_data(9) = SWL_SCALAR_QUERY("Select sum(" & Customer_Sales_Archive_Field & ") from " & Customer_Balance_Table_Name & " where " & Region_Field & " like '%Accra%' and " & District_Field & " like '%Accra%'")



        'Catch ex As Exception

        'End Try

        'Declare a datatable
        Dim dt As DataTable
        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPolygonLayer
            stateLayer = TryCast(Map1.Layers(0), MapPolygonLayer)
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a polygon layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                For Each row As DataRow In dt.Rows
                    'Dim males As Double = row("Sales")
                    'males = males * 100
                    'row("Debtors") = males


                    row("Customers") = Customer_Count(i)

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

#End Region



    'Private Sub showLabel()

    '    Try



    '        Dim cbAnalysisText As String
    '        cbAnalysisText = cbAnalysis.SelectedItem
    '        If Map1.Layers.Count > 0 Then
    '            Dim stateLayer As MapPolygonLayer
    '            'this loop iterates all the items and compares to the desired layer
    '            ' For x = 0 To Map1.Layers.Count - 1

    '            ' selIndx = Map1.Layers(x).LegendText

    '            'If Map1.Layers(x).IsSelected = True Then
    '            ' selectedIdx = x
    '            stateLayer = CType(Map1.Layers(0), MapPolygonLayer)


    '            Dim labelLayer As IMapLabelLayer
    '            labelLayer = New MapLabelLayer

    '            Dim cat As ILabelCategory
    '            cat = labelLayer.Symbology.Categories(0)

    '            If ShowLabelToolStripMenuItem.Checked = True Then
    '                If cbAnalysisText = "Sales" Then
    '                    cat.Expression = "[DISTRICT]" & vbNewLine & "GHS [" &
    '                                        cbAnalysisText & "]"
    '                Else
    '                    cat.Expression = "[DISTRICT]" & vbNewLine & "[" &
    '                        cbAnalysisText & "]"
    '                End If

    '                cat.Symbolizer.BackColorEnabled = True
    '                cat.Symbolizer.BackColor = Color.Transparent
    '                cat.Symbolizer.Orientation = ContentAlignment.MiddleCenter
    '                cat.Symbolizer.Alignment = StringAlignment.Center
    '                stateLayer.ShowLabels = True
    '                stateLayer.LabelLayer = labelLayer
    '            Else

    '                cat.Expression = "[DISTRICT]"
    '                cat.Symbolizer.Orientation = ContentAlignment.MiddleCenter
    '                cat.Symbolizer.Alignment = StringAlignment.Center
    '                cat.Symbolizer.FontSize = 8
    '                stateLayer.ShowLabels = True
    '                stateLayer.LabelLayer = labelLayer
    '            End If


    '        End If


    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub

    Private Sub dgvAttributeTable_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAttributeTable.CellContentClick

    End Sub

    Private Sub dgvAttributeTable_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAttributeTable.CellDoubleClick
        Selected_District_for_Drill_Down = dgvAttributeTable.Item(1, dgvAttributeTable.SelectedCells(0).RowIndex).Value.ToString
        Selected_District_Captial_for_Drill_Down = dgvAttributeTable.Item(2, dgvAttributeTable.SelectedCells(0).RowIndex).Value.ToString

        If dgvAttributeTable.Item(2, dgvAttributeTable.SelectedCells(0).RowIndex).Value.ToString = "Ashanti" Then
            Selected_District_for_Drill_Down = dgvAttributeTable.Item(3, dgvAttributeTable.SelectedCells(0).RowIndex).Value.ToString
            Selected_District_Captial_for_Drill_Down = dgvAttributeTable.Item(4, dgvAttributeTable.SelectedCells(0).RowIndex).Value.ToString


        End If

        Dim f As New District_Drilll_Downvb
        f.Show()

    End Sub

    Private Sub SalesMapToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalesMapToolStripMenuItem.Click

    End Sub

    Private Sub sSales_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sSales.Click

        Legend1.ClearSelection()

        Try
            Map1.ProjectionModeReproject = ActionMode.Always
            If sSales.Checked = True Then
                SMS_sender.Cursor = Cursors.WaitCursor

                hsSales = Map1.AddLayer(DocumentsDir & "\map\Towns\" & reg & ".shp")
                hsSales.LegendText = "Sales spot analysis"

                'determine index of added layer
                layerIndex = Map1.Layers.IndexOf(hsSales)
                ' pass index for synchronization
                spotSynchIndex = layerIndex

                'apply symbology
                ZoomToLayer(Map1, layerIndex)
                selectedIdx = layerIndex
                ApplyPointLayerStyle(layerIndex, "Sales")

                'activate synchronization menu
                If mSpotSynch.Enabled = False Then
                    mSpotSynch.Enabled = True
                End If

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
                'synchronization(deactivation)
                'hotSpot_Synch_Deactivation()
                sa_sales = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#Region "Functions"
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
                            scheme.EditorSettings.StartColor = (Color.Blue.ToTransparent(0.8))
                            scheme.EditorSettings.EndColor = (Color.Blue.ToTransparent(0.8))
                        Case "Sales"
                            scheme.EditorSettings.UseColorRange = True
                            scheme.EditorSettings.StartColor = (Color.Red.ToTransparent(0.8))
                            scheme.EditorSettings.EndColor = (Color.Red.ToTransparent(0.8))
                        Case "Debtors"
                            scheme.EditorSettings.UseColorRange = True
                            scheme.EditorSettings.StartColor = (Color.Green.ToTransparent(0.8))
                            scheme.EditorSettings.EndColor = (Color.Green.ToTransparent(0.8))
                    End Select


                    'Set the ClassificationType for the PolygonScheme via EditotSettings
                    scheme.EditorSettings.ClassificationType = ClassificationType.Quantities
                    scheme.EditorSettings.IntervalMethod = IntervalMethod.EqualInterval
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
    Sub ResetShapefile(ByVal index As Integer)
        Dim i As Integer
        Dim rowcount As String
        Dim Shapefile_Town As String
        Dim Shapefile_District As String

        Dim dt As DataTable

        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPointLayer
            stateLayer = TryCast(Map1.Layers(Index), MapPointLayer) 'change index
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a point layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                rowcount = dt.Rows.Count


                For Each row As DataRow In dt.Rows
                    row("Customers") = "0"
                Next

                dgvAttributeTable.DataSource = dt

                'saving to the shapefile
                stateLayer.DataSet.Save()
                stateLayer.UnSelectAll()
            End If
        Else
            MessageBox.Show("Please add a layer to the map.")
        End If




    End Sub
    Sub Insert_This_Town_Into_Shape_File(ByVal Value_For_Insertion As String, ByVal District_Parameter As String, ByVal Town_Parameter As String, ByVal Index As Integer)
        Dim i As Integer
        Dim rowcount As String
        Dim Shapefile_Town As String
        Dim Shapefile_District As String

        Dim dt As DataTable

        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPointLayer
            stateLayer = TryCast(Map1.Layers(Index), MapPointLayer) 'change index
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a point layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                rowcount = dt.Rows.Count


                For Each row As DataRow In dt.Rows
                    'Dim males As Double = row("Sales")
                    'males = males * 100
                    'row("Debtors") = males
                    Try
                        Shapefile_Town = row("TOWN")
                        Shapefile_District = row("DISTRICT")
                        ' region = row("REGION")
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try



                    If Shapefile_District = District_Parameter And Shapefile_Town = Town_Parameter Then
                        row("Customers") = Value_For_Insertion
                        'row("Debtors") = Debtors_Data(i)

                    Else

                        '   row("Customers") = "0"
                    End If





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


    End Sub


    Sub Optimized_Load_Shape_File_Point_With_Focus_Data(ByVal index As Integer)
        Dim i As Integer

        '******************************************************************************************
        'Reading from shapefile.
       
        Dim RowCount As Integer
        Dim Value_for_insert As String
        Dim This_Town As String
        Dim This_District As String

        RowCount = Rows_In_Backend_Database

        '***************Reset Shape File to Zero
        ResetShapefile(index)

        Dim queryString As String = _
         "SELECT " & Region_Field & ", " & District_Field & ", " & Customer_Location_Field & " from " & Customer_Table_Name


        If IsSQLSERVER Then
            Using connection As New SqlConnection(CREATE_CONNECTION_STRING(IsACCESS, IsSQLSERVER, MSACESSFILE, MSACCESSFILE_Path, SQLSERVER_NAME, SQL_Initial_Catalog, SQLUSERID, SQLPassword, IntegratedSecurity))
                Dim command As SqlCommand = connection.CreateCommand()
                command.CommandText = queryString
                Try
                    connection.Open()
                    Dim dataReader As SqlDataReader = _
                     command.ExecuteReader()
                    Do While dataReader.Read()
                        This_District = dataReader(1).ToString
                        This_Town = dataReader(2).ToString

                        Value_for_insert = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & District_Field & " = '" & This_District & "' and " & Customer_Location_Field & " = '" & This_Town & "'")
                        Insert_This_Town_Into_Shape_File(Value_for_insert, This_District, This_Town, index)

                        ToolStripProgressBar1.Value = (i / RowCount) * 100

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
                        This_District = dataReader(1).ToString
                        This_Town = dataReader(2).ToString

                        Value_for_insert = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & District_Field & " = '" & This_District & "' and " & Customer_Location_Field & " = '" & This_Town & "'")
                        Insert_This_Town_Into_Shape_File(Value_for_insert, This_District, This_Town, index)

                        ToolStripProgressBar1.Value = (i / RowCount) * 100

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
        '                This_District = dataReader(1).ToString
        '                This_Town = dataReader(2).ToString

        '                Value_for_insert = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & District_Field & " = '" & This_District & "' and " & Customer_Location_Field & " = '" & This_Town & "'")
        '                Insert_This_Town_Into_Shape_File(Value_for_insert, This_District, This_Town, index)


        '                ToolStripProgressBar1.Value = (i / RowCount) * 100

        '                i += 1
        '            Loop

        '            dataReader.Close()

        '        Catch ex As Exception
        '            MessageBox.Show(ex.Message)
        '        End Try

        '        connection.Close()
        '    End Using



        'End If



        ToolStripProgressBar1.Value = 0




    End Sub

    Function Fetch_Rows_In_Backend_Database() As Integer

        Return CInt(SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name))

    End Function

    Function Fetch_Rows_In_Shape_File(ByVal Index1 As Integer) As Integer
        Dim dt As DataTable

        If Map1.Layers.Count > 0 Then
            Dim stateLayer As MapPointLayer
            stateLayer = TryCast(Map1.Layers(Index1), MapPointLayer) 'change index
            If stateLayer Is Nothing Then
                MessageBox.Show("The layer is not a point layer.")
            Else
                'Get the shapefile's attribute table to our datatable dt
                dt = stateLayer.DataSet.DataTable

                Return dt.Rows.Count

            End If
        Else
            MessageBox.Show("Please add a layer to the map.")
        End If
    End Function

    Sub Load_Shape_File_Point_With_Focus_Data(ByVal index As Integer)
        Dim i As Integer

        '******************************************************************************************
        'Reading from shapefile.
        Dim towns As String
        Dim district As String
        Dim region As String
        Dim RowCount As Integer

        RowCount = Rows_In_Shapefile

        Dim dt As DataTable

        If Map1.Layers.Count > 0 Then

            If Map1.Layers.SelectedLayer.GetType.ToString = "DotSpatial.Controls.MapPolygonLayer" Then
                Dim stateLayer As MapPolygonLayer
                stateLayer = TryCast(Map1.Layers(index), MapPolygonLayer) 'change index
                If stateLayer Is Nothing Then
                    MessageBox.Show("The layer is not a selected.")
                Else
                    'Get the shapefile's attribute table to our datatable dt
                    dt = stateLayer.DataSet.DataTable
                    For Each row As DataRow In dt.Rows
                        Try
                            'Dim males As Double = row("Sales")
                            'males = males * 100
                            'row("Debtors") = males


                            district = Trim(row("DISTRICT"))
                            ' region = row("REGION")
                        Catch ex As Exception
                        End Try


                        Dim intCount As String
                        intCount = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & District_Field & " = '" & district & "'")
                        ' row("Sales") =
                        row("Customers") = intCount
                        'row("Debtors") = Debtors_Data(i)


                        Try
                            ToolStripProgressBar1.Value = (i / RowCount) * 100

                        Catch ex As Exception

                        End Try


                        i += 1

                    Next


                    '   "Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & Region_Field & "= 'A/R'"
                    'Set the datagridview datasource from datatable dt
                    dgvAttributeTable.DataSource = dt

                    'saving to the shapefile
                    stateLayer.DataSet.Save()
                    stateLayer.UnSelectAll()
                End If
            End If

                If Map1.Layers.SelectedLayer.GetType.ToString = "DotSpatial.Controls.MapPointLayer" Then
                    Dim stateLayer As MapPointLayer
                    stateLayer = TryCast(Map1.Layers(index), MapPointLayer) 'change index
                    If stateLayer Is Nothing Then
                        MessageBox.Show("The layer is not selected.")
                    Else
                        'Get the shapefile's attribute table to our datatable dt
                    dt = stateLayer.DataSet.DataTable
                    For Each row As DataRow In dt.Rows
                        Try
                            'Dim males As Double = row("Sales")
                            'males = males * 100
                            'row("Debtors") = males

                            towns = Trim(row("TOWN"))
                            district = Trim(row("DISTRICT"))
                            ' region = row("REGION")
                        Catch ex As Exception
                        End Try


                        Dim intCount As String
                        intCount = SWL_SCALAR_QUERY("Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & District_Field & " = '" & district & "' and " & Customer_Location_Field & " = '" & towns & "'")
                        ' row("Sales") =
                        row("Customers") = intCount
                        'row("Debtors") = Debtors_Data(i)


                        Try
                            ToolStripProgressBar1.Value = (i / RowCount) * 100
                        Catch ex As Exception

                        End Try


                        i += 1

                    Next


                    '   "Select count(" & Customer_Code_Field & ") from " & Customer_Table_Name & " where " & Region_Field & " like '%Ashanti%' and " & Region_Field & "= 'A/R'"
                    'Set the datagridview datasource from datatable dt
                    dgvAttributeTable.DataSource = dt

                    'saving to the shapefile
                    stateLayer.DataSet.Save()
                    stateLayer.UnSelectAll()
                End If
            End If


                    RowCount = Rows_In_Shapefile


                    

            Else
                MessageBox.Show("Please add a layer to the map.")
            End If

        '***********************************************************************************

        ToolStripProgressBar1.Value = 0
    End Sub

#End Region

    Private Sub sCustomers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sCustomers.Click

        Legend1.ClearSelection()
        Try

            Map1.ProjectionModeReproject = ActionMode.Always
            If sCustomers.Checked = True Then
                SMS_sender.Cursor = Cursors.WaitCursor

                hsCustomers = Map1.AddLayer(DocumentsDir & "\map\Towns\" & reg & ".shp")
                hsCustomers.LegendText = "Customers spot analysis"

                'determine index of added layer
                layerIndex = Map1.Layers.IndexOf(hsCustomers)
                ' pass index for synchronization
                spotSynchIndex = layerIndex
                selectedIdx = layerIndex
                ApplyPointLayerStyle(layerIndex, "Customers")
                'apply symbology
                ZoomToLayer(Map1, layerIndex)



                ''activate synchronization menu
                If mSpotSynch.Enabled = False Then
                    mSpotSynch.Enabled = True
                End If

                SMS_sender.Cursor = Cursors.Default
                Map1.Layers(layerIndex).IsSelected = True
                'If layerIndex <> 0 Then
                '    Map1.Layers(layerIndex - 1).IsSelected = False
                'End If

                sa_customers = True
            Else
                Map1.Layers.Remove(hsSales)
                'synchronization deactivation
                'hotSpot_Synch_Deactivation()
                sa_customers = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub sDebtors_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sDebtors.Click

        Legend1.ClearSelection()

        Try
            Map1.ProjectionModeReproject = ActionMode.Always
            If sDebtors.Checked = True Then
                SMS_sender.Cursor = Cursors.WaitCursor

                hsDebtors = Map1.AddLayer(DocumentsDir & "\map\Towns\" & reg & ".shp")
                hsDebtors.LegendText = "Debtors spot analysis"

                'determine index of added layer
                layerIndex = Map1.Layers.IndexOf(hsDebtors)
                ' pass index for synchronization
                spotSynchIndex = layerIndex
                selectedIdx = layerIndex
                'apply symbology
                ZoomToLayer(Map1, layerIndex)
                ApplyPointLayerStyle(layerIndex, "Debtors")

                ''activate synchronization menu
                If mSpotSynch.Enabled = False Then
                    mSpotSynch.Enabled = True
                End If

                SMS_sender.Cursor = Cursors.Default
                Map1.Layers(layerIndex).IsSelected = True
                'If layerIndex <> 0 Then
                '    Map1.Layers(layerIndex - 1).IsSelected = False
                'End If

                sa_debtors = True
            Else
                Map1.Layers.Remove(hsDebtors)
                'synchronization deactivation
                'hotSpot_Synch_Deactivation()
                sa_debtors = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub





    Private Sub ZoomInToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZoomInToolStripMenuItem.Click
        Map1.FunctionMode = FunctionMode.ZoomIn

    End Sub

    Private Sub ZoomOutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Map1.FunctionMode = FunctionMode.ZoomOut

    End Sub

    Private Sub SelectToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Map1.FunctionMode = FunctionMode.Select
    End Sub

    Private Sub ZoomToNextToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Map1.ZoomToNext()
    End Sub

    Private Sub ZoomToPreviousToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Map1.ZoomToPrevious()
    End Sub

    Private Sub ZoomToExtentToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Map1.ZoomToMaxExtent()
    End Sub

    Private Sub PanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Map1.FunctionMode = FunctionMode.Pan
    End Sub

    Private Sub InfoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Map1.FunctionMode = FunctionMode.Info
    End Sub

    Private Sub PointerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Map1.FunctionMode = FunctionMode.None
    End Sub

    Private Sub UnselectToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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
        If DistrictbaseMap = False Then
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

    Private Sub btAttribute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub mSpotSynch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mSpotSynch.Click
        Rows_In_Backend_Database = Fetch_Rows_In_Backend_Database()
        Rows_In_Shapefile = Fetch_Rows_In_Shape_File(spotSynchIndex)



        SMS_sender.Cursor = Cursors.WaitCursor

        If Rows_In_Backend_Database > Rows_In_Shapefile Then
            Load_Shape_File_Point_With_Focus_Data(spotSynchIndex)
        Else
            Optimized_Load_Shape_File_Point_With_Focus_Data(spotSynchIndex)
        End If



        'Optimized_Load_Shape_File_Point_With_Focus_Data(spotSynchIndex)
        'apply symbology
        If sa_customers = True Then
            ApplyPointLayerStyle(spotSynchIndex, "Customers")
        End If
        If sa_sales = True Then
            ApplyPointLayerStyle(layerIndex, "Sales")
        End If
        If sa_debtors = True Then
            ApplyPointLayerStyle(layerIndex, "Debtors")
        End If
        SMS_sender.Cursor = Cursors.Default

    End Sub


    Private Sub ApplyPolygonLayerStyle(ByVal Map As Map, ByVal PolygonLayer As IMapPolygonLayer, ByVal LabelColumnName As String, ByVal LayerFillColor As System.Drawing.Color, ByVal LayerOutlineColor As System.Drawing.Color, ByVal Opacity As Single)


        Try

      

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
        If DistrictbaseMap = True Then
            scheme.EditorSettings.StartColor = (Color.LightYellow.ToTransparent(0.3))
                scheme.EditorSettings.EndColor = (Color.DarkRed.ToTransparent(0.8))
            scheme.EditorSettings.UseGradient = False

        Else

            scheme.EditorSettings.StartColor = (Color.LightYellow)
            scheme.EditorSettings.EndColor = (Color.DarkRed)
            scheme.EditorSettings.UseGradient = True

        End If
        'Set the ClassificationType for the PolygonScheme via EditotSettings
        scheme.EditorSettings.ClassificationType = ClassificationType.Quantities

            scheme.EditorSettings.IntervalMethod = IntervalMethod.EqualInterval
            ' scheme.EditorSettings.IntervalSnapMethod = IntervalSnapMethod.DataValue
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
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    
    Private Sub cbAnalysis_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbAnalysis.SelectedIndexChanged
        Dim analysis As String

        analysis = cbAnalysis.Items(cbAnalysis.SelectedIndex)

        ApplyPolygonLayerStyle(Me.Map1, pSales, analysis, Color.ForestGreen, Color.DarkSeaGreen, 0.3)
    End Sub
    Dim clkInt As Int16 = 0
    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        If DistrictbaseMap = True Then
            If DistrictbaseMap = True Then
                SMS_sender.Cursor = Cursors.WaitCursor
                AddBaseMap(typeOfBaseMap, False)
                SMS_sender.Cursor = Cursors.Default
                ToolStripButton1.Text = "View base map"
                DistrictbaseMap = False
            Else
                If DistrictbaseMap = True Then
                    SMS_sender.Cursor = Cursors.WaitCursor
                    AddBaseMap(typeOfBaseMap, True)
                    SMS_sender.Cursor = Cursors.Default
                    ToolStripButton1.Text = "Hide base map"
                    DistrictbaseMap = True
                End If
            End If

        End If

    End Sub

    Private Sub Map1_Load(sender As System.Object, e As System.EventArgs) Handles Map1.Load

    End Sub

    Private Sub ToolStrip1_ItemClicked(sender As System.Object, e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

    End Sub

    Private Sub ToolStripDropDownButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripDropDownButton2.Click

    End Sub

    Private Sub dgvAttributeTable_Click(sender As Object, e As System.EventArgs) Handles dgvAttributeTable.Click

    End Sub
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
            If DistrictbaseMap = True And selectedIdx = 0 Then
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
                                value1 = feature.DataRow.Item("DISTRICT").ToString

                                Display_Clients_in_District_Spot(value1)
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

                                value1 = feature.DataRow.Item("TOWN").ToString

                                'MessageBox.Show(value1.ToString)

                                '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                                Display_Clients_in_Town_Spot(value1)

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
                            '  ToolTip1.ShowAlways = True

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
            MsgBox("Error in Pastel Connection")
        End Try
    End Sub
    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If info = True Then
            info = False
            Map1.FunctionMode = FunctionMode.None
            ToolTip1.RemoveAll()
            ToolTip1.ShowAlways = False
            ToolStripButton3.BackColor = Color.Red
            ' SplitContainer2.Panel2Collapsed = True
        Else

            Map1.FunctionMode = FunctionMode.Select
            info = True
            ToolStripButton3.BackColor = Color.DarkRed
            SplitContainer2.Panel2Collapsed = False


        End If
    End Sub

    Private Sub mRegionalSynch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'SMS_sender.Cursor = Cursors.WaitCursor
        'Load_Shape_File_With_Focus_Data()

        'If Not IsNothing(pSales) Then
        '    ApplyPolygonLayerStyle(Me.Map1, pSales, "Sales", Color.ForestGreen, Color.DarkSeaGreen, 0.3)
        'End If
        'If Not IsNothing(pCustomers) Then
        '    ApplyPolygonLayerStyle(Me.Map1, pCustomers, "Customers", Color.ForestGreen, Color.DarkSeaGreen, 0.3)
        'End If
        'If Not IsNothing(pDebtors) Then
        '    ApplyPolygonLayerStyle(Me.Map1, pDebtors, "Debtors", Color.ForestGreen, Color.DarkSeaGreen, 0.3)
        'End If
        'Map1.Refresh()
        'SMS_sender.Cursor = Cursors.Default
    End Sub

    Private Sub aCustomers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles aCustomers.Click
        Legend1.ClearSelection()

        Try
            If aCustomers.Checked = True Then
                'set reprojection mode always true so that the reprojection message will not appear
                Map1.ProjectionModeReproject = ActionMode.Always
                pCustomers = Map1.AddLayer(DocumentsDir & "\map\Districts\" & p_RegionSelected & ".shp")
                pCustomers.LegendText = "Customers analysis"



                Dim MyFont As Font
                MyFont = New Font("Tahoma", 8.0)

                'reproject the map
                'pCustomers.Projection = KnownCoordinateSystems.Projected.World.WebMercator

                '8888888888888888888888888888888888888888888888
                ' Load_Shape_File_With_Focus_Data()
                '*************************************************



                'determine index
                layerIndex = Map1.Layers.IndexOf(pCustomers)
                districtSynchIndex = layerIndex
                selectedIdx = layerIndex
                'zoom to layer
                ZoomToLayer(Map1, layerIndex)
                'apply symbology
                ApplyPolygonLayerStyle(Me.Map1, pCustomers, "Customers", Color.ForestGreen, Color.DarkSeaGreen, 0.3)

                Map1.Layers(layerIndex).IsSelected = True



                'If layerIndex <> 0 Then
                '    Map1.Layers(layerIndex - 1).IsSelected = False
                'End If
                'activate synchronization menu
                'If mRegionalSynch.Enabled = False Then
                '    mRegionalSynch.Enabled = True
                'End If


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
                pDebtors = Map1.AddLayer(DocumentsDir & "\map\Districts\" & p_RegionSelected & ".shp")
                pDebtors.LegendText = "Debtors analysis"

                Dim MyFont As Font
                MyFont = New Font("Tahoma", 8.0)

                'reproject the map
                'pDebtors.Projection = KnownCoordinateSystems.Projected.World.WebMercator

                'determine index
                layerIndex = Map1.Layers.IndexOf(pDebtors)
                districtSynchIndex = layerIndex
                selectedIdx = layerIndex
                'zoom to layer
                ZoomToLayer(Map1, layerIndex)

                'apply symbology
                ApplyPolygonLayerStyle(Me.Map1, pDebtors, "Debtors", Color.ForestGreen, Color.DarkSeaGreen, 0.3)
                Map1.Layers(layerIndex).IsSelected = True
                'If layerIndex <> 0 Then
                '    Map1.Layers(layerIndex - 1).IsSelected = False
                'End If
                'activate synchronization menu
                'If mRegionalSynch.Enabled = False Then
                '    mRegionalSynch.Enabled = True
                'End If
            Else
                Map1.Layers.Remove(pDebtors)
                Region_Synch_Deactivation()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub aSales_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles aSales.Click

        Legend1.ClearSelection()
        Try
            If aSales.Checked = True Then

                'Map1.Projection = KnownCoordinateSystems.Projected.World.WebMercator
                'set reprojection mode always true so that the reprojection message will not appear
                Map1.ProjectionModeReproject = ActionMode.Always
                pSales = Map1.AddLayer(DocumentsDir & "\map\Districts\" & p_RegionSelected & ".shp")
                pSales.LegendText = "Sales analysis"


                'reproject the map



                'determine index
                layerIndex = Map1.Layers.IndexOf(pSales)
                'passing index to the synchronization index
                districtSynchIndex = layerIndex
                selectedIdx = layerIndex
                'zoom to layer
                ZoomToLayer(Map1, layerIndex)
                'apply symbology
                ApplyPolygonLayerStyle(Me.Map1, pSales, "Sales", Color.ForestGreen, Color.DarkSeaGreen, 0.3)

                Map1.Layers(layerIndex).IsSelected = True
                'If layerIndex <> 0 Then
                '    Map1.Layers(layerIndex - 1).IsSelected = False
                'End If

                'activate synchronization menu
                'If mRegionalSynch.Enabled = False Then
                '    mRegionalSynch.Enabled = True
                'End If
            Else
                Map1.Layers.Remove(pSales)
                Region_Synch_Deactivation()
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        mapContainer.Panel1Collapsed = True
    End Sub

    Private Sub OptimizedSpotToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles OptimizedSpotToolStripMenuItem.Click
        '  Optimized_Load_Shape_File_Point_With_Focus_Data(spotSynchIndex)
    End Sub

    Private Sub mDistrictSynch_Click(sender As System.Object, e As System.EventArgs) Handles mDistrictSynch.Click
        SMS_sender.Cursor = Cursors.WaitCursor
        Load_Shape_File_Point_With_Focus_Data(districtSynchIndex)

        If Not IsNothing(pSales) Then
            ApplyPolygonLayerStyle(Me.Map1, pSales, "Sales", Color.ForestGreen, Color.DarkSeaGreen, 0.3)
        End If
        If Not IsNothing(pCustomers) Then
            ApplyPolygonLayerStyle(Me.Map1, pCustomers, "Customers", Color.ForestGreen, Color.DarkSeaGreen, 0.3)
        End If
        If Not IsNothing(pDebtors) Then
            ApplyPolygonLayerStyle(Me.Map1, pDebtors, "Debtors", Color.ForestGreen, Color.DarkSeaGreen, 0.3)
        End If
        Map1.Refresh()
        SMS_sender.Cursor = Cursors.Default
    End Sub

    Private Sub PanToolStripMenuItem_Click_1(sender As System.Object, e As System.EventArgs) Handles PanToolStripMenuItem.Click
        Map1.FunctionMode = FunctionMode.Pan
    End Sub

    Private Sub btnPan_Click(sender As System.Object, e As System.EventArgs) Handles btnPan.Click
        Map1.FunctionMode = FunctionMode.Pan
    End Sub

    Private Sub UnselectToolStripMenuItem_Click_1(sender As System.Object, e As System.EventArgs) Handles UnselectToolStripMenuItem.Click
        Map1.ClearSelection()
    End Sub

    Private Sub SelectToolStripMenuItem_Click_1(sender As System.Object, e As System.EventArgs) Handles SelectToolStripMenuItem.Click
        Map1.FunctionMode = FunctionMode.Select
    End Sub

    Private Sub PointerToolStripMenuItem_Click_1(sender As System.Object, e As System.EventArgs) Handles PointerToolStripMenuItem.Click
        Map1.FunctionMode = FunctionMode.None
    End Sub

    Private Sub InfoToolStripMenuItem_Click_1(sender As System.Object, e As System.EventArgs) Handles InfoToolStripMenuItem.Click

    End Sub

    Private Sub btnRefresh_Click(sender As System.Object, e As System.EventArgs) Handles btnRefresh.Click
        Map1.ClearSelection()
    End Sub

    Private Sub ToolStripButton8_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton8.Click
        SplitContainer2.Panel2Collapsed = True
    End Sub

    Sub Clear_BatchBox()
        batchbox.Rows.Clear()
        Batch_Count = 0

        Send_List_Title = ""

        ' txtMessage.Text = ""
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

        '            dataReader.Close()

        '        Catch ex As Exception
        '            MessageBox.Show(ex.Message)
        '        End Try

        '        connection.Close()
        '    End Using
        'End If




    End Sub


    Sub Display_Clients_in_Town_Spot(ByVal TownSpot As String)
        Load_MySettings_Variables()

        Clear_BatchBox()

        Dim Email As String

        Dim queryString As String = _
   "SELECT " & Customer_Field_Name & "," & Customer_Contact_Field1 & "," & Customer_Contact_Field2 & "," & Customer_Email_Field & " from " & Customer_Table_Name & " where " & Customer_Location_Field & " = '" & TownSpot & "'"


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
        '                '   id = dataReader(4).ToString

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
        '            MessageBox.Show("Error in Pastel Connection for Display Code")
        '        End Try

        '        connection.Close()
        '    End Using
        'End If




    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
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
        ' lblAccountName.Text = SMS_Account_Name

        'lblSMSCreditBalance.Text = SMS_Current_Account_Balance      ' Display_SMS_Account_Details()

        'Display Send Status in Tab
        Set_Text_Of_TabControl1(2, "Sent (" & SMS_Successufully_Sent & ")")
        Set_Text_Of_TabControl1(3, "Not Sent (" & SMS_Unsuccessfully_Sent & ")")

    End Sub

    Function Adroit_SMS_Report() As String
        Dim Report As String
        Report = SMS_Account_Name & " Sent: " & SMS_Successufully_Sent & ", " & Check_SMS_Account_Credits(SMS_Account_Name, SMS_Account_Password) & ", " & "Not Sent: " & SMS_Unsuccessfully_Sent
        Return Report
    End Function


    Private Sub UPPERWESTToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles UPPERWESTToolStripMenuItem.Click
        upper_west()


        'apply symbology
        ApplyPolygonLayerStyle(Me.Map1, pCustomers, "Customers", Color.ForestGreen, Color.DarkSeaGreen, 0.3)
        'determine index
        layerIndex = Map1.Layers.IndexOf(pCustomers)
        'zoom to layer
        ZoomToLayer(Map1, layerIndex)
        Map1.Layers(layerIndex).IsSelected = True

    End Sub
End Class

