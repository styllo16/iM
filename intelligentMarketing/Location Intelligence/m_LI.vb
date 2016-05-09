Imports DotSpatial.Controls
Imports DotSpatial.Symbology
Imports DotSpatial.Data

Imports DotSpatial.Serialization
Imports DotSpatial.Controls.BruTileLayer
Imports BruTile
Imports System.Drawing
Imports DotSpatial.Projections
Imports DotSpatial.Topology

Module m_LI

    Dim mapSatellite As BruTileLayer
    Dim p_Region, p_Town As String
    Dim p_District, p_DistrictCode As String
    Dim p_form_is_lunched As Boolean = False
    Dim p_mapName, p_shapefilepath As String
    Dim p_Northing, p_Eastings As Double
    Dim DocumentsDir As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
    Dim p_PolygonLayerMap As IMapPolygonLayer
    Dim p_PointLayerMap As IMapPointLayer
    Dim p_IndexOfSelectedLayer As Integer
    Dim p_layerOpacity As Integer = 1
    Public Enum MapType
        district = 2
        regional = 1
        towns = 3
        stats = 4
    End Enum
    Public Enum LayerType
        polygon = 1
        point = 2
        line = 3


    End Enum





#Region "Properties"
    ''' <summary>
    ''' Set the opacity of the layer to a value
    ''' </summary>
    ''' <value>value between 0-1</value>
    ''' <returns></returns>
    ''' <remarks>Eg. layerOpacity = 0.2</remarks>

    Public Property layerOpacity As Int16
        Get
            layerOpacity = p_layerOpacity
        End Get
        Set(ByVal value As Int16)
            p_layerOpacity = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets that the form is being lunched for the first time
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property form_is_lunched As Boolean
        Get
            form_is_lunched = p_form_is_lunched
        End Get
        Set(ByVal value As Boolean)
            p_form_is_lunched = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the index of the selected layer added to the map
    ''' </summary>
    ''' <value></value>
    ''' <returns>integer</returns>
    ''' <remarks></remarks>
    Public Property IndexOfSelectedLayer As Int16
        Get
            IndexOfSelectedLayer = p_IndexOfSelectedLayer
        End Get
        Set(ByVal value As Int16)
            p_IndexOfSelectedLayer = value
        End Set
    End Property

    ''' <summary>
    ''' this contains mappolygonlayer
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property myPolygonLayer As IMapPolygonLayer
        Get
            myPolygonLayer = p_PolygonLayerMap
        End Get
        Set(ByVal value As IMapPolygonLayer)
            p_PolygonLayerMap = value
        End Set
    End Property
    ''' <summary>
    ''' This contains MapPointLayer
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property myPointLayer As IMapPointLayer
        Get
            myPointLayer = p_PointLayerMap
        End Get
        Set(ByVal value As IMapPointLayer)
            p_PointLayerMap = value
        End Set
    End Property
    ''' <summary>
    ''' Easting value or latitude of point
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Eastings As Double

        Get
            Eastings = p_Eastings
        End Get
        Set(ByVal value As Double)
            p_Eastings = value
        End Set


    End Property
    ''' <summary>
    ''' Northing or longititude value of point
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Northings As Double
        Get
            Northings = p_Northing
        End Get
        Set(ByVal value As Double)
            p_Northing = value
        End Set
    End Property
    ''' <summary>
    ''' Name of town
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Town As String
        Get
            Town = p_Town
        End Get
        Set(ByVal value As String)
            p_Town = value
        End Set
    End Property
    ''' <summary>
    ''' District code
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DistrictCode As String
        Get
            DistrictCode = p_DistrictCode
        End Get
        Set(ByVal value As String)
            p_DistrictCode = value
        End Set
    End Property
    Public Property shapefilepath As String
        Get
            shapefilepath = p_shapefilepath
        End Get
        Set(ByVal value As String)
            p_shapefilepath = value
        End Set
    End Property
   
    ''' <summary>
    ''' Shapefile name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MapName As String

        Get
            MapName = p_mapName
        End Get
        Set(ByVal value As String)
            p_mapName = value
        End Set
    End Property
    Public Property District As String
        Get
            District = p_District
        End Get
        Set(ByVal value As String)
            p_District = value
        End Set


    End Property

    Public Property Region As String
        Get
            Region = p_Region

        End Get
        Set(ByVal value As String)
            p_Region = value
        End Set
    End Property



#End Region

    ''' <summary>
    ''' This  a sub that adds  on-line base map
    ''' </summary>
    ''' <param name="maptype">the type of map to display</param>
    ''' <param name="Add">whether to add a map or remove</param>
    ''' <remarks></remarks>
    Public Sub AddBaseMap(ByVal Map As DotSpatial.Controls.Map, ByVal maptype As String, ByVal Add As Boolean,
                           ByVal baseMap As Boolean)


        Select Case maptype

            Case "Bing road map"

                'End If
                'mapSatellite.Reproject(KnownCoordinateSystems.Projected.UtmWgs1984.WGS1984UTMZone30N)

                If Add = True Then
                    If baseMap = True Then
                        Map.Layers.Remove(mapSatellite)
                    End If
                    'this code reproject the map


                    'this create the map
                    'change projection of map to webmercator
                    'changes the legend text
                    'insert the map at index layer 0

                    mapSatellite = BruTileLayer.CreateBingRoadsLayer
                    Map.ProjectionModeReproject = ActionMode.Always
                    Map.Projection = KnownCoordinateSystems.Projected.World.WebMercator
                    mapSatellite.LegendText = "Base map"
                    Map.Layers.Insert(0, mapSatellite) ' Map.Layers.Add(mapSatellite)

                Else
                    Map.Layers.Remove(mapSatellite)
                End If

            Case "Bing satellite"


                If Add = True Then
                    If baseMap = True Then
                        Map.Layers.Remove(mapSatellite)
                    End If
                    'this code reproject the map
                    Map.ProjectionModeReproject = ActionMode.Always

                    'this create the map
                    'change projection of map to webmercator
                    'changes the legend text
                    'insert the map at index layer 0
                    mapSatellite = BruTileLayer.CreateGoogleTerrainLayer
                    Map.Projection = KnownCoordinateSystems.Projected.World.WebMercator
                    mapSatellite.LegendText = "Base map"
                    Map.Layers.Insert(0, mapSatellite)
                Else
                    Map.Layers.Remove(mapSatellite)
                End If


            Case "Google map"


                If Add = True Then
                    If baseMap = True Then
                        Map.Layers.Remove(mapSatellite)
                    End If
                    'this code reproject the map
                    Map.ProjectionModeReproject = ActionMode.Always

                    'this create the map
                    'change projection of map to webmercator
                    'changes the legend text
                    'insert the map at index layer 0
                    mapSatellite = BruTileLayer.CreateGoogleMapLayer
                    Map.Projection = KnownCoordinateSystems.Projected.World.WebMercator
                    mapSatellite.LegendText = "Base map"
                    Map.Layers.Insert(0, mapSatellite)
                Else
                    Map.Layers.Remove(mapSatellite)
                End If

            Case "Open street"


                If Add = True Then
                    If baseMap = True Then
                        Map.Layers.Remove(mapSatellite)
                    End If
                    'this code reproject the map
                    Map.ProjectionModeReproject = ActionMode.Always

                    'this create the map
                    'change projection of map to webmercator
                    'changes the legend text
                    'insert the map at index layer 0
                    mapSatellite = BruTileLayer.CreateOsmLayer
                    Map.Projection = KnownCoordinateSystems.Projected.World.WebMercator
                    Map.Layers.Insert(0, mapSatellite)
                    mapSatellite.LegendText = "Base map"

                Else
                    Map.Layers.Remove(mapSatellite)
                End If

        End Select

    End Sub

    ''' <summary>
    ''' Add shapefile to the map
    ''' </summary>
    ''' <param name="Map">name of your map</param>
    ''' <param name="Legend">name of Legend</param>
    ''' <param name="shapeFileName">Name of shape file</param>
    ''' <param name="map_type">eg. MapType.District</param>
    '''  <param name="layer_type">eg. LayerType.point</param>
    ''' <remarks></remarks>
    Public Sub AddLocalMap(ByVal Map As DotSpatial.Controls.Map, ByVal Legend As DotSpatial.Controls.Legend, ByVal shapeFileName As String,
                           ByVal map_type As Int16, ByVal layer_type As Int16)

        Legend.ClearSelection()

        'determine index
        Dim layerIndex As Int16
        Dim direct As String

        If map_type = MapType.district Then
            direct = "\map\Districts\"
        ElseIf map_type = MapType.regional Then
            direct = "\map\"

        ElseIf map_type = MapType.towns Then
            direct = "\map\Towns\"

        ElseIf map_type = MapType.stats Then
            direct = "\map\stats\"
        End If

        If IsNothing(direct) Then
            MsgBox("Maptype not stated")
            Exit Sub
        End If
        p_shapefilepath = DocumentsDir & direct & shapeFileName & ".shp"
        ' Try

        'set reprojection mode always true so that the reprojection message will not appear
        Map.ProjectionModeReproject = ActionMode.Always

        If layer_type = LayerType.polygon Then
            ' p_PolygonLayerMap = 
            p_PolygonLayerMap = Map.AddLayer(p_shapefilepath)



            'Exit Sub
            p_PolygonLayerMap.LegendText = shapeFileName




            '*************************************************

            layerIndex = Map.Layers.IndexOf(p_PolygonLayerMap)

            'apply colour
            ApplyPolygonLayerColour(Map, p_PolygonLayerMap, Color.LightYellow, Color.LightGray, p_layerOpacity)


        ElseIf layer_type = LayerType.point Then

            p_PointLayerMap = Map.AddLayer(DocumentsDir & direct & shapeFileName & ".shp")
            p_PointLayerMap.LegendText = shapeFileName

            layerIndex = Map.Layers.IndexOf(p_PointLayerMap)


        End If


        'zoom to layer
        ZoomToLayer(Map, layerIndex)
        Map.Layers(layerIndex).IsSelected = True




        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

    End Sub
   
    ''' <summary>
    ''' Applys a colour to the polygon layer
    ''' </summary>
    ''' <param name="Map">name of map control</param>
    ''' <param name="PolygonLayer">name of polygon layer to apply colour</param>
    ''' <param name="LayerFillColor">Fill color. System.Drawing.Color</param>
    ''' <param name="LayerOutlineColor">Outline color. System.Drawing.Color</param>
    ''' <param name="Opacity">Transparency of the fill color. 0=transparent, 1=opaque. Single</param>
    ''' <remarks>This is a utility function for quickly applying colours and opacity to polygon layers</remarks>

    Public Sub ApplyPolygonLayerColour(ByVal Map As Map, ByVal PolygonLayer As IMapPolygonLayer, ByVal LayerFillColor As System.Drawing.Color, ByVal LayerOutlineColor As System.Drawing.Color, ByVal Opacity As Single)

        
        '***********************************************************************************
        ' Now(we) 'll get a reference to the PolygonLayer's Symbolizer and apply styling changes
        Dim MySymbolizer As PolygonSymbolizer = PolygonLayer.Symbolizer
        With MySymbolizer
            .SetFillColor(LayerFillColor.ToTransparent(Opacity))
            .SetOutline(Color.LightGray, 1)
        End With

        '*******************************************

    End Sub
    ''' <summary>
    ''' Displays the labels of the selected layer
    ''' </summary>
    ''' <param name="LabelColumnName">the name of the column which contains the information to be displayed</param>
    ''' <param name="PolygonLayer">the subject polygon layer </param>
    ''' <remarks>this is only for polygons</remarks>
    Public Sub AddLabel(ByVal LabelColumnName As String, ByVal PolygonLayer As IMapPolygonLayer)
        ' **************************************************************************
        '                      Labels
        Dim labelLayer As IMapLabelLayer = New MapLabelLayer()
        Dim category As ILabelCategory = labelLayer.Symbology.Categories(0)


        category.Expression = "[" & LabelColumnName & "]"

        'category.Symbolizer.BackColorEnabled = True
        ' category.Symbolizer.BackColor = Color.FromArgb(128, Color.LightBlue)
        category.Symbolizer.BorderVisible = False
        category.Symbolizer.FontStyle = FontStyle.Bold
        category.Symbolizer.FontColor = Color.Black
        category.Symbolizer.Orientation = ContentAlignment.MiddleCenter
        category.Symbolizer.Alignment = StringAlignment.Center
        PolygonLayer.ShowLabels = True
        PolygonLayer.LabelLayer = labelLayer
    End Sub


    ''' <summary>
    ''' Zoom to layer
    ''' </summary>
    ''' <param name="LayerIndex">index of layer </param>
    ''' <remarks></remarks>
    Public Sub ZoomToLayer(ByVal Map As DotSpatial.Controls.Map, ByVal LayerIndex As Integer)

        Try

            If LayerIndex <= Map.Layers.Count And LayerIndex > -1 Then
                Dim MyLayer As IMapLayer = Map.Layers(LayerIndex)
                If Not MyLayer Is Nothing Then
                    Map.ViewExtents = MyLayer.DataSet.Extent
                End If

            Else
                Exit Sub
                'MessageBox.Show("Select a layer")
            End If

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub

    ''' <summary>
    ''' Determines the index of the last layer added to the map
    ''' </summary>
    ''' <param name="map">name of map control</param>
    ''' <remarks>after executing this sub proceedure, the return value is stored in IndexOfSelectedLayer property</remarks>
    Public Sub FindIndexOfSelectedLayer(ByVal Map As DotSpatial.Controls.Map)

        Try

        
        Dim i As Int16 = 0
        p_IndexOfSelectedLayer = Nothing
        If Map.Layers.Count > 0 Then
            For x = i To Map.Layers.Count - 1
                If Map.Layers(x).IsSelected = True Then
                    p_IndexOfSelectedLayer = x
                End If
            Next
            If p_IndexOfSelectedLayer = Nothing Then
                MsgBox("No layer is selected")
            End If

        Else
            MsgBox("Add a layer to the map")
        End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    ''' <summary>
    ''' Function that convert long hand Region name into short abbreviation
    ''' </summary>
    ''' <param name="Region">eg. Greater Accra</param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Function Conver_Region_to_ShortHand(ByVal Region As String) As String
        Dim s_region As String
        Try

            Select Case Region

                Case "Greater Accra"
                    s_region = "GAR"

                    Return s_region
                Case "Ashanti"

                    s_region = "AR"
                    Return s_region

                Case "Eastern"


                    s_region = "ER"

                    Return s_region
                Case "Northern"
                    s_region = "NR"
                    Return s_region
                Case "Western"

                    s_region = "WR"

                    Return s_region
                Case "Central"

                    s_region = "CR"

                    Return s_region
                Case "Brong Ahafo"

                    s_region = "BAR"

                    Return s_region
                Case "Upper East"

                    s_region = "UER"

                    Return s_region
                Case "Upper West"

                    s_region = "UWR"

                    Return s_region
                Case "Volta"

                    s_region = "VR"
                    Return s_region
                Case Else
                    MsgBox("Region not found. Check spelling")
            End Select

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    ''' <summary>
    ''' This activate checkbox
    ''' </summary>
    ''' <param name="dgv">name of datagrid</param>
    ''' <param name="check1">sms check box</param>
    ''' <param name="check2">email chekc box</param>
    ''' <remarks>eg activate_checkbox(batch_box,smscheck,emailcheck)</remarks>
    Public Sub activate_checkbox(ByVal dgv As DataGridView, ByVal check1 As CheckBox, ByVal check2 As CheckBox)

        Dim total_row As Int16
        For total_row = 0 To dgv.RowCount - 1
            If check1.Enabled = False Then
                If dgv.Item(3, total_row).Value = True Then
                    check1.Enabled = True

                End If
            End If
            If check2.Enabled = False Then
                If dgv.Item(5, total_row).Value = True Then
                    check2.Enabled = True
                    Exit For

                End If
            End If
        Next




    End Sub
    ''' <summary>
    ''' Ticks or untick all checkboxes in the datagrid column
    ''' </summary>
    ''' <param name="checkbox">the name of checkbox</param>
    ''' <param name="column1">the column that contain the checkbox</param>
    ''' <param name="column2">the column that has to be check before the boxes are checked</param>
    ''' <param name="dgv"></param>
    ''' <remarks></remarks>
    Public Sub check_checkbox(ByVal checkbox As CheckBox, ByVal column1 As Integer, ByVal column2 As Integer, ByVal dgv As DataGridView)

        Dim total_row As Int16

        If checkbox.CheckState = CheckState.Unchecked Then
            checkbox.Checked = False
            For total_row = 0 To dgv.RowCount - 1

                dgv.Item(column1, total_row).Value = False
            Next
            Exit Sub
        Else
            checkbox.Checked = True
            For total_row = 0 To dgv.RowCount - 1
                If dgv.Item(column2, total_row).Value <> "" Then
                    dgv.Item(column1, total_row).Value = True
                End If
            Next
            Exit Sub
        End If
    End Sub
End Module
