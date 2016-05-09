Imports DotSpatial.Controls
Imports DotSpatial.Symbology

Imports DotSpatial.Data
Imports DotSpatial.Serialization
Imports DotSpatial.Controls.BruTileLayer
Imports BruTile
Imports System.Drawing
Imports DotSpatial.Projections
Imports DotSpatial.Topology

Public Class validation

    Dim coord As Coordinate
    Dim baseMap As Boolean

    Dim baseMapType As String
    Dim i As Int16
    Dim featureLayer As FeatureLayer
    Dim featureSet, EditFeatureSet As FeatureSet
    Dim currentFeature As IFeature
    'to differentiate the right and left mouse click
    Dim pointmouseClick As Boolean = False
    Dim region_and_code(24, 24) As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click


        RightPanel.Visible = False
    End Sub



    Private Sub mGmap_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mGmap.Click
        Try


            If mGmap.Checked = True Then

                'If baseMap = True Then
                '    Map1.Layers.Remove(mapSatellite)
                'End If
                SMS_sender.Cursor = Cursors.WaitCursor
                baseMapType = "Google map"
                m_LI.AddBaseMap(Map1, baseMapType, True, True)

                SMS_sender.Cursor = Cursors.Default

                mGmap.Checked = True
                mBRmap.Checked = False
                mBSmap.Checked = False
                mOsm.Checked = False





            Else
                SMS_sender.Cursor = Cursors.WaitCursor
                m_LI.AddBaseMap(Map1, baseMapType, True, False)
                SMS_sender.Cursor = Cursors.Default
                mGmap.Checked = False


            End If


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub mBRmap_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mBRmap.Click
        If mBRmap.Checked = True Then


            'If baseMap = True Then
            '    Map1.Layers.Remove(mapSatellite)
            'End If
            SMS_sender.Cursor = Cursors.WaitCursor
            baseMapType = "Bing road map"
            m_LI.AddBaseMap(Map1, baseMapType, True, True)

            SMS_sender.Cursor = Cursors.Default

            mBRmap.Checked = True
            mGmap.Checked = False
            mBSmap.Checked = False
            mOsm.Checked = False



        Else


            SMS_sender.Cursor = Cursors.WaitCursor
            m_LI.AddBaseMap(Map1, baseMapType, True, False)
            SMS_sender.Cursor = Cursors.Default

            mBRmap.Checked = False


        End If
    End Sub

    Private Sub mBSmap_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mBSmap.Click
        If mBSmap.Checked = True Then


            'If baseMap = True Then
            '    Map1.Layers.Remove(mapSatellite)
            'End If
            SMS_sender.Cursor = Cursors.WaitCursor
            baseMapType = "Bing satellite"
            m_LI.AddBaseMap(Map1, baseMapType, True, True)

            SMS_sender.Cursor = Cursors.Default

            mBSmap.Checked = True
            mBRmap.Checked = False
            mGmap.Checked = False
            mOsm.Checked = False


        Else
            SMS_sender.Cursor = Cursors.WaitCursor
            m_LI.AddBaseMap(Map1, baseMapType, True, False)
            SMS_sender.Cursor = Cursors.Default
            mBSmap.Checked = False


        End If

    End Sub

    Private Sub mOsm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mOsm.Click
        If mOsm.Checked = True Then


            'If baseMap = True Then
            '    Map1.Layers.Remove(mapSatellite)
            'End If
            SMS_sender.Cursor = Cursors.WaitCursor
            baseMapType = "Open street"
            m_LI.AddBaseMap(Map1, baseMapType, True, True)

            SMS_sender.Cursor = Cursors.Default

            mBRmap.Checked = False
            mBSmap.Checked = False
            mGmap.Checked = False
            mOsm.Checked = True



        Else
            SMS_sender.Cursor = Cursors.WaitCursor
            m_LI.AddBaseMap(Map1, baseMapType, True, False)
            SMS_sender.Cursor = Cursors.Default
            mOsm.Checked = False



        End If

    End Sub

    Private Sub LIEdit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RightPanel.Visible = False


        'codes for loading datatable without opening it as a layer
        'Dim DocumentsDir As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        'Dim direct As String = "\map\Towns\"
        'Dim shapeFilename As String = "GA"
        'Dim s As FeatureSet
        's = featureSet.Open(DocumentsDir & direct & shapeFilename & ".shp")
        'dgv.DataSource = s.DataTable

        '***************************************************
    End Sub


#Region "Add Region's map"
    Private Sub GreaterAccraToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GreaterAccraToolStripMenuItem.Click


        m_LI.MapName = GreaterAccraToolStripMenuItem.Text



        m_LI.AddLocalMap(Map1, Legend1, MapName, MapType.district, LayerType.polygon)

    End Sub

    Private Sub WesternToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WesternToolStripMenuItem.Click
        m_LI.MapName = WesternToolStripMenuItem.Text

        m_LI.AddLocalMap(Map1, Legend1, MapName, MapType.district, LayerType.polygon)
    End Sub

    Private Sub CentralToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CentralToolStripMenuItem.Click

        m_LI.MapName = CentralToolStripMenuItem.Text

        m_LI.AddLocalMap(Map1, Legend1, MapName, MapType.district, LayerType.polygon)

    End Sub

    Private Sub AshantiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AshantiToolStripMenuItem.Click


        m_LI.MapName = AshantiToolStripMenuItem.Text

        m_LI.AddLocalMap(Map1, Legend1, MapName, MapType.district, LayerType.polygon)

    End Sub

    Private Sub VoltaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VoltaToolStripMenuItem.Click
        m_LI.MapName = VoltaToolStripMenuItem.Text

        m_LI.AddLocalMap(Map1, Legend1, MapName, MapType.district, LayerType.polygon)
    End Sub

    Private Sub EasternToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EasternToolStripMenuItem.Click

        m_LI.MapName = EasternToolStripMenuItem.Text

        m_LI.AddLocalMap(Map1, Legend1, MapName, MapType.district, LayerType.polygon)

    End Sub

    Private Sub BrongAhafoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BrongAhafoToolStripMenuItem.Click

        m_LI.MapName = BrongAhafoToolStripMenuItem.Text

        m_LI.AddLocalMap(Map1, Legend1, MapName, MapType.district, LayerType.polygon)

    End Sub

    Private Sub NorthernToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NorthernToolStripMenuItem.Click

        m_LI.MapName = NorthernToolStripMenuItem.Text

        m_LI.AddLocalMap(Map1, Legend1, MapName, MapType.district, LayerType.polygon)

    End Sub

    Private Sub UpperEastToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpperEastToolStripMenuItem.Click

        m_LI.MapName = UpperEastToolStripMenuItem.Text

        m_LI.AddLocalMap(Map1, Legend1, MapName, MapType.district, LayerType.polygon)

    End Sub

    Private Sub UpperWestToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpperWestToolStripMenuItem.Click

        m_LI.MapName = UpperWestToolStripMenuItem.Text

        m_LI.AddLocalMap(Map1, Legend1, MapName, MapType.district, LayerType.polygon)

    End Sub
#End Region

    Private Sub btnZoomIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZoomIn.Click
        Map1.FunctionMode = FunctionMode.ZoomIn
        pointmouseClick = False
    End Sub

    Private Sub btnZoomOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnZoomOut.Click
        Map1.FunctionMode = FunctionMode.ZoomOut

        pointmouseClick = False
    End Sub

    Private Sub btnPan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPan.Click
        Map1.FunctionMode = FunctionMode.Pan
        pointmouseClick = False
    End Sub

    Private Sub btnPointer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPointer.Click
        Map1.FunctionMode = FunctionMode.None
        pointmouseClick = False
    End Sub

    Private Sub LIEdit_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        If form_is_lunched = True Then
            form_is_lunched = False
            Exit Sub
        End If
        ZoomToLayer(Map1, Map1.Layers.Count - 1)
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        ZoomToLayer(Map1, Map1.Layers.Count - 1)

    End Sub



    Private Sub btnStartEdit_ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStartEdit.ButtonClick
        If RightPanel.Visible = False Then
            RightPanel.Visible = True
            btnStartEdit.Text = "Stop edit"
            'codes for loading datatable without opening it as a layer
            Dim DocumentsDir As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Dim direct As String = "\map\Districts\"
            Dim shapeFilename As String = "Districts"

            EditFeatureSet = featureSet.Open(DocumentsDir & direct & shapeFilename & ".shp")
            'dgv.DataSource = s.DataTable
            txtRegion.Enabled = True
        Else
            RightPanel.Visible = False
            btnStartEdit.Text = "Start edit"
            pointmouseClick = False
            Map1.Cursor = Cursors.Default
            btnEdit.Enabled = False
            EditFeatureSet.Close()
        End If
    End Sub

    Private Sub Map1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Map1.MouseDown
        'Select Case shapeType
        '    Case "Point"
        Try


            If e.Button = MouseButtons.Left Then
                If (pointmouseClick) Then
                    'This method is used to convert the screen coordinate to map coordinate
                    'e.location is the mouse click point on the map control
                    coord = Map1.PixelToProj(e.Location)
                    txtE.Text = coord.X
                    txtN.Text = coord.Y
                    '*******************************************************


                    '  MsgBox(coord.ToString)

                    ' Exit Sub

                    'featureSet.DataTable.Rows[featureSet.DataTable.Rows.Count - 1][0] = featureSet.DataTable.Rows.Count;
                    '//featureSet.DataTable.Rows[featureSet.DataTable.Rows.Count - 1][1] = "123";              
                    featureLayer = Map1.Layers(Map1.Layers.IndexOf(myPointLayer))
                    featureSet = featureLayer.DataSet

                    Dim point As DotSpatial.Topology.Point = New DotSpatial.Topology.Point(coord)
                    Dim feature As Feature = New Feature(point)
                    currentFeature = featureSet.AddFeature(feature)



                    featureSet.InitializeVertices()
                    featureLayer.AssignFastDrawnStates()
                    '   Map1.MapFrame.Invalidate() ' ===================>this code will crash while running
                    Map1.Invalidate()


                    Map1.Refresh()
                    Map1.ResetBuffer()


                    '  Map1.MapFrame.Invalidate()
                    ' Map1.Invalidate()

                    '********************************




                End If
            Else
                Map1.Cursor = Cursors.Default
                pointmouseClick = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub spGA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles spGA.Click

        Try



            'Change the cursor style
            Map1.Cursor = Cursors.Cross

            AddLocalMap(Map1, Legend1, "GAR", MapType.towns, LayerType.point)
            txtRegion.Text = "GAR"


            ''******************************************
            'mypointlayer is ImapLayer in the module which stores the Layer added
            ' Dim pointLayer As MapPointLayer = myPointLayer

            ''Create a new symbolizer
            Dim symbol As New PointSymbolizer(Color.Red, DotSpatial.Symbology.PointShape.Ellipse, 3)

            'Set the symbolizer to the point layer
            myPointLayer.Symbolizer = symbol

            btnEdit.Enabled = True

            pointmouseClick = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        'validating the the town text box is not empty
        If txtTown.Text Is Nothing Then
            MsgBox("Cannot save" & vbCrLf & "Enter the name of the town")
            Exit Sub
        End If

        '***************************************************************************************
        'featureLayer = Map1.Layers(Map1.Layers.IndexOf(myPointLayer))
        'featureSet = featureLayer.DataSet

        'Dim point As DotSpatial.Topology.Point = New DotSpatial.Topology.Point(coord)
        'Dim feature As Feature = New Feature(point)
        'Dim currentFeature As IFeature = featureSet.AddFeature(feature)
        '****************************************************************************************

        'type the necessary codes to load all the districts in a combo
        'then read the District code and  Region info
        'then save it


        m_LI.Town = txtTown.Text.Trim
        m_LI.Region = txtRegion.Text.Trim
        m_LI.DistrictCode = txtDistcode.Text.Trim
        m_LI.District = cbDist.SelectedItem


        currentFeature.DataRow("TOWN") = m_LI.Town
        currentFeature.DataRow("DISTRICT") = m_LI.District
        currentFeature.DataRow("REGION") = m_LI.Region
        currentFeature.DataRow("DIST_CODE") = m_LI.DistrictCode



        featureSet.Save()
        featureSet.Close()
        ' EditFeatureSet.Close()

        featureSet.InitializeVertices()
        featureLayer.AssignFastDrawnStates()


        Map1.Refresh()
        Map1.ResetBuffer()

        txtTown.Clear()
    End Sub




    Private Sub txtRegion_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRegion.SelectedIndexChanged
        Dim comboRegion As String
        comboRegion = txtRegion.SelectedItem




        ' Dim district_code As Int16 = 1
        cbDist.Items.Clear()
        For Each row As DataRow In EditFeatureSet.DataTable.Rows

            If comboRegion = row("REGION") Then


                cbDist.Items.Add(row("DISTRICT"))
                'saving distric code and reading it once
                'If district_code = 1 Then
                '    DistrictCode = row("DIST_CODE")
                '    district_code = 0
                'End If

            End If

        Next

        'txtDistcode.Text = DistrictCode
        cbDist.SelectedIndex = 0
        cbDist.AutoCompleteSource = AutoCompleteSource.ListItems
        cbDist.AutoCompleteMode = AutoCompleteMode.Append
    End Sub

    Private Sub cbDist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbDist.SelectedIndexChanged

        'setting District in module to the selected combo value
        m_LI.District = cbDist.SelectedItem
        txtDistcode.Clear()

        For Each row As DataRow In EditFeatureSet.DataTable.Rows


            If m_LI.District = row("DISTRICT") Then



                m_LI.DistrictCode = row("DIST_CODE")
                Exit For

            End If

        Next
        txtDistcode.Text = DistrictCode
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click

        'add Label
        AddLabel("DISTRICT", myPolygonLayer)

    End Sub


    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        pointmouseClick = True
    End Sub

    Private Sub spW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles spW.Click
        Try



            'Change the cursor style
            Map1.Cursor = Cursors.Cross

            AddLocalMap(Map1, Legend1, "WR", MapType.towns, LayerType.point)
            txtRegion.Text = "WR"


            ''******************************************
            'mypointlayer is ImapLayer in the module which stores the Layer added
            ' Dim pointLayer As MapPointLayer = myPointLayer

            ''Create a new symbolizer
            Dim symbol As New PointSymbolizer(Color.Red, DotSpatial.Symbology.PointShape.Ellipse, 3)

            'Set the symbolizer to the point layer
            myPointLayer.Symbolizer = symbol

            btnEdit.Enabled = True
            pointmouseClick = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub spC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles spC.Click
        Try



            'Change the cursor style
            Map1.Cursor = Cursors.Cross

            AddLocalMap(Map1, Legend1, "CR", MapType.towns, LayerType.point)
            txtRegion.Text = "CR"


            ''******************************************
            'mypointlayer is ImapLayer in the module which stores the Layer added
            ' Dim pointLayer As MapPointLayer = myPointLayer

            ''Create a new symbolizer
            Dim symbol As New PointSymbolizer(Color.Red, DotSpatial.Symbology.PointShape.Ellipse, 3)

            'Set the symbolizer to the point layer
            myPointLayer.Symbolizer = symbol

            btnEdit.Enabled = True
            pointmouseClick = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub spA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles spA.Click
        Try



            'Change the cursor style
            Map1.Cursor = Cursors.Cross

            AddLocalMap(Map1, Legend1, "AR", MapType.towns, LayerType.point)
            txtRegion.Text = "AR"


            ''******************************************
            'mypointlayer is ImapLayer in the module which stores the Layer added
            ' Dim pointLayer As MapPointLayer = myPointLayer

            ''Create a new symbolizer
            Dim symbol As New PointSymbolizer(Color.Red, DotSpatial.Symbology.PointShape.Ellipse, 3)

            'Set the symbolizer to the point layer
            myPointLayer.Symbolizer = symbol

            btnEdit.Enabled = True
            pointmouseClick = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub spV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles spV.Click
        Try



            'Change the cursor style
            Map1.Cursor = Cursors.Cross

            AddLocalMap(Map1, Legend1, "VR", MapType.towns, LayerType.point)
            txtRegion.Text = "VR"


            ''******************************************
            'mypointlayer is ImapLayer in the module which stores the Layer added
            ' Dim pointLayer As MapPointLayer = myPointLayer

            ''Create a new symbolizer
            Dim symbol As New PointSymbolizer(Color.Red, DotSpatial.Symbology.PointShape.Ellipse, 3)

            'Set the symbolizer to the point layer
            myPointLayer.Symbolizer = symbol

            btnEdit.Enabled = True
            pointmouseClick = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub spE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles spE.Click
        Try



            'Change the cursor style
            Map1.Cursor = Cursors.Cross

            AddLocalMap(Map1, Legend1, "ER", MapType.towns, LayerType.point)
            txtRegion.Text = "ER"


            ''******************************************
            'mypointlayer is ImapLayer in the module which stores the Layer added
            ' Dim pointLayer As MapPointLayer = myPointLayer

            ''Create a new symbolizer
            Dim symbol As New PointSymbolizer(Color.Red, DotSpatial.Symbology.PointShape.Ellipse, 3)

            'Set the symbolizer to the point layer
            myPointLayer.Symbolizer = symbol

            btnEdit.Enabled = True
            pointmouseClick = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub spBA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles spBA.Click
        Try



            'Change the cursor style
            Map1.Cursor = Cursors.Cross

            AddLocalMap(Map1, Legend1, "BAR", MapType.towns, LayerType.point)
            txtRegion.Text = "BAR"


            ''******************************************
            'mypointlayer is ImapLayer in the module which stores the Layer added
            ' Dim pointLayer As MapPointLayer = myPointLayer

            ''Create a new symbolizer
            Dim symbol As New PointSymbolizer(Color.Red, DotSpatial.Symbology.PointShape.Ellipse, 3)

            'Set the symbolizer to the point layer
            myPointLayer.Symbolizer = symbol

            btnEdit.Enabled = True
            pointmouseClick = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub spN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles spN.Click
        Try



            'Change the cursor style
            Map1.Cursor = Cursors.Cross

            AddLocalMap(Map1, Legend1, "NR", MapType.towns, LayerType.point)
            txtRegion.Text = "NR"


            ''******************************************
            'mypointlayer is ImapLayer in the module which stores the Layer added
            ' Dim pointLayer As MapPointLayer = myPointLayer

            ''Create a new symbolizer
            Dim symbol As New PointSymbolizer(Color.Red, DotSpatial.Symbology.PointShape.Ellipse, 3)

            'Set the symbolizer to the point layer
            myPointLayer.Symbolizer = symbol

            btnEdit.Enabled = True
            pointmouseClick = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub spUE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles spUE.Click
        Try



            'Change the cursor style
            Map1.Cursor = Cursors.Cross

            AddLocalMap(Map1, Legend1, "UER", MapType.towns, LayerType.point)
            txtRegion.Text = "UER"


            ''******************************************
            'mypointlayer is ImapLayer in the module which stores the Layer added
            ' Dim pointLayer As MapPointLayer = myPointLayer

            ''Create a new symbolizer
            Dim symbol As New PointSymbolizer(Color.Red, DotSpatial.Symbology.PointShape.Ellipse, 3)

            'Set the symbolizer to the point layer
            myPointLayer.Symbolizer = symbol

            btnEdit.Enabled = True
            pointmouseClick = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub spUW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles spUW.Click
        Try



            'Change the cursor style
            Map1.Cursor = Cursors.Cross

            AddLocalMap(Map1, Legend1, "UWR", MapType.towns, LayerType.point)
            txtRegion.Text = "UWR"



            ''Create a new symbolizer
            Dim symbol As New PointSymbolizer(Color.Red, DotSpatial.Symbology.PointShape.Ellipse, 3)

            'Set the symbolizer to the point layer
            myPointLayer.Symbolizer = symbol

            btnEdit.Enabled = True
            pointmouseClick = True

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub txtTown_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTown.TextChanged

    End Sub
End Class