Imports DotSpatial.Controls
Imports DotSpatial.Symbology

Imports DotSpatial.Data
Imports DotSpatial.Serialization
Imports DotSpatial.Controls.BruTileLayer
Imports BruTile
Imports System.Drawing
Imports DotSpatial.Projections
Imports DotSpatial.Topology

Public Class frmEditTownName

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try

       
        If txtNTown.Text Is Nothing Then
            MsgBox("New town textbox empty" & vbCrLf & "Enter the correct spelling of town")
            Exit Sub
        End If

        Dim EditFeatureSet As FeatureSet
        Dim DocumentsDir As String = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim direct As String = "\map\Towns\"
        Dim shapeFilename As String = m_LI.Region

        
        EditFeatureSet = FeatureSet.Open(DocumentsDir & direct & shapeFilename & ".shp")

        m_LI.District = txtDistrictOfTown.Text.Trim
        m_LI.Town = txtOTown.Text.Trim
        m_LI.Region = txtRegionOfTown.Text.Trim

        For Each row As DataRow In EditFeatureSet.DataTable.Rows


            If m_LI.District = row("DISTRICT") Then



                If m_LI.Town = row("TOWN") Then

                    row("TOWN") = txtNTown.Text.Trim
                End If

                Exit For

            End If

        Next

        EditFeatureSet.Save()
            EditFeatureSet.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub frmEditTownName_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

       
                txtRegionOfTown.Text = m_LI.Region
            m_LI.Region = Conver_Region_to_ShortHand(m_LI.Region)




        
        txtDistrictOfTown.Text = m_LI.District
        txtOTown.Text = m_LI.Town
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        txtDistrictOfTown.Clear()
        txtNTown.Clear()
        txtOTown.Clear()

    End Sub
End Class