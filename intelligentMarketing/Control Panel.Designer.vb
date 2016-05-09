<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Control_Panel
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btnRevenueCollection = New System.Windows.Forms.Button()
        Me.btnLocationIntelligence = New System.Windows.Forms.Button()
        Me.btnDigitalMarketing = New System.Windows.Forms.Button()
        Me.btnBusinessIntelligence = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(1244, 795)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(369, 25)
        Me.Label1.TabIndex = 25
        Me.Label1.Text = "Call 0268 313 343, 0541 235 271 "
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.TableLayoutPanel1.ColumnCount = 5
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 288.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 288.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.btnRevenueCollection, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.btnLocationIntelligence, 3, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.btnDigitalMarketing, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.btnBusinessIntelligence, 3, 3)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 62)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 5
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(934, 251)
        Me.TableLayoutPanel1.TabIndex = 26
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.ToolStrip1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(934, 62)
        Me.Panel1.TabIndex = 27
        '
        'ToolStrip1
        '
        Me.ToolStrip1.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripLabel1})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(934, 62)
        Me.ToolStrip1.TabIndex = 1
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.Font = New System.Drawing.Font("Segoe UI", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripLabel1.ForeColor = System.Drawing.Color.White
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(178, 59)
        Me.ToolStripLabel1.Text = "Control Panel"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.PictureBox1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 313)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(934, 90)
        Me.Panel2.TabIndex = 28
        '
        'btnRevenueCollection
        '
        Me.btnRevenueCollection.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnRevenueCollection.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnRevenueCollection.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRevenueCollection.Image = Global.intelligentMarketing.My.Resources.Resources.debLogo2
        Me.btnRevenueCollection.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnRevenueCollection.Location = New System.Drawing.Point(172, 18)
        Me.btnRevenueCollection.Name = "btnRevenueCollection"
        Me.btnRevenueCollection.Size = New System.Drawing.Size(282, 94)
        Me.btnRevenueCollection.TabIndex = 23
        Me.btnRevenueCollection.Text = "Revenue Collection"
        Me.btnRevenueCollection.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnRevenueCollection.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnRevenueCollection.UseVisualStyleBackColor = True
        '
        'btnLocationIntelligence
        '
        Me.btnLocationIntelligence.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnLocationIntelligence.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLocationIntelligence.Image = Global.intelligentMarketing.My.Resources.Resources.li2
        Me.btnLocationIntelligence.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnLocationIntelligence.Location = New System.Drawing.Point(480, 18)
        Me.btnLocationIntelligence.Name = "btnLocationIntelligence"
        Me.btnLocationIntelligence.Size = New System.Drawing.Size(282, 94)
        Me.btnLocationIntelligence.TabIndex = 21
        Me.btnLocationIntelligence.Text = "Location Intelligence"
        Me.btnLocationIntelligence.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnLocationIntelligence.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnLocationIntelligence.UseVisualStyleBackColor = True
        '
        'btnDigitalMarketing
        '
        Me.btnDigitalMarketing.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnDigitalMarketing.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDigitalMarketing.Image = Global.intelligentMarketing.My.Resources.Resources.smsLogo2
        Me.btnDigitalMarketing.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDigitalMarketing.Location = New System.Drawing.Point(172, 138)
        Me.btnDigitalMarketing.Name = "btnDigitalMarketing"
        Me.btnDigitalMarketing.Size = New System.Drawing.Size(282, 94)
        Me.btnDigitalMarketing.TabIndex = 20
        Me.btnDigitalMarketing.Text = "Digital Marketing"
        Me.btnDigitalMarketing.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnDigitalMarketing.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnDigitalMarketing.UseVisualStyleBackColor = True
        '
        'btnBusinessIntelligence
        '
        Me.btnBusinessIntelligence.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btnBusinessIntelligence.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBusinessIntelligence.Image = Global.intelligentMarketing.My.Resources.Resources.BiLogo2
        Me.btnBusinessIntelligence.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnBusinessIntelligence.Location = New System.Drawing.Point(480, 138)
        Me.btnBusinessIntelligence.Name = "btnBusinessIntelligence"
        Me.btnBusinessIntelligence.Size = New System.Drawing.Size(282, 94)
        Me.btnBusinessIntelligence.TabIndex = 22
        Me.btnBusinessIntelligence.Text = "Business Intelligence"
        Me.btnBusinessIntelligence.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnBusinessIntelligence.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnBusinessIntelligence.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Right
        Me.PictureBox1.Image = Global.intelligentMarketing.My.Resources.Resources.logo1
        Me.PictureBox1.Location = New System.Drawing.Point(682, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(252, 90)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'Control_Panel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(934, 403)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Control_Panel"
        Me.Text = "Control_Panel"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnRevenueCollection As System.Windows.Forms.Button
    Friend WithEvents btnBusinessIntelligence As System.Windows.Forms.Button
    Friend WithEvents btnDigitalMarketing As System.Windows.Forms.Button
    Friend WithEvents btnLocationIntelligence As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents ToolStripLabel1 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
End Class
