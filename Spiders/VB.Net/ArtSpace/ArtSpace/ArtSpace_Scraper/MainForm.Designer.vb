<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
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
        Me.components = New System.ComponentModel.Container()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.WorkDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ThumbImageDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FullImageDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ArtistDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TagsDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.WorkURLDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PriceDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SizeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EditionSizeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PublishDateDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.InstitutionCnameDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PartnerURLDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PartnerDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ListingURLDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ArtSpaceFieldsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Button3 = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ArtSpaceFieldsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.WorkDataGridViewTextBoxColumn, Me.ThumbImageDataGridViewTextBoxColumn, Me.FullImageDataGridViewTextBoxColumn, Me.ArtistDataGridViewTextBoxColumn, Me.TagsDataGridViewTextBoxColumn, Me.WorkURLDataGridViewTextBoxColumn, Me.PriceDataGridViewTextBoxColumn, Me.SizeDataGridViewTextBoxColumn, Me.EditionSizeDataGridViewTextBoxColumn, Me.PublishDateDataGridViewTextBoxColumn, Me.InstitutionCnameDataGridViewTextBoxColumn, Me.PartnerURLDataGridViewTextBoxColumn, Me.PartnerDataGridViewTextBoxColumn, Me.ListingURLDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.ArtSpaceFieldsBindingSource
        Me.DataGridView1.Location = New System.Drawing.Point(8, 40)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(1085, 464)
        Me.DataGridView1.TabIndex = 0
        '
        'WorkDataGridViewTextBoxColumn
        '
        Me.WorkDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.WorkDataGridViewTextBoxColumn.DataPropertyName = "Work"
        Me.WorkDataGridViewTextBoxColumn.HeaderText = "Work"
        Me.WorkDataGridViewTextBoxColumn.Name = "WorkDataGridViewTextBoxColumn"
        '
        'ThumbImageDataGridViewTextBoxColumn
        '
        Me.ThumbImageDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.ThumbImageDataGridViewTextBoxColumn.DataPropertyName = "ThumbImage"
        Me.ThumbImageDataGridViewTextBoxColumn.HeaderText = "ThumbImage"
        Me.ThumbImageDataGridViewTextBoxColumn.Name = "ThumbImageDataGridViewTextBoxColumn"
        '
        'FullImageDataGridViewTextBoxColumn
        '
        Me.FullImageDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.FullImageDataGridViewTextBoxColumn.DataPropertyName = "FullImage"
        Me.FullImageDataGridViewTextBoxColumn.HeaderText = "FullImage"
        Me.FullImageDataGridViewTextBoxColumn.Name = "FullImageDataGridViewTextBoxColumn"
        '
        'ArtistDataGridViewTextBoxColumn
        '
        Me.ArtistDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.ArtistDataGridViewTextBoxColumn.DataPropertyName = "Artist"
        Me.ArtistDataGridViewTextBoxColumn.HeaderText = "Artist"
        Me.ArtistDataGridViewTextBoxColumn.Name = "ArtistDataGridViewTextBoxColumn"
        '
        'TagsDataGridViewTextBoxColumn
        '
        Me.TagsDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.TagsDataGridViewTextBoxColumn.DataPropertyName = "Tags"
        Me.TagsDataGridViewTextBoxColumn.HeaderText = "Tags"
        Me.TagsDataGridViewTextBoxColumn.Name = "TagsDataGridViewTextBoxColumn"
        '
        'WorkURLDataGridViewTextBoxColumn
        '
        Me.WorkURLDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.WorkURLDataGridViewTextBoxColumn.DataPropertyName = "WorkURL"
        Me.WorkURLDataGridViewTextBoxColumn.HeaderText = "WorkURL"
        Me.WorkURLDataGridViewTextBoxColumn.Name = "WorkURLDataGridViewTextBoxColumn"
        '
        'PriceDataGridViewTextBoxColumn
        '
        Me.PriceDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.PriceDataGridViewTextBoxColumn.DataPropertyName = "Price"
        Me.PriceDataGridViewTextBoxColumn.HeaderText = "Price"
        Me.PriceDataGridViewTextBoxColumn.Name = "PriceDataGridViewTextBoxColumn"
        '
        'SizeDataGridViewTextBoxColumn
        '
        Me.SizeDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.SizeDataGridViewTextBoxColumn.DataPropertyName = "Size"
        Me.SizeDataGridViewTextBoxColumn.HeaderText = "Size"
        Me.SizeDataGridViewTextBoxColumn.Name = "SizeDataGridViewTextBoxColumn"
        '
        'EditionSizeDataGridViewTextBoxColumn
        '
        Me.EditionSizeDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.EditionSizeDataGridViewTextBoxColumn.DataPropertyName = "EditionSize"
        Me.EditionSizeDataGridViewTextBoxColumn.HeaderText = "EditionSize"
        Me.EditionSizeDataGridViewTextBoxColumn.Name = "EditionSizeDataGridViewTextBoxColumn"
        '
        'PublishDateDataGridViewTextBoxColumn
        '
        Me.PublishDateDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.PublishDateDataGridViewTextBoxColumn.DataPropertyName = "PublishDate"
        Me.PublishDateDataGridViewTextBoxColumn.HeaderText = "PublishDate"
        Me.PublishDateDataGridViewTextBoxColumn.Name = "PublishDateDataGridViewTextBoxColumn"
        '
        'InstitutionCnameDataGridViewTextBoxColumn
        '
        Me.InstitutionCnameDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.InstitutionCnameDataGridViewTextBoxColumn.DataPropertyName = "InstitutionCname"
        Me.InstitutionCnameDataGridViewTextBoxColumn.HeaderText = "InstitutionCname"
        Me.InstitutionCnameDataGridViewTextBoxColumn.Name = "InstitutionCnameDataGridViewTextBoxColumn"
        '
        'PartnerURLDataGridViewTextBoxColumn
        '
        Me.PartnerURLDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.PartnerURLDataGridViewTextBoxColumn.DataPropertyName = "PartnerURL"
        Me.PartnerURLDataGridViewTextBoxColumn.HeaderText = "PartnerURL"
        Me.PartnerURLDataGridViewTextBoxColumn.Name = "PartnerURLDataGridViewTextBoxColumn"
        '
        'PartnerDataGridViewTextBoxColumn
        '
        Me.PartnerDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.PartnerDataGridViewTextBoxColumn.DataPropertyName = "Partner"
        Me.PartnerDataGridViewTextBoxColumn.HeaderText = "Partner"
        Me.PartnerDataGridViewTextBoxColumn.Name = "PartnerDataGridViewTextBoxColumn"
        '
        'ListingURLDataGridViewTextBoxColumn
        '
        Me.ListingURLDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.ListingURLDataGridViewTextBoxColumn.DataPropertyName = "ListingURL"
        Me.ListingURLDataGridViewTextBoxColumn.HeaderText = "ListingURL"
        Me.ListingURLDataGridViewTextBoxColumn.Name = "ListingURLDataGridViewTextBoxColumn"
        '
        'ArtSpaceFieldsBindingSource
        '
        Me.ArtSpaceFieldsBindingSource.DataSource = GetType(ScrapeLib.ArtSpaceFields)
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 518)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1103, 22)
        Me.StatusStrip1.TabIndex = 1
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(0, 17)
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(107, 7)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(89, 27)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Scrape"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(202, 7)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(89, 27)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Save"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1800000
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(12, 7)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(89, 27)
        Me.Button3.TabIndex = 4
        Me.Button3.Text = "Get URLs"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1103, 540)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Name = "MainForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Artspace Scraper"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ArtSpaceFieldsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents WorkDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ThumbImageDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FullImageDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ArtistDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TagsDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents WorkURLDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PriceDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SizeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EditionSizeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PublishDateDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents InstitutionCnameDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PartnerURLDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PartnerDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ListingURLDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ArtSpaceFieldsBindingSource As System.Windows.Forms.BindingSource

End Class
