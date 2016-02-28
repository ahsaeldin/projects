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
        Me.ScrapingDataGridView = New System.Windows.Forms.DataGridView()
        Me.Keyword = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ScrapeButton = New System.Windows.Forms.Button()
        Me.SaveButton = New System.Windows.Forms.Button()
        Me.UseProxiesCheckBox = New System.Windows.Forms.CheckBox()
        Me.ScrapingStatusStrip = New System.Windows.Forms.StatusStrip()
        Me.MainFormToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.SavingTimer = New System.Windows.Forms.Timer(Me.components)
        Me.ScrapingBackgroundWorker = New System.ComponentModel.BackgroundWorker()
        Me.ScrapeFromSourceButton = New System.Windows.Forms.Button()
        Me.StopButton = New System.Windows.Forms.Button()
        Me.DelayCheckBox = New System.Windows.Forms.CheckBox()
        Me.DelayNumericUpDown = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ReScrapeFailedCheckBox = New System.Windows.Forms.CheckBox()
        Me.DeleteBackupsCheckBox = New System.Windows.Forms.CheckBox()
        Me.ExtactResultsNoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.InURLResutlsNoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.InTitleResultsNoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.InURLinTitleSumDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SearchPagesNoDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GoogleResultsStatsFieldsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.ScrapedToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.FailedToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        CType(Me.ScrapingDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ScrapingStatusStrip.SuspendLayout()
        CType(Me.DelayNumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GoogleResultsStatsFieldsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ScrapingDataGridView
        '
        Me.ScrapingDataGridView.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ScrapingDataGridView.AutoGenerateColumns = False
        Me.ScrapingDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.ScrapingDataGridView.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Keyword, Me.ExtactResultsNoDataGridViewTextBoxColumn, Me.InURLResutlsNoDataGridViewTextBoxColumn, Me.InTitleResultsNoDataGridViewTextBoxColumn, Me.InURLinTitleSumDataGridViewTextBoxColumn, Me.SearchPagesNoDataGridViewTextBoxColumn})
        Me.ScrapingDataGridView.DataSource = Me.GoogleResultsStatsFieldsBindingSource
        Me.ScrapingDataGridView.Location = New System.Drawing.Point(13, 12)
        Me.ScrapingDataGridView.Name = "ScrapingDataGridView"
        Me.ScrapingDataGridView.Size = New System.Drawing.Size(591, 247)
        Me.ScrapingDataGridView.TabIndex = 0
        '
        'Keyword
        '
        Me.Keyword.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Keyword.DataPropertyName = "Keyword"
        Me.Keyword.HeaderText = "Keyword"
        Me.Keyword.Name = "Keyword"
        '
        'ScrapeButton
        '
        Me.ScrapeButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ScrapeButton.Location = New System.Drawing.Point(13, 299)
        Me.ScrapeButton.Name = "ScrapeButton"
        Me.ScrapeButton.Size = New System.Drawing.Size(77, 28)
        Me.ScrapeButton.TabIndex = 1
        Me.ScrapeButton.Text = "&Scrape"
        Me.ScrapeButton.UseVisualStyleBackColor = True
        '
        'SaveButton
        '
        Me.SaveButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.SaveButton.Location = New System.Drawing.Point(185, 299)
        Me.SaveButton.Name = "SaveButton"
        Me.SaveButton.Size = New System.Drawing.Size(77, 28)
        Me.SaveButton.TabIndex = 2
        Me.SaveButton.Text = "&Save"
        Me.SaveButton.UseVisualStyleBackColor = True
        '
        'UseProxiesCheckBox
        '
        Me.UseProxiesCheckBox.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.UseProxiesCheckBox.AutoSize = True
        Me.UseProxiesCheckBox.Checked = True
        Me.UseProxiesCheckBox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.UseProxiesCheckBox.Location = New System.Drawing.Point(13, 276)
        Me.UseProxiesCheckBox.Name = "UseProxiesCheckBox"
        Me.UseProxiesCheckBox.Size = New System.Drawing.Size(82, 17)
        Me.UseProxiesCheckBox.TabIndex = 3
        Me.UseProxiesCheckBox.Text = "Use Proxies"
        Me.UseProxiesCheckBox.UseVisualStyleBackColor = True
        '
        'ScrapingStatusStrip
        '
        Me.ScrapingStatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MainFormToolStripStatusLabel, Me.ScrapedToolStripStatusLabel, Me.FailedToolStripStatusLabel})
        Me.ScrapingStatusStrip.Location = New System.Drawing.Point(0, 337)
        Me.ScrapingStatusStrip.Name = "ScrapingStatusStrip"
        Me.ScrapingStatusStrip.Size = New System.Drawing.Size(615, 22)
        Me.ScrapingStatusStrip.TabIndex = 4
        Me.ScrapingStatusStrip.Text = "StatusStrip1"
        '
        'MainFormToolStripStatusLabel
        '
        Me.MainFormToolStripStatusLabel.Name = "MainFormToolStripStatusLabel"
        Me.MainFormToolStripStatusLabel.Size = New System.Drawing.Size(0, 17)
        '
        'SavingTimer
        '
        Me.SavingTimer.Enabled = True
        Me.SavingTimer.Interval = 3600000
        '
        'ScrapingBackgroundWorker
        '
        Me.ScrapingBackgroundWorker.WorkerReportsProgress = True
        Me.ScrapingBackgroundWorker.WorkerSupportsCancellation = True
        '
        'ScrapeFromSourceButton
        '
        Me.ScrapeFromSourceButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ScrapeFromSourceButton.Location = New System.Drawing.Point(268, 299)
        Me.ScrapeFromSourceButton.Name = "ScrapeFromSourceButton"
        Me.ScrapeFromSourceButton.Size = New System.Drawing.Size(119, 28)
        Me.ScrapeFromSourceButton.TabIndex = 5
        Me.ScrapeFromSourceButton.Text = "&Scrape From Source"
        Me.ScrapeFromSourceButton.UseVisualStyleBackColor = True
        Me.ScrapeFromSourceButton.Visible = False
        '
        'StopButton
        '
        Me.StopButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.StopButton.Location = New System.Drawing.Point(99, 299)
        Me.StopButton.Name = "StopButton"
        Me.StopButton.Size = New System.Drawing.Size(77, 28)
        Me.StopButton.TabIndex = 6
        Me.StopButton.Text = "&Stop"
        Me.StopButton.UseVisualStyleBackColor = True
        '
        'DelayCheckBox
        '
        Me.DelayCheckBox.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.DelayCheckBox.AutoSize = True
        Me.DelayCheckBox.Checked = True
        Me.DelayCheckBox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.DelayCheckBox.Location = New System.Drawing.Point(431, 276)
        Me.DelayCheckBox.Name = "DelayCheckBox"
        Me.DelayCheckBox.Size = New System.Drawing.Size(53, 17)
        Me.DelayCheckBox.TabIndex = 7
        Me.DelayCheckBox.Text = "Delay"
        Me.DelayCheckBox.UseVisualStyleBackColor = True
        '
        'DelayNumericUpDown
        '
        Me.DelayNumericUpDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.DelayNumericUpDown.Location = New System.Drawing.Point(481, 273)
        Me.DelayNumericUpDown.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.DelayNumericUpDown.Name = "DelayNumericUpDown"
        Me.DelayNumericUpDown.Size = New System.Drawing.Size(56, 20)
        Me.DelayNumericUpDown.TabIndex = 8
        Me.DelayNumericUpDown.Value = New Decimal(New Integer() {120, 0, 0, 0})
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(544, 277)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(47, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Seconds"
        '
        'ReScrapeFailedCheckBox
        '
        Me.ReScrapeFailedCheckBox.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ReScrapeFailedCheckBox.AutoSize = True
        Me.ReScrapeFailedCheckBox.Checked = True
        Me.ReScrapeFailedCheckBox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ReScrapeFailedCheckBox.Location = New System.Drawing.Point(268, 276)
        Me.ReScrapeFailedCheckBox.Name = "ReScrapeFailedCheckBox"
        Me.ReScrapeFailedCheckBox.Size = New System.Drawing.Size(157, 17)
        Me.ReScrapeFailedCheckBox.TabIndex = 10
        Me.ReScrapeFailedCheckBox.Text = "Re-Scrape Failed Keywords"
        Me.ReScrapeFailedCheckBox.UseVisualStyleBackColor = True
        '
        'DeleteBackupsCheckBox
        '
        Me.DeleteBackupsCheckBox.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.DeleteBackupsCheckBox.AutoSize = True
        Me.DeleteBackupsCheckBox.Checked = True
        Me.DeleteBackupsCheckBox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.DeleteBackupsCheckBox.Location = New System.Drawing.Point(101, 276)
        Me.DeleteBackupsCheckBox.Name = "DeleteBackupsCheckBox"
        Me.DeleteBackupsCheckBox.Size = New System.Drawing.Size(161, 17)
        Me.DeleteBackupsCheckBox.TabIndex = 11
        Me.DeleteBackupsCheckBox.Text = "Delete Backups Before Start"
        Me.DeleteBackupsCheckBox.UseVisualStyleBackColor = True
        '
        'ExtactResultsNoDataGridViewTextBoxColumn
        '
        Me.ExtactResultsNoDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.ExtactResultsNoDataGridViewTextBoxColumn.DataPropertyName = "ExtactResultsNo"
        Me.ExtactResultsNoDataGridViewTextBoxColumn.HeaderText = "ExtactResultsNo"
        Me.ExtactResultsNoDataGridViewTextBoxColumn.Name = "ExtactResultsNoDataGridViewTextBoxColumn"
        '
        'InURLResutlsNoDataGridViewTextBoxColumn
        '
        Me.InURLResutlsNoDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.InURLResutlsNoDataGridViewTextBoxColumn.DataPropertyName = "inURLResutlsNo"
        Me.InURLResutlsNoDataGridViewTextBoxColumn.HeaderText = "inURLResutlsNo"
        Me.InURLResutlsNoDataGridViewTextBoxColumn.Name = "InURLResutlsNoDataGridViewTextBoxColumn"
        '
        'InTitleResultsNoDataGridViewTextBoxColumn
        '
        Me.InTitleResultsNoDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.InTitleResultsNoDataGridViewTextBoxColumn.DataPropertyName = "inTitleResultsNo"
        Me.InTitleResultsNoDataGridViewTextBoxColumn.HeaderText = "inTitleResultsNo"
        Me.InTitleResultsNoDataGridViewTextBoxColumn.Name = "InTitleResultsNoDataGridViewTextBoxColumn"
        '
        'InURLinTitleSumDataGridViewTextBoxColumn
        '
        Me.InURLinTitleSumDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.InURLinTitleSumDataGridViewTextBoxColumn.DataPropertyName = "inURLinTitleSum"
        Me.InURLinTitleSumDataGridViewTextBoxColumn.HeaderText = "inURLinTitleSum"
        Me.InURLinTitleSumDataGridViewTextBoxColumn.Name = "InURLinTitleSumDataGridViewTextBoxColumn"
        '
        'SearchPagesNoDataGridViewTextBoxColumn
        '
        Me.SearchPagesNoDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.SearchPagesNoDataGridViewTextBoxColumn.DataPropertyName = "SearchPagesNo"
        Me.SearchPagesNoDataGridViewTextBoxColumn.HeaderText = "SearchPagesNo"
        Me.SearchPagesNoDataGridViewTextBoxColumn.Name = "SearchPagesNoDataGridViewTextBoxColumn"
        '
        'GoogleResultsStatsFieldsBindingSource
        '
        Me.GoogleResultsStatsFieldsBindingSource.DataSource = GetType(ScrapeLib.GoogleResultsStatsFields)
        '
        'ScrapedToolStripStatusLabel
        '
        Me.ScrapedToolStripStatusLabel.Name = "ScrapedToolStripStatusLabel"
        Me.ScrapedToolStripStatusLabel.Size = New System.Drawing.Size(0, 17)
        '
        'FailedToolStripStatusLabel
        '
        Me.FailedToolStripStatusLabel.Name = "FailedToolStripStatusLabel"
        Me.FailedToolStripStatusLabel.Size = New System.Drawing.Size(0, 17)
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(615, 359)
        Me.Controls.Add(Me.DeleteBackupsCheckBox)
        Me.Controls.Add(Me.ReScrapeFailedCheckBox)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DelayNumericUpDown)
        Me.Controls.Add(Me.DelayCheckBox)
        Me.Controls.Add(Me.StopButton)
        Me.Controls.Add(Me.ScrapeFromSourceButton)
        Me.Controls.Add(Me.ScrapingStatusStrip)
        Me.Controls.Add(Me.UseProxiesCheckBox)
        Me.Controls.Add(Me.SaveButton)
        Me.Controls.Add(Me.ScrapeButton)
        Me.Controls.Add(Me.ScrapingDataGridView)
        Me.Name = "MainForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Google Results Stats Scraper"
        CType(Me.ScrapingDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ScrapingStatusStrip.ResumeLayout(False)
        Me.ScrapingStatusStrip.PerformLayout()
        CType(Me.DelayNumericUpDown, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GoogleResultsStatsFieldsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ScrapingDataGridView As System.Windows.Forms.DataGridView
    Friend WithEvents ScrapeButton As System.Windows.Forms.Button
    Friend WithEvents SaveButton As System.Windows.Forms.Button
    Friend WithEvents UseProxiesCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents ScrapingStatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents MainFormToolStripStatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents SavingTimer As System.Windows.Forms.Timer
    Friend WithEvents ScrapingBackgroundWorker As System.ComponentModel.BackgroundWorker
    Friend WithEvents ScrapeFromSourceButton As System.Windows.Forms.Button
    Friend WithEvents StopButton As System.Windows.Forms.Button
    Friend WithEvents GoogleResultsStatsFieldsBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Keyword As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ExtactResultsNoDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents InURLResutlsNoDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents InTitleResultsNoDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents InURLinTitleSumDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SearchPagesNoDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DelayCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents DelayNumericUpDown As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ReScrapeFailedCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents DeleteBackupsCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents ScrapedToolStripStatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents FailedToolStripStatusLabel As System.Windows.Forms.ToolStripStatusLabel

End Class
