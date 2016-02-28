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
        Me.ScrapeButton = New System.Windows.Forms.Button()
        Me.SaveButton = New System.Windows.Forms.Button()
        Me.UseProxiesCheckBox = New System.Windows.Forms.CheckBox()
        Me.ScrapingStatusStrip = New System.Windows.Forms.StatusStrip()
        Me.MainFormToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ProcessedToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.FailedToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.SavingTimer = New System.Windows.Forms.Timer(Me.components)
        Me.ScrapingBackgroundWorker = New System.ComponentModel.BackgroundWorker()
        Me.ScrapeFromSourceButton = New System.Windows.Forms.Button()
        Me.StopButton = New System.Windows.Forms.Button()
        Me.DelayCheckBox = New System.Windows.Forms.CheckBox()
        Me.DelayNumericUpDown = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ReScrapeFailedCheckBox = New System.Windows.Forms.CheckBox()
        Me.DeleteBackupsCheckBox = New System.Windows.Forms.CheckBox()
        Me.DataTabControl = New System.Windows.Forms.TabControl()
        Me.CrawlTabPage = New System.Windows.Forms.TabPage()
        Me.CrawlingDataGridView = New System.Windows.Forms.DataGridView()
        Me.NameDataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ListingURLDataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.URLsFieldsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.ScrapingTabPage = New System.Windows.Forms.TabPage()
        Me.CrawlButton = New System.Windows.Forms.Button()
        Me.CrawlingBackgroundWorker = New System.ComponentModel.BackgroundWorker()
        Me.HouzzFieldsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.NameDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PhoneDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.WebsiteDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ContactDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TypeDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AddressDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FaxDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FacebookDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TwitterDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LinkedInDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EmailsDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AdvertURLDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.ScrapingDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ScrapingStatusStrip.SuspendLayout()
        CType(Me.DelayNumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.DataTabControl.SuspendLayout()
        Me.CrawlTabPage.SuspendLayout()
        CType(Me.CrawlingDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.URLsFieldsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ScrapingTabPage.SuspendLayout()
        CType(Me.HouzzFieldsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ScrapingDataGridView
        '
        Me.ScrapingDataGridView.AllowUserToAddRows = False
        Me.ScrapingDataGridView.AutoGenerateColumns = False
        Me.ScrapingDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.ScrapingDataGridView.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.NameDataGridViewTextBoxColumn, Me.PhoneDataGridViewTextBoxColumn, Me.WebsiteDataGridViewTextBoxColumn, Me.ContactDataGridViewTextBoxColumn, Me.TypeDataGridViewTextBoxColumn, Me.AddressDataGridViewTextBoxColumn, Me.FaxDataGridViewTextBoxColumn, Me.FacebookDataGridViewTextBoxColumn, Me.TwitterDataGridViewTextBoxColumn, Me.LinkedInDataGridViewTextBoxColumn, Me.EmailsDataGridViewTextBoxColumn, Me.AdvertURLDataGridViewTextBoxColumn})
        Me.ScrapingDataGridView.DataSource = Me.HouzzFieldsBindingSource
        Me.ScrapingDataGridView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ScrapingDataGridView.Location = New System.Drawing.Point(3, 3)
        Me.ScrapingDataGridView.Name = "ScrapingDataGridView"
        Me.ScrapingDataGridView.Size = New System.Drawing.Size(577, 223)
        Me.ScrapingDataGridView.TabIndex = 0
        '
        'ScrapeButton
        '
        Me.ScrapeButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ScrapeButton.Location = New System.Drawing.Point(96, 299)
        Me.ScrapeButton.Name = "ScrapeButton"
        Me.ScrapeButton.Size = New System.Drawing.Size(77, 28)
        Me.ScrapeButton.TabIndex = 1
        Me.ScrapeButton.Text = "&Scrape"
        Me.ScrapeButton.UseVisualStyleBackColor = True
        '
        'SaveButton
        '
        Me.SaveButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.SaveButton.Location = New System.Drawing.Point(268, 299)
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
        Me.ScrapingStatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MainFormToolStripStatusLabel, Me.ProcessedToolStripStatusLabel, Me.FailedToolStripStatusLabel})
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
        'ProcessedToolStripStatusLabel
        '
        Me.ProcessedToolStripStatusLabel.Name = "ProcessedToolStripStatusLabel"
        Me.ProcessedToolStripStatusLabel.Size = New System.Drawing.Size(0, 17)
        '
        'FailedToolStripStatusLabel
        '
        Me.FailedToolStripStatusLabel.Name = "FailedToolStripStatusLabel"
        Me.FailedToolStripStatusLabel.Size = New System.Drawing.Size(0, 17)
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
        Me.ScrapeFromSourceButton.Location = New System.Drawing.Point(351, 299)
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
        Me.StopButton.Location = New System.Drawing.Point(182, 299)
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
        Me.Label1.Size = New System.Drawing.Size(49, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Seconds"
        '
        'ReScrapeFailedCheckBox
        '
        Me.ReScrapeFailedCheckBox.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ReScrapeFailedCheckBox.AutoSize = True
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
        'DataTabControl
        '
        Me.DataTabControl.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataTabControl.Controls.Add(Me.CrawlTabPage)
        Me.DataTabControl.Controls.Add(Me.ScrapingTabPage)
        Me.DataTabControl.Location = New System.Drawing.Point(12, 12)
        Me.DataTabControl.Name = "DataTabControl"
        Me.DataTabControl.SelectedIndex = 0
        Me.DataTabControl.Size = New System.Drawing.Size(591, 255)
        Me.DataTabControl.TabIndex = 12
        '
        'CrawlTabPage
        '
        Me.CrawlTabPage.Controls.Add(Me.CrawlingDataGridView)
        Me.CrawlTabPage.Location = New System.Drawing.Point(4, 22)
        Me.CrawlTabPage.Name = "CrawlTabPage"
        Me.CrawlTabPage.Padding = New System.Windows.Forms.Padding(3)
        Me.CrawlTabPage.Size = New System.Drawing.Size(583, 229)
        Me.CrawlTabPage.TabIndex = 0
        Me.CrawlTabPage.Text = "Crawler"
        Me.CrawlTabPage.UseVisualStyleBackColor = True
        '
        'CrawlingDataGridView
        '
        Me.CrawlingDataGridView.AutoGenerateColumns = False
        Me.CrawlingDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.CrawlingDataGridView.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.NameDataGridViewTextBoxColumn1, Me.ListingURLDataGridViewTextBoxColumn1})
        Me.CrawlingDataGridView.DataSource = Me.URLsFieldsBindingSource
        Me.CrawlingDataGridView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CrawlingDataGridView.Location = New System.Drawing.Point(3, 3)
        Me.CrawlingDataGridView.Name = "CrawlingDataGridView"
        Me.CrawlingDataGridView.Size = New System.Drawing.Size(577, 223)
        Me.CrawlingDataGridView.TabIndex = 0
        '
        'NameDataGridViewTextBoxColumn1
        '
        Me.NameDataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.NameDataGridViewTextBoxColumn1.DataPropertyName = "Name"
        Me.NameDataGridViewTextBoxColumn1.HeaderText = "Name"
        Me.NameDataGridViewTextBoxColumn1.Name = "NameDataGridViewTextBoxColumn1"
        '
        'ListingURLDataGridViewTextBoxColumn1
        '
        Me.ListingURLDataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.ListingURLDataGridViewTextBoxColumn1.DataPropertyName = "ListingURL"
        Me.ListingURLDataGridViewTextBoxColumn1.HeaderText = "ListingURL"
        Me.ListingURLDataGridViewTextBoxColumn1.Name = "ListingURLDataGridViewTextBoxColumn1"
        '
        'URLsFieldsBindingSource
        '
        Me.URLsFieldsBindingSource.DataSource = GetType(ScrapeLib.URLsFields)
        '
        'ScrapingTabPage
        '
        Me.ScrapingTabPage.Controls.Add(Me.ScrapingDataGridView)
        Me.ScrapingTabPage.Location = New System.Drawing.Point(4, 22)
        Me.ScrapingTabPage.Name = "ScrapingTabPage"
        Me.ScrapingTabPage.Padding = New System.Windows.Forms.Padding(3)
        Me.ScrapingTabPage.Size = New System.Drawing.Size(583, 229)
        Me.ScrapingTabPage.TabIndex = 1
        Me.ScrapingTabPage.Text = "Scraper"
        Me.ScrapingTabPage.UseVisualStyleBackColor = True
        '
        'CrawlButton
        '
        Me.CrawlButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CrawlButton.Location = New System.Drawing.Point(13, 299)
        Me.CrawlButton.Name = "CrawlButton"
        Me.CrawlButton.Size = New System.Drawing.Size(77, 28)
        Me.CrawlButton.TabIndex = 13
        Me.CrawlButton.Text = "&Crawl"
        Me.CrawlButton.UseVisualStyleBackColor = True
        '
        'CrawlingBackgroundWorker
        '
        Me.CrawlingBackgroundWorker.WorkerReportsProgress = True
        Me.CrawlingBackgroundWorker.WorkerSupportsCancellation = True
        '
        'HouzzFieldsBindingSource
        '
        Me.HouzzFieldsBindingSource.DataSource = GetType(ScrapeLib.HouzzFields)
        '
        'NameDataGridViewTextBoxColumn
        '
        Me.NameDataGridViewTextBoxColumn.DataPropertyName = "Name"
        Me.NameDataGridViewTextBoxColumn.HeaderText = "Name"
        Me.NameDataGridViewTextBoxColumn.Name = "NameDataGridViewTextBoxColumn"
        '
        'PhoneDataGridViewTextBoxColumn
        '
        Me.PhoneDataGridViewTextBoxColumn.DataPropertyName = "Phone"
        Me.PhoneDataGridViewTextBoxColumn.HeaderText = "Phone"
        Me.PhoneDataGridViewTextBoxColumn.Name = "PhoneDataGridViewTextBoxColumn"
        '
        'WebsiteDataGridViewTextBoxColumn
        '
        Me.WebsiteDataGridViewTextBoxColumn.DataPropertyName = "Website"
        Me.WebsiteDataGridViewTextBoxColumn.HeaderText = "Website"
        Me.WebsiteDataGridViewTextBoxColumn.Name = "WebsiteDataGridViewTextBoxColumn"
        '
        'ContactDataGridViewTextBoxColumn
        '
        Me.ContactDataGridViewTextBoxColumn.DataPropertyName = "Contact"
        Me.ContactDataGridViewTextBoxColumn.HeaderText = "Contact"
        Me.ContactDataGridViewTextBoxColumn.Name = "ContactDataGridViewTextBoxColumn"
        '
        'TypeDataGridViewTextBoxColumn
        '
        Me.TypeDataGridViewTextBoxColumn.DataPropertyName = "Type"
        Me.TypeDataGridViewTextBoxColumn.HeaderText = "Type"
        Me.TypeDataGridViewTextBoxColumn.Name = "TypeDataGridViewTextBoxColumn"
        '
        'AddressDataGridViewTextBoxColumn
        '
        Me.AddressDataGridViewTextBoxColumn.DataPropertyName = "Address"
        Me.AddressDataGridViewTextBoxColumn.HeaderText = "Address"
        Me.AddressDataGridViewTextBoxColumn.Name = "AddressDataGridViewTextBoxColumn"
        '
        'FaxDataGridViewTextBoxColumn
        '
        Me.FaxDataGridViewTextBoxColumn.DataPropertyName = "Fax"
        Me.FaxDataGridViewTextBoxColumn.HeaderText = "Fax"
        Me.FaxDataGridViewTextBoxColumn.Name = "FaxDataGridViewTextBoxColumn"
        '
        'FacebookDataGridViewTextBoxColumn
        '
        Me.FacebookDataGridViewTextBoxColumn.DataPropertyName = "Facebook"
        Me.FacebookDataGridViewTextBoxColumn.HeaderText = "Facebook"
        Me.FacebookDataGridViewTextBoxColumn.Name = "FacebookDataGridViewTextBoxColumn"
        '
        'TwitterDataGridViewTextBoxColumn
        '
        Me.TwitterDataGridViewTextBoxColumn.DataPropertyName = "Twitter"
        Me.TwitterDataGridViewTextBoxColumn.HeaderText = "Twitter"
        Me.TwitterDataGridViewTextBoxColumn.Name = "TwitterDataGridViewTextBoxColumn"
        '
        'LinkedInDataGridViewTextBoxColumn
        '
        Me.LinkedInDataGridViewTextBoxColumn.DataPropertyName = "LinkedIn"
        Me.LinkedInDataGridViewTextBoxColumn.HeaderText = "LinkedIn"
        Me.LinkedInDataGridViewTextBoxColumn.Name = "LinkedInDataGridViewTextBoxColumn"
        '
        'EmailsDataGridViewTextBoxColumn
        '
        Me.EmailsDataGridViewTextBoxColumn.DataPropertyName = "Emails"
        Me.EmailsDataGridViewTextBoxColumn.HeaderText = "Emails"
        Me.EmailsDataGridViewTextBoxColumn.Name = "EmailsDataGridViewTextBoxColumn"
        '
        'AdvertURLDataGridViewTextBoxColumn
        '
        Me.AdvertURLDataGridViewTextBoxColumn.DataPropertyName = "AdvertURL"
        Me.AdvertURLDataGridViewTextBoxColumn.HeaderText = "AdvertURL"
        Me.AdvertURLDataGridViewTextBoxColumn.Name = "AdvertURLDataGridViewTextBoxColumn"
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(615, 359)
        Me.Controls.Add(Me.CrawlButton)
        Me.Controls.Add(Me.DataTabControl)
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
        Me.Name = "MainForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Youtube Channel Data Scraper"
        CType(Me.ScrapingDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ScrapingStatusStrip.ResumeLayout(False)
        Me.ScrapingStatusStrip.PerformLayout()
        CType(Me.DelayNumericUpDown, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DataTabControl.ResumeLayout(False)
        Me.CrawlTabPage.ResumeLayout(False)
        CType(Me.CrawlingDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.URLsFieldsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ScrapingTabPage.ResumeLayout(False)
        CType(Me.HouzzFieldsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents DelayCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents DelayNumericUpDown As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ReScrapeFailedCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents DeleteBackupsCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents ProcessedToolStripStatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents FailedToolStripStatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents DataTabControl As System.Windows.Forms.TabControl
    Friend WithEvents CrawlTabPage As System.Windows.Forms.TabPage
    Friend WithEvents ScrapingTabPage As System.Windows.Forms.TabPage
    Friend WithEvents CrawlButton As System.Windows.Forms.Button
    Friend WithEvents CrawlingDataGridView As System.Windows.Forms.DataGridView
    Friend WithEvents CrawlingBackgroundWorker As System.ComponentModel.BackgroundWorker
    Friend WithEvents NameDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ListingURLDataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents URLsFieldsBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents NameDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PhoneDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents WebsiteDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ContactDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TypeDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AddressDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FaxDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FacebookDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TwitterDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LinkedInDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents EmailsDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AdvertURLDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents HouzzFieldsBindingSource As System.Windows.Forms.BindingSource

End Class
