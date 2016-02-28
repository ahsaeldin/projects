<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SettingsForm
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
        Me.AutoScrapeCheckBox = New System.Windows.Forms.CheckBox()
        Me.AutoImportNumericUpDown = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SaveButton = New System.Windows.Forms.Button()
        Me.DropBoxPublicFolderTextBox = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.FolderBrowserDia = New System.Windows.Forms.FolderBrowserDialog()
        Me.BrowseButton = New System.Windows.Forms.Button()
        Me.DropBoxUrlLabel = New System.Windows.Forms.Label()
        Me.DropBoxTextBoxUrl = New System.Windows.Forms.TextBox()
        Me.PercentageNumericUpDown = New System.Windows.Forms.NumericUpDown()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.ConstantChangeNumericUpDown = New System.Windows.Forms.NumericUpDown()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.HongKongWarehouseRadioButton = New System.Windows.Forms.RadioButton()
        Me.UnitedKingdomWarehouseRadioButton = New System.Windows.Forms.RadioButton()
        Me.UnitedStatesWarehouseRadioButton = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.ShippingCheckBox = New System.Windows.Forms.CheckBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.MinPriceNumericUpDown = New System.Windows.Forms.NumericUpDown()
        Me.MaxPriceNumericUpDown2 = New System.Windows.Forms.NumericUpDown()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.RangeCheckBox = New System.Windows.Forms.CheckBox()
        CType(Me.AutoImportNumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PercentageNumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ConstantChangeNumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.MinPriceNumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MaxPriceNumericUpDown2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'AutoScrapeCheckBox
        '
        Me.AutoScrapeCheckBox.AutoSize = True
        Me.AutoScrapeCheckBox.Location = New System.Drawing.Point(12, 362)
        Me.AutoScrapeCheckBox.Name = "AutoScrapeCheckBox"
        Me.AutoScrapeCheckBox.Size = New System.Drawing.Size(311, 17)
        Me.AutoScrapeCheckBox.TabIndex = 0
        Me.AutoScrapeCheckBox.Text = "While scraping, automatically send data To Magento every:"
        Me.AutoScrapeCheckBox.UseVisualStyleBackColor = True
        Me.AutoScrapeCheckBox.Visible = False
        '
        'AutoImportNumericUpDown
        '
        Me.AutoImportNumericUpDown.Location = New System.Drawing.Point(329, 359)
        Me.AutoImportNumericUpDown.Minimum = New Decimal(New Integer() {5, 0, 0, 0})
        Me.AutoImportNumericUpDown.Name = "AutoImportNumericUpDown"
        Me.AutoImportNumericUpDown.Size = New System.Drawing.Size(50, 20)
        Me.AutoImportNumericUpDown.TabIndex = 1
        Me.AutoImportNumericUpDown.Value = New Decimal(New Integer() {5, 0, 0, 0})
        Me.AutoImportNumericUpDown.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(385, 362)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "minutes."
        Me.Label1.Visible = False
        '
        'SaveButton
        '
        Me.SaveButton.Location = New System.Drawing.Point(349, 262)
        Me.SaveButton.Name = "SaveButton"
        Me.SaveButton.Size = New System.Drawing.Size(93, 28)
        Me.SaveButton.TabIndex = 3
        Me.SaveButton.Text = "&Save"
        Me.SaveButton.UseVisualStyleBackColor = True
        '
        'DropBoxPublicFolderTextBox
        '
        Me.DropBoxPublicFolderTextBox.Location = New System.Drawing.Point(138, 16)
        Me.DropBoxPublicFolderTextBox.Name = "DropBoxPublicFolderTextBox"
        Me.DropBoxPublicFolderTextBox.ReadOnly = True
        Me.DropBoxPublicFolderTextBox.Size = New System.Drawing.Size(276, 20)
        Me.DropBoxPublicFolderTextBox.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(18, 19)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(115, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "DropBox Public Folder:"
        '
        'BrowseButton
        '
        Me.BrowseButton.Location = New System.Drawing.Point(419, 15)
        Me.BrowseButton.Name = "BrowseButton"
        Me.BrowseButton.Size = New System.Drawing.Size(27, 20)
        Me.BrowseButton.TabIndex = 6
        Me.BrowseButton.Text = "..."
        Me.BrowseButton.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.BrowseButton.UseVisualStyleBackColor = True
        '
        'DropBoxUrlLabel
        '
        Me.DropBoxUrlLabel.AutoSize = True
        Me.DropBoxUrlLabel.Location = New System.Drawing.Point(18, 74)
        Me.DropBoxUrlLabel.Name = "DropBoxUrlLabel"
        Me.DropBoxUrlLabel.Size = New System.Drawing.Size(98, 13)
        Me.DropBoxUrlLabel.TabIndex = 8
        Me.DropBoxUrlLabel.Text = "DropBox Public Url:"
        '
        'DropBoxTextBoxUrl
        '
        Me.DropBoxTextBoxUrl.Location = New System.Drawing.Point(138, 71)
        Me.DropBoxTextBoxUrl.Name = "DropBoxTextBoxUrl"
        Me.DropBoxTextBoxUrl.Size = New System.Drawing.Size(359, 20)
        Me.DropBoxTextBoxUrl.TabIndex = 7
        '
        'PercentageNumericUpDown
        '
        Me.PercentageNumericUpDown.Location = New System.Drawing.Point(147, 19)
        Me.PercentageNumericUpDown.Name = "PercentageNumericUpDown"
        Me.PercentageNumericUpDown.Size = New System.Drawing.Size(50, 20)
        Me.PercentageNumericUpDown.TabIndex = 9
        Me.PercentageNumericUpDown.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(15, 21)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(111, 17)
        Me.RadioButton1.TabIndex = 12
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Price Increase %:"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(15, 45)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(106, 17)
        Me.RadioButton2.TabIndex = 13
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "Price Increase $:"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'ConstantChangeNumericUpDown
        '
        Me.ConstantChangeNumericUpDown.Location = New System.Drawing.Point(147, 43)
        Me.ConstantChangeNumericUpDown.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.ConstantChangeNumericUpDown.Minimum = New Decimal(New Integer() {5, 0, 0, 0})
        Me.ConstantChangeNumericUpDown.Name = "ConstantChangeNumericUpDown"
        Me.ConstantChangeNumericUpDown.Size = New System.Drawing.Size(50, 20)
        Me.ConstantChangeNumericUpDown.TabIndex = 14
        Me.ConstantChangeNumericUpDown.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(135, 96)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(394, 13)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "For ex: https://dl.dropboxusercontent.com/u/*******/TmartScraperOutput.csv"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(135, 44)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(194, 13)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "For ex: D:\My Dropbox\Dropbox\Public"
        '
        'HongKongWarehouseRadioButton
        '
        Me.HongKongWarehouseRadioButton.AutoSize = True
        Me.HongKongWarehouseRadioButton.Location = New System.Drawing.Point(23, 19)
        Me.HongKongWarehouseRadioButton.Name = "HongKongWarehouseRadioButton"
        Me.HongKongWarehouseRadioButton.Size = New System.Drawing.Size(135, 17)
        Me.HongKongWarehouseRadioButton.TabIndex = 17
        Me.HongKongWarehouseRadioButton.TabStop = True
        Me.HongKongWarehouseRadioButton.Text = "Hong Kong Warehouse"
        Me.HongKongWarehouseRadioButton.UseVisualStyleBackColor = True
        '
        'UnitedKingdomWarehouseRadioButton
        '
        Me.UnitedKingdomWarehouseRadioButton.AutoSize = True
        Me.UnitedKingdomWarehouseRadioButton.Location = New System.Drawing.Point(23, 65)
        Me.UnitedKingdomWarehouseRadioButton.Name = "UnitedKingdomWarehouseRadioButton"
        Me.UnitedKingdomWarehouseRadioButton.Size = New System.Drawing.Size(157, 17)
        Me.UnitedKingdomWarehouseRadioButton.TabIndex = 18
        Me.UnitedKingdomWarehouseRadioButton.TabStop = True
        Me.UnitedKingdomWarehouseRadioButton.Text = "United Kingdom Warehouse"
        Me.UnitedKingdomWarehouseRadioButton.UseVisualStyleBackColor = True
        '
        'UnitedStatesWarehouseRadioButton
        '
        Me.UnitedStatesWarehouseRadioButton.AutoSize = True
        Me.UnitedStatesWarehouseRadioButton.Location = New System.Drawing.Point(23, 42)
        Me.UnitedStatesWarehouseRadioButton.Name = "UnitedStatesWarehouseRadioButton"
        Me.UnitedStatesWarehouseRadioButton.Size = New System.Drawing.Size(148, 17)
        Me.UnitedStatesWarehouseRadioButton.TabIndex = 19
        Me.UnitedStatesWarehouseRadioButton.TabStop = True
        Me.UnitedStatesWarehouseRadioButton.Text = "United States Warehouse"
        Me.UnitedStatesWarehouseRadioButton.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TextBox1)
        Me.GroupBox1.Controls.Add(Me.ShippingCheckBox)
        Me.GroupBox1.Controls.Add(Me.HongKongWarehouseRadioButton)
        Me.GroupBox1.Controls.Add(Me.UnitedStatesWarehouseRadioButton)
        Me.GroupBox1.Controls.Add(Me.UnitedKingdomWarehouseRadioButton)
        Me.GroupBox1.Location = New System.Drawing.Point(221, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(320, 101)
        Me.GroupBox1.TabIndex = 20
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Price Type"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.DropBoxPublicFolderTextBox)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.BrowseButton)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.DropBoxTextBoxUrl)
        Me.GroupBox2.Controls.Add(Me.DropBoxUrlLabel)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 119)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(531, 137)
        Me.GroupBox2.TabIndex = 21
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Dropbox Settings"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.RangeCheckBox)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.MaxPriceNumericUpDown2)
        Me.GroupBox3.Controls.Add(Me.MinPriceNumericUpDown)
        Me.GroupBox3.Controls.Add(Me.RadioButton1)
        Me.GroupBox3.Controls.Add(Me.PercentageNumericUpDown)
        Me.GroupBox3.Controls.Add(Me.RadioButton2)
        Me.GroupBox3.Controls.Add(Me.ConstantChangeNumericUpDown)
        Me.GroupBox3.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(203, 100)
        Me.GroupBox3.TabIndex = 22
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Price Change"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(264, 20)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(50, 20)
        Me.TextBox1.TabIndex = 18
        Me.TextBox1.Visible = False
        '
        'ShippingCheckBox
        '
        Me.ShippingCheckBox.AutoSize = True
        Me.ShippingCheckBox.Location = New System.Drawing.Point(210, 46)
        Me.ShippingCheckBox.Name = "ShippingCheckBox"
        Me.ShippingCheckBox.Size = New System.Drawing.Size(104, 17)
        Me.ShippingCheckBox.TabIndex = 17
        Me.ShippingCheckBox.Text = "Shipping Cost $:"
        Me.ShippingCheckBox.UseVisualStyleBackColor = True
        Me.ShippingCheckBox.Visible = False
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(448, 262)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(93, 28)
        Me.Button1.TabIndex = 23
        Me.Button1.Text = "&Save && Close"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'MinPriceNumericUpDown
        '
        Me.MinPriceNumericUpDown.Location = New System.Drawing.Point(80, 69)
        Me.MinPriceNumericUpDown.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.MinPriceNumericUpDown.Name = "MinPriceNumericUpDown"
        Me.MinPriceNumericUpDown.Size = New System.Drawing.Size(50, 20)
        Me.MinPriceNumericUpDown.TabIndex = 15
        '
        'MaxPriceNumericUpDown2
        '
        Me.MaxPriceNumericUpDown2.Location = New System.Drawing.Point(147, 69)
        Me.MaxPriceNumericUpDown2.Maximum = New Decimal(New Integer() {1000000000, 0, 0, 0})
        Me.MaxPriceNumericUpDown2.Name = "MaxPriceNumericUpDown2"
        Me.MaxPriceNumericUpDown2.Size = New System.Drawing.Size(50, 20)
        Me.MaxPriceNumericUpDown2.TabIndex = 16
        Me.MaxPriceNumericUpDown2.Value = New Decimal(New Integer() {1000, 0, 0, 0})
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(134, 72)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(11, 13)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "-"
        '
        'RangeCheckBox
        '
        Me.RangeCheckBox.AutoSize = True
        Me.RangeCheckBox.Checked = True
        Me.RangeCheckBox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.RangeCheckBox.Location = New System.Drawing.Point(15, 71)
        Me.RangeCheckBox.Name = "RangeCheckBox"
        Me.RangeCheckBox.Size = New System.Drawing.Size(57, 17)
        Me.RangeCheckBox.TabIndex = 20
        Me.RangeCheckBox.Text = "Range"
        Me.RangeCheckBox.UseVisualStyleBackColor = True
        '
        'SettingsForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(549, 299)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.SaveButton)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.AutoImportNumericUpDown)
        Me.Controls.Add(Me.AutoScrapeCheckBox)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SettingsForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Settings"
        CType(Me.AutoImportNumericUpDown, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PercentageNumericUpDown, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ConstantChangeNumericUpDown, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.MinPriceNumericUpDown, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MaxPriceNumericUpDown2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents AutoScrapeCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents AutoImportNumericUpDown As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents SaveButton As System.Windows.Forms.Button
    Friend WithEvents DropBoxPublicFolderTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents FolderBrowserDia As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents BrowseButton As System.Windows.Forms.Button
    Friend WithEvents DropBoxUrlLabel As System.Windows.Forms.Label
    Friend WithEvents DropBoxTextBoxUrl As System.Windows.Forms.TextBox
    Friend WithEvents PercentageNumericUpDown As System.Windows.Forms.NumericUpDown
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents ConstantChangeNumericUpDown As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents HongKongWarehouseRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents UnitedKingdomWarehouseRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents UnitedStatesWarehouseRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ShippingCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents MaxPriceNumericUpDown2 As System.Windows.Forms.NumericUpDown
    Friend WithEvents MinPriceNumericUpDown As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents RangeCheckBox As System.Windows.Forms.CheckBox
End Class
