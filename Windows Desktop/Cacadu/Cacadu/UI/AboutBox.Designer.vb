Namespace UI

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class AboutForm
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

        Friend WithEvents TableLayoutPanel As System.Windows.Forms.TableLayoutPanel
        Friend WithEvents LogoPictureBox As System.Windows.Forms.PictureBox
        Friend WithEvents LabelProductName As System.Windows.Forms.Label
        Friend WithEvents LabelVersion As System.Windows.Forms.Label
        Friend WithEvents LabelCopyright As System.Windows.Forms.Label

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.  
        'Do not modify it using the code editor.
        <System.Diagnostics.DebuggerStepThrough()> _
        Private Sub InitializeComponent()
            Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AboutForm))
            Me.TableLayoutPanel = New System.Windows.Forms.TableLayoutPanel()
            Me.LogoPictureBox = New System.Windows.Forms.PictureBox()
            Me.LabelProductName = New System.Windows.Forms.Label()
            Me.LabelVersion = New System.Windows.Forms.Label()
            Me.LabelCopyright = New System.Windows.Forms.Label()
            Me.OKButton = New System.Windows.Forms.Button()
            Me.RegisteredToLabel = New System.Windows.Forms.Label()
            Me.LabelCompanyName = New System.Windows.Forms.Label()
            Me.CompanyNameHyperLinkEdit = New DevExpress.XtraEditors.HyperLinkEdit()
            Me.TableLayoutPanel.SuspendLayout()
            CType(Me.LogoPictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.CompanyNameHyperLinkEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'TableLayoutPanel
            '
            Me.TableLayoutPanel.ColumnCount = 2
            Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 33.0!))
            Me.TableLayoutPanel.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 67.0!))
            Me.TableLayoutPanel.Controls.Add(Me.LogoPictureBox, 0, 0)
            Me.TableLayoutPanel.Controls.Add(Me.LabelProductName, 1, 0)
            Me.TableLayoutPanel.Controls.Add(Me.LabelVersion, 1, 1)
            Me.TableLayoutPanel.Controls.Add(Me.LabelCopyright, 1, 2)
            Me.TableLayoutPanel.Controls.Add(Me.OKButton, 1, 6)
            Me.TableLayoutPanel.Controls.Add(Me.RegisteredToLabel, 1, 5)
            Me.TableLayoutPanel.Controls.Add(Me.LabelCompanyName, 1, 3)
            Me.TableLayoutPanel.Controls.Add(Me.CompanyNameHyperLinkEdit, 1, 4)
            Me.TableLayoutPanel.Dock = System.Windows.Forms.DockStyle.Fill
            Me.TableLayoutPanel.Location = New System.Drawing.Point(9, 9)
            Me.TableLayoutPanel.Name = "TableLayoutPanel"
            Me.TableLayoutPanel.RowCount = 7
            Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.581868!))
            Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.581868!))
            Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.581868!))
            Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 7.774723!))
            Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 44.31593!))
            Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.581868!))
            Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 9.581868!))
            Me.TableLayoutPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
            Me.TableLayoutPanel.Size = New System.Drawing.Size(396, 258)
            Me.TableLayoutPanel.TabIndex = 0
            '
            'LogoPictureBox
            '
            Me.LogoPictureBox.Dock = System.Windows.Forms.DockStyle.Fill
            Me.LogoPictureBox.Image = CType(resources.GetObject("LogoPictureBox.Image"), System.Drawing.Image)
            Me.LogoPictureBox.Location = New System.Drawing.Point(3, 3)
            Me.LogoPictureBox.Name = "LogoPictureBox"
            Me.TableLayoutPanel.SetRowSpan(Me.LogoPictureBox, 7)
            Me.LogoPictureBox.Size = New System.Drawing.Size(124, 252)
            Me.LogoPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
            Me.LogoPictureBox.TabIndex = 0
            Me.LogoPictureBox.TabStop = False
            '
            'LabelProductName
            '
            Me.LabelProductName.Dock = System.Windows.Forms.DockStyle.Fill
            Me.LabelProductName.Location = New System.Drawing.Point(136, 0)
            Me.LabelProductName.Margin = New System.Windows.Forms.Padding(6, 0, 3, 0)
            Me.LabelProductName.MaximumSize = New System.Drawing.Size(0, 17)
            Me.LabelProductName.Name = "LabelProductName"
            Me.LabelProductName.Size = New System.Drawing.Size(257, 17)
            Me.LabelProductName.TabIndex = 0
            Me.LabelProductName.Text = "Product Name"
            Me.LabelProductName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'LabelVersion
            '
            Me.LabelVersion.Dock = System.Windows.Forms.DockStyle.Fill
            Me.LabelVersion.Location = New System.Drawing.Point(136, 24)
            Me.LabelVersion.Margin = New System.Windows.Forms.Padding(6, 0, 3, 0)
            Me.LabelVersion.MaximumSize = New System.Drawing.Size(0, 17)
            Me.LabelVersion.Name = "LabelVersion"
            Me.LabelVersion.Size = New System.Drawing.Size(257, 17)
            Me.LabelVersion.TabIndex = 0
            Me.LabelVersion.Text = "Version"
            Me.LabelVersion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'LabelCopyright
            '
            Me.LabelCopyright.Dock = System.Windows.Forms.DockStyle.Fill
            Me.LabelCopyright.Location = New System.Drawing.Point(136, 48)
            Me.LabelCopyright.Margin = New System.Windows.Forms.Padding(6, 0, 3, 0)
            Me.LabelCopyright.MaximumSize = New System.Drawing.Size(0, 17)
            Me.LabelCopyright.Name = "LabelCopyright"
            Me.LabelCopyright.Size = New System.Drawing.Size(257, 17)
            Me.LabelCopyright.TabIndex = 0
            Me.LabelCopyright.Text = "Copyright"
            Me.LabelCopyright.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'OKButton
            '
            Me.OKButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.OKButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.OKButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            Me.OKButton.Location = New System.Drawing.Point(318, 233)
            Me.OKButton.Name = "OKButton"
            Me.OKButton.Size = New System.Drawing.Size(75, 22)
            Me.OKButton.TabIndex = 0
            Me.OKButton.Text = "&OK"
            '
            'RegisteredToLabel
            '
            Me.RegisteredToLabel.AutoSize = True
            Me.RegisteredToLabel.Dock = System.Windows.Forms.DockStyle.Fill
            Me.RegisteredToLabel.Location = New System.Drawing.Point(133, 206)
            Me.RegisteredToLabel.Name = "RegisteredToLabel"
            Me.RegisteredToLabel.Size = New System.Drawing.Size(260, 24)
            Me.RegisteredToLabel.TabIndex = 1
            Me.RegisteredToLabel.Text = "Registered To:"
            '
            'LabelCompanyName
            '
            Me.LabelCompanyName.Dock = System.Windows.Forms.DockStyle.Fill
            Me.LabelCompanyName.Location = New System.Drawing.Point(136, 72)
            Me.LabelCompanyName.Margin = New System.Windows.Forms.Padding(6, 0, 3, 0)
            Me.LabelCompanyName.MaximumSize = New System.Drawing.Size(0, 17)
            Me.LabelCompanyName.Name = "LabelCompanyName"
            Me.LabelCompanyName.Size = New System.Drawing.Size(257, 17)
            Me.LabelCompanyName.TabIndex = 0
            Me.LabelCompanyName.Text = "Company Name"
            Me.LabelCompanyName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            '
            'CompanyNameHyperLinkEdit
            '
            Me.CompanyNameHyperLinkEdit.EditValue = "http://www.conderella.com"
            Me.CompanyNameHyperLinkEdit.Location = New System.Drawing.Point(133, 95)
            Me.CompanyNameHyperLinkEdit.Name = "CompanyNameHyperLinkEdit"
            Me.CompanyNameHyperLinkEdit.Properties.Appearance.BackColor = System.Drawing.Color.Transparent
            Me.CompanyNameHyperLinkEdit.Properties.Appearance.Options.UseBackColor = True
            Me.CompanyNameHyperLinkEdit.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
            Me.CompanyNameHyperLinkEdit.Size = New System.Drawing.Size(150, 18)
            Me.CompanyNameHyperLinkEdit.TabIndex = 2
            '
            'AboutBox
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.CancelButton = Me.OKButton
            Me.ClientSize = New System.Drawing.Size(414, 276)
            Me.Controls.Add(Me.TableLayoutPanel)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "AboutBox"
            Me.Padding = New System.Windows.Forms.Padding(9)
            Me.ShowInTaskbar = False
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.Text = "AboutBox1"
            Me.TableLayoutPanel.ResumeLayout(False)
            Me.TableLayoutPanel.PerformLayout()
            CType(Me.LogoPictureBox, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.CompanyNameHyperLinkEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)

        End Sub
        Friend WithEvents LabelCompanyName As System.Windows.Forms.Label
        Friend WithEvents RegisteredToLabel As System.Windows.Forms.Label
        Friend WithEvents OKButton As System.Windows.Forms.Button
        Friend WithEvents CompanyNameHyperLinkEdit As DevExpress.XtraEditors.HyperLinkEdit

    End Class

End Namespace