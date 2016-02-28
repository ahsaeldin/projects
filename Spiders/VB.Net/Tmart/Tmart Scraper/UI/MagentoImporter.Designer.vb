<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MagentoImporter
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
        Me.MagentoImporterWebBrowser = New System.Windows.Forms.WebBrowser()
        Me.SuspendLayout()
        '
        'MagentoImporterWebBrowser
        '
        Me.MagentoImporterWebBrowser.Dock = System.Windows.Forms.DockStyle.Fill
        Me.MagentoImporterWebBrowser.Location = New System.Drawing.Point(0, 0)
        Me.MagentoImporterWebBrowser.MinimumSize = New System.Drawing.Size(20, 20)
        Me.MagentoImporterWebBrowser.Name = "MagentoImporterWebBrowser"
        Me.MagentoImporterWebBrowser.Size = New System.Drawing.Size(472, 453)
        Me.MagentoImporterWebBrowser.TabIndex = 0
        '
        'MagentoImporter
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(472, 453)
        Me.Controls.Add(Me.MagentoImporterWebBrowser)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "MagentoImporter"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Magento Importer"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents MagentoImporterWebBrowser As System.Windows.Forms.WebBrowser
End Class
