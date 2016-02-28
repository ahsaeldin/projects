Namespace UI

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class AddNewGroupForm
        Inherits DevExpress.XtraEditors.XtraForm

        'Form overrides dispose to clean up the component list.
        <System.Diagnostics.DebuggerNonUserCode()> _
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.  
        'Do not modify it using the code editor.
        <System.Diagnostics.DebuggerStepThrough()> _
        Private Sub InitializeComponent()
            Me.NewGroupNameTextEdit = New DevExpress.XtraEditors.TextEdit()
            Me.AddGroupSimpleButton = New DevExpress.XtraEditors.SimpleButton()
            CType(Me.NewGroupNameTextEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'NewGroupNameTextEdit
            '
            Me.NewGroupNameTextEdit.Location = New System.Drawing.Point(12, 12)
            Me.NewGroupNameTextEdit.Name = "NewGroupNameTextEdit"
            Me.NewGroupNameTextEdit.Size = New System.Drawing.Size(215, 20)
            Me.NewGroupNameTextEdit.TabIndex = 0
            '
            'AddGroupSimpleButton
            '
            Me.AddGroupSimpleButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.AddGroupSimpleButton.Enabled = False
            Me.AddGroupSimpleButton.Location = New System.Drawing.Point(150, 38)
            Me.AddGroupSimpleButton.Name = "AddGroupSimpleButton"
            Me.AddGroupSimpleButton.Size = New System.Drawing.Size(77, 25)
            Me.AddGroupSimpleButton.TabIndex = 1
            Me.AddGroupSimpleButton.Text = "&Create"
            '
            'AddNewGroupForm
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.CancelButton = Me.AddGroupSimpleButton
            Me.ClientSize = New System.Drawing.Size(232, 70)
            Me.Controls.Add(Me.AddGroupSimpleButton)
            Me.Controls.Add(Me.NewGroupNameTextEdit)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "AddNewGroupForm"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.Text = "Add New Group"
            CType(Me.NewGroupNameTextEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)

        End Sub
        Friend WithEvents NewGroupNameTextEdit As DevExpress.XtraEditors.TextEdit
        Friend WithEvents AddGroupSimpleButton As DevExpress.XtraEditors.SimpleButton
    End Class

End Namespace