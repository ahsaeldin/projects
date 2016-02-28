<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmKeyStroke
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmKeyStroke))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.OK_Button = New System.Windows.Forms.Button
        Me.Cancel_Button = New System.Windows.Forms.Button
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.LabDesc = New System.Windows.Forms.Label
        Me.ComKeyAction = New System.Windows.Forms.ComboBox
        Me.ComKeys = New System.Windows.Forms.ComboBox
        Me.LabKeyAction = New System.Windows.Forms.Label
        Me.LabKeys = New System.Windows.Forms.Label
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(135, 127)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(12, 12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(38, 35)
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'LabDesc
        '
        Me.LabDesc.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabDesc.Location = New System.Drawing.Point(55, 12)
        Me.LabDesc.Name = "LabDesc"
        Me.LabDesc.Size = New System.Drawing.Size(239, 39)
        Me.LabDesc.TabIndex = 2
        Me.LabDesc.Text = "Select the keyboard action type and the corresponding keystroke."
        '
        'ComKeyAction
        '
        Me.ComKeyAction.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComKeyAction.FormattingEnabled = True
        Me.ComKeyAction.Items.AddRange(New Object() {"KeyDown", "KeyUp"})
        Me.ComKeyAction.Location = New System.Drawing.Point(122, 65)
        Me.ComKeyAction.Name = "ComKeyAction"
        Me.ComKeyAction.Size = New System.Drawing.Size(157, 21)
        Me.ComKeyAction.TabIndex = 3
        '
        'ComKeys
        '
        Me.ComKeys.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComKeys.FormattingEnabled = True
        Me.ComKeys.Location = New System.Drawing.Point(122, 92)
        Me.ComKeys.Name = "ComKeys"
        Me.ComKeys.Size = New System.Drawing.Size(157, 21)
        Me.ComKeys.TabIndex = 4
        '
        'LabKeyAction
        '
        Me.LabKeyAction.AutoSize = True
        Me.LabKeyAction.Location = New System.Drawing.Point(55, 65)
        Me.LabKeyAction.Name = "LabKeyAction"
        Me.LabKeyAction.Size = New System.Drawing.Size(37, 13)
        Me.LabKeyAction.TabIndex = 5
        Me.LabKeyAction.Text = "Action"
        '
        'LabKeys
        '
        Me.LabKeys.AutoSize = True
        Me.LabKeys.Location = New System.Drawing.Point(55, 100)
        Me.LabKeys.Name = "LabKeys"
        Me.LabKeys.Size = New System.Drawing.Size(55, 13)
        Me.LabKeys.TabIndex = 6
        Me.LabKeys.Text = "Keystroke"
        '
        'FrmKeyStroke
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(293, 168)
        Me.Controls.Add(Me.LabKeys)
        Me.Controls.Add(Me.LabKeyAction)
        Me.Controls.Add(Me.ComKeys)
        Me.Controls.Add(Me.ComKeyAction)
        Me.Controls.Add(Me.LabDesc)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmKeyStroke"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Keyboard Action"
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents LabDesc As System.Windows.Forms.Label
    Friend WithEvents ComKeyAction As System.Windows.Forms.ComboBox
    Friend WithEvents ComKeys As System.Windows.Forms.ComboBox
    Friend WithEvents LabKeyAction As System.Windows.Forms.Label
    Friend WithEvents LabKeys As System.Windows.Forms.Label

End Class
