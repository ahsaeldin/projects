<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmBackGround
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmBackGround))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.TxtTitle = New System.Windows.Forms.TextBox()
        Me.LabelWinTitle = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LableDesc = New System.Windows.Forms.Label()
        Me.LabelTop = New System.Windows.Forms.Label()
        Me.LabelLeft = New System.Windows.Forms.Label()
        Me.LabelWidth = New System.Windows.Forms.Label()
        Me.LabelHeight = New System.Windows.Forms.Label()
        Me.mstTop = New System.Windows.Forms.MaskedTextBox()
        Me.mstHeight = New System.Windows.Forms.MaskedTextBox()
        Me.mstLeft = New System.Windows.Forms.MaskedTextBox()
        Me.mstWidth = New System.Windows.Forms.MaskedTextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.tmrGetTitle = New System.Windows.Forms.Timer(Me.components)
        Me.TxtWait = New System.Windows.Forms.TextBox()
        Me.ChkWaitForWindow = New System.Windows.Forms.CheckBox()
        Me.LabelProcessName = New System.Windows.Forms.Label()
        Me.LabelProcessPath = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(222, 346)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 6
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
        Me.Cancel_Button.TabIndex = 2
        Me.Cancel_Button.Text = "Cancel"
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Enabled = False
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 1
        Me.OK_Button.Text = "OK"
        '
        'TxtTitle
        '
        Me.TxtTitle.Location = New System.Drawing.Point(15, 93)
        Me.TxtTitle.Multiline = True
        Me.TxtTitle.Name = "TxtTitle"
        Me.TxtTitle.Size = New System.Drawing.Size(349, 63)
        Me.TxtTitle.TabIndex = 0
        '
        'LabelWinTitle
        '
        Me.LabelWinTitle.AutoSize = True
        Me.LabelWinTitle.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelWinTitle.Location = New System.Drawing.Point(12, 67)
        Me.LabelWinTitle.Name = "LabelWinTitle"
        Me.LabelWinTitle.Size = New System.Drawing.Size(152, 14)
        Me.LabelWinTitle.TabIndex = 51
        Me.LabelWinTitle.Text = "Background Window Title:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(388, 23)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(327, 91)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = resources.GetString("Label2.Text")
        '
        'LableDesc
        '
        Me.LableDesc.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LableDesc.Location = New System.Drawing.Point(12, 9)
        Me.LableDesc.Name = "LableDesc"
        Me.LableDesc.Size = New System.Drawing.Size(352, 58)
        Me.LableDesc.TabIndex = 50
        Me.LableDesc.Text = "Use this form to specify the properties of the background window in which the sub" & _
            "sequent actions in the macro list will executed within"
        '
        'LabelTop
        '
        Me.LabelTop.AutoSize = True
        Me.LabelTop.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelTop.Location = New System.Drawing.Point(20, 16)
        Me.LabelTop.Name = "LabelTop"
        Me.LabelTop.Size = New System.Drawing.Size(34, 16)
        Me.LabelTop.TabIndex = 52
        Me.LabelTop.Text = "Top "
        '
        'LabelLeft
        '
        Me.LabelLeft.AutoSize = True
        Me.LabelLeft.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelLeft.Location = New System.Drawing.Point(20, 45)
        Me.LabelLeft.Name = "LabelLeft"
        Me.LabelLeft.Size = New System.Drawing.Size(29, 16)
        Me.LabelLeft.TabIndex = 54
        Me.LabelLeft.Text = "Left"
        '
        'LabelWidth
        '
        Me.LabelWidth.AutoSize = True
        Me.LabelWidth.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelWidth.Location = New System.Drawing.Point(162, 16)
        Me.LabelWidth.Name = "LabelWidth"
        Me.LabelWidth.Size = New System.Drawing.Size(91, 16)
        Me.LabelWidth.TabIndex = 53
        Me.LabelWidth.Text = "Window Width"
        '
        'LabelHeight
        '
        Me.LabelHeight.AutoSize = True
        Me.LabelHeight.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelHeight.Location = New System.Drawing.Point(162, 41)
        Me.LabelHeight.Name = "LabelHeight"
        Me.LabelHeight.Size = New System.Drawing.Size(94, 16)
        Me.LabelHeight.TabIndex = 55
        Me.LabelHeight.Text = "Window Height"
        '
        'mstTop
        '
        Me.mstTop.BeepOnError = True
        Me.mstTop.Location = New System.Drawing.Point(57, 15)
        Me.mstTop.Margin = New System.Windows.Forms.Padding(0)
        Me.mstTop.Mask = "#00000"
        Me.mstTop.Name = "mstTop"
        Me.mstTop.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.mstTop.Size = New System.Drawing.Size(62, 20)
        Me.mstTop.TabIndex = 1
        Me.mstTop.Text = "0"
        '
        'mstHeight
        '
        Me.mstHeight.BeepOnError = True
        Me.mstHeight.Location = New System.Drawing.Point(270, 37)
        Me.mstHeight.Margin = New System.Windows.Forms.Padding(0)
        Me.mstHeight.Mask = "00000"
        Me.mstHeight.Name = "mstHeight"
        Me.mstHeight.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.mstHeight.Size = New System.Drawing.Size(65, 20)
        Me.mstHeight.TabIndex = 4
        Me.mstHeight.Text = "100"
        '
        'mstLeft
        '
        Me.mstLeft.BeepOnError = True
        Me.mstLeft.Location = New System.Drawing.Point(57, 44)
        Me.mstLeft.Margin = New System.Windows.Forms.Padding(0)
        Me.mstLeft.Mask = "#00000"
        Me.mstLeft.Name = "mstLeft"
        Me.mstLeft.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.mstLeft.Size = New System.Drawing.Size(62, 20)
        Me.mstLeft.TabIndex = 2
        Me.mstLeft.Text = "0"
        '
        'mstWidth
        '
        Me.mstWidth.BeepOnError = True
        Me.mstWidth.Location = New System.Drawing.Point(270, 12)
        Me.mstWidth.Margin = New System.Windows.Forms.Padding(0)
        Me.mstWidth.Mask = "00000"
        Me.mstWidth.Name = "mstWidth"
        Me.mstWidth.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.mstWidth.Size = New System.Drawing.Size(65, 20)
        Me.mstWidth.TabIndex = 3
        Me.mstWidth.Text = "100"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.mstTop)
        Me.GroupBox1.Controls.Add(Me.mstWidth)
        Me.GroupBox1.Controls.Add(Me.LabelTop)
        Me.GroupBox1.Controls.Add(Me.mstLeft)
        Me.GroupBox1.Controls.Add(Me.LabelLeft)
        Me.GroupBox1.Controls.Add(Me.mstHeight)
        Me.GroupBox1.Controls.Add(Me.LabelWidth)
        Me.GroupBox1.Controls.Add(Me.LabelHeight)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 162)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(352, 76)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label1.Location = New System.Drawing.Point(51, 244)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(313, 49)
        Me.Label1.TabIndex = 56
        Me.Label1.Text = "To copy title, dimensions and position of any window to this form, click the wind" & _
            "ow title bar while holding the Alt Key."
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(12, 244)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(33, 33)
        Me.PictureBox1.TabIndex = 23
        Me.PictureBox1.TabStop = False
        '
        'tmrGetTitle
        '
        Me.tmrGetTitle.Enabled = True
        '
        'TxtWait
        '
        Me.TxtWait.Location = New System.Drawing.Point(263, 310)
        Me.TxtWait.Name = "TxtWait"
        Me.TxtWait.Size = New System.Drawing.Size(101, 20)
        Me.TxtWait.TabIndex = 4
        Me.TxtWait.Text = "0"
        '
        'ChkWaitForWindow
        '
        Me.ChkWaitForWindow.AutoSize = True
        Me.ChkWaitForWindow.Location = New System.Drawing.Point(12, 313)
        Me.ChkWaitForWindow.Name = "ChkWaitForWindow"
        Me.ChkWaitForWindow.Size = New System.Drawing.Size(246, 17)
        Me.ChkWaitForWindow.TabIndex = 3
        Me.ChkWaitForWindow.Text = "Wait for the window to appear (milliseconds) :"
        Me.ChkWaitForWindow.UseVisualStyleBackColor = True
        '
        'LabelProcessName
        '
        Me.LabelProcessName.AutoSize = True
        Me.LabelProcessName.Location = New System.Drawing.Point(174, 68)
        Me.LabelProcessName.Name = "LabelProcessName"
        Me.LabelProcessName.Size = New System.Drawing.Size(71, 13)
        Me.LabelProcessName.TabIndex = 57
        Me.LabelProcessName.Text = "ProcessName"
        Me.LabelProcessName.Visible = False
        '
        'LabelProcessPath
        '
        Me.LabelProcessPath.AutoSize = True
        Me.LabelProcessPath.Location = New System.Drawing.Point(251, 68)
        Me.LabelProcessPath.Name = "LabelProcessPath"
        Me.LabelProcessPath.Size = New System.Drawing.Size(66, 13)
        Me.LabelProcessPath.TabIndex = 58
        Me.LabelProcessPath.Text = "ProcessPath"
        Me.LabelProcessPath.Visible = False
        '
        'FrmBackGround
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(378, 387)
        Me.Controls.Add(Me.LabelProcessPath)
        Me.Controls.Add(Me.LabelProcessName)
        Me.Controls.Add(Me.TxtWait)
        Me.Controls.Add(Me.ChkWaitForWindow)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.LableDesc)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.LabelWinTitle)
        Me.Controls.Add(Me.TxtTitle)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmBackGround"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Background Window"
        Me.TopMost = True
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents TxtTitle As System.Windows.Forms.TextBox
    Friend WithEvents LabelWinTitle As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LableDesc As System.Windows.Forms.Label
    Friend WithEvents LabelTop As System.Windows.Forms.Label
    Friend WithEvents LabelLeft As System.Windows.Forms.Label
    Friend WithEvents LabelWidth As System.Windows.Forms.Label
    Friend WithEvents LabelHeight As System.Windows.Forms.Label
    Friend WithEvents mstTop As System.Windows.Forms.MaskedTextBox
    Friend WithEvents mstHeight As System.Windows.Forms.MaskedTextBox
    Friend WithEvents mstLeft As System.Windows.Forms.MaskedTextBox
    Friend WithEvents mstWidth As System.Windows.Forms.MaskedTextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents tmrGetTitle As System.Windows.Forms.Timer
    Friend WithEvents TxtWait As System.Windows.Forms.TextBox
    Friend WithEvents ChkWaitForWindow As System.Windows.Forms.CheckBox
    Friend WithEvents LabelProcessName As System.Windows.Forms.Label
    Friend WithEvents LabelProcessPath As System.Windows.Forms.Label

End Class
