<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMouse
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMouse))
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.OK_Button = New System.Windows.Forms.Button
        Me.Cancel_Button = New System.Windows.Forms.Button
        Me.lblDesc = New System.Windows.Forms.Label
        Me.lblMouseActTyp = New System.Windows.Forms.Label
        Me.lblMousePositions = New System.Windows.Forms.Label
        Me.lblX = New System.Windows.Forms.Label
        Me.lblY = New System.Windows.Forms.Label
        Me.picBMouseIcon = New System.Windows.Forms.PictureBox
        Me.mstX = New System.Windows.Forms.MaskedTextBox
        Me.mstY = New System.Windows.Forms.MaskedTextBox
        Me.comMouseEvents = New System.Windows.Forms.ComboBox
        Me.tmrCoordinates = New System.Windows.Forms.Timer(Me.components)
        Me.lblHowToCptPos = New System.Windows.Forms.Label
        Me.lblMouseWhlVal = New System.Windows.Forms.Label
        Me.mstWheelValue = New System.Windows.Forms.MaskedTextBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.picBMouseIcon, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(236, 248)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 4
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 4
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
        'lblDesc
        '
        Me.lblDesc.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblDesc.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.lblDesc.Location = New System.Drawing.Point(82, 12)
        Me.lblDesc.Name = "lblDesc"
        Me.lblDesc.Size = New System.Drawing.Size(298, 76)
        Me.lblDesc.TabIndex = 2
        Me.lblDesc.Text = "وصف هنا"
        '
        'lblMouseActTyp
        '
        Me.lblMouseActTyp.AutoSize = True
        Me.lblMouseActTyp.Location = New System.Drawing.Point(79, 107)
        Me.lblMouseActTyp.Name = "lblMouseActTyp"
        Me.lblMouseActTyp.Size = New System.Drawing.Size(102, 13)
        Me.lblMouseActTyp.TabIndex = 3
        Me.lblMouseActTyp.Text = "Mouse Action Type:"
        '
        'lblMousePositions
        '
        Me.lblMousePositions.AutoSize = True
        Me.lblMousePositions.Location = New System.Drawing.Point(79, 142)
        Me.lblMousePositions.Name = "lblMousePositions"
        Me.lblMousePositions.Size = New System.Drawing.Size(83, 13)
        Me.lblMousePositions.TabIndex = 4
        Me.lblMousePositions.Text = "Mouse Positions"
        '
        'lblX
        '
        Me.lblX.AutoSize = True
        Me.lblX.Location = New System.Drawing.Point(166, 142)
        Me.lblX.Name = "lblX"
        Me.lblX.Size = New System.Drawing.Size(20, 13)
        Me.lblX.TabIndex = 5
        Me.lblX.Text = "X :"
        '
        'lblY
        '
        Me.lblY.AutoSize = True
        Me.lblY.Location = New System.Drawing.Point(271, 142)
        Me.lblY.Name = "lblY"
        Me.lblY.Size = New System.Drawing.Size(20, 13)
        Me.lblY.TabIndex = 6
        Me.lblY.Text = "Y :"
        '
        'picBMouseIcon
        '
        Me.picBMouseIcon.Image = CType(resources.GetObject("picBMouseIcon.Image"), System.Drawing.Image)
        Me.picBMouseIcon.Location = New System.Drawing.Point(12, 12)
        Me.picBMouseIcon.Name = "picBMouseIcon"
        Me.picBMouseIcon.Size = New System.Drawing.Size(64, 64)
        Me.picBMouseIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.picBMouseIcon.TabIndex = 10
        Me.picBMouseIcon.TabStop = False
        '
        'mstX
        '
        Me.mstX.BeepOnError = True
        Me.mstX.Location = New System.Drawing.Point(187, 135)
        Me.mstX.Mask = "0000"
        Me.mstX.Name = "mstX"
        Me.mstX.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.mstX.Size = New System.Drawing.Size(77, 20)
        Me.mstX.TabIndex = 1
        Me.mstX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'mstY
        '
        Me.mstY.BeepOnError = True
        Me.mstY.Location = New System.Drawing.Point(290, 135)
        Me.mstY.Mask = "0000"
        Me.mstY.Name = "mstY"
        Me.mstY.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.mstY.Size = New System.Drawing.Size(77, 20)
        Me.mstY.TabIndex = 2
        Me.mstY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'comMouseEvents
        '
        Me.comMouseEvents.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.comMouseEvents.FormattingEnabled = True
        Me.comMouseEvents.Items.AddRange(New Object() {"LeftMouseDown", "LeftMouseUp", "MiddleMouseDown", "MiddleMouseUp", "MouseMove", "MouseWheel", "RightMouseDown", "RightMouseUp"})
        Me.comMouseEvents.Location = New System.Drawing.Point(187, 99)
        Me.comMouseEvents.Name = "comMouseEvents"
        Me.comMouseEvents.Size = New System.Drawing.Size(138, 21)
        Me.comMouseEvents.TabIndex = 0
        '
        'tmrCoordinates
        '
        Me.tmrCoordinates.Enabled = True
        '
        'lblHowToCptPos
        '
        Me.lblHowToCptPos.Location = New System.Drawing.Point(78, 173)
        Me.lblHowToCptPos.Name = "lblHowToCptPos"
        Me.lblHowToCptPos.Size = New System.Drawing.Size(296, 42)
        Me.lblHowToCptPos.TabIndex = 14
        Me.lblHowToCptPos.Text = "(Click anywhere while holding Alt key to start tracking the mouse coordinates the" & _
            "n click again while holding Alt key to capture the coordinates.)"
        '
        'lblMouseWhlVal
        '
        Me.lblMouseWhlVal.AutoSize = True
        Me.lblMouseWhlVal.Location = New System.Drawing.Point(78, 229)
        Me.lblMouseWhlVal.Name = "lblMouseWhlVal"
        Me.lblMouseWhlVal.Size = New System.Drawing.Size(104, 13)
        Me.lblMouseWhlVal.TabIndex = 15
        Me.lblMouseWhlVal.Text = "Mouse Wheel Value:"
        Me.lblMouseWhlVal.Visible = False
        '
        'mstWheelValue
        '
        Me.mstWheelValue.BeepOnError = True
        Me.mstWheelValue.Location = New System.Drawing.Point(187, 222)
        Me.mstWheelValue.Margin = New System.Windows.Forms.Padding(0)
        Me.mstWheelValue.Mask = "00000"
        Me.mstWheelValue.Name = "mstWheelValue"
        Me.mstWheelValue.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.mstWheelValue.Size = New System.Drawing.Size(77, 20)
        Me.mstWheelValue.TabIndex = 3
        Me.mstWheelValue.Visible = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(39, 173)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(33, 33)
        Me.PictureBox1.TabIndex = 17
        Me.PictureBox1.TabStop = False
        '
        'frmMouse
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(394, 289)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.mstWheelValue)
        Me.Controls.Add(Me.lblMouseWhlVal)
        Me.Controls.Add(Me.lblHowToCptPos)
        Me.Controls.Add(Me.lblDesc)
        Me.Controls.Add(Me.comMouseEvents)
        Me.Controls.Add(Me.mstY)
        Me.Controls.Add(Me.mstX)
        Me.Controls.Add(Me.picBMouseIcon)
        Me.Controls.Add(Me.lblY)
        Me.Controls.Add(Me.lblX)
        Me.Controls.Add(Me.lblMousePositions)
        Me.Controls.Add(Me.lblMouseActTyp)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMouse"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Mouse Command"
        Me.TopMost = True
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.picBMouseIcon, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents lblDesc As System.Windows.Forms.Label
    Friend WithEvents lblMouseActTyp As System.Windows.Forms.Label
    Friend WithEvents lblMousePositions As System.Windows.Forms.Label
    Friend WithEvents lblX As System.Windows.Forms.Label
    Friend WithEvents lblY As System.Windows.Forms.Label
    Friend WithEvents picBMouseIcon As System.Windows.Forms.PictureBox
    Friend WithEvents mstX As System.Windows.Forms.MaskedTextBox
    Friend WithEvents mstY As System.Windows.Forms.MaskedTextBox
    Friend WithEvents comMouseEvents As System.Windows.Forms.ComboBox
    Friend WithEvents tmrCoordinates As System.Windows.Forms.Timer
    Friend WithEvents lblHowToCptPos As System.Windows.Forms.Label
    Friend WithEvents lblMouseWhlVal As System.Windows.Forms.Label
    Friend WithEvents mstWheelValue As System.Windows.Forms.MaskedTextBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox

End Class
