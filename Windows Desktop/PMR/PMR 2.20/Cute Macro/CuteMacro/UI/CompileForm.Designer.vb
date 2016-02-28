<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CompileForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CompileForm))
        Me.PlyOptGroup = New System.Windows.Forms.GroupBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblSlow = New System.Windows.Forms.Label()
        Me.lblVerySlow = New System.Windows.Forms.Label()
        Me.txtRepeatTimes = New System.Windows.Forms.TextBox()
        Me.lblRepeatTimes = New System.Windows.Forms.Label()
        Me.lblPlayBackSpeed = New System.Windows.Forms.Label()
        Me.PlaSpdSlider = New System.Windows.Forms.TrackBar()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.chkEnableEditing = New System.Windows.Forms.CheckBox()
        Me.chkCompileSelected = New System.Windows.Forms.CheckBox()
        Me.chkNotify = New System.Windows.Forms.CheckBox()
        Me.BtnOK = New System.Windows.Forms.Button()
        Me.BtnCancel = New System.Windows.Forms.Button()
        Me.PlyOptGroup.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PlaSpdSlider, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PlyOptGroup
        '
        Me.PlyOptGroup.Controls.Add(Me.PictureBox2)
        Me.PlyOptGroup.Controls.Add(Me.Label1)
        Me.PlyOptGroup.Controls.Add(Me.PictureBox1)
        Me.PlyOptGroup.Controls.Add(Me.Label6)
        Me.PlyOptGroup.Controls.Add(Me.Label4)
        Me.PlyOptGroup.Controls.Add(Me.Label3)
        Me.PlyOptGroup.Controls.Add(Me.lblSlow)
        Me.PlyOptGroup.Controls.Add(Me.lblVerySlow)
        Me.PlyOptGroup.Controls.Add(Me.txtRepeatTimes)
        Me.PlyOptGroup.Controls.Add(Me.lblRepeatTimes)
        Me.PlyOptGroup.Controls.Add(Me.lblPlayBackSpeed)
        Me.PlyOptGroup.Controls.Add(Me.PlaSpdSlider)
        Me.PlyOptGroup.Location = New System.Drawing.Point(12, 12)
        Me.PlyOptGroup.Name = "PlyOptGroup"
        Me.PlyOptGroup.Size = New System.Drawing.Size(369, 184)
        Me.PlyOptGroup.TabIndex = 3
        Me.PlyOptGroup.TabStop = False
        Me.PlyOptGroup.Text = "Playbacking Options"
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(10, 136)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(33, 33)
        Me.PictureBox2.TabIndex = 15
        Me.PictureBox2.TabStop = False
        '
        'Label1
        '
        Me.Label1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label1.Location = New System.Drawing.Point(49, 135)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(300, 40)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "The executable macro file needs the .NET Framework 3.5 or .NET Framework 4.0 to b" & _
            "e installed in the machine in which it is executing."
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(10, 96)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(33, 33)
        Me.PictureBox1.TabIndex = 13
        Me.PictureBox1.TabStop = False
        '
        'Label6
        '
        Me.Label6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label6.Location = New System.Drawing.Point(49, 95)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(314, 40)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "To terminate the playback of the macro, you can hit (Ctrl+Alt+Del) or (Ctrl+Shift" & _
            "+F2) ."
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(321, 77)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(28, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Fast"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(237, 77)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Normal"
        '
        'lblSlow
        '
        Me.lblSlow.AutoSize = True
        Me.lblSlow.Location = New System.Drawing.Point(166, 77)
        Me.lblSlow.Name = "lblSlow"
        Me.lblSlow.Size = New System.Drawing.Size(29, 13)
        Me.lblSlow.TabIndex = 5
        Me.lblSlow.Text = "Slow"
        '
        'lblVerySlow
        '
        Me.lblVerySlow.AutoSize = True
        Me.lblVerySlow.Location = New System.Drawing.Point(78, 77)
        Me.lblVerySlow.Name = "lblVerySlow"
        Me.lblVerySlow.Size = New System.Drawing.Size(54, 13)
        Me.lblVerySlow.TabIndex = 4
        Me.lblVerySlow.Text = "Very Slow"
        '
        'txtRepeatTimes
        '
        Me.txtRepeatTimes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRepeatTimes.Location = New System.Drawing.Point(96, 19)
        Me.txtRepeatTimes.Name = "txtRepeatTimes"
        Me.txtRepeatTimes.Size = New System.Drawing.Size(96, 20)
        Me.txtRepeatTimes.TabIndex = 3
        Me.txtRepeatTimes.Text = "1"
        '
        'lblRepeatTimes
        '
        Me.lblRepeatTimes.AutoSize = True
        Me.lblRepeatTimes.Location = New System.Drawing.Point(6, 26)
        Me.lblRepeatTimes.Name = "lblRepeatTimes"
        Me.lblRepeatTimes.Size = New System.Drawing.Size(77, 13)
        Me.lblRepeatTimes.TabIndex = 2
        Me.lblRepeatTimes.Text = "Repeat times :"
        '
        'lblPlayBackSpeed
        '
        Me.lblPlayBackSpeed.AutoSize = True
        Me.lblPlayBackSpeed.Location = New System.Drawing.Point(6, 50)
        Me.lblPlayBackSpeed.Name = "lblPlayBackSpeed"
        Me.lblPlayBackSpeed.Size = New System.Drawing.Size(88, 13)
        Me.lblPlayBackSpeed.TabIndex = 1
        Me.lblPlayBackSpeed.Text = "Playback speed :"
        '
        'PlaSpdSlider
        '
        Me.PlaSpdSlider.Location = New System.Drawing.Point(90, 45)
        Me.PlaSpdSlider.Maximum = 3
        Me.PlaSpdSlider.Name = "PlaSpdSlider"
        Me.PlaSpdSlider.Size = New System.Drawing.Size(259, 45)
        Me.PlaSpdSlider.TabIndex = 0
        Me.PlaSpdSlider.Value = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(559, 89)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(53, 13)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Very Fast"
        '
        'chkEnableEditing
        '
        Me.chkEnableEditing.AutoSize = True
        Me.chkEnableEditing.Checked = True
        Me.chkEnableEditing.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkEnableEditing.Location = New System.Drawing.Point(461, 167)
        Me.chkEnableEditing.Name = "chkEnableEditing"
        Me.chkEnableEditing.Size = New System.Drawing.Size(343, 17)
        Me.chkEnableEditing.TabIndex = 12
        Me.chkEnableEditing.Text = "Enable editing of the exectuable file using Perfect Macro Recorder"
        Me.chkEnableEditing.UseVisualStyleBackColor = True
        Me.chkEnableEditing.Visible = False
        '
        'chkCompileSelected
        '
        Me.chkCompileSelected.AutoSize = True
        Me.chkCompileSelected.Enabled = False
        Me.chkCompileSelected.Location = New System.Drawing.Point(576, 269)
        Me.chkCompileSelected.Name = "chkCompileSelected"
        Me.chkCompileSelected.Size = New System.Drawing.Size(228, 17)
        Me.chkCompileSelected.TabIndex = 11
        Me.chkCompileSelected.Text = "Compile selected commands only in the list"
        Me.chkCompileSelected.UseVisualStyleBackColor = True
        Me.chkCompileSelected.Visible = False
        '
        'chkNotify
        '
        Me.chkNotify.AutoSize = True
        Me.chkNotify.Location = New System.Drawing.Point(514, 329)
        Me.chkNotify.Name = "chkNotify"
        Me.chkNotify.Size = New System.Drawing.Size(165, 17)
        Me.chkNotify.TabIndex = 10
        Me.chkNotify.Text = "Notify me of playback ending"
        Me.chkNotify.UseVisualStyleBackColor = True
        Me.chkNotify.Visible = False
        '
        'BtnOK
        '
        Me.BtnOK.Location = New System.Drawing.Point(218, 202)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(80, 26)
        Me.BtnOK.TabIndex = 7
        Me.BtnOK.Text = "&Compile"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'BtnCancel
        '
        Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancel.Location = New System.Drawing.Point(301, 202)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(80, 26)
        Me.BtnCancel.TabIndex = 6
        Me.BtnCancel.Text = "C&ancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'CompileForm
        '
        Me.AcceptButton = Me.BtnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnCancel
        Me.ClientSize = New System.Drawing.Size(388, 237)
        Me.ControlBox = False
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.PlyOptGroup)
        Me.Controls.Add(Me.chkEnableEditing)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.chkNotify)
        Me.Controls.Add(Me.chkCompileSelected)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "CompileForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Convert a macro to exe file"
        Me.PlyOptGroup.ResumeLayout(False)
        Me.PlyOptGroup.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PlaSpdSlider, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PlyOptGroup As System.Windows.Forms.GroupBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents chkEnableEditing As System.Windows.Forms.CheckBox
    Friend WithEvents chkCompileSelected As System.Windows.Forms.CheckBox
    Friend WithEvents chkNotify As System.Windows.Forms.CheckBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblSlow As System.Windows.Forms.Label
    Friend WithEvents lblVerySlow As System.Windows.Forms.Label
    Friend WithEvents txtRepeatTimes As System.Windows.Forms.TextBox
    Friend WithEvents lblRepeatTimes As System.Windows.Forms.Label
    Friend WithEvents lblPlayBackSpeed As System.Windows.Forms.Label
    Friend WithEvents PlaSpdSlider As System.Windows.Forms.TrackBar
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
