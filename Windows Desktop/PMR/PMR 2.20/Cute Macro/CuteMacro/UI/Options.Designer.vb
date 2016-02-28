<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CMOptionsForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CMOptionsForm))
        Me.GenOptGroup = New System.Windows.Forms.GroupBox()
        Me.chkTips = New System.Windows.Forms.CheckBox()
        Me.chkAutoRecord = New System.Windows.Forms.CheckBox()
        Me.chkStartup = New System.Windows.Forms.CheckBox()
        Me.RecOptGroup = New System.Windows.Forms.GroupBox()
        Me.chkKeyActs = New System.Windows.Forms.CheckBox()
        Me.chkAllMouActs = New System.Windows.Forms.CheckBox()
        Me.chkMouMov = New System.Windows.Forms.CheckBox()
        Me.PlyOptGroup = New System.Windows.Forms.GroupBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.LBLTip = New System.Windows.Forms.Label()
        Me.ChkSmartPlayback = New System.Windows.Forms.CheckBox()
        Me.chkNotify = New System.Windows.Forms.CheckBox()
        Me.chkCloseCM = New System.Windows.Forms.CheckBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblSlow = New System.Windows.Forms.Label()
        Me.lblVerySlow = New System.Windows.Forms.Label()
        Me.txtRepeatTimes = New System.Windows.Forms.TextBox()
        Me.lblRepeatTimes = New System.Windows.Forms.Label()
        Me.lblPlayBackSpeed = New System.Windows.Forms.Label()
        Me.PlaSpdSlider = New System.Windows.Forms.TrackBar()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.KeyShortGroup = New System.Windows.Forms.GroupBox()
        Me.BtnReset = New System.Windows.Forms.Button()
        Me.txtStopPly = New System.Windows.Forms.TextBox()
        Me.txtStrPly = New System.Windows.Forms.TextBox()
        Me.txtStopRec = New System.Windows.Forms.TextBox()
        Me.txtStrRec = New System.Windows.Forms.TextBox()
        Me.lblStopPlayback = New System.Windows.Forms.Label()
        Me.lblStartPlayback = New System.Windows.Forms.Label()
        Me.lblAbortRecording = New System.Windows.Forms.Label()
        Me.lblStrRec = New System.Windows.Forms.Label()
        Me.BtnCancel = New System.Windows.Forms.Button()
        Me.BtnOK = New System.Windows.Forms.Button()
        Me.lblDefFoldPath = New System.Windows.Forms.Label()
        Me.lblDefMacName = New System.Windows.Forms.Label()
        Me.DefFoldGroup = New System.Windows.Forms.GroupBox()
        Me.lblPath = New System.Windows.Forms.Label()
        Me.chkDefFold = New System.Windows.Forms.CheckBox()
        Me.btnDefFoldPath = New System.Windows.Forms.Button()
        Me.txtDefMacName = New System.Windows.Forms.TextBox()
        Me.FoldBrowDia = New System.Windows.Forms.FolderBrowserDialog()
        Me.OptionsTab = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.TabShortCuts = New System.Windows.Forms.TabPage()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GenOptGroup.SuspendLayout()
        Me.RecOptGroup.SuspendLayout()
        Me.PlyOptGroup.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PlaSpdSlider, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.KeyShortGroup.SuspendLayout()
        Me.DefFoldGroup.SuspendLayout()
        Me.OptionsTab.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabShortCuts.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GenOptGroup
        '
        Me.GenOptGroup.Controls.Add(Me.chkTips)
        Me.GenOptGroup.Controls.Add(Me.chkAutoRecord)
        Me.GenOptGroup.Controls.Add(Me.chkStartup)
        Me.GenOptGroup.Location = New System.Drawing.Point(3, 21)
        Me.GenOptGroup.Name = "GenOptGroup"
        Me.GenOptGroup.Size = New System.Drawing.Size(200, 93)
        Me.GenOptGroup.TabIndex = 0
        Me.GenOptGroup.TabStop = False
        Me.GenOptGroup.Text = "General Options"
        '
        'chkTips
        '
        Me.chkTips.AutoSize = True
        Me.chkTips.Location = New System.Drawing.Point(6, 63)
        Me.chkTips.Name = "chkTips"
        Me.chkTips.Size = New System.Drawing.Size(99, 17)
        Me.chkTips.TabIndex = 3
        Me.chkTips.Text = "Show help, tips"
        Me.chkTips.UseVisualStyleBackColor = True
        '
        'chkAutoRecord
        '
        Me.chkAutoRecord.AutoSize = True
        Me.chkAutoRecord.Location = New System.Drawing.Point(6, 40)
        Me.chkAutoRecord.Name = "chkAutoRecord"
        Me.chkAutoRecord.Size = New System.Drawing.Size(188, 17)
        Me.chkAutoRecord.TabIndex = 2
        Me.chkAutoRecord.Text = "Auto start recording upon startup"
        Me.chkAutoRecord.UseVisualStyleBackColor = True
        '
        'chkStartup
        '
        Me.chkStartup.AutoSize = True
        Me.chkStartup.Location = New System.Drawing.Point(6, 19)
        Me.chkStartup.Name = "chkStartup"
        Me.chkStartup.Size = New System.Drawing.Size(152, 17)
        Me.chkStartup.TabIndex = 0
        Me.chkStartup.Text = "Starts at Windows startup"
        Me.chkStartup.UseVisualStyleBackColor = True
        '
        'RecOptGroup
        '
        Me.RecOptGroup.Controls.Add(Me.chkKeyActs)
        Me.RecOptGroup.Controls.Add(Me.chkAllMouActs)
        Me.RecOptGroup.Controls.Add(Me.chkMouMov)
        Me.RecOptGroup.Location = New System.Drawing.Point(209, 21)
        Me.RecOptGroup.Name = "RecOptGroup"
        Me.RecOptGroup.Size = New System.Drawing.Size(163, 93)
        Me.RecOptGroup.TabIndex = 1
        Me.RecOptGroup.TabStop = False
        Me.RecOptGroup.Text = "Recording Options"
        '
        'chkKeyActs
        '
        Me.chkKeyActs.AutoSize = True
        Me.chkKeyActs.Location = New System.Drawing.Point(6, 63)
        Me.chkKeyActs.Name = "chkKeyActs"
        Me.chkKeyActs.Size = New System.Drawing.Size(153, 17)
        Me.chkKeyActs.TabIndex = 2
        Me.chkKeyActs.Text = "Record keyboard activities"
        Me.chkKeyActs.UseVisualStyleBackColor = True
        '
        'chkAllMouActs
        '
        Me.chkAllMouActs.AutoSize = True
        Me.chkAllMouActs.Location = New System.Drawing.Point(6, 40)
        Me.chkAllMouActs.Name = "chkAllMouActs"
        Me.chkAllMouActs.Size = New System.Drawing.Size(152, 17)
        Me.chkAllMouActs.TabIndex = 1
        Me.chkAllMouActs.Text = "Record all mouse activities"
        Me.chkAllMouActs.UseVisualStyleBackColor = True
        '
        'chkMouMov
        '
        Me.chkMouMov.AutoSize = True
        Me.chkMouMov.Location = New System.Drawing.Point(6, 19)
        Me.chkMouMov.Name = "chkMouMov"
        Me.chkMouMov.Size = New System.Drawing.Size(152, 17)
        Me.chkMouMov.TabIndex = 0
        Me.chkMouMov.Text = "Record mouse movements"
        Me.chkMouMov.UseVisualStyleBackColor = True
        '
        'PlyOptGroup
        '
        Me.PlyOptGroup.BackColor = System.Drawing.SystemColors.Control
        Me.PlyOptGroup.Controls.Add(Me.PictureBox2)
        Me.PlyOptGroup.Controls.Add(Me.LBLTip)
        Me.PlyOptGroup.Controls.Add(Me.ChkSmartPlayback)
        Me.PlyOptGroup.Controls.Add(Me.chkNotify)
        Me.PlyOptGroup.Controls.Add(Me.chkCloseCM)
        Me.PlyOptGroup.Controls.Add(Me.Label5)
        Me.PlyOptGroup.Controls.Add(Me.Label4)
        Me.PlyOptGroup.Controls.Add(Me.Label3)
        Me.PlyOptGroup.Controls.Add(Me.lblSlow)
        Me.PlyOptGroup.Controls.Add(Me.lblVerySlow)
        Me.PlyOptGroup.Controls.Add(Me.txtRepeatTimes)
        Me.PlyOptGroup.Controls.Add(Me.lblRepeatTimes)
        Me.PlyOptGroup.Controls.Add(Me.lblPlayBackSpeed)
        Me.PlyOptGroup.Controls.Add(Me.PlaSpdSlider)
        Me.PlyOptGroup.Location = New System.Drawing.Point(5, 6)
        Me.PlyOptGroup.Name = "PlyOptGroup"
        Me.PlyOptGroup.Size = New System.Drawing.Size(369, 231)
        Me.PlyOptGroup.TabIndex = 2
        Me.PlyOptGroup.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(9, 108)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(33, 33)
        Me.PictureBox2.TabIndex = 16
        Me.PictureBox2.TabStop = False
        '
        'LBLTip
        '
        Me.LBLTip.Location = New System.Drawing.Point(48, 108)
        Me.LBLTip.Name = "LBLTip"
        Me.LBLTip.Size = New System.Drawing.Size(315, 50)
        Me.LBLTip.TabIndex = 14
        Me.LBLTip.Text = "Choosing any playback speed but Normal, makes Perfect Macro Recorder ignores the " & _
            "delay items in the macro script and use and use instead of it a built in delay v" & _
            "alues for these speeds."
        '
        'ChkSmartPlayback
        '
        Me.ChkSmartPlayback.AutoSize = True
        Me.ChkSmartPlayback.Location = New System.Drawing.Point(6, 207)
        Me.ChkSmartPlayback.Name = "ChkSmartPlayback"
        Me.ChkSmartPlayback.Size = New System.Drawing.Size(133, 17)
        Me.ChkSmartPlayback.TabIndex = 11
        Me.ChkSmartPlayback.Text = "Enable smart playback"
        Me.ChkSmartPlayback.UseVisualStyleBackColor = True
        '
        'chkNotify
        '
        Me.chkNotify.AutoSize = True
        Me.chkNotify.Location = New System.Drawing.Point(6, 161)
        Me.chkNotify.Name = "chkNotify"
        Me.chkNotify.Size = New System.Drawing.Size(165, 17)
        Me.chkNotify.TabIndex = 9
        Me.chkNotify.Text = "Notify me of playback ending"
        Me.chkNotify.UseVisualStyleBackColor = True
        '
        'chkCloseCM
        '
        Me.chkCloseCM.AutoSize = True
        Me.chkCloseCM.Location = New System.Drawing.Point(6, 184)
        Me.chkCloseCM.Name = "chkCloseCM"
        Me.chkCloseCM.Size = New System.Drawing.Size(241, 17)
        Me.chkCloseCM.TabIndex = 9
        Me.chkCloseCM.Text = "Close Perfect Macro Recorder after playback"
        Me.chkCloseCM.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(310, 81)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(53, 13)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Very Fast"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(267, 81)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(28, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Fast"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(203, 81)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Normal"
        '
        'lblSlow
        '
        Me.lblSlow.AutoSize = True
        Me.lblSlow.Location = New System.Drawing.Point(148, 81)
        Me.lblSlow.Name = "lblSlow"
        Me.lblSlow.Size = New System.Drawing.Size(29, 13)
        Me.lblSlow.TabIndex = 5
        Me.lblSlow.Text = "Slow"
        '
        'lblVerySlow
        '
        Me.lblVerySlow.AutoSize = True
        Me.lblVerySlow.Location = New System.Drawing.Point(78, 81)
        Me.lblVerySlow.Name = "lblVerySlow"
        Me.lblVerySlow.Size = New System.Drawing.Size(54, 13)
        Me.lblVerySlow.TabIndex = 4
        Me.lblVerySlow.Text = "Very Slow"
        '
        'txtRepeatTimes
        '
        Me.txtRepeatTimes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtRepeatTimes.Location = New System.Drawing.Point(89, 19)
        Me.txtRepeatTimes.Name = "txtRepeatTimes"
        Me.txtRepeatTimes.Size = New System.Drawing.Size(96, 20)
        Me.txtRepeatTimes.TabIndex = 0
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
        Me.lblPlayBackSpeed.Location = New System.Drawing.Point(6, 54)
        Me.lblPlayBackSpeed.Name = "lblPlayBackSpeed"
        Me.lblPlayBackSpeed.Size = New System.Drawing.Size(88, 13)
        Me.lblPlayBackSpeed.TabIndex = 1
        Me.lblPlayBackSpeed.Text = "Playback speed :"
        '
        'PlaSpdSlider
        '
        Me.PlaSpdSlider.BackColor = System.Drawing.SystemColors.Control
        Me.PlaSpdSlider.Location = New System.Drawing.Point(90, 49)
        Me.PlaSpdSlider.Maximum = 4
        Me.PlaSpdSlider.Name = "PlaSpdSlider"
        Me.PlaSpdSlider.Size = New System.Drawing.Size(259, 45)
        Me.PlaSpdSlider.TabIndex = 0
        Me.PlaSpdSlider.Value = 2
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(424, 34)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(165, 121)
        Me.PictureBox1.TabIndex = 12
        Me.PictureBox1.TabStop = False
        '
        'KeyShortGroup
        '
        Me.KeyShortGroup.Controls.Add(Me.BtnReset)
        Me.KeyShortGroup.Controls.Add(Me.txtStopPly)
        Me.KeyShortGroup.Controls.Add(Me.txtStrPly)
        Me.KeyShortGroup.Controls.Add(Me.txtStopRec)
        Me.KeyShortGroup.Controls.Add(Me.txtStrRec)
        Me.KeyShortGroup.Controls.Add(Me.lblStopPlayback)
        Me.KeyShortGroup.Controls.Add(Me.lblStartPlayback)
        Me.KeyShortGroup.Controls.Add(Me.lblAbortRecording)
        Me.KeyShortGroup.Controls.Add(Me.lblStrRec)
        Me.KeyShortGroup.Location = New System.Drawing.Point(8, 14)
        Me.KeyShortGroup.Name = "KeyShortGroup"
        Me.KeyShortGroup.Size = New System.Drawing.Size(369, 219)
        Me.KeyShortGroup.TabIndex = 3
        Me.KeyShortGroup.TabStop = False
        '
        'BtnReset
        '
        Me.BtnReset.Location = New System.Drawing.Point(252, 176)
        Me.BtnReset.Name = "BtnReset"
        Me.BtnReset.Size = New System.Drawing.Size(113, 25)
        Me.BtnReset.TabIndex = 4
        Me.BtnReset.Text = "Reset To Defaults"
        Me.BtnReset.UseVisualStyleBackColor = True
        '
        'txtStopPly
        '
        Me.txtStopPly.Location = New System.Drawing.Point(226, 138)
        Me.txtStopPly.Name = "txtStopPly"
        Me.txtStopPly.Size = New System.Drawing.Size(139, 20)
        Me.txtStopPly.TabIndex = 3
        '
        'txtStrPly
        '
        Me.txtStrPly.Location = New System.Drawing.Point(224, 97)
        Me.txtStrPly.Name = "txtStrPly"
        Me.txtStrPly.Size = New System.Drawing.Size(139, 20)
        Me.txtStrPly.TabIndex = 2
        '
        'txtStopRec
        '
        Me.txtStopRec.Location = New System.Drawing.Point(224, 61)
        Me.txtStopRec.Name = "txtStopRec"
        Me.txtStopRec.Size = New System.Drawing.Size(139, 20)
        Me.txtStopRec.TabIndex = 1
        '
        'txtStrRec
        '
        Me.txtStrRec.Location = New System.Drawing.Point(224, 25)
        Me.txtStrRec.Name = "txtStrRec"
        Me.txtStrRec.Size = New System.Drawing.Size(139, 20)
        Me.txtStrRec.TabIndex = 0
        '
        'lblStopPlayback
        '
        Me.lblStopPlayback.AutoSize = True
        Me.lblStopPlayback.Location = New System.Drawing.Point(6, 145)
        Me.lblStopPlayback.Name = "lblStopPlayback"
        Me.lblStopPlayback.Size = New System.Drawing.Size(81, 13)
        Me.lblStopPlayback.TabIndex = 3
        Me.lblStopPlayback.Text = "Stop playback :"
        '
        'lblStartPlayback
        '
        Me.lblStartPlayback.AutoSize = True
        Me.lblStartPlayback.Location = New System.Drawing.Point(6, 104)
        Me.lblStartPlayback.Name = "lblStartPlayback"
        Me.lblStartPlayback.Size = New System.Drawing.Size(170, 13)
        Me.lblStartPlayback.TabIndex = 2
        Me.lblStartPlayback.Text = "Start / Pause / Resume Playback :"
        '
        'lblAbortRecording
        '
        Me.lblAbortRecording.AutoSize = True
        Me.lblAbortRecording.Location = New System.Drawing.Point(6, 68)
        Me.lblAbortRecording.Name = "lblAbortRecording"
        Me.lblAbortRecording.Size = New System.Drawing.Size(87, 13)
        Me.lblAbortRecording.TabIndex = 1
        Me.lblAbortRecording.Text = "Stop Recording :"
        '
        'lblStrRec
        '
        Me.lblStrRec.AutoSize = True
        Me.lblStrRec.Location = New System.Drawing.Point(6, 32)
        Me.lblStrRec.Name = "lblStrRec"
        Me.lblStrRec.Size = New System.Drawing.Size(176, 13)
        Me.lblStrRec.TabIndex = 0
        Me.lblStrRec.Text = "Start / Pause / Resume Recording :"
        '
        'BtnCancel
        '
        Me.BtnCancel.CausesValidation = False
        Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnCancel.Location = New System.Drawing.Point(81, 4)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(72, 26)
        Me.BtnCancel.TabIndex = 2
        Me.BtnCancel.Text = "&Cancel"
        Me.BtnCancel.UseVisualStyleBackColor = True
        '
        'BtnOK
        '
        Me.BtnOK.Location = New System.Drawing.Point(3, 3)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(72, 26)
        Me.BtnOK.TabIndex = 1
        Me.BtnOK.Text = "&OK"
        Me.BtnOK.UseVisualStyleBackColor = True
        '
        'lblDefFoldPath
        '
        Me.lblDefFoldPath.AutoSize = True
        Me.lblDefFoldPath.Location = New System.Drawing.Point(6, 85)
        Me.lblDefFoldPath.Name = "lblDefFoldPath"
        Me.lblDefFoldPath.Size = New System.Drawing.Size(107, 13)
        Me.lblDefFoldPath.TabIndex = 7
        Me.lblDefFoldPath.Text = "Default Folder Path :"
        '
        'lblDefMacName
        '
        Me.lblDefMacName.AutoSize = True
        Me.lblDefMacName.Location = New System.Drawing.Point(6, 49)
        Me.lblDefMacName.Name = "lblDefMacName"
        Me.lblDefMacName.Size = New System.Drawing.Size(111, 13)
        Me.lblDefMacName.TabIndex = 8
        Me.lblDefMacName.Text = "Default Macro Name :"
        '
        'DefFoldGroup
        '
        Me.DefFoldGroup.Controls.Add(Me.lblPath)
        Me.DefFoldGroup.Controls.Add(Me.chkDefFold)
        Me.DefFoldGroup.Controls.Add(Me.btnDefFoldPath)
        Me.DefFoldGroup.Controls.Add(Me.txtDefMacName)
        Me.DefFoldGroup.Controls.Add(Me.lblDefMacName)
        Me.DefFoldGroup.Controls.Add(Me.lblDefFoldPath)
        Me.DefFoldGroup.Location = New System.Drawing.Point(3, 120)
        Me.DefFoldGroup.Name = "DefFoldGroup"
        Me.DefFoldGroup.Size = New System.Drawing.Size(368, 117)
        Me.DefFoldGroup.TabIndex = 3
        Me.DefFoldGroup.TabStop = False
        '
        'lblPath
        '
        Me.lblPath.AutoEllipsis = True
        Me.lblPath.BackColor = System.Drawing.Color.White
        Me.lblPath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPath.Location = New System.Drawing.Point(115, 79)
        Me.lblPath.Name = "lblPath"
        Me.lblPath.Size = New System.Drawing.Size(211, 19)
        Me.lblPath.TabIndex = 12
        '
        'chkDefFold
        '
        Me.chkDefFold.AutoSize = True
        Me.chkDefFold.Location = New System.Drawing.Point(6, 19)
        Me.chkDefFold.Name = "chkDefFold"
        Me.chkDefFold.Size = New System.Drawing.Size(295, 17)
        Me.chkDefFold.TabIndex = 0
        Me.chkDefFold.Text = "Save my macros to a default folder with a default name "
        Me.chkDefFold.UseVisualStyleBackColor = True
        '
        'btnDefFoldPath
        '
        Me.btnDefFoldPath.Enabled = False
        Me.btnDefFoldPath.Location = New System.Drawing.Point(332, 78)
        Me.btnDefFoldPath.Name = "btnDefFoldPath"
        Me.btnDefFoldPath.Size = New System.Drawing.Size(30, 20)
        Me.btnDefFoldPath.TabIndex = 11
        Me.btnDefFoldPath.Text = "..."
        Me.btnDefFoldPath.UseVisualStyleBackColor = True
        '
        'txtDefMacName
        '
        Me.txtDefMacName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDefMacName.Enabled = False
        Me.txtDefMacName.Location = New System.Drawing.Point(115, 42)
        Me.txtDefMacName.Name = "txtDefMacName"
        Me.txtDefMacName.Size = New System.Drawing.Size(211, 20)
        Me.txtDefMacName.TabIndex = 10
        Me.txtDefMacName.Text = "default.pmac"
        '
        'OptionsTab
        '
        Me.OptionsTab.Controls.Add(Me.TabPage1)
        Me.OptionsTab.Controls.Add(Me.TabPage2)
        Me.OptionsTab.Controls.Add(Me.TabShortCuts)
        Me.OptionsTab.Location = New System.Drawing.Point(12, 12)
        Me.OptionsTab.Name = "OptionsTab"
        Me.OptionsTab.SelectedIndex = 0
        Me.OptionsTab.Size = New System.Drawing.Size(392, 270)
        Me.OptionsTab.TabIndex = 13
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.GenOptGroup)
        Me.TabPage1.Controls.Add(Me.DefFoldGroup)
        Me.TabPage1.Controls.Add(Me.RecOptGroup)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(384, 244)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "General"
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage2.Controls.Add(Me.PlyOptGroup)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(384, 244)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Playbacking Options"
        '
        'TabShortCuts
        '
        Me.TabShortCuts.BackColor = System.Drawing.SystemColors.Control
        Me.TabShortCuts.Controls.Add(Me.KeyShortGroup)
        Me.TabShortCuts.Location = New System.Drawing.Point(4, 22)
        Me.TabShortCuts.Name = "TabShortCuts"
        Me.TabShortCuts.Size = New System.Drawing.Size(384, 244)
        Me.TabShortCuts.TabIndex = 2
        Me.TabShortCuts.Text = "Keyboard Shortcuts"
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.BtnCancel)
        Me.Panel1.Controls.Add(Me.BtnOK)
        Me.Panel1.Location = New System.Drawing.Point(250, 288)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(157, 33)
        Me.Panel1.TabIndex = 14
        '
        'CMOptionsForm
        '
        Me.AcceptButton = Me.BtnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnCancel
        Me.ClientSize = New System.Drawing.Size(413, 332)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.OptionsTab)
        Me.Controls.Add(Me.PictureBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CMOptionsForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Perfect Macro Recorder Options"
        Me.GenOptGroup.ResumeLayout(False)
        Me.GenOptGroup.PerformLayout()
        Me.RecOptGroup.ResumeLayout(False)
        Me.RecOptGroup.PerformLayout()
        Me.PlyOptGroup.ResumeLayout(False)
        Me.PlyOptGroup.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PlaSpdSlider, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.KeyShortGroup.ResumeLayout(False)
        Me.KeyShortGroup.PerformLayout()
        Me.DefFoldGroup.ResumeLayout(False)
        Me.DefFoldGroup.PerformLayout()
        Me.OptionsTab.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TabShortCuts.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GenOptGroup As System.Windows.Forms.GroupBox
    Friend WithEvents chkStartup As System.Windows.Forms.CheckBox
    Friend WithEvents chkAutoRecord As System.Windows.Forms.CheckBox
    Friend WithEvents RecOptGroup As System.Windows.Forms.GroupBox
    Friend WithEvents PlyOptGroup As System.Windows.Forms.GroupBox
    Friend WithEvents KeyShortGroup As System.Windows.Forms.GroupBox
    Friend WithEvents chkMouMov As System.Windows.Forms.CheckBox
    Friend WithEvents chkKeyActs As System.Windows.Forms.CheckBox
    Friend WithEvents chkAllMouActs As System.Windows.Forms.CheckBox
    Friend WithEvents PlaSpdSlider As System.Windows.Forms.TrackBar
    Friend WithEvents lblPlayBackSpeed As System.Windows.Forms.Label
    Friend WithEvents txtRepeatTimes As System.Windows.Forms.TextBox
    Friend WithEvents lblRepeatTimes As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblSlow As System.Windows.Forms.Label
    Friend WithEvents lblVerySlow As System.Windows.Forms.Label
    Friend WithEvents chkNotify As System.Windows.Forms.CheckBox
    Friend WithEvents chkCloseCM As System.Windows.Forms.CheckBox
    Friend WithEvents chkTips As System.Windows.Forms.CheckBox
    Friend WithEvents lblAbortRecording As System.Windows.Forms.Label
    Friend WithEvents lblStrRec As System.Windows.Forms.Label
    Friend WithEvents txtStrRec As System.Windows.Forms.TextBox
    Friend WithEvents lblStopPlayback As System.Windows.Forms.Label
    Friend WithEvents lblStartPlayback As System.Windows.Forms.Label
    Friend WithEvents txtStopPly As System.Windows.Forms.TextBox
    Friend WithEvents txtStrPly As System.Windows.Forms.TextBox
    Friend WithEvents txtStopRec As System.Windows.Forms.TextBox
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents lblDefFoldPath As System.Windows.Forms.Label
    Friend WithEvents lblDefMacName As System.Windows.Forms.Label
    Friend WithEvents DefFoldGroup As System.Windows.Forms.GroupBox
    Friend WithEvents btnDefFoldPath As System.Windows.Forms.Button
    Friend WithEvents txtDefMacName As System.Windows.Forms.TextBox
    Friend WithEvents chkDefFold As System.Windows.Forms.CheckBox
    Friend WithEvents lblPath As System.Windows.Forms.Label
    Friend WithEvents FoldBrowDia As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents ChkSmartPlayback As System.Windows.Forms.CheckBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents OptionsTab As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TabShortCuts As System.Windows.Forms.TabPage
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents LBLTip As System.Windows.Forms.Label
    Friend WithEvents BtnReset As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
End Class
