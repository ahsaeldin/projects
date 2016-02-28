Imports System.ComponentModel

Namespace ErrorReporter

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class ErrorReportingForm
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
            Me.ErrorReportingFormBottomPanel = New System.Windows.Forms.Panel()
            Me.LabelHeadingTextLabel = New System.Windows.Forms.Label()
            Me.AppIconPictureBox = New System.Windows.Forms.PictureBox()
            Me.ErrorReportingFormTopPanel = New System.Windows.Forms.Panel()
            Me.ErrorDescTextBox = New System.Windows.Forms.TextBox()
            Me.AttachmentPathTextBox = New System.Windows.Forms.TextBox()
            Me.AttachmentButton = New System.Windows.Forms.Button()
            Me.AttachmentLabel = New System.Windows.Forms.Label()
            Me.CommentsTextBox = New System.Windows.Forms.TextBox()
            Me.AutoSendCheckBox = New System.Windows.Forms.CheckBox()
            Me.DoNotSendButton = New System.Windows.Forms.Button()
            Me.SendErrorReportButton = New System.Windows.Forms.Button()
            Me.PlzTellUsLabel = New System.Windows.Forms.Label()
            Me.Statusbar = New System.Windows.Forms.StatusStrip()
            Me.SendingStatusToolStripProgressBar = New System.Windows.Forms.ToolStripProgressBar()
            Me.LabelStartusbar = New System.Windows.Forms.ToolStripStatusLabel()
            Me.ErrorReportingFormBottomPanel.SuspendLayout()
            CType(Me.AppIconPictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.ErrorReportingFormTopPanel.SuspendLayout()
            Me.Statusbar.SuspendLayout()
            Me.SuspendLayout()
            '
            'ErrorReportingFormBottomPanel
            '
            Me.ErrorReportingFormBottomPanel.BackColor = System.Drawing.Color.White
            Me.ErrorReportingFormBottomPanel.Controls.Add(Me.LabelHeadingTextLabel)
            Me.ErrorReportingFormBottomPanel.Controls.Add(Me.AppIconPictureBox)
            Me.ErrorReportingFormBottomPanel.Dock = System.Windows.Forms.DockStyle.Top
            Me.ErrorReportingFormBottomPanel.Location = New System.Drawing.Point(0, 0)
            Me.ErrorReportingFormBottomPanel.Name = "ErrorReportingFormBottomPanel"
            Me.ErrorReportingFormBottomPanel.Size = New System.Drawing.Size(524, 66)
            Me.ErrorReportingFormBottomPanel.TabIndex = 0
            '
            'LabelHeadingTextLabel
            '
            Me.LabelHeadingTextLabel.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.LabelHeadingTextLabel.Location = New System.Drawing.Point(9, 12)
            Me.LabelHeadingTextLabel.Name = "LabelHeadingTextLabel"
            Me.LabelHeadingTextLabel.Size = New System.Drawing.Size(452, 40)
            Me.LabelHeadingTextLabel.TabIndex = 0
            Me.LabelHeadingTextLabel.Text = "Run to see me."
            '
            'AppIconPictureBox
            '
            Me.AppIconPictureBox.Location = New System.Drawing.Point(467, 10)
            Me.AppIconPictureBox.Name = "AppIconPictureBox"
            Me.AppIconPictureBox.Size = New System.Drawing.Size(48, 48)
            Me.AppIconPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
            Me.AppIconPictureBox.TabIndex = 0
            Me.AppIconPictureBox.TabStop = False
            '
            'ErrorReportingFormTopPanel
            '
            Me.ErrorReportingFormTopPanel.BackColor = System.Drawing.Color.WhiteSmoke
            Me.ErrorReportingFormTopPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.ErrorReportingFormTopPanel.Controls.Add(Me.ErrorDescTextBox)
            Me.ErrorReportingFormTopPanel.Controls.Add(Me.AttachmentPathTextBox)
            Me.ErrorReportingFormTopPanel.Controls.Add(Me.AttachmentButton)
            Me.ErrorReportingFormTopPanel.Controls.Add(Me.AttachmentLabel)
            Me.ErrorReportingFormTopPanel.Controls.Add(Me.CommentsTextBox)
            Me.ErrorReportingFormTopPanel.Controls.Add(Me.AutoSendCheckBox)
            Me.ErrorReportingFormTopPanel.Controls.Add(Me.DoNotSendButton)
            Me.ErrorReportingFormTopPanel.Controls.Add(Me.SendErrorReportButton)
            Me.ErrorReportingFormTopPanel.Controls.Add(Me.PlzTellUsLabel)
            Me.ErrorReportingFormTopPanel.Dock = System.Windows.Forms.DockStyle.Fill
            Me.ErrorReportingFormTopPanel.Location = New System.Drawing.Point(0, 66)
            Me.ErrorReportingFormTopPanel.Name = "ErrorReportingFormTopPanel"
            Me.ErrorReportingFormTopPanel.Size = New System.Drawing.Size(524, 323)
            Me.ErrorReportingFormTopPanel.TabIndex = 1
            '
            'ErrorDescTextBox
            '
            Me.ErrorDescTextBox.ForeColor = System.Drawing.Color.Black
            Me.ErrorDescTextBox.Location = New System.Drawing.Point(11, 16)
            Me.ErrorDescTextBox.Multiline = True
            Me.ErrorDescTextBox.Name = "ErrorDescTextBox"
            Me.ErrorDescTextBox.ReadOnly = True
            Me.ErrorDescTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
            Me.ErrorDescTextBox.Size = New System.Drawing.Size(503, 90)
            Me.ErrorDescTextBox.TabIndex = 14
            Me.ErrorDescTextBox.Text = """ Excpetion here"""
            '
            'AttachmentPathTextBox
            '
            Me.AttachmentPathTextBox.BackColor = System.Drawing.Color.White
            Me.AttachmentPathTextBox.Location = New System.Drawing.Point(85, 232)
            Me.AttachmentPathTextBox.Name = "AttachmentPathTextBox"
            Me.AttachmentPathTextBox.ReadOnly = True
            Me.AttachmentPathTextBox.Size = New System.Drawing.Size(401, 20)
            Me.AttachmentPathTextBox.TabIndex = 13
            '
            'AttachmentButton
            '
            Me.AttachmentButton.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.AttachmentButton.Location = New System.Drawing.Point(487, 232)
            Me.AttachmentButton.Name = "AttachmentButton"
            Me.AttachmentButton.Size = New System.Drawing.Size(24, 20)
            Me.AttachmentButton.TabIndex = 12
            Me.AttachmentButton.Text = "..."
            Me.AttachmentButton.TextAlign = System.Drawing.ContentAlignment.TopCenter
            Me.AttachmentButton.UseVisualStyleBackColor = True
            '
            'AttachmentLabel
            '
            Me.AttachmentLabel.AutoSize = True
            Me.AttachmentLabel.Location = New System.Drawing.Point(11, 239)
            Me.AttachmentLabel.Name = "AttachmentLabel"
            Me.AttachmentLabel.Size = New System.Drawing.Size(72, 13)
            Me.AttachmentLabel.TabIndex = 9
            Me.AttachmentLabel.Text = "Attachments:"
            '
            'CommentsTextBox
            '
            Me.CommentsTextBox.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
            Me.CommentsTextBox.Location = New System.Drawing.Point(14, 131)
            Me.CommentsTextBox.Multiline = True
            Me.CommentsTextBox.Name = "CommentsTextBox"
            Me.CommentsTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
            Me.CommentsTextBox.Size = New System.Drawing.Size(497, 90)
            Me.CommentsTextBox.TabIndex = 7
            Me.CommentsTextBox.Text = "If you wish, you can add your comments here."
            '
            'AutoSendCheckBox
            '
            Me.AutoSendCheckBox.AutoSize = True
            Me.AutoSendCheckBox.Location = New System.Drawing.Point(14, 275)
            Me.AutoSendCheckBox.Name = "AutoSendCheckBox"
            Me.AutoSendCheckBox.Size = New System.Drawing.Size(248, 17)
            Me.AutoSendCheckBox.TabIndex = 5
            Me.AutoSendCheckBox.Text = "Send errors report &automatically in the future."
            Me.AutoSendCheckBox.UseVisualStyleBackColor = True
            '
            'DoNotSendButton
            '
            Me.DoNotSendButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.DoNotSendButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            Me.DoNotSendButton.Location = New System.Drawing.Point(410, 266)
            Me.DoNotSendButton.Name = "DoNotSendButton"
            Me.DoNotSendButton.Size = New System.Drawing.Size(104, 26)
            Me.DoNotSendButton.TabIndex = 3
            Me.DoNotSendButton.Text = "&Don't Send"
            Me.DoNotSendButton.UseVisualStyleBackColor = True
            '
            'SendErrorReportButton
            '
            Me.SendErrorReportButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            Me.SendErrorReportButton.Location = New System.Drawing.Point(300, 266)
            Me.SendErrorReportButton.Name = "SendErrorReportButton"
            Me.SendErrorReportButton.Size = New System.Drawing.Size(104, 26)
            Me.SendErrorReportButton.TabIndex = 2
            Me.SendErrorReportButton.Text = "&Send Error Report"
            Me.SendErrorReportButton.UseVisualStyleBackColor = True
            '
            'PlzTellUsLabel
            '
            Me.PlzTellUsLabel.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.PlzTellUsLabel.Location = New System.Drawing.Point(11, 109)
            Me.PlzTellUsLabel.Name = "PlzTellUsLabel"
            Me.PlzTellUsLabel.Size = New System.Drawing.Size(503, 21)
            Me.PlzTellUsLabel.TabIndex = 1
            Me.PlzTellUsLabel.Text = "Run to see me."
            '
            'Statusbar
            '
            Me.Statusbar.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SendingStatusToolStripProgressBar, Me.LabelStartusbar})
            Me.Statusbar.Location = New System.Drawing.Point(0, 367)
            Me.Statusbar.Name = "Statusbar"
            Me.Statusbar.Size = New System.Drawing.Size(524, 22)
            Me.Statusbar.TabIndex = 2
            Me.Statusbar.Text = "StatusStrip"
            '
            'SendingStatusToolStripProgressBar
            '
            Me.SendingStatusToolStripProgressBar.MarqueeAnimationSpeed = 50
            Me.SendingStatusToolStripProgressBar.Name = "SendingStatusToolStripProgressBar"
            Me.SendingStatusToolStripProgressBar.Size = New System.Drawing.Size(100, 16)
            Me.SendingStatusToolStripProgressBar.Step = 1
            Me.SendingStatusToolStripProgressBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee
            Me.SendingStatusToolStripProgressBar.Visible = False
            '
            'LabelStartusbar
            '
            Me.LabelStartusbar.ForeColor = System.Drawing.Color.Blue
            Me.LabelStartusbar.Name = "LabelStartusbar"
            Me.LabelStartusbar.Size = New System.Drawing.Size(186, 17)
            Me.LabelStartusbar.Text = "Error Report Sending Status here..."
            Me.LabelStartusbar.Visible = False
            '
            'ErrorReportingForm
            '
            Me.AcceptButton = Me.SendErrorReportButton
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.CancelButton = Me.DoNotSendButton
            Me.ClientSize = New System.Drawing.Size(524, 389)
            Me.Controls.Add(Me.Statusbar)
            Me.Controls.Add(Me.ErrorReportingFormTopPanel)
            Me.Controls.Add(Me.ErrorReportingFormBottomPanel)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "ErrorReportingForm"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.ErrorReportingFormBottomPanel.ResumeLayout(False)
            CType(Me.AppIconPictureBox, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ErrorReportingFormTopPanel.ResumeLayout(False)
            Me.ErrorReportingFormTopPanel.PerformLayout()
            Me.Statusbar.ResumeLayout(False)
            Me.Statusbar.PerformLayout()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Public WithEvents ErrorReportingFormBottomPanel As System.Windows.Forms.Panel
        Public WithEvents ErrorReportingFormTopPanel As System.Windows.Forms.Panel
        Public WithEvents AppIconPictureBox As System.Windows.Forms.PictureBox
        Public WithEvents LabelHeadingTextLabel As System.Windows.Forms.Label
        Public WithEvents PlzTellUsLabel As System.Windows.Forms.Label
        Public WithEvents DoNotSendButton As System.Windows.Forms.Button
        Public WithEvents SendErrorReportButton As System.Windows.Forms.Button
        Public WithEvents AutoSendCheckBox As System.Windows.Forms.CheckBox
        Public WithEvents CommentsTextBox As System.Windows.Forms.TextBox
        Public WithEvents AttachmentLabel As System.Windows.Forms.Label
        Public WithEvents AttachmentButton As System.Windows.Forms.Button
        Public WithEvents Statusbar As System.Windows.Forms.StatusStrip
        Public WithEvents LabelStartusbar As System.Windows.Forms.ToolStripStatusLabel
        Public WithEvents SendingStatusToolStripProgressBar As System.Windows.Forms.ToolStripProgressBar
        Public WithEvents AttachmentPathTextBox As System.Windows.Forms.TextBox
        Public WithEvents ErrorDescTextBox As System.Windows.Forms.TextBox

    End Class
End Namespace