Namespace UI

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class OptionsForm
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
            Me.WindowStartupCheckEdit = New DevExpress.XtraEditors.CheckEdit()
            Me.DisableSplashScreenCheckEdit = New DevExpress.XtraEditors.CheckEdit()
            Me.SendErrorReportsCheckEdit = New DevExpress.XtraEditors.CheckEdit()
            Me.CacaduEnabledCheckEdit = New DevExpress.XtraEditors.CheckEdit()
            Me.OnClosingRadioGroup = New DevExpress.XtraEditors.RadioGroup()
            Me.OnClosingLabelControl = New DevExpress.XtraEditors.LabelControl()
            Me.CancelSimpleButton = New DevExpress.XtraEditors.SimpleButton()
            Me.OKSimpleButton = New DevExpress.XtraEditors.SimpleButton()
            Me.AutoStartTriggersCheckEdit = New DevExpress.XtraEditors.CheckEdit()
            Me.AutoCloseAlertsCheckEdit = New DevExpress.XtraEditors.CheckEdit()
            Me.DefaultpositionForCacaduAlertComboBoxEdit = New DevExpress.XtraEditors.ComboBoxEdit()
            Me.DefaultpositionForCacaduAlertLabelControl = New DevExpress.XtraEditors.LabelControl()
            Me.AskMeFirstBeforeResumingTriggersCheckEdit = New DevExpress.XtraEditors.CheckEdit()
            CType(Me.WindowStartupCheckEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.DisableSplashScreenCheckEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.SendErrorReportsCheckEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.CacaduEnabledCheckEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.OnClosingRadioGroup.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.AutoStartTriggersCheckEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.AutoCloseAlertsCheckEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.DefaultpositionForCacaduAlertComboBoxEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.AskMeFirstBeforeResumingTriggersCheckEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'WindowStartupCheckEdit
            '
            Me.WindowStartupCheckEdit.Location = New System.Drawing.Point(12, 37)
            Me.WindowStartupCheckEdit.Name = "WindowStartupCheckEdit"
            Me.WindowStartupCheckEdit.Properties.Caption = "Starts at Windows startup"
            Me.WindowStartupCheckEdit.Size = New System.Drawing.Size(152, 19)
            Me.WindowStartupCheckEdit.TabIndex = 0
            '
            'DisableSplashScreenCheckEdit
            '
            Me.DisableSplashScreenCheckEdit.Location = New System.Drawing.Point(12, 62)
            Me.DisableSplashScreenCheckEdit.Name = "DisableSplashScreenCheckEdit"
            Me.DisableSplashScreenCheckEdit.Properties.Caption = "Disable Splash Screen"
            Me.DisableSplashScreenCheckEdit.Size = New System.Drawing.Size(131, 19)
            Me.DisableSplashScreenCheckEdit.TabIndex = 1
            '
            'SendErrorReportsCheckEdit
            '
            Me.SendErrorReportsCheckEdit.Location = New System.Drawing.Point(12, 87)
            Me.SendErrorReportsCheckEdit.Name = "SendErrorReportsCheckEdit"
            Me.SendErrorReportsCheckEdit.Properties.Caption = "Send errors report automatically in the future"
            Me.SendErrorReportsCheckEdit.Size = New System.Drawing.Size(251, 19)
            Me.SendErrorReportsCheckEdit.TabIndex = 2
            '
            'CacaduEnabledCheckEdit
            '
            Me.CacaduEnabledCheckEdit.Location = New System.Drawing.Point(12, 12)
            Me.CacaduEnabledCheckEdit.Name = "CacaduEnabledCheckEdit"
            Me.CacaduEnabledCheckEdit.Properties.Caption = "Cacadu Enabled"
            Me.CacaduEnabledCheckEdit.Size = New System.Drawing.Size(110, 19)
            Me.CacaduEnabledCheckEdit.TabIndex = 9
            '
            'OnClosingRadioGroup
            '
            Me.OnClosingRadioGroup.Location = New System.Drawing.Point(29, 209)
            Me.OnClosingRadioGroup.Name = "OnClosingRadioGroup"
            Me.OnClosingRadioGroup.Properties.Appearance.BackColor = System.Drawing.Color.Transparent
            Me.OnClosingRadioGroup.Properties.Appearance.Options.UseBackColor = True
            Me.OnClosingRadioGroup.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
            Me.OnClosingRadioGroup.Properties.Items.AddRange(New DevExpress.XtraEditors.Controls.RadioGroupItem() {New DevExpress.XtraEditors.Controls.RadioGroupItem(Nothing, "Send to System Tray"), New DevExpress.XtraEditors.Controls.RadioGroupItem(Nothing, "Close Cacadu")})
            Me.OnClosingRadioGroup.Size = New System.Drawing.Size(249, 57)
            Me.OnClosingRadioGroup.TabIndex = 10
            '
            'OnClosingLabelControl
            '
            Me.OnClosingLabelControl.Location = New System.Drawing.Point(12, 190)
            Me.OnClosingLabelControl.Name = "OnClosingLabelControl"
            Me.OnClosingLabelControl.Size = New System.Drawing.Size(55, 13)
            Me.OnClosingLabelControl.TabIndex = 11
            Me.OnClosingLabelControl.Text = "On Closing:"
            '
            'CancelSimpleButton
            '
            Me.CancelSimpleButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.CancelSimpleButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
            Me.CancelSimpleButton.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.CancelSimpleButton.Location = New System.Drawing.Point(306, 308)
            Me.CancelSimpleButton.Name = "CancelSimpleButton"
            Me.CancelSimpleButton.Size = New System.Drawing.Size(88, 23)
            Me.CancelSimpleButton.TabIndex = 13
            Me.CancelSimpleButton.Text = "&Cancel"
            '
            'OKSimpleButton
            '
            Me.OKSimpleButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.OKSimpleButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
            Me.OKSimpleButton.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.OKSimpleButton.Location = New System.Drawing.Point(212, 308)
            Me.OKSimpleButton.Name = "OKSimpleButton"
            Me.OKSimpleButton.Size = New System.Drawing.Size(88, 23)
            Me.OKSimpleButton.TabIndex = 14
            Me.OKSimpleButton.Text = "&OK"
            '
            'AutoStartTriggersCheckEdit
            '
            Me.AutoStartTriggersCheckEdit.Location = New System.Drawing.Point(12, 112)
            Me.AutoStartTriggersCheckEdit.Name = "AutoStartTriggersCheckEdit"
            Me.AutoStartTriggersCheckEdit.Properties.Caption = "Automatically resume triggers at startup."
            Me.AutoStartTriggersCheckEdit.Size = New System.Drawing.Size(251, 19)
            Me.AutoStartTriggersCheckEdit.TabIndex = 15
            '
            'AutoCloseAlertsCheckEdit
            '
            Me.AutoCloseAlertsCheckEdit.Location = New System.Drawing.Point(12, 165)
            Me.AutoCloseAlertsCheckEdit.Name = "AutoCloseAlertsCheckEdit"
            Me.AutoCloseAlertsCheckEdit.Properties.Caption = "Automatically close alerts."
            Me.AutoCloseAlertsCheckEdit.Size = New System.Drawing.Size(251, 19)
            Me.AutoCloseAlertsCheckEdit.TabIndex = 16
            '
            'DefaultpositionForCacaduAlertComboBoxEdit
            '
            Me.DefaultpositionForCacaduAlertComboBoxEdit.EditValue = "BottomRight"
            Me.DefaultpositionForCacaduAlertComboBoxEdit.Location = New System.Drawing.Point(313, 276)
            Me.DefaultpositionForCacaduAlertComboBoxEdit.Name = "DefaultpositionForCacaduAlertComboBoxEdit"
            Me.DefaultpositionForCacaduAlertComboBoxEdit.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
            Me.DefaultpositionForCacaduAlertComboBoxEdit.Properties.Items.AddRange(New Object() {"TopLeft", "TopRight", "BottomLeft", "BottomRight"})
            Me.DefaultpositionForCacaduAlertComboBoxEdit.Size = New System.Drawing.Size(80, 20)
            Me.DefaultpositionForCacaduAlertComboBoxEdit.TabIndex = 18
            '
            'DefaultpositionForCacaduAlertLabelControl
            '
            Me.DefaultpositionForCacaduAlertLabelControl.Location = New System.Drawing.Point(12, 279)
            Me.DefaultpositionForCacaduAlertLabelControl.Name = "DefaultpositionForCacaduAlertLabelControl"
            Me.DefaultpositionForCacaduAlertLabelControl.Size = New System.Drawing.Size(295, 13)
            Me.DefaultpositionForCacaduAlertLabelControl.TabIndex = 19
            Me.DefaultpositionForCacaduAlertLabelControl.Text = "Default position for Cacadu alerts (not includes alert actions):" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
            '
            'AskMeFirstBeforeResumingTriggersCheckEdit
            '
            Me.AskMeFirstBeforeResumingTriggersCheckEdit.Location = New System.Drawing.Point(12, 140)
            Me.AskMeFirstBeforeResumingTriggersCheckEdit.Name = "AskMeFirstBeforeResumingTriggersCheckEdit"
            Me.AskMeFirstBeforeResumingTriggersCheckEdit.Properties.Caption = "Ask me first before resuming triggers at startup."
            Me.AskMeFirstBeforeResumingTriggersCheckEdit.Size = New System.Drawing.Size(251, 19)
            Me.AskMeFirstBeforeResumingTriggersCheckEdit.TabIndex = 20
            '
            'OptionsForm
            '
            Me.AcceptButton = Me.OKSimpleButton
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.CancelButton = Me.CancelSimpleButton
            Me.ClientSize = New System.Drawing.Size(406, 343)
            Me.Controls.Add(Me.AskMeFirstBeforeResumingTriggersCheckEdit)
            Me.Controls.Add(Me.DefaultpositionForCacaduAlertLabelControl)
            Me.Controls.Add(Me.DefaultpositionForCacaduAlertComboBoxEdit)
            Me.Controls.Add(Me.AutoCloseAlertsCheckEdit)
            Me.Controls.Add(Me.AutoStartTriggersCheckEdit)
            Me.Controls.Add(Me.OKSimpleButton)
            Me.Controls.Add(Me.CancelSimpleButton)
            Me.Controls.Add(Me.OnClosingLabelControl)
            Me.Controls.Add(Me.OnClosingRadioGroup)
            Me.Controls.Add(Me.CacaduEnabledCheckEdit)
            Me.Controls.Add(Me.SendErrorReportsCheckEdit)
            Me.Controls.Add(Me.DisableSplashScreenCheckEdit)
            Me.Controls.Add(Me.WindowStartupCheckEdit)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "OptionsForm"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.Text = "Options"
            CType(Me.WindowStartupCheckEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.DisableSplashScreenCheckEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.SendErrorReportsCheckEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.CacaduEnabledCheckEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.OnClosingRadioGroup.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.AutoStartTriggersCheckEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.AutoCloseAlertsCheckEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.DefaultpositionForCacaduAlertComboBoxEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.AskMeFirstBeforeResumingTriggersCheckEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents WindowStartupCheckEdit As DevExpress.XtraEditors.CheckEdit
        Friend WithEvents DisableSplashScreenCheckEdit As DevExpress.XtraEditors.CheckEdit
        Friend WithEvents SendErrorReportsCheckEdit As DevExpress.XtraEditors.CheckEdit
        Friend WithEvents CacaduEnabledCheckEdit As DevExpress.XtraEditors.CheckEdit
        Friend WithEvents OnClosingRadioGroup As DevExpress.XtraEditors.RadioGroup
        Friend WithEvents OnClosingLabelControl As DevExpress.XtraEditors.LabelControl
        Friend WithEvents CancelSimpleButton As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents OKSimpleButton As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents AutoStartTriggersCheckEdit As DevExpress.XtraEditors.CheckEdit
        Friend WithEvents AutoCloseAlertsCheckEdit As DevExpress.XtraEditors.CheckEdit
        Friend WithEvents DefaultpositionForCacaduAlertComboBoxEdit As DevExpress.XtraEditors.ComboBoxEdit
        Friend WithEvents DefaultpositionForCacaduAlertLabelControl As DevExpress.XtraEditors.LabelControl
        Friend WithEvents AskMeFirstBeforeResumingTriggersCheckEdit As DevExpress.XtraEditors.CheckEdit
    End Class

End Namespace