Namespace UI

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class AddNewTaskForm
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
            Me.GroupNameLabel = New DevExpress.XtraEditors.LabelControl()
            Me.AddToGroupComboBoxEdit = New DevExpress.XtraEditors.ComboBoxEdit()
            Me.WaitBetweenActionsCheckEdit = New DevExpress.XtraEditors.CheckEdit()
            Me.TaskEnabledCheckEdit = New DevExpress.XtraEditors.CheckEdit()
            Me.TaskNameLabel = New DevExpress.XtraEditors.LabelControl()
            Me.TaskNameTextEdit = New DevExpress.XtraEditors.TextEdit()
            Me.OkButtonOfAddNewTask = New DevExpress.XtraEditors.SimpleButton()
            Me.CancelButtonOfAddNewTask = New DevExpress.XtraEditors.SimpleButton()
            CType(Me.AddToGroupComboBoxEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.WaitBetweenActionsCheckEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.TaskEnabledCheckEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.TaskNameTextEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'GroupNameLabel
            '
            Me.GroupNameLabel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.GroupNameLabel.Location = New System.Drawing.Point(7, 44)
            Me.GroupNameLabel.Name = "GroupNameLabel"
            Me.GroupNameLabel.Size = New System.Drawing.Size(70, 13)
            Me.GroupNameLabel.TabIndex = 7
            Me.GroupNameLabel.Text = "Add To Group:"
            '
            'AddToGroupComboBoxEdit
            '
            Me.AddToGroupComboBoxEdit.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.AddToGroupComboBoxEdit.Location = New System.Drawing.Point(83, 41)
            Me.AddToGroupComboBoxEdit.Name = "AddToGroupComboBoxEdit"
            Me.AddToGroupComboBoxEdit.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
            Me.AddToGroupComboBoxEdit.Properties.Sorted = True
            Me.AddToGroupComboBoxEdit.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
            Me.AddToGroupComboBoxEdit.Size = New System.Drawing.Size(198, 20)
            Me.AddToGroupComboBoxEdit.TabIndex = 2
            '
            'WaitBetweenActionsCheckEdit
            '
            Me.WaitBetweenActionsCheckEdit.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.WaitBetweenActionsCheckEdit.Location = New System.Drawing.Point(5, 80)
            Me.WaitBetweenActionsCheckEdit.Name = "WaitBetweenActionsCheckEdit"
            Me.WaitBetweenActionsCheckEdit.Properties.Caption = "Wait Between Actions"
            Me.WaitBetweenActionsCheckEdit.Size = New System.Drawing.Size(160, 19)
            Me.WaitBetweenActionsCheckEdit.TabIndex = 3
            '
            'TaskEnabledCheckEdit
            '
            Me.TaskEnabledCheckEdit.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.TaskEnabledCheckEdit.EditValue = True
            Me.TaskEnabledCheckEdit.Location = New System.Drawing.Point(5, 105)
            Me.TaskEnabledCheckEdit.Name = "TaskEnabledCheckEdit"
            Me.TaskEnabledCheckEdit.Properties.Caption = "Enabled"
            Me.TaskEnabledCheckEdit.Size = New System.Drawing.Size(160, 19)
            Me.TaskEnabledCheckEdit.TabIndex = 4
            '
            'TaskNameLabel
            '
            Me.TaskNameLabel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.TaskNameLabel.Location = New System.Drawing.Point(7, 15)
            Me.TaskNameLabel.Name = "TaskNameLabel"
            Me.TaskNameLabel.Size = New System.Drawing.Size(56, 13)
            Me.TaskNameLabel.TabIndex = 4
            Me.TaskNameLabel.Text = "Task Name:"
            '
            'TaskNameTextEdit
            '
            Me.TaskNameTextEdit.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.TaskNameTextEdit.Location = New System.Drawing.Point(83, 12)
            Me.TaskNameTextEdit.Name = "TaskNameTextEdit"
            Me.TaskNameTextEdit.Size = New System.Drawing.Size(198, 20)
            Me.TaskNameTextEdit.TabIndex = 1
            '
            'OkButtonOfAddNewTask
            '
            Me.OkButtonOfAddNewTask.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.OkButtonOfAddNewTask.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
            Me.OkButtonOfAddNewTask.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.OkButtonOfAddNewTask.Enabled = False
            Me.OkButtonOfAddNewTask.Location = New System.Drawing.Point(125, 140)
            Me.OkButtonOfAddNewTask.Name = "OkButtonOfAddNewTask"
            Me.OkButtonOfAddNewTask.Size = New System.Drawing.Size(75, 23)
            Me.OkButtonOfAddNewTask.TabIndex = 5
            Me.OkButtonOfAddNewTask.Text = "OK"
            '
            'CancelButtonOfAddNewTask
            '
            Me.CancelButtonOfAddNewTask.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.CancelButtonOfAddNewTask.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
            Me.CancelButtonOfAddNewTask.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.CancelButtonOfAddNewTask.Location = New System.Drawing.Point(206, 140)
            Me.CancelButtonOfAddNewTask.Name = "CancelButtonOfAddNewTask"
            Me.CancelButtonOfAddNewTask.Size = New System.Drawing.Size(75, 23)
            Me.CancelButtonOfAddNewTask.TabIndex = 6
            Me.CancelButtonOfAddNewTask.Text = "Cancel"
            '
            'AddNewTaskForm
            '
            Me.AcceptButton = Me.OkButtonOfAddNewTask
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.CancelButton = Me.CancelButtonOfAddNewTask
            Me.ClientSize = New System.Drawing.Size(287, 169)
            Me.Controls.Add(Me.CancelButtonOfAddNewTask)
            Me.Controls.Add(Me.OkButtonOfAddNewTask)
            Me.Controls.Add(Me.TaskNameTextEdit)
            Me.Controls.Add(Me.TaskNameLabel)
            Me.Controls.Add(Me.TaskEnabledCheckEdit)
            Me.Controls.Add(Me.WaitBetweenActionsCheckEdit)
            Me.Controls.Add(Me.AddToGroupComboBoxEdit)
            Me.Controls.Add(Me.GroupNameLabel)
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "AddNewTaskForm"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.Text = "Add New Task"
            CType(Me.AddToGroupComboBoxEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.WaitBetweenActionsCheckEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.TaskEnabledCheckEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.TaskNameTextEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents GroupNameLabel As DevExpress.XtraEditors.LabelControl
        Friend WithEvents AddToGroupComboBoxEdit As DevExpress.XtraEditors.ComboBoxEdit
        Friend WithEvents WaitBetweenActionsCheckEdit As DevExpress.XtraEditors.CheckEdit
        Friend WithEvents TaskEnabledCheckEdit As DevExpress.XtraEditors.CheckEdit
        Friend WithEvents TaskNameLabel As DevExpress.XtraEditors.LabelControl
        Friend WithEvents TaskNameTextEdit As DevExpress.XtraEditors.TextEdit
        Friend WithEvents OkButtonOfAddNewTask As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents CancelButtonOfAddNewTask As DevExpress.XtraEditors.SimpleButton
    End Class

End Namespace
