Namespace UI
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class TriggersForm
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
            Dim ListViewGroup5 As System.Windows.Forms.ListViewGroup = New System.Windows.Forms.ListViewGroup("Schedule Triggers", System.Windows.Forms.HorizontalAlignment.Left)
            Dim ListViewGroup6 As System.Windows.Forms.ListViewGroup = New System.Windows.Forms.ListViewGroup("Event Triggers", System.Windows.Forms.HorizontalAlignment.Left)
            Me.AddScheduleTriggerSimpleButton = New DevExpress.XtraEditors.SimpleButton()
            Me.AddEventTriggerSimpleButton = New DevExpress.XtraEditors.SimpleButton()
            Me.EditTriggerSimpleButton = New DevExpress.XtraEditors.SimpleButton()
            Me.DeleteTriggerSimpleButton = New DevExpress.XtraEditors.SimpleButton()
            Me.CloseSimpleButton = New DevExpress.XtraEditors.SimpleButton()
            Me.TriggersListView = New System.Windows.Forms.ListView()
            Me.TypeColumn = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
            Me.StateColumn = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
            Me.DescriptionColumn = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
            Me.PauseAllSimpleButton = New DevExpress.XtraEditors.SimpleButton()
            Me.ResumeAllSimpleButton = New DevExpress.XtraEditors.SimpleButton()
            Me.PauseSimpleButton = New DevExpress.XtraEditors.SimpleButton()
            Me.ResumeSimpleButton = New DevExpress.XtraEditors.SimpleButton()
            Me.TriggersFromStatusStrip = New System.Windows.Forms.StatusStrip()
            Me.TriggerFormLeftToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
            Me.TriggersFromStatusStrip.SuspendLayout()
            Me.SuspendLayout()
            '
            'AddScheduleTriggerSimpleButton
            '
            Me.AddScheduleTriggerSimpleButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.AddScheduleTriggerSimpleButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
            Me.AddScheduleTriggerSimpleButton.Location = New System.Drawing.Point(417, 7)
            Me.AddScheduleTriggerSimpleButton.Name = "AddScheduleTriggerSimpleButton"
            Me.AddScheduleTriggerSimpleButton.Size = New System.Drawing.Size(109, 23)
            Me.AddScheduleTriggerSimpleButton.TabIndex = 1
            Me.AddScheduleTriggerSimpleButton.Text = "&Add Schedule Trigger"
            '
            'AddEventTriggerSimpleButton
            '
            Me.AddEventTriggerSimpleButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.AddEventTriggerSimpleButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
            Me.AddEventTriggerSimpleButton.Location = New System.Drawing.Point(417, 259)
            Me.AddEventTriggerSimpleButton.Name = "AddEventTriggerSimpleButton"
            Me.AddEventTriggerSimpleButton.Size = New System.Drawing.Size(109, 23)
            Me.AddEventTriggerSimpleButton.TabIndex = 8
            Me.AddEventTriggerSimpleButton.Text = "Add &Event Trigger"
            Me.AddEventTriggerSimpleButton.Visible = False
            '
            'EditTriggerSimpleButton
            '
            Me.EditTriggerSimpleButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.EditTriggerSimpleButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
            Me.EditTriggerSimpleButton.Enabled = False
            Me.EditTriggerSimpleButton.Location = New System.Drawing.Point(417, 36)
            Me.EditTriggerSimpleButton.Name = "EditTriggerSimpleButton"
            Me.EditTriggerSimpleButton.Size = New System.Drawing.Size(109, 23)
            Me.EditTriggerSimpleButton.TabIndex = 2
            Me.EditTriggerSimpleButton.Text = "E&dit Trigger"
            '
            'DeleteTriggerSimpleButton
            '
            Me.DeleteTriggerSimpleButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.DeleteTriggerSimpleButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
            Me.DeleteTriggerSimpleButton.Enabled = False
            Me.DeleteTriggerSimpleButton.Location = New System.Drawing.Point(417, 65)
            Me.DeleteTriggerSimpleButton.Name = "DeleteTriggerSimpleButton"
            Me.DeleteTriggerSimpleButton.Size = New System.Drawing.Size(109, 23)
            Me.DeleteTriggerSimpleButton.TabIndex = 3
            Me.DeleteTriggerSimpleButton.Text = "&Delete Trigger"
            '
            'CloseSimpleButton
            '
            Me.CloseSimpleButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.CloseSimpleButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
            Me.CloseSimpleButton.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.CloseSimpleButton.Location = New System.Drawing.Point(417, 288)
            Me.CloseSimpleButton.Name = "CloseSimpleButton"
            Me.CloseSimpleButton.Size = New System.Drawing.Size(109, 23)
            Me.CloseSimpleButton.TabIndex = 9
            Me.CloseSimpleButton.Text = "&Close"
            '
            'TriggersListView
            '
            Me.TriggersListView.Activation = System.Windows.Forms.ItemActivation.OneClick
            Me.TriggersListView.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.TriggersListView.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.TypeColumn, Me.StateColumn, Me.DescriptionColumn})
            Me.TriggersListView.FullRowSelect = True
            ListViewGroup5.Header = "Schedule Triggers"
            ListViewGroup5.Name = "scheduleTriggersListViewGroup"
            ListViewGroup6.Header = "Event Triggers"
            ListViewGroup6.Name = "eventTriggersListViewGroup"
            Me.TriggersListView.Groups.AddRange(New System.Windows.Forms.ListViewGroup() {ListViewGroup5, ListViewGroup6})
            Me.TriggersListView.Location = New System.Drawing.Point(6, 7)
            Me.TriggersListView.Name = "TriggersListView"
            Me.TriggersListView.ShowItemToolTips = True
            Me.TriggersListView.Size = New System.Drawing.Size(405, 304)
            Me.TriggersListView.TabIndex = 0
            Me.TriggersListView.TileSize = New System.Drawing.Size(377, 25)
            Me.TriggersListView.UseCompatibleStateImageBehavior = False
            Me.TriggersListView.View = System.Windows.Forms.View.Details
            '
            'TypeColumn
            '
            Me.TypeColumn.Text = "Type"
            Me.TypeColumn.Width = 110
            '
            'StateColumn
            '
            Me.StateColumn.DisplayIndex = 2
            Me.StateColumn.Text = "State"
            Me.StateColumn.Width = 200
            '
            'DescriptionColumn
            '
            Me.DescriptionColumn.DisplayIndex = 1
            Me.DescriptionColumn.Text = "Description"
            Me.DescriptionColumn.Width = 88
            '
            'PauseAllSimpleButton
            '
            Me.PauseAllSimpleButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.PauseAllSimpleButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
            Me.PauseAllSimpleButton.Location = New System.Drawing.Point(417, 168)
            Me.PauseAllSimpleButton.Name = "PauseAllSimpleButton"
            Me.PauseAllSimpleButton.Size = New System.Drawing.Size(109, 23)
            Me.PauseAllSimpleButton.TabIndex = 6
            Me.PauseAllSimpleButton.Text = "Pa&use All"
            '
            'ResumeAllSimpleButton
            '
            Me.ResumeAllSimpleButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.ResumeAllSimpleButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
            Me.ResumeAllSimpleButton.Location = New System.Drawing.Point(417, 197)
            Me.ResumeAllSimpleButton.Name = "ResumeAllSimpleButton"
            Me.ResumeAllSimpleButton.Size = New System.Drawing.Size(109, 23)
            Me.ResumeAllSimpleButton.TabIndex = 7
            Me.ResumeAllSimpleButton.Text = "Re&sume All"
            '
            'PauseSimpleButton
            '
            Me.PauseSimpleButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.PauseSimpleButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
            Me.PauseSimpleButton.Enabled = False
            Me.PauseSimpleButton.Location = New System.Drawing.Point(417, 110)
            Me.PauseSimpleButton.Name = "PauseSimpleButton"
            Me.PauseSimpleButton.Size = New System.Drawing.Size(109, 23)
            Me.PauseSimpleButton.TabIndex = 4
            Me.PauseSimpleButton.Text = "&Pause "
            '
            'ResumeSimpleButton
            '
            Me.ResumeSimpleButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.ResumeSimpleButton.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.HotFlat
            Me.ResumeSimpleButton.Enabled = False
            Me.ResumeSimpleButton.Location = New System.Drawing.Point(417, 139)
            Me.ResumeSimpleButton.Name = "ResumeSimpleButton"
            Me.ResumeSimpleButton.Size = New System.Drawing.Size(109, 23)
            Me.ResumeSimpleButton.TabIndex = 5
            Me.ResumeSimpleButton.Text = "&Resume"
            '
            'TriggersFromStatusStrip
            '
            Me.TriggersFromStatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TriggerFormLeftToolStripStatusLabel})
            Me.TriggersFromStatusStrip.Location = New System.Drawing.Point(0, 314)
            Me.TriggersFromStatusStrip.Name = "TriggersFromStatusStrip"
            Me.TriggersFromStatusStrip.Size = New System.Drawing.Size(531, 22)
            Me.TriggersFromStatusStrip.TabIndex = 10
            Me.TriggersFromStatusStrip.Text = "StatusStrip1"
            '
            'TriggerFormLeftToolStripStatusLabel
            '
            Me.TriggerFormLeftToolStripStatusLabel.BackColor = System.Drawing.Color.Transparent
            Me.TriggerFormLeftToolStripStatusLabel.Name = "TriggerFormLeftToolStripStatusLabel"
            Me.TriggerFormLeftToolStripStatusLabel.Size = New System.Drawing.Size(0, 17)
            '
            'TriggersForm
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.CancelButton = Me.CloseSimpleButton
            Me.ClientSize = New System.Drawing.Size(531, 336)
            Me.Controls.Add(Me.TriggersFromStatusStrip)
            Me.Controls.Add(Me.ResumeSimpleButton)
            Me.Controls.Add(Me.PauseSimpleButton)
            Me.Controls.Add(Me.ResumeAllSimpleButton)
            Me.Controls.Add(Me.PauseAllSimpleButton)
            Me.Controls.Add(Me.TriggersListView)
            Me.Controls.Add(Me.CloseSimpleButton)
            Me.Controls.Add(Me.DeleteTriggerSimpleButton)
            Me.Controls.Add(Me.EditTriggerSimpleButton)
            Me.Controls.Add(Me.AddEventTriggerSimpleButton)
            Me.Controls.Add(Me.AddScheduleTriggerSimpleButton)
            Me.KeyPreview = True
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "TriggersForm"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.Text = "Triggers"
            Me.TriggersFromStatusStrip.ResumeLayout(False)
            Me.TriggersFromStatusStrip.PerformLayout()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents AddScheduleTriggerSimpleButton As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents AddEventTriggerSimpleButton As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents EditTriggerSimpleButton As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents DeleteTriggerSimpleButton As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents CloseSimpleButton As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents TriggersListView As System.Windows.Forms.ListView
        Friend WithEvents TypeColumn As System.Windows.Forms.ColumnHeader
        Friend WithEvents DescriptionColumn As System.Windows.Forms.ColumnHeader
        Friend WithEvents PauseAllSimpleButton As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents ResumeAllSimpleButton As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents StateColumn As System.Windows.Forms.ColumnHeader
        Friend WithEvents PauseSimpleButton As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents ResumeSimpleButton As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents TriggersFromStatusStrip As System.Windows.Forms.StatusStrip
        Friend WithEvents TriggerFormLeftToolStripStatusLabel As System.Windows.Forms.ToolStripStatusLabel
    End Class
End Namespace

