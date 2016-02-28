Namespace UI
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class VariablesAndConstantsForm
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
            Me.components = New System.ComponentModel.Container()
            Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
            Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
            Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
            Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
            Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
            Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
            Me.VarsAndConsXtraTabControl = New DevExpress.XtraTab.XtraTabControl()
            Me.TasksVariablesXtraTabPage = New DevExpress.XtraTab.XtraTabPage()
            Me.VariablesDataGridView = New System.Windows.Forms.DataGridView()
            Me.noColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.VariableNameColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.VariableValueColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.TaskNameColumn = New System.Windows.Forms.DataGridViewComboBoxColumn()
            Me.TasksBindingSource = New System.Windows.Forms.BindingSource(Me.components)
            Me.taskIdColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.GridViewContextMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
            Me.AddToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.DeleteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.Task_variablesBindingSource = New System.Windows.Forms.BindingSource(Me.components)
            Me.GlobalConstantsXtraTabPage = New DevExpress.XtraTab.XtraTabPage()
            Me.GlobalDataGridView = New System.Windows.Forms.DataGridView()
            Me.nameColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.valueColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.typeColumn = New System.Windows.Forms.DataGridViewComboBoxColumn()
            Me.GlobalLookupBindingSource = New System.Windows.Forms.BindingSource(Me.components)
            Me.GlobalContextMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
            Me.AddGlobalToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.EditGlobalToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.DeleteGlobalToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
            Me.GlobalsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
            Me.TasksTableAdapter = New TecDAL.TectonicDataSetTableAdapters.tasksTableAdapter()
            Me.Task_variablesTableAdapter = New TecDAL.TectonicDataSetTableAdapters.task_variablesTableAdapter()
            Me.TableAdapterManager = New TecDAL.TectonicDataSetTableAdapters.TableAdapterManager()
            Me.GlobalsTableAdapter = New TecDAL.TectonicDataSetTableAdapters.globalsTableAdapter()
            Me.VariablesAndConstantsStatusStrip = New System.Windows.Forms.StatusStrip()
            Me.StatusToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
            CType(Me.VarsAndConsXtraTabControl, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.VarsAndConsXtraTabControl.SuspendLayout()
            Me.TasksVariablesXtraTabPage.SuspendLayout()
            CType(Me.VariablesDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.TasksBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.GridViewContextMenuStrip.SuspendLayout()
            CType(Me.Task_variablesBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.GlobalConstantsXtraTabPage.SuspendLayout()
            CType(Me.GlobalDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.GlobalLookupBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.GlobalContextMenuStrip.SuspendLayout()
            CType(Me.GlobalsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.VariablesAndConstantsStatusStrip.SuspendLayout()
            Me.SuspendLayout()
            '
            'VarsAndConsXtraTabControl
            '
            Me.VarsAndConsXtraTabControl.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            Me.VarsAndConsXtraTabControl.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
            Me.VarsAndConsXtraTabControl.Location = New System.Drawing.Point(0, 13)
            Me.VarsAndConsXtraTabControl.Name = "VarsAndConsXtraTabControl"
            Me.VarsAndConsXtraTabControl.SelectedTabPage = Me.TasksVariablesXtraTabPage
            Me.VarsAndConsXtraTabControl.Size = New System.Drawing.Size(477, 380)
            Me.VarsAndConsXtraTabControl.TabIndex = 0
            Me.VarsAndConsXtraTabControl.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.TasksVariablesXtraTabPage, Me.GlobalConstantsXtraTabPage})
            '
            'TasksVariablesXtraTabPage
            '
            Me.TasksVariablesXtraTabPage.Controls.Add(Me.VariablesDataGridView)
            Me.TasksVariablesXtraTabPage.Name = "TasksVariablesXtraTabPage"
            Me.TasksVariablesXtraTabPage.Size = New System.Drawing.Size(471, 352)
            Me.TasksVariablesXtraTabPage.Text = "Tasks Variables"
            '
            'VariablesDataGridView
            '
            Me.VariablesDataGridView.AllowUserToResizeRows = False
            Me.VariablesDataGridView.AutoGenerateColumns = False
            Me.VariablesDataGridView.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
            DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
            DataGridViewCellStyle1.Font = New System.Drawing.Font("Tahoma", 8.0!)
            DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
            DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
            DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
            DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
            Me.VariablesDataGridView.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
            Me.VariablesDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.VariablesDataGridView.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.noColumn, Me.VariableNameColumn, Me.VariableValueColumn, Me.TaskNameColumn, Me.taskIdColumn})
            Me.VariablesDataGridView.ContextMenuStrip = Me.GridViewContextMenuStrip
            Me.VariablesDataGridView.DataSource = Me.Task_variablesBindingSource
            DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
            DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
            DataGridViewCellStyle2.Font = New System.Drawing.Font("Tahoma", 8.0!)
            DataGridViewCellStyle2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(53, Byte), Integer))
            DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
            DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
            DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
            Me.VariablesDataGridView.DefaultCellStyle = DataGridViewCellStyle2
            Me.VariablesDataGridView.Dock = System.Windows.Forms.DockStyle.Fill
            Me.VariablesDataGridView.GridColor = System.Drawing.Color.LightBlue
            Me.VariablesDataGridView.Location = New System.Drawing.Point(0, 0)
            Me.VariablesDataGridView.Name = "VariablesDataGridView"
            DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
            DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
            DataGridViewCellStyle3.Font = New System.Drawing.Font("Tahoma", 8.0!)
            DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
            DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
            DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
            DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
            Me.VariablesDataGridView.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
            Me.VariablesDataGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
            Me.VariablesDataGridView.Size = New System.Drawing.Size(471, 352)
            Me.VariablesDataGridView.TabIndex = 0
            '
            'noColumn
            '
            Me.noColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
            Me.noColumn.DataPropertyName = "no"
            Me.noColumn.FillWeight = 10.0!
            Me.noColumn.HeaderText = "No"
            Me.noColumn.Name = "noColumn"
            Me.noColumn.Visible = False
            '
            'VariableNameColumn
            '
            Me.VariableNameColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
            Me.VariableNameColumn.DataPropertyName = "variable_name"
            Me.VariableNameColumn.FillWeight = 30.0!
            Me.VariableNameColumn.HeaderText = "Variable Name"
            Me.VariableNameColumn.Name = "VariableNameColumn"
            '
            'VariableValueColumn
            '
            Me.VariableValueColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
            Me.VariableValueColumn.DataPropertyName = "variable_value"
            Me.VariableValueColumn.FillWeight = 30.0!
            Me.VariableValueColumn.HeaderText = "Value"
            Me.VariableValueColumn.Name = "VariableValueColumn"
            '
            'TaskNameColumn
            '
            Me.TaskNameColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
            Me.TaskNameColumn.DataPropertyName = "task_id"
            Me.TaskNameColumn.DataSource = Me.TasksBindingSource
            Me.TaskNameColumn.FillWeight = 40.0!
            Me.TaskNameColumn.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            Me.TaskNameColumn.HeaderText = "Task Name"
            Me.TaskNameColumn.Name = "TaskNameColumn"
            '
            'taskIdColumn
            '
            Me.taskIdColumn.DataPropertyName = "task_Id"
            Me.taskIdColumn.HeaderText = "TaskId"
            Me.taskIdColumn.Name = "taskIdColumn"
            Me.taskIdColumn.ReadOnly = True
            Me.taskIdColumn.Visible = False
            '
            'GridViewContextMenuStrip
            '
            Me.GridViewContextMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AddToolStripMenuItem, Me.EditToolStripMenuItem, Me.DeleteToolStripMenuItem})
            Me.GridViewContextMenuStrip.Name = "GridViewContextMenuStrip"
            Me.GridViewContextMenuStrip.Size = New System.Drawing.Size(108, 70)
            '
            'AddToolStripMenuItem
            '
            Me.AddToolStripMenuItem.Name = "AddToolStripMenuItem"
            Me.AddToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
            Me.AddToolStripMenuItem.Text = "&Add"
            '
            'EditToolStripMenuItem
            '
            Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
            Me.EditToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
            Me.EditToolStripMenuItem.Text = "&Edit"
            '
            'DeleteToolStripMenuItem
            '
            Me.DeleteToolStripMenuItem.Name = "DeleteToolStripMenuItem"
            Me.DeleteToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
            Me.DeleteToolStripMenuItem.Text = "&Delete"
            '
            'GlobalConstantsXtraTabPage
            '
            Me.GlobalConstantsXtraTabPage.Controls.Add(Me.GlobalDataGridView)
            Me.GlobalConstantsXtraTabPage.Name = "GlobalConstantsXtraTabPage"
            Me.GlobalConstantsXtraTabPage.Size = New System.Drawing.Size(471, 352)
            Me.GlobalConstantsXtraTabPage.Text = "Global Members"
            '
            'GlobalDataGridView
            '
            Me.GlobalDataGridView.AllowUserToResizeRows = False
            Me.GlobalDataGridView.AutoGenerateColumns = False
            Me.GlobalDataGridView.BorderStyle = System.Windows.Forms.BorderStyle.None
            DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
            DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
            DataGridViewCellStyle4.Font = New System.Drawing.Font("Tahoma", 8.0!)
            DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
            DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
            DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
            DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
            Me.GlobalDataGridView.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
            Me.GlobalDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.GlobalDataGridView.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.nameColumn, Me.valueColumn, Me.typeColumn})
            Me.GlobalDataGridView.ContextMenuStrip = Me.GlobalContextMenuStrip
            Me.GlobalDataGridView.DataSource = Me.GlobalsBindingSource
            DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
            DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
            DataGridViewCellStyle5.Font = New System.Drawing.Font("Tahoma", 8.0!)
            DataGridViewCellStyle5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(32, Byte), Integer), CType(CType(31, Byte), Integer), CType(CType(53, Byte), Integer))
            DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
            DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
            DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
            Me.GlobalDataGridView.DefaultCellStyle = DataGridViewCellStyle5
            Me.GlobalDataGridView.Dock = System.Windows.Forms.DockStyle.Fill
            Me.GlobalDataGridView.GridColor = System.Drawing.Color.LightBlue
            Me.GlobalDataGridView.Location = New System.Drawing.Point(0, 0)
            Me.GlobalDataGridView.Name = "GlobalDataGridView"
            DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
            DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Control
            DataGridViewCellStyle6.Font = New System.Drawing.Font("Tahoma", 8.0!)
            DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText
            DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
            DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
            DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
            Me.GlobalDataGridView.RowHeadersDefaultCellStyle = DataGridViewCellStyle6
            Me.GlobalDataGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
            Me.GlobalDataGridView.Size = New System.Drawing.Size(471, 352)
            Me.GlobalDataGridView.TabIndex = 1
            '
            'nameColumn
            '
            Me.nameColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
            Me.nameColumn.DataPropertyName = "name"
            Me.nameColumn.FillWeight = 30.0!
            Me.nameColumn.HeaderText = "Name"
            Me.nameColumn.Name = "nameColumn"
            '
            'valueColumn
            '
            Me.valueColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
            Me.valueColumn.DataPropertyName = "value"
            Me.valueColumn.FillWeight = 30.0!
            Me.valueColumn.HeaderText = "Value"
            Me.valueColumn.Name = "valueColumn"
            '
            'typeColumn
            '
            Me.typeColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
            Me.typeColumn.DataPropertyName = "type"
            Me.typeColumn.DataSource = Me.GlobalLookupBindingSource
            Me.typeColumn.FillWeight = 40.0!
            Me.typeColumn.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            Me.typeColumn.HeaderText = "Type"
            Me.typeColumn.Name = "typeColumn"
            Me.typeColumn.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
            Me.typeColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
            '
            'GlobalContextMenuStrip
            '
            Me.GlobalContextMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AddGlobalToolStripMenuItem, Me.EditGlobalToolStripMenuItem, Me.DeleteGlobalToolStripMenuItem})
            Me.GlobalContextMenuStrip.Name = "GridViewContextMenuStrip"
            Me.GlobalContextMenuStrip.Size = New System.Drawing.Size(108, 70)
            '
            'AddGlobalToolStripMenuItem
            '
            Me.AddGlobalToolStripMenuItem.Name = "AddGlobalToolStripMenuItem"
            Me.AddGlobalToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
            Me.AddGlobalToolStripMenuItem.Text = "&Add"
            '
            'EditGlobalToolStripMenuItem
            '
            Me.EditGlobalToolStripMenuItem.Name = "EditGlobalToolStripMenuItem"
            Me.EditGlobalToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
            Me.EditGlobalToolStripMenuItem.Text = "&Edit"
            '
            'DeleteGlobalToolStripMenuItem
            '
            Me.DeleteGlobalToolStripMenuItem.Name = "DeleteGlobalToolStripMenuItem"
            Me.DeleteGlobalToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
            Me.DeleteGlobalToolStripMenuItem.Text = "&Delete"
            '
            'TasksTableAdapter
            '
            Me.TasksTableAdapter.ClearBeforeFill = True
            '
            'Task_variablesTableAdapter
            '
            Me.Task_variablesTableAdapter.ClearBeforeFill = True
            '
            'TableAdapterManager
            '
            Me.TableAdapterManager.actionsTableAdapter = Nothing
            Me.TableAdapterManager.BackupDataSetBeforeUpdate = False
            Me.TableAdapterManager.global_lookupTableAdapter = Nothing
            Me.TableAdapterManager.globalsTableAdapter = Me.GlobalsTableAdapter
            Me.TableAdapterManager.groupsTableAdapter = Nothing
            Me.TableAdapterManager.historyTableAdapter = Nothing
            Me.TableAdapterManager.settingsTableAdapter = Nothing
            Me.TableAdapterManager.sqlite_sequenceTableAdapter = Nothing
            Me.TableAdapterManager.task_variablesTableAdapter = Me.Task_variablesTableAdapter
            Me.TableAdapterManager.tasks_statesTableAdapter = Nothing
            Me.TableAdapterManager.tasksTableAdapter = Me.TasksTableAdapter
            Me.TableAdapterManager.triggersTableAdapter = Nothing
            Me.TableAdapterManager.UpdateOrder = TecDAL.TectonicDataSetTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
            '
            'GlobalsTableAdapter
            '
            Me.GlobalsTableAdapter.ClearBeforeFill = True
            '
            'VariablesAndConstantsStatusStrip
            '
            Me.VariablesAndConstantsStatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StatusToolStripStatusLabel})
            Me.VariablesAndConstantsStatusStrip.Location = New System.Drawing.Point(0, 383)
            Me.VariablesAndConstantsStatusStrip.Name = "VariablesAndConstantsStatusStrip"
            Me.VariablesAndConstantsStatusStrip.Size = New System.Drawing.Size(474, 22)
            Me.VariablesAndConstantsStatusStrip.TabIndex = 1
            Me.VariablesAndConstantsStatusStrip.Text = "StatusStrip1"
            '
            'StatusToolStripStatusLabel
            '
            Me.StatusToolStripStatusLabel.BackColor = System.Drawing.Color.FromArgb(CType(CType(241, Byte), Integer), CType(CType(237, Byte), Integer), CType(CType(237, Byte), Integer))
            Me.StatusToolStripStatusLabel.Name = "StatusToolStripStatusLabel"
            Me.StatusToolStripStatusLabel.Size = New System.Drawing.Size(0, 17)
            '
            'VariablesAndConstantsForm
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(474, 405)
            Me.Controls.Add(Me.VariablesAndConstantsStatusStrip)
            Me.Controls.Add(Me.VarsAndConsXtraTabControl)
            Me.KeyPreview = True
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "VariablesAndConstantsForm"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.Text = "Variables and Constants"
            CType(Me.VarsAndConsXtraTabControl, System.ComponentModel.ISupportInitialize).EndInit()
            Me.VarsAndConsXtraTabControl.ResumeLayout(False)
            Me.TasksVariablesXtraTabPage.ResumeLayout(False)
            CType(Me.VariablesDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.TasksBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
            Me.GridViewContextMenuStrip.ResumeLayout(False)
            CType(Me.Task_variablesBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
            Me.GlobalConstantsXtraTabPage.ResumeLayout(False)
            CType(Me.GlobalDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.GlobalLookupBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
            Me.GlobalContextMenuStrip.ResumeLayout(False)
            CType(Me.GlobalsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
            Me.VariablesAndConstantsStatusStrip.ResumeLayout(False)
            Me.VariablesAndConstantsStatusStrip.PerformLayout()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents VarsAndConsXtraTabControl As DevExpress.XtraTab.XtraTabControl
        Friend WithEvents TasksVariablesXtraTabPage As DevExpress.XtraTab.XtraTabPage
        Friend WithEvents GlobalConstantsXtraTabPage As DevExpress.XtraTab.XtraTabPage
        Friend WithEvents VariablesDataGridView As System.Windows.Forms.DataGridView
        Friend WithEvents Task_variablesBindingSource As System.Windows.Forms.BindingSource
        Friend WithEvents TasksBindingSource As System.Windows.Forms.BindingSource
        Friend WithEvents TasksTableAdapter As TecDAL.TectonicDataSetTableAdapters.tasksTableAdapter
        Friend WithEvents Task_variablesTableAdapter As TecDAL.TectonicDataSetTableAdapters.task_variablesTableAdapter
        Friend WithEvents TableAdapterManager As TecDAL.TectonicDataSetTableAdapters.TableAdapterManager
        Friend WithEvents GridViewContextMenuStrip As System.Windows.Forms.ContextMenuStrip
        Friend WithEvents AddToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents EditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents DeleteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents VariablesAndConstantsStatusStrip As System.Windows.Forms.StatusStrip
        Friend WithEvents StatusToolStripStatusLabel As System.Windows.Forms.ToolStripStatusLabel
        Friend WithEvents GlobalDataGridView As System.Windows.Forms.DataGridView
        Friend WithEvents GlobalsTableAdapter As TecDAL.TectonicDataSetTableAdapters.globalsTableAdapter
        Friend WithEvents GlobalsBindingSource As System.Windows.Forms.BindingSource
        Friend WithEvents GlobalLookupBindingSource As System.Windows.Forms.BindingSource
        Friend WithEvents nameColumn As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents valueColumn As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents typeColumn As System.Windows.Forms.DataGridViewComboBoxColumn
        Friend WithEvents GlobalContextMenuStrip As System.Windows.Forms.ContextMenuStrip
        Friend WithEvents AddGlobalToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents EditGlobalToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents DeleteGlobalToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
        Friend WithEvents noColumn As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents VariableNameColumn As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents VariableValueColumn As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents TaskNameColumn As System.Windows.Forms.DataGridViewComboBoxColumn
        Friend WithEvents taskIdColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    End Class
End Namespace

