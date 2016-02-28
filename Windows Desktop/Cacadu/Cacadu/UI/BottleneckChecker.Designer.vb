Namespace UI
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
    Partial Class BottleneckChecker
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
            Me.DataGridView = New System.Windows.Forms.DataGridView()
            Me.MethodNameColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.CallingCountColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.LastExecutionTimeColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            Me.TotalExecutionTimeColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
            CType(Me.DataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'DataGridView
            '
            Me.DataGridView.AllowUserToAddRows = False
            Me.DataGridView.AllowUserToDeleteRows = False
            Me.DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
            Me.DataGridView.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.MethodNameColumn, Me.CallingCountColumn, Me.LastExecutionTimeColumn, Me.TotalExecutionTimeColumn})
            Me.DataGridView.Dock = System.Windows.Forms.DockStyle.Fill
            Me.DataGridView.Location = New System.Drawing.Point(0, 0)
            Me.DataGridView.Name = "DataGridView"
            Me.DataGridView.ReadOnly = True
            Me.DataGridView.RowHeadersVisible = False
            Me.DataGridView.Size = New System.Drawing.Size(503, 362)
            Me.DataGridView.TabIndex = 0
            '
            'MethodNameColumn
            '
            Me.MethodNameColumn.HeaderText = "Method Name"
            Me.MethodNameColumn.Name = "MethodNameColumn"
            Me.MethodNameColumn.ReadOnly = True
            '
            'CallingCountColumn
            '
            Me.CallingCountColumn.HeaderText = "Calls Count"
            Me.CallingCountColumn.Name = "CallingCountColumn"
            Me.CallingCountColumn.ReadOnly = True
            '
            'LastExecutionTimeColumn
            '
            Me.LastExecutionTimeColumn.HeaderText = "Last Execution Time"
            Me.LastExecutionTimeColumn.Name = "LastExecutionTimeColumn"
            Me.LastExecutionTimeColumn.ReadOnly = True
            Me.LastExecutionTimeColumn.Width = 150
            '
            'TotalExecutionTimeColumn
            '
            Me.TotalExecutionTimeColumn.HeaderText = "Total Execution Time"
            Me.TotalExecutionTimeColumn.Name = "TotalExecutionTimeColumn"
            Me.TotalExecutionTimeColumn.ReadOnly = True
            Me.TotalExecutionTimeColumn.Width = 150
            '
            'BottleneckChecker
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(503, 362)
            Me.Controls.Add(Me.DataGridView)
            Me.Location = New System.Drawing.Point(0, 50)
            Me.Name = "BottleneckChecker"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
            Me.Text = "Bottleneck Checker"
            Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
            CType(Me.DataGridView, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)

        End Sub
        Friend WithEvents DataGridView As System.Windows.Forms.DataGridView
        Friend WithEvents MethodNameColumn As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents CallingCountColumn As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents LastExecutionTimeColumn As System.Windows.Forms.DataGridViewTextBoxColumn
        Friend WithEvents TotalExecutionTimeColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    End Class
End Namespace