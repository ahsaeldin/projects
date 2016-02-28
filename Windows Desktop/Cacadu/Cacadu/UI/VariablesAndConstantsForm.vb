Imports Cacadore
Imports Cacadu.Connectors
Imports Cacadu.Connectors.VariablesAndConstantsConnector

Namespace UI
    Friend Class VariablesAndConstantsForm

#Region "Properties"
        Public Property IsFormLoaded As Boolean = False
#End Region

#Region "Form Events"
        Private Sub VariablesAndConstantsForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
            SetIconToCacaduIcon(Me)
            BindDataGridToTaskVariables()
            BindDataGridColumnCellToTasks()
            BindDataGridToGlobalConstant()
            BindDataGridColumnCellToGlobals()
        End Sub

        Private Sub VariablesAndConstantsForm_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
            IsFormLoaded = True
            MakeSureNoColumnIsNotVisible()
        End Sub

        Private Sub VariablesAndConstantsForm_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
            Dim validationResult As Boolean = False
            validationResult = ValidateAndUpdateVariables()
            e.Cancel = Not validationResult
            If Not validationResult Then SetSelectedTabPage(0)
        End Sub

        Private Sub VariablesAndConstantsForm_KeyDown(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
            If e.KeyCode = Keys.Escape Then Me.Close()
        End Sub
#End Region

#Region "Context Menu Strip"
        Private Sub EditToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EditToolStripMenuItem.Click
            VariablesDataGridView.BeginEdit(True)
        End Sub

        Private Sub DeleteToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
            DeleteSelecetdRows(VariablesDataGridView)
        End Sub

        Private Sub AddToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AddToolStripMenuItem.Click
            FocusBlankRow(VariablesDataGridView)
        End Sub

        Private Sub AddGlobalToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AddGlobalToolStripMenuItem.Click
            FocusBlankRow(GlobalDataGridView)
        End Sub

        Private Sub EditGlobalToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EditGlobalToolStripMenuItem.Click
            GlobalDataGridView.BeginEdit(True)
        End Sub

        Private Sub DeleteGlobalToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DeleteGlobalToolStripMenuItem.Click
            DeleteSelecetdRows(GlobalDataGridView)
        End Sub
#End Region

#Region "The two DataGridView"
        Private Sub VariablesDataGridView_CellValidating(sender As Object, e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles VariablesDataGridView.CellValidating
            If VariablesDataGridView.IsCurrentCellDirty AndAlso IsFormLoaded And Not VariablesDataGridView.CurrentRow.IsNewRow Then
                If e.ColumnIndex = 1 Then
                    ValidateVariableNameCell(sender, e)
                ElseIf e.ColumnIndex = 3 Then
                    ValidateTaskNameCombo(sender, e)
                End If
            End If
        End Sub

        Private Sub VariablesDataGridView_DataError(sender As System.Object, e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles VariablesDataGridView.DataError
#If DEBUG Then
            Balora.Alerter.REP(e.Exception.Message, e.Exception, True, True)
#Else
            Balora.Alerter.REP(e.Exception.Message, e.Exception, False, True)
#End If
        End Sub

        Private Sub GlobalDataGridView_CellValidating(sender As Object, e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles GlobalDataGridView.CellValidating
            ValidateGlobalNameCell(sender, e)
            ValidateGlobalTypeComboCell(sender, e)
        End Sub

        Private Sub GlobalConstantDataGridView_DataError(sender As System.Object, e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles GlobalDataGridView.DataError
#If DEBUG Then
            Balora.Alerter.REP(e.Exception.Message, e.Exception, True, True)
#Else
            Balora.Alerter.REP(e.Exception.Message, e.Exception, False, True)
#End If
        End Sub

        Private Sub VariablesDataGridView_RowsRemoved(sender As Object, e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles VariablesDataGridView.RowsRemoved
            If IsFormLoaded Then ValidateAndUpdateVariables()
        End Sub
#End Region

#Region "XtraTab Control"
        Private Sub VarsAndConsXtraTabControl_SelectedPageChanging(sender As System.Object, e As DevExpress.XtraTab.TabPageChangingEventArgs) Handles VarsAndConsXtraTabControl.SelectedPageChanging
            If IsFormLoaded Then UpdateStatusBar("", Color.Black)
        End Sub
#End Region
    End Class
End Namespace