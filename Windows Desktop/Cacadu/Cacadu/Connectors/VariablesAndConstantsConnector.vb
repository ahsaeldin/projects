Imports Cacadore
Imports Cacadu.UI
Imports TecDAL

Namespace Connectors
    Friend Class VariablesAndConstantsConnector
        Shared Sub BindDataGridToTaskVariables()
            With VariablesAndConstantsForm.Task_variablesBindingSource
                .DataSource = TecDAL.TypedDataSetCrud.TectonicDataSetObj
                .DataMember = "task_variables"
            End With
        End Sub

        Shared Sub BindDataGridToGlobalConstant()
            With VariablesAndConstantsForm.GlobalDataGridView
                .DataSource = TecDAL.TypedDataSetCrud.TectonicDataSetObj
                .DataMember = "globals"
            End With
        End Sub

        Shared Sub BindDataGridColumnCellToGlobals()
            With VariablesAndConstantsForm
                .GlobalLookupBindingSource.DataSource = TecDAL.TypedDataSetCrud.TectonicDataSetObj
                .GlobalLookupBindingSource.DataMember = "global_lookup"
                .typeColumn.ValueMember = "no"
                .typeColumn.DisplayMember = "name"
            End With
        End Sub

        Shared Sub BindDataGridColumnCellToTasks()
            With VariablesAndConstantsForm
                .TasksBindingSource.DataSource = TecDAL.TypedDataSetCrud.TectonicDataSetObj
                .TasksBindingSource.DataMember = "tasks"
                .TaskNameColumn.ValueMember = "id"
                .TaskNameColumn.DisplayMember = "Name"
            End With
        End Sub

        Shared Sub MakeSureNoColumnIsNotVisible()
            VariablesAndConstantsForm.VariablesDataGridView.Columns("noColumn").Visible = False
        End Sub

        Shared Function ValidateAndUpdateVariables() As Boolean
            With VariablesAndConstantsForm
                Dim isValidated As Boolean = .Validate()
                If isValidated Then
                    .Task_variablesBindingSource.EndEdit()
                    Dim asdfsafasd = TecDAL.TypedDataSetCrud.Connection
                    'AhSaElDin: 3-10-2012: Added next line as a fix for this bug 
                    'https://conderella.com/wi/client/index.php?issue=464
                    .TableAdapterManager.Connection = TecDAL.TypedDataSetCrud.Connection
                    .TableAdapterManager.UpdateAll(TecDAL.TypedDataSetCrud.TectonicDataSetObj)
                End If
                Return isValidated
            End With
        End Function

        Shared Sub DeleteSelecetdRows(grid As DataGridView)
            For Each selRow As DataGridViewRow In grid.SelectedRows
                grid.Rows.Remove(selRow)
            Next
            ValidateAndUpdateVariables()
        End Sub

        Shared Sub SetSelectedTabPage(tabPageIndex As Integer)
            VariablesAndConstantsForm.VarsAndConsXtraTabControl.SelectedTabPageIndex = tabPageIndex
        End Sub

        Shared Sub UpdateStatusBar(state As String, color As Color)
            VariablesAndConstantsForm.StatusToolStripStatusLabel.Text = state
            VariablesAndConstantsForm.StatusToolStripStatusLabel.ForeColor = color
        End Sub

        Shared Sub ValidateTaskNameCombo(sender As Object, e As System.Windows.Forms.DataGridViewCellValidatingEventArgs)
            Dim varGridView As DataGridView = CType(sender, DataGridView)
            Dim variableNameCell As String = varGridView.Rows(e.RowIndex).Cells(1).Value.ToString
            Dim newTaskIdCell As String = Commands.GetTaskIdByName(e.FormattedValue.ToString)
            Dim oldTaskIdCell As String = varGridView.Rows(e.RowIndex).Cells(4).Value.ToString
            Dim noChange As Boolean = Strings.Equals(newTaskIdCell, oldTaskIdCell)

            'If newTaskIdCell <> "" And Not noChange AndAlso variableNameCell <> "" AndAlso oldTaskIdCell <> "" Then
            If newTaskIdCell <> "" And Not noChange AndAlso variableNameCell <> "" Then
                Dim validationResult2 = ValidateVariableNameCellAgainstNotation(variableNameCell)
                Dim validationResult As Boolean = ValidateVariableNameCellAgainstTaskVariables(newTaskIdCell, variableNameCell)
                If Not validationResult OrElse Not validationResult2 Then
                    e.Cancel = True
                    Exit Sub
                End If
            End If

            If variableNameCell <> "" Then
                varGridView.Rows(e.RowIndex).Cells(1).Value = variableNameCell
                UpdateStatusBar("Variables Updated.", Color.Blue)
            End If

        End Sub

        Shared Function ValidateVariableNameCell(sender As Object, e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) As Boolean
            Dim varGridView As DataGridView = CType(sender, DataGridView)
            Dim newVariableNameCell As String = e.FormattedValue.ToString

            If IsDBNull(varGridView.Rows(e.RowIndex).Cells(3).Value) Then Return True

            Dim taskId As String = varGridView.Rows(e.RowIndex).Cells(4).Value.ToString
            Dim taskName As String = varGridView.Rows(e.RowIndex).Cells(3).Value.ToString
            Dim oldVariableNameCell As String = varGridView.Rows(e.RowIndex).Cells(1).Value.ToString
            Dim noChange As Boolean = Strings.Equals(newVariableNameCell, oldVariableNameCell)

            'If newVariableNameCell <> "" And Not noChange And oldVariableNameCell <> "" Then
            If newVariableNameCell <> "" And Not noChange Then

                If Not ValidateVariableNameCellAgainstNotation(newVariableNameCell) OrElse
                   Not ValidateVariableNameCellAgainstTaskVariables(taskId, newVariableNameCell) Then

                    e.Cancel = True
                    Return False

                End If

                If newVariableNameCell <> "" Then
                    'Why=>http://stackoverflow.com/questions/4742960/datagridview-validation-changing-cell-value
                    varGridView.EditingControl.Text = newVariableNameCell
                    ValidateAndUpdateVariables()
                    UpdateStatusBar("Variables Updated.", Color.Blue)
                End If
            End If
            Return True
        End Function

        Shared Function ValidateVariableNameCellAgainstNotation(ByRef variableNameCell As String) As Boolean
            Dim isVarHasNotation As Boolean
            Dim errorMessage As String = vbNullString
            isVarHasNotation = TaskVariables.AddVariableNotation(variableNameCell, errorMessage)
            If Not isVarHasNotation Then
                UpdateStatusBar(errorMessage, Color.Red)
                Return False
            End If
            Return True
        End Function

        Shared Function ValidateVariableNameCellAgainstTaskVariables(taskId As String, variableNameCell As String) As Boolean
            Dim isRegisteredVariable As Boolean = Commands.IsRegisteredTaskVariable(taskId, variableNameCell)
            UpdateStatusBar("A variable with the same is already registered for this task.", Color.Red)
            Return Not isRegisteredVariable
        End Function

        Shared Function TryAddNotation(ByRef globalMemberName As String, globalMemberType As Integer, ByRef errorMessage As String) As Boolean
            Dim addingNotationResult As Boolean = False
            Select Case globalMemberType
                Case IGlobals.GlobalType.Constant
                    addingNotationResult = Globals.AddGlobalConstantNotation(globalMemberName, errorMessage)
                Case IGlobals.GlobalType.Variable
                    addingNotationResult = Globals.AddGlobalVariableNotation(globalMemberName, errorMessage)
            End Select
            Return addingNotationResult
        End Function

        Shared Function CheckIfValidationFailed(addingNotationResult As Boolean,
                                                e As System.Windows.Forms.DataGridViewCellValidatingEventArgs,
                                                errorMessage As String) As Boolean
            If Not addingNotationResult Then
                UpdateStatusBar(errorMessage, Color.Red)
                e.Cancel = True
                Return False
            End If
            Return True
        End Function

        Shared Sub ValidateGlobalNameCell(sender As Object, e As System.Windows.Forms.DataGridViewCellValidatingEventArgs)
            Dim globalGridView As DataGridView = CType(sender, DataGridView)
            If globalGridView.IsCurrentCellDirty Then

                If e.ColumnIndex = 0 Then

                    Dim errorMessage As String = vbNullString
                    Dim addingNotationResult As Boolean = False
                    Dim globalMemberName As String = e.FormattedValue.ToString

                    'If adding a new row then exit sub
                    If IsDBNull(globalGridView.Rows(e.RowIndex).Cells(2).Value) Then Exit Sub

                    Dim globalMemberType As Integer = CInt(globalGridView.Rows(e.RowIndex).Cells(2).Value)

                    addingNotationResult = TryAddNotation(globalMemberName, globalMemberType, errorMessage)

                    If Not CheckIfValidationFailed(addingNotationResult, e, errorMessage) Then Exit Sub

                    globalGridView.EditingControl.Text = globalMemberName
                    UpdateStatusBar("Global members updated.", Color.Blue)

                End If

            End If
        End Sub

        Shared Sub ValidateGlobalTypeComboCell(sender As Object, e As System.Windows.Forms.DataGridViewCellValidatingEventArgs)
            Dim globalGridView As DataGridView = CType(sender, DataGridView)
            If globalGridView.IsCurrentCellDirty Then

                If e.ColumnIndex = 2 Then

                    Dim globalMemberType As Integer
                    Dim errorMessage As String = vbNullString
                    Dim addingNotationResult As Boolean = False

                    If e.FormattedValue.ToString = "Variable" Then
                        globalMemberType = 1
                    ElseIf e.FormattedValue.ToString = "Constant" Then
                        globalMemberType = 0
                    End If

                    Dim globalMemberName As String = globalGridView.Rows(e.RowIndex).Cells(0).Value.ToString

                    'If adding a new row then exit sub
                    If globalMemberName = "" Then Exit Sub

                    addingNotationResult = TryAddNotation(globalMemberName, globalMemberType, errorMessage)

                    If Not CheckIfValidationFailed(addingNotationResult, e, errorMessage) Then Exit Sub

                    globalGridView.Rows(e.RowIndex).Cells(0).Value = globalMemberName
                    UpdateStatusBar("Global members updated.", Color.Blue)
                End If

            End If
        End Sub

        Shared Sub FocusBlankRow(grid As DataGridView)
            grid.CurrentCell = grid(0, grid.Rows.Count - 1)
        End Sub
    End Class
End Namespace

