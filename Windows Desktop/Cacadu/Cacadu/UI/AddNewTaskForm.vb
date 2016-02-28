Imports Cacadu.Connectors
Namespace UI
    Friend Class AddNewTaskForm
        Private Sub AddNewTask_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
            AddNewTaskFormConnector.OnFormLoad()
        End Sub

        Private Sub OkButtonOfAddNewTask_Click(sender As System.Object, e As System.EventArgs) Handles OkButtonOfAddNewTask.Click
            AddNewTaskFormConnector.OnFormOK(Me.WaitBetweenActionsCheckEdit.Checked,
                                             Me.AddToGroupComboBoxEdit.Text,
                                             Me.TaskNameTextEdit.Text,
                                             Me.TaskEnabledCheckEdit.Checked)
            Me.Dispose()
        End Sub

        Private Sub TaskNameTextEdit_KeyDown(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles TaskNameTextEdit.KeyDown
            If e.Control And e.KeyData = Keys.A Then TaskNameTextEdit.SelectAll()
        End Sub

        Private Sub AddToGroupComboBoxEdit_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles AddToGroupComboBoxEdit.SelectedIndexChanged
            If AddToGroupComboBoxEdit.SelectedIndex = -1 Then
                OkButtonOfAddNewTask.Enabled = False
            Else
                OkButtonOfAddNewTask.Enabled = True
            End If
        End Sub

        Private Sub CancelButtonOfAddNewTask_Click(sender As System.Object, e As System.EventArgs) Handles CancelButtonOfAddNewTask.Click
            Me.Dispose()
        End Sub
    End Class
End Namespace
