Imports Cacadu.Connectors.AddNewGroupConnector

Namespace UI
    Friend Class AddNewGroupForm
        Private Sub AddNewGroupForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
            SetIconToCacaduIcon(Me)
        End Sub

        Private Sub NewGroupNameTextEdit_EditValueChanged(sender As Object, e As System.EventArgs) Handles NewGroupNameTextEdit.EditValueChanged
            ToggleCreateButtonEnable()
        End Sub

        Private Sub AddGroupSimpleButton_Click(sender As System.Object, e As System.EventArgs) Handles AddGroupSimpleButton.Click
            Connectors.MainFormConnector.AddNewGroup(NewGroupNameTextEdit.EditValue.ToString)
        End Sub

        Private Sub AddNewGroupForm_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
            Me.Dispose()
        End Sub
    End Class
End Namespace
