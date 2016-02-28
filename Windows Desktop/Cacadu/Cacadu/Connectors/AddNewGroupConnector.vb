Imports Cacadu.UI

Namespace Connectors
    Friend Class AddNewGroupConnector
        Shared Sub ToggleCreateButtonEnable()
            With AddNewGroupForm
                If IsNothing(.NewGroupNameTextEdit.EditValue) Then
                    .AddGroupSimpleButton.Enabled = False
                Else
                    .AddGroupSimpleButton.Enabled = True
                End If
            End With
        End Sub
    End Class
End Namespace

