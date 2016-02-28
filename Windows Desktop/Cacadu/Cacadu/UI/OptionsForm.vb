Imports Cacadu.Connectors.OptionsFormConnector

Namespace UI
    Friend Class OptionsForm
        Private Sub OptionsForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
            LoadOptionForm()
        End Sub

        Private Sub AutoStartTriggersCheckEdit_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles AutoStartTriggersCheckEdit.CheckedChanged
            If AutoStartTriggersCheckEdit.Checked Then AskMeFirstBeforeResumingTriggersCheckEdit.Checked = False
        End Sub

        Private Sub AskMeFirstBeforeResumingTriggersCheckEdit_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles AskMeFirstBeforeResumingTriggersCheckEdit.CheckedChanged
            If AskMeFirstBeforeResumingTriggersCheckEdit.Checked Then AutoStartTriggersCheckEdit.Checked = False
        End Sub

        Private Sub OKSimpleButton_Click(sender As System.Object, e As System.EventArgs) Handles OKSimpleButton.Click
            SaveOptions()
            Me.Dispose()
        End Sub

        Private Sub CancelSimpleButton_Click(sender As System.Object, e As System.EventArgs) Handles CancelSimpleButton.Click
            Me.Dispose()
        End Sub
    End Class
End Namespace
