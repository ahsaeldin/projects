Imports System.Windows.Forms

Friend Class FrmDelay
    Private Sub txtDelayValue_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDelayValue.Validating
        validateMyTxtBoxes(txtDelayValue)
    End Sub

    Private Sub txtDelayValue_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDelayValue.TextChanged
        If Len(txtDelayValue.Text) >= 8 Then
            txtDelayValue.Text = Strings.Left(txtDelayValue.Text, 7)
            txtDelayValue.SelectionStart = Len(txtDelayValue.Text)
        End If
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class