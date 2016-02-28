Friend Class StopRecordingForm
    Private Sub ButStopRecording_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButStopRecording.Click
        '//changed next if clause to use the form tag rather than the button text in order to exclude the button.
        'If ButStopRecording.Text = "&Stop Recording" Then
        '    MainForm.StopMacro()
        '    MainForm.ShowMe()
        'ElseIf ButStopRecording.Text = "&Resume Record" Then
        '    MainForm.ResRec()
        'End If
        If Me.Tag = "&Stop Recording" Or Me.Tag = "" Then
            MainForm.StopMacro()
            MainForm.ShowMe()
        ElseIf Me.Tag = "&Resume Record" Then
            MainForm.ResRec()
        End If
    End Sub
End Class