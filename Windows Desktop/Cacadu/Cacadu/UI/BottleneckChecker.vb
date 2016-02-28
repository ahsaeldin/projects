Namespace UI
    Friend Class BottleneckChecker
        Shared Property IsShown As Boolean = False

        Private Sub BottleneckChecker_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
            IsShown = True
        End Sub

        Private Sub BottleneckChecker_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
            IsShown = False
        End Sub
    End Class
End Namespace
