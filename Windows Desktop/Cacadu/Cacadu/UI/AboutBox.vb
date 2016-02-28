Imports Cacadu.IntelliProtectorService.IntelliProtector

Namespace UI
    Friend NotInheritable Class AboutForm
        Private Sub AboutBox_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
            ' Set the title of the form.
            Dim ApplicationTitle As String
            If My.Application.Info.Title <> "" Then
                ApplicationTitle = My.Application.Info.Title
            Else
                ApplicationTitle = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
            End If
            Me.Text = String.Format("About {0}", ApplicationTitle)
            ' Initialize all of the text displayed on the About Box.
            ' TODO: Customize the application's assembly information in the "Application" pane of the project 
            '    properties dialog (under the "Project" menu).
            Me.LabelProductName.Text = My.Application.Info.ProductName
            Me.LabelVersion.Text = String.Format("Version {0}", My.Application.Info.Version.ToString)
            Me.LabelCopyright.Text = My.Application.Info.Copyright
            Me.LabelCompanyName.Text = "Conderella"
            Dim customerName As String = IntelliprotectorHelper.GetCustomerName
            Me.RegisteredToLabel.Text = "Registered To: " & customerName
        End Sub

        Private Sub AboutBox_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
            Me.Dispose()
        End Sub
    End Class
End Namespace

