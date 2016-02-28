Public Class MagentoImporter
    Private Sub MagentoImporter_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        MagentoImporterWebBrowser.Navigate("http://1001deals.com/magmi/web/magmi_run.php?mode=update&profile=dropy&engine=magmi_productimportengine:Magmi_ProductImportEngine")
    End Sub

    Private Sub MagentoImporterWebBrowser_DocumentCompleted(sender As Object, e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles MagentoImporterWebBrowser.DocumentCompleted
        If MagentoImporterWebBrowser.ReadyState = WebBrowserReadyState.Complete Then
            Dim fileName As String = DateTime.Now.ToString("G").Replace(":", "-").Replace("/", "-")
            Dim filePath As String = GetAppFolderPath() + "Magento-Import [" + fileName + "]" + ".html"
            My.Computer.FileSystem.WriteAllText(filePath, MagentoImporterWebBrowser.DocumentText, False)
            ChangeStatus("Importing to Magento done, check report at::" + filePath)
            With MainForm
                .MagentoOutputBindingSource.Clear()
            End With
        End If
    End Sub
End Class