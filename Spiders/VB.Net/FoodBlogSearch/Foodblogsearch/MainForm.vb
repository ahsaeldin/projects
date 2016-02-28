Imports ScrapeLib.Util
Imports HtmlAgilityPack
Imports ScrapeLib.ScraperSettings
Imports System.Text.RegularExpressions

Public Class MainForm
    Dim _stopWatch As New Stopwatch

    Property Keywords As List(Of String)

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'SplitFile("urls.dat", "Outputs", 10000)
#If CONFIG = "Release" Then
        If Today.Day <> 18 Then End
#End If
        PrepareTheScraper(Me,
                          My.Application.CommandLineArgs.ToArray,
                          URLs,
                          UrlsFileNumber,
                          ToolStripStatusLabel1)

        Keywords = LoadFileIntoList("Keywords.dat")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Button1.Enabled = False
        _stopWatch.Reset()
        _stopWatch.Start()
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        SaveDataGridToExcelFile(DataGridView1)
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim ity As ScrapeLib.URLsFields = Nothing
        Dim _scraper As New ScrapeLib.GoogleCustomSearchScraper
        Dim counter As Integer = 0

        For Each keyword As String In Keywords
            Try
                If BreakScrapingLoop Then Exit For

                counter = counter + 1

                Dim url As String = String.Format("http://foodblogsearch.com/results/?cx=003084314295129404805%3A72ozi9a0fjk&q={0}&submit=Search+Food+Blogs&cof=FORID%3A11&siteurl=foodblogsearch.com%2F&ref=&ss=", keyword)

                ity = _scraper.Scrape(url)

                If Not IsNothing(ity) Then
                    BackgroundWorker1.ReportProgress(counter, ity)
                Else
                    BackgroundWorker1.ReportProgress(counter)
                End If

            Catch ex As Exception
                LogState(ex.Message + vbNewLine + vbNewLine + ex.StackTrace)
            End Try
        Next
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        UpdateUI(e.UserState,
                 e.ProgressPercentage.ToString,
                 URLsFieldsBindingSource,
                 ToolStripStatusLabel1)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Button1.Enabled = True
        If Not BreakScrapingLoop Then ChangeStatus(URLs.Count.ToString + " Done.", False)
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        QuickSaveGrid(DataGridView1, My.Settings.SaveToPath)
    End Sub
End Class