Imports ScrapeLib.Util
Imports ScrapeLib.ScraperSettings
Imports System.Text.RegularExpressions
Imports System.Net

Public Class MainForm
#Region "Properties"
    Property IsThereNext As Boolean = False
    Property ProductsURLs As New List(Of String)
    Property ScrapingStopWatcher As New Stopwatch
    Public Property IsDocumentCompleted As Boolean = False
#End Region

#Region "Controls Events Handlers"
    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Try
            PrepareTheScraper(Me,
                              My.Application.CommandLineArgs.ToArray,
                              URLs,
                              UrlsFileNumber,
                              MainFormToolStripStatusLabel)

        Catch ex As Exception
            MsgBox("MainForm_Shown::" + ex.Message)
        End Try
    End Sub

    Private Sub ScrapeButton_Click(sender As Object, e As EventArgs) Handles ScrapeButton.Click
        StartScraping()
    End Sub

    Function CreateWebResponse(url As String) As HttpWebResponse
        Dim myRequest As HttpWebRequest = CType(HttpWebRequest.Create(url), HttpWebRequest)

        Dim _proxy As New WebProxy("173.245.220.233:17668")

        myRequest.Proxy = _proxy

        Dim myResponse As HttpWebResponse = DirectCast(myRequest.GetResponse(), HttpWebResponse)

        Return myResponse
    End Function

    Private Sub ScrapingBackgroundWorker_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles ScrapingBackgroundWorker.DoWork
        Dim counter As Integer = 0
        Dim ity As ScrapeLib.OfficeworkFields = Nothing
        Dim _scraper As New ScrapeLib.OfficeworksScraper

        CheckScrapingDelay()

        IsThereNext = False
        IsDocumentCompleted = False
        For Each URL As String In URLs
            Try
                ProductsURLs = New List(Of String)

                If BreakScrapingLoop Then BreakScrapingLoop = False : Exit For

                ScrapingBackgroundWorker.ReportProgress(0, "Collecting products URLs...")

                Me.BeginInvoke(Sub()
                                   'Me.WebBrowser1.DocumentStream = CreateWebResponse(URL).GetResponseStream
                                   Me.WebBrowser1.Navigate(URL)
                               End Sub)

                Do Until IsDocumentCompleted
                    Application.DoEvents()
                Loop
                CheckScrapingDelay()

                'Set PageSize 50:-=
                IsDocumentCompleted = False
                Me.BeginInvoke(Sub()
                                   'Me.WebBrowser1.DocumentStream = CreateWebResponse("javascript:setPageSize(50);").GetResponseStream
                                   Me.WebBrowser1.Navigate("javascript:setPageSize(50);")
                               End Sub)
                Do Until IsDocumentCompleted
                    Application.DoEvents()
                Loop
                CheckScrapingDelay()

                Dim index As Integer = 1
                Do While IsThereNext
                    IsThereNext = False
                    IsDocumentCompleted = False
                    index = index + 1
                    Dim NextURL As String = String.Format("javascript:setPage({0})", index.ToString)
                    Me.BeginInvoke(Sub()
                                       'Me.WebBrowser1.DocumentStream = CreateWebResponse(NextURL).GetResponseStream
                                       Me.WebBrowser1.Navigate(NextURL)
                                   End Sub)
                    Do Until IsDocumentCompleted
                        Application.DoEvents()
                    Loop
                    CheckScrapingDelay()
                Loop

                For Each prodURL As String In ProductsURLs
                    My.Computer.FileSystem.WriteAllText("Prod_URLs.txt", prodURL.Trim + vbNewLine, True)
                Next

                Dim fileContents As String()
                fileContents = My.Computer.FileSystem.ReadAllText("Prod_URLs.txt").Split(vbNewLine)
                ProductsURLs.AddRange(fileContents)

                counter = counter + 1
                ScrapingBackgroundWorker.ReportProgress(0, "Scraping products pages...")
                For Each prodURL In ProductsURLs
                    ity = _scraper.Scrape(prodURL.Trim)

                    If Not IsNothing(ity) Then
                        ScrapingBackgroundWorker.ReportProgress(counter, ity)
                    Else
                        My.Computer.FileSystem.WriteAllText("FailedLinks.txt", prodURL.Trim + vbNewLine, True)
                        ScrapingBackgroundWorker.ReportProgress(counter)
                    End If
                    CheckScrapingDelay()
                Next

            Catch ex As Exception
                My.Computer.FileSystem.WriteAllText("FailedLinks.txt", URL.Trim + vbNewLine, True)
                LogState(ex.Message + vbNewLine + vbNewLine + ex.StackTrace)
            End Try
        Next
    End Sub

    Private Sub WebBrowser1_Navigated(sender As Object, e As WebBrowserNavigatedEventArgs) Handles WebBrowser1.Navigated
        If WebBrowser1.ReadyState = WebBrowserReadyState.Complete Then
            Dim pageSource As String = WebBrowser1.DocumentText.Replace(vbCr, "")
            pageSource = pageSource.Replace(vbLf, "")
            pageSource = pageSource.Replace(vbTab, "")
            Dim regexy As New Regex("<li\sclass=%title\-v1%>.+?<a\shref=%(.+?)%.+?</li>".Replace("%", Chr(34)))
            Dim matchzed As MatchCollection = regexy.Matches(pageSource)
            For Each matcy As Match In matchzed
                Dim link As String = "http://www.officeworks.com.au" + matcy.Groups(1).Value.Trim
                If Not ProductsURLs.Contains(link) Then ProductsURLs.Add(link)
            Next
            Dim nextURLMatcher As New Regex("href=%(javascript:setPage\(\d+\))%.+?>Next<".Replace("%", Chr(34)))
            Dim matchedURLs As MatchCollection = nextURLMatcher.Matches(pageSource)
            If matchedURLs.Count > 0 Then
                IsThereNext = True
            Else
                IsThereNext = False
            End If
            IsDocumentCompleted = True
        End If
    End Sub

    Private Sub ScrapingBackgroundWorker_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles ScrapingBackgroundWorker.ProgressChanged
        If e.ProgressPercentage = 0 Then
            ChangeStatus(e.UserState)
        Else
            UpdateUI(e.UserState,
                     e.ProgressPercentage.ToString,
                     OfficeworkFieldsBindingSource,
                     MainFormToolStripStatusLabel)
            If Not IsNothing(e.UserState) Then ScrollGridToEnd(ScrapingDataGridView)
        End If
    End Sub

    Private Sub ScrapingBackgroundWorker_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles ScrapingBackgroundWorker.RunWorkerCompleted
        ScrapingStopWatcher.Stop()
        ScrapeButton.Enabled = True
        LastScrapingSessionPeriod = ScrapingStopWatcher.Elapsed
        ChangeStatus(URLs.Count.ToString + " Records Scraped in " + LastScrapingSessionPeriod.ToString("hh\:mm\:ss"), False)
        ScrapedToolStripStatusLabel.Text = "Scraped#" + ScrapingDataGridView.RowCount.ToString
        If URLs.Count <> ScrapingDataGridView.RowCount Then FailedToolStripStatusLabel.Text = "Failed#" + Str(URLs.Count - ScrapingDataGridView.RowCount)
    End Sub

    Private Sub SavingTimer_Tick(sender As Object, e As EventArgs) Handles SavingTimer.Tick
        QuickSaveGrid(ScrapingDataGridView, My.Settings.SaveToPath)
    End Sub

    Private Sub SaveButton_Click(sender As Object, e As EventArgs) Handles SaveButton.Click
        SaveDataGridToExcelFile(ScrapingDataGridView)
    End Sub

    Private Sub StopButton_Click(sender As Object, e As EventArgs) Handles StopButton.Click
        If ScrapingBackgroundWorker.IsBusy Then BreakScrapingLoop = True
    End Sub

    Private Sub DeleteBackupsCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles DeleteBackupsCheckBox.CheckedChanged
        DeleteBackupsBeforeStart = DeleteBackupsCheckBox.Checked
    End Sub
#End Region

#Region "Helpers"
    Sub CheckBackupDirectory()
        If DeleteBackupsBeforeStart AndAlso IO.Directory.Exists(My.Settings.SaveToPath) Then IO.Directory.Delete(My.Settings.SaveToPath, True)
        If Not IO.Directory.Exists(My.Settings.SaveToPath) Then IO.Directory.CreateDirectory(My.Settings.SaveToPath)
    End Sub

    Sub StartScraping()
        CheckBackupDirectory()
        ScrapingStopWatcher.Reset()
        ScrapingStopWatcher.Start()
        BreakScrapingLoop = False
        ScrapeButton.Enabled = False
        ScrapeFromSourceButton.Enabled = False
        LastScrapingSessionPeriod = New TimeSpan
        UseProxies = UseProxiesCheckBox.Checked
        ReScrapeFailedURLs = ReScrapeFailedCheckBox.Checked
        If DelayCheckBox.Checked Then Delay = DelayNumericUpDown.Value
        If IO.File.Exists(FailedURLsFileName) Then IO.File.Delete(FailedURLsFileName)
        ScrapingBackgroundWorker.RunWorkerAsync()
    End Sub
#End Region
End Class