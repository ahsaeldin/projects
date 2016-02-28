Imports ScrapeLib.Util
Imports ScrapeLib.ScraperSettings
Imports System.Text.RegularExpressions

Public Class MainForm
#Region "Properties"
    Property ProcessingStopWatcher As New Stopwatch
#End Region

#Region "Controls Events Handlers"
    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Try
            PrepareTheApplication(Me,
                              My.Application.CommandLineArgs.ToArray,
                              URLs,
                              Keywords,
                              UrlsFileNumber,
                              MainFormToolStripStatusLabel)

        Catch ex As Exception
            MsgBox("MainForm_Shown::" + ex.Message)
        End Try
    End Sub

    Private Sub ScrapeButton_Click(sender As Object, e As EventArgs) Handles ScrapeButton.Click
        StartScraping()
    End Sub

    Private Sub CrawlButton_Click(sender As Object, e As EventArgs) Handles CrawlButton.Click
        StartCrawling()
    End Sub

#Region "Scraping"
    Private Sub ScrapingBackgroundWorker_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles ScrapingBackgroundWorker.DoWork
        Dim counter As Integer = 0
        Dim ity As ScrapeLib.HouzzFields = Nothing
        Dim _scraper As New ScrapeLib.HouzzScraper

        For Each URL As String In URLs
            Try
                counter = counter + 1

                If BreakScrapingLoop Then BreakScrapingLoop = False : Exit For

                ity = _scraper.ScrapeAdvert(URL)

                If Not IsNothing(ity) And ity.Name <> vbNullString Then
                    ScrapingBackgroundWorker.ReportProgress(counter, ity)
                Else
                    My.Computer.FileSystem.WriteAllText(FailedURLsFileName, URL.Trim + vbNewLine, True)
                    ScrapingBackgroundWorker.ReportProgress(counter)
                End If

                CheckScrapingDelay()

            Catch ex As Exception
                My.Computer.FileSystem.WriteAllText(FailedURLsFileName, URL.Trim + vbNewLine, True)
                LogState(ex.Message + vbNewLine + vbNewLine + ex.StackTrace)
            End Try
        Next
    End Sub

    Private Sub ScrapingBackgroundWorker_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles ScrapingBackgroundWorker.ProgressChanged
        UpdateUI(e.UserState,
                 e.ProgressPercentage.ToString,
                 HouzzFieldsBindingSource,
                 MainFormToolStripStatusLabel)
        ScrollGridToEnd(ScrapingDataGridView)
    End Sub

    Private Sub ScrapingBackgroundWorker_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles ScrapingBackgroundWorker.RunWorkerCompleted
        InComplete()
        ChangeStatus(URLs.Count.ToString + " Records Scraped in " + LastProcessingSessionPeriod.ToString("hh\:mm\:ss"), False)
        ProcessedToolStripStatusLabel.Text = "Scraped#" + ScrapingDataGridView.RowCount.ToString
        If URLs.Count <> ScrapingDataGridView.RowCount Then FailedToolStripStatusLabel.Text = "Failed#" + Str(URLs.Count - ScrapingDataGridView.RowCount)
    End Sub
#End Region

#Region "Crawling"
    Function ExtractLinks(docy As HtmlAgilityPack.HtmlDocument) As List(Of ScrapeLib.URLsFields)
        Dim itys As New List(Of ScrapeLib.URLsFields)
        Dim advertsNodes = docy.DocumentNode.SelectNodes("//a").Where(Function(x) x.GetAttributeValue("class", "").Contains("pro-title") AndAlso Not x.GetAttributeValue("href", "").Contains("javascript")).Select(Function(y) y.GetAttributeValue("href", ""))
        For Each node In advertsNodes
            itys.Add(New ScrapeLib.URLsFields With {.ListingURL = node})
        Next
        Return itys
    End Function

    Sub UpdateGrid(ByRef itys As List(Of ScrapeLib.URLsFields),
                   pageDoc As HtmlAgilityPack.HtmlDocument,
                   counter As Integer,
                   pageURL As String)
        itys = ExtractLinks(pageDoc)
        If Not IsNothing(itys) AndAlso itys.Count > 0 Then
            CrawlingBackgroundWorker.ReportProgress(counter, itys.ToArray)
        Else
            My.Computer.FileSystem.WriteAllText("FailedCrawlingPages.txt", pageURL + vbNewLine, True)
        End If
    End Sub

    Sub asdfasdf()

    End Sub

    Private Sub CrawlingBackgroundWorker_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles CrawlingBackgroundWorker.DoWork
        Dim _scraper As New ScrapeLib.Scraper
        Dim itys As List(Of ScrapeLib.URLsFields) = Nothing

        Dim lst As List(Of String) = ScrapeLib.Util.LoadFileIntoList("areas.txt")

        Dim counter As Integer = 0
        For Each area As String In lst
            Try
                counter = counter + 1
                itys = New List(Of ScrapeLib.URLsFields)
                Dim pageDoc As HtmlAgilityPack.HtmlDocument = Nothing
                Dim pageSource As String = _scraper.DownloadSource(area)
                If Not IsNothing(pageSource) Then pageDoc = CreateHtmlDocument(pageSource)

                If pageSource = vbNullString OrElse IsNothing(pageDoc) Then
                    My.Computer.FileSystem.WriteAllText("FailedCrawlingPages.txt", area + vbNewLine, True)
                    Continue For
                End If

                UpdateGrid(itys, pageDoc, counter, area)
                If pageSource.Contains("Next Page") Then
                    Do While True
                        Dim nextNodes = pageDoc.DocumentNode.SelectNodes("//a").Where(Function(x) x.InnerText.Contains("Next Page"))
                        If IsNothing(nextNodes) OrElse nextNodes.Count = 0 Then
                            Exit Do
                        Else
                            Dim nextURL = nextNodes.First.GetAttributeValue("href", "")
                            pageDoc = Nothing
                            pageSource = vbNullString
                            pageSource = _scraper.DownloadSource(nextURL)
                            If Not IsNothing(pageSource) Then pageDoc = CreateHtmlDocument(pageSource)

                            If pageSource = vbNullString OrElse IsNothing(pageDoc) Then
                                My.Computer.FileSystem.WriteAllText("FailedCrawlingPages.txt", nextURL + vbNewLine, True)
                                Exit Do
                            End If

                            itys = New List(Of ScrapeLib.URLsFields)
                            UpdateGrid(itys, pageDoc, counter, nextURL)
                        End If
                    Loop
                End If
            Catch ex As Exception
                My.Computer.FileSystem.WriteAllText("FailedCrawlingPages.txt", area + vbNewLine, True)
                LogState(ex.Message + vbNewLine + vbNewLine + ex.StackTrace)
            End Try
        Next

        'Dim pageStep As Integer = 0
        'Dim targetURL As String = vbNullString
        'For i As Integer = 1 To 333
        '    Try
        '        itys = New List(Of ScrapeLib.URLsFields)

        '        If pageStep = 0 Then
        '            targetURL = "http://www.houzz.com/professionals/s/designer/c"
        '        Else
        '            pageStep = pageStep + 15
        '            targetURL = "http://www.houzz.com/professionals/s/designer/p/" + pageStep
        '        End If

        '        Dim pageDoc As HtmlAgilityPack.HtmlDocument = Nothing
        '        Dim pageSource As String = _scraper.DownloadSource(targetURL)
        '        If Not IsNothing(pageSource) Then pageDoc = CreateHtmlDocument(pageSource)

        '        If pageSource = vbNullString OrElse IsNothing(pageDoc) Then
        '            My.Computer.FileSystem.WriteAllText("FailedCrawlingPages.txt", targetURL + vbNewLine, True)
        '            Continue For
        '        End If

        '        Dim advertsNodes = pageDoc.DocumentNode.SelectNodes("//a").Where(Function(x) x.GetAttributeValue("class", "").Contains("pro-title") AndAlso Not x.GetAttributeValue("href", "").Contains("javascript")).Select(Function(y) y.GetAttributeValue("href", ""))
        '        For Each node In advertsNodes
        '            itys.Add(New ScrapeLib.URLsFields With {.ListingURL = node})
        '        Next

        '        If Not IsNothing(advertsNodes) AndAlso advertsNodes.Count > 0 Then
        '            CrawlingBackgroundWorker.ReportProgress(i, itys.ToArray)
        '        Else
        '            My.Computer.FileSystem.WriteAllText("FailedCrawlingPages.txt", targetURL + vbNewLine, True)
        '        End If
        '    Catch ex As Exception
        '        My.Computer.FileSystem.WriteAllText("FailedCrawlingPages.txt", targetURL + vbNewLine, True)
        '        LogState(ex.Message + vbNewLine + vbNewLine + ex.StackTrace)
        '    End Try
        'Next
    End Sub

    Private Sub CrawlingBackgroundWorker_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles CrawlingBackgroundWorker.ProgressChanged
        UpdateMany(e.UserState,
                   e.ProgressPercentage.ToString,
                   URLsFieldsBindingSource,
                   MainFormToolStripStatusLabel)
        ScrollGridToEnd(CrawlingDataGridView)
        ProcessedToolStripStatusLabel.Text = "Crawled#" + (CrawlingDataGridView.RowCount - 1).ToString
    End Sub

    Private Sub CrawlingBackgroundWorker_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles CrawlingBackgroundWorker.RunWorkerCompleted
        InComplete()
        ChangeStatus(CrawlingDataGridView.RowCount.ToString + " Records collected in " + LastProcessingSessionPeriod.ToString("hh\:mm\:ss"), False)
        ProcessedToolStripStatusLabel.Text = "Crawled#" + CrawlingDataGridView.RowCount.ToString
    End Sub
#End Region

    Private Sub SavingTimer_Tick(sender As Object, e As EventArgs) Handles SavingTimer.Tick
        If DataTabControl.SelectedIndex = 0 Then
            QuickSaveGrid(CrawlingDataGridView, My.Settings.SaveToPath)
        Else
            QuickSaveGrid(ScrapingDataGridView, My.Settings.SaveToPath)
        End If
    End Sub

    Private Sub SaveButton_Click(sender As Object, e As EventArgs) Handles SaveButton.Click
        If DataTabControl.SelectedIndex = 0 Then
            SaveDataGridToExcelFile(CrawlingDataGridView)
        Else
            SaveDataGridToExcelFile(ScrapingDataGridView)
        End If
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

    Sub PreStarting()
        If Not IO.Directory.Exists(My.Settings.SaveToPath) Then IO.Directory.CreateDirectory(My.Settings.SaveToPath)
        CheckBackupDirectory()
        ProcessingStopWatcher.Reset()
        ProcessingStopWatcher.Start()
        ScrapeButton.Enabled = False
        ScrapeFromSourceButton.Enabled = False
        CrawlButton.Enabled = False
        UseProxies = UseProxiesCheckBox.Checked
        If IO.File.Exists(FailedURLsFileName) Then IO.File.Delete(FailedURLsFileName)
        If IO.File.Exists(FailedKeywords) Then IO.File.Delete(FailedKeywords)
    End Sub

    Sub StartScraping()
        PreStarting()
        BreakScrapingLoop = False
        ScrapeButton.Enabled = False
        LastProcessingSessionPeriod = New TimeSpan
        ReScrapeFailedURLs = ReScrapeFailedCheckBox.Checked
        If DelayCheckBox.Checked Then Delay = DelayNumericUpDown.Value
        ScrapingBackgroundWorker.RunWorkerAsync()
    End Sub

    Sub StartCrawling()
        PreStarting()
        CrawlingBackgroundWorker.RunWorkerAsync()
    End Sub

    Sub InComplete()
        ProcessingStopWatcher.Stop()
        ScrapeButton.Enabled = True
        CrawlButton.Enabled = True
        LastProcessingSessionPeriod = ProcessingStopWatcher.Elapsed
    End Sub
#End Region
End Class