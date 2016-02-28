Imports ScrapeLib.Util
Imports ScrapeLib.ScraperSettings
Imports System.Text.RegularExpressions

Public Class MainForm
#Region "Properties"
    Property ScrapingStopWatcher As New Stopwatch
#End Region

#Region "Controls Events Handlers"
    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ScrapeLib.ScraperSettings.URLsFileName = "Keywords.txt"
        FailedURLsFileName = "FailedKeywords.txt"
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

    Private Sub ScrapingBackgroundWorker_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles ScrapingBackgroundWorker.DoWork
        Dim counter As Integer = 0
        Dim ity As ScrapeLib.GoogleResultsStatsFields = Nothing
        Dim _scraper As New ScrapeLib.GoogleResultsStats

        For Each keyword As String In URLs
            Try
                counter = counter + 1

                If BreakScrapingLoop Then BreakScrapingLoop = False : Exit For

                ity = _scraper.ScrapeStats(keyword)

                If Not IsNothing(ity) Then
                    ScrapingBackgroundWorker.ReportProgress(counter, ity)
                Else
                    My.Computer.FileSystem.WriteAllText("FailedKeywords.txt", keyword.Trim + vbNewLine, True)
                    ScrapingBackgroundWorker.ReportProgress(counter)
                End If

                CheckScrapingDelay()

            Catch ex As Exception
                My.Computer.FileSystem.WriteAllText("FailedKeywords.txt", keyword.Trim + vbNewLine, True)
                LogState(ex.Message + vbNewLine + vbNewLine + ex.StackTrace)
            End Try
        Next

        Dim isThereFailed As Boolean = False
        If IO.File.Exists("FailedKeywords.txt") Then
            isThereFailed = True
            URLs = LoadFileIntoList("FailedKeywords.txt")
            If IO.File.Exists(FailedURLsFileName) Then IO.File.Delete(FailedURLsFileName)
        End If


        If isThereFailed AndAlso URLs.Count > 0 AndAlso ReScrapeFailedURLs Then

            counter = 0
            ChangeStatus("ReScraping Failed Keyowrds.")

            For Each Keyword As String In URLs
                Try
                    counter = counter + 1

                    If BreakScrapingLoop Then BreakScrapingLoop = False : Exit For

                    ity = _scraper.ScrapeStats(Keyword)

                    If Not IsNothing(ity) Then
                        ScrapingBackgroundWorker.ReportProgress(counter, ity)
                    Else
                        My.Computer.FileSystem.WriteAllText("FailedKeywords.txt", Keyword.Trim + vbNewLine, True)
                        ScrapingBackgroundWorker.ReportProgress(counter)
                    End If

                    CheckScrapingDelay()

                Catch ex As Exception
                    My.Computer.FileSystem.WriteAllText("FailedKeywords.txt", Keyword.Trim + vbNewLine, True)
                    LogState(ex.Message + vbNewLine + vbNewLine + ex.StackTrace)
                End Try
            Next
        End If
    End Sub

    Private Sub ScrapingBackgroundWorker_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles ScrapingBackgroundWorker.ProgressChanged
        UpdateUI(e.UserState,
                 e.ProgressPercentage.ToString,
                 GoogleResultsStatsFieldsBindingSource,
                 MainFormToolStripStatusLabel)
        ScrollGridToEnd(ScrapingDataGridView)
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