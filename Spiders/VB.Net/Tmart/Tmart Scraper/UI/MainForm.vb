Imports System.Net
Imports System.Text
Imports MyScraper.Scraper
Imports Microsoft.VisualBasic.FileIO
Imports System.Text.RegularExpressions

Friend Class MainForm

#Region "Properties"
    'Property StartFromCategory As String
    Property IsSubComplete As Boolean = False
    Property IsMainComplete As Boolean = False
    Property MagentoFileDropboxUrl() As String
    Property MagentoFileDropboxPath() As String
    Property TreeNodesList As New List(Of TreeNode)
    Property ExcludedDepartmentsDict As New List(Of String)
    Property MagentoCsvFileName As String = "TmartScraperOutput.csv"
#End Region

#Region "Form Handlers"
    Private Sub MainForm_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        FillCatList()

        ReCreateLogFile()

        Dim exCats() As String = My.Computer.FileSystem.ReadAllText("Excluded Cats.dat").Split(vbNewLine)
        For Each cat In exCats
            ExcludedDepartmentsDict.Add(cat.Trim)
        Next

        If IO.File.Exists("DescriptionFooter.dat") Then
            DecriptionFooter = My.Computer.FileSystem.ReadAllText("DescriptionFooter.dat")
        End If

        If IO.File.Exists("ExcludedProducts.dat") Then
            Dim execProds() As String = My.Computer.FileSystem.ReadAllText("ExcludedProducts.dat").Split(vbNewLine)
            For Each prod As String In execProds
                ExcludedProducts.Add(prod.Trim)
            Next
        End If

        'Get site map to retrieve all available departments.
        ChangeStatus("Wait, loading available departments...", False)
        Navigate(MainWebBrowser, SiteMapURL)
        WaitForBrowserCompleted(MainWebBrowser)
        Application.DoEvents()

        'Get main departments & Display main departments count
        ScrapeDepartments(MainWebBrowser.Document, "siteMapList", "level0")

        If Not IsNothing(DepartmentDict) AndAlso DepartmentDict.Count > 0 Then
            Dim mainDepartmentsCount As Integer = DepartmentDict.Count
            ChangeStatus(String.Format("{0} departments found in {1}.", mainDepartmentsCount, TargetWebSiteName))
            ScrapeButton.Enabled = True
        ElseIf IsNothing(DepartmentDict) Then
            ChangeStatus("Departments not found")
        ElseIf DepartmentDict.Count = 0 Then
            ChangeStatus("Error: 0 departments found")
        End If

        If IO.File.Exists(GetAppFolderPath() + "magen.dat") Then
            PauseButton.Text = "Continue"
            PauseButton.Enabled = True
            FillMagentoDatagridByCsvFile(GetAppFolderPath() + "magen.dat", MagentoOutputBindingSource)
            MarkAllDoneNodes()
            DeSerilizeDataStructure()
        End If

        If IO.File.Exists(GetAppFolderPath() + "offery.dat") Then
            FilliOfferDatagridByCsvFile(GetAppFolderPath() + "offery.dat", iOfferOutputBindingSource)
        End If
    End Sub

    Shared Function ReadCSVFile(csvFilePath As String) As Microsoft.VisualBasic.FileIO.TextFieldParser
        Dim TextFileReader As Microsoft.VisualBasic.FileIO.TextFieldParser = New Microsoft.VisualBasic.FileIO.TextFieldParser(csvFilePath)
        TextFileReader.TextFieldType = FileIO.FieldType.Delimited
        TextFileReader.SetDelimiters(",")
        Return TextFileReader
    End Function

    Shared Sub FilliOfferDatagridByCsvFile(csvFilePath As String, gridBindingSource As BindingSource)
        Dim rowIndex As Integer = 0
        Dim currentRow As String() = Nothing
        Dim TextFileReader As Microsoft.VisualBasic.FileIO.TextFieldParser = ReadCSVFile(csvFilePath)

        While Not TextFileReader.EndOfData
            currentRow = TextFileReader.ReadFields()
            If Not currentRow Is Nothing Then
                If rowIndex = 0 Then
                    'You can play with headers here.
                End If

                If rowIndex > 0 Then
                    Dim magen As iOfferOutput = iOfferOutput.CreateiOfferObject(currentRow)
                    gridBindingSource.Add(magen)
                End If
            End If
            rowIndex = rowIndex + 1
        End While

        TextFileReader.Close()
    End Sub

    Shared Sub FillMagentoDatagridByCsvFile(csvFilePath As String, gridBindingSource As BindingSource)
        Dim rowIndex As Integer = 0
        Dim currentRow As String() = Nothing
        Dim TextFileReader As Microsoft.VisualBasic.FileIO.TextFieldParser = ReadCSVFile(csvFilePath)

        While Not TextFileReader.EndOfData
            currentRow = TextFileReader.ReadFields()
            If Not currentRow Is Nothing Then
                If rowIndex = 0 Then
                    'You can play with headers here.
                End If

                If rowIndex > 0 Then
                    Dim magen As MagentoOutput = MagentoOutput.CreateMagentoObject(currentRow)
                    gridBindingSource.Add(magen)
                End If
            End If
            rowIndex = rowIndex + 1
        End While
        TextFileReader.Close()
    End Sub

    Private Sub MainForm_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        ExchangeRateFromEuroToUsd = ScrapeEuroToUsdConversionRate()
    End Sub

    Private Sub WebBrowsers_DocumentCompleted(sender As System.Object, e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles MainWebBrowser.DocumentCompleted
        ReportBrowserComplete(CType(sender, WebBrowser))
    End Sub

    Private Sub ScrapeButton_Click(sender As System.Object, e As System.EventArgs) Handles ScrapeButton.Click
        If Not CheckMagentoSettings() Then Exit Sub
        ClearGrids()
        My.Settings.LastNodeScrapedFullPath = vbNullString
        My.Settings.Save()
        MarkDepartAsUnDoneByDepartmentPath()
        PauseButton.Text = "Pause"
        PauseButton.Enabled = True
        If IO.File.Exists("offery.dat") Then IO.File.Delete("offery.dat")
        If IO.File.Exists("magen.dat") Then IO.File.Delete("magen.dat")
        If IO.File.Exists("ScraProds.dat") Then IO.File.Delete("ScraProds.dat")
        If IO.File.Exists("prodsList.dat") Then IO.File.Delete("prodsList.dat")
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ScrapeAllNow(sender)
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        If e.UserState.GetType Is GetType(String) Then
            Dim message = e.UserState
            ChangeStatus(message, False)
        Else
            Dim outputsArr As List(Of Object) = e.UserState
            Try
                MagentoOutputBindingSource.Add(outputsArr(1))
                iOfferOutputBindingSource.Add(outputsArr(2))
            Catch ex As Exception
                LogState(ex.Message)
                Console.WriteLine("asdf")
            End Try
            outputsArr = Nothing 'may be this do and error I don't know 'ko
        End If
        ShownToolStripStatusLabel.Text = ", Scraped Products: " + DataGridView1.Rows.Count.ToString
        'NonUsaPricedToolStripStatusLabel.Text = ", Excluded Warehouses Products: " + excludedWarehousesItems.Count.ToString
        'OutOfStockItemsToolStripStatusLabel.Text = ", Out Of Stock Products: " + OutOfStockItems.Count.ToString
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As System.Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        SendToMagento()
    End Sub

    Private Sub DepartmentsTreeView_MouseDown(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles DepartmentsTreeView.MouseDown
        Dim clickedNode As TreeNode = GetClickedNode(e.X, e.Y)
        If Not IsNothing(clickedNode) Then
            DepartmentsTreeView.SelectedNode = clickedNode
#If CONFIG = "Debug" Then
            If clickedNode.Nodes.Count = 0 Then
                ScrapeNowToolStripMenuItem.Visible = True
            Else
                ScrapeNowToolStripMenuItem.Visible = False
            End If
#End If
        End If
    End Sub

    Private Sub OpenDepartmentToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles OpenDepartmentToolStripMenuItem.Click
        Dim clickedNode As TreeNode = GetClickedNode()
        If Not IsNothing(clickedNode) Then
            Dim url As String = clickedNode.Tag
            If url <> vbNullString Then
                Process.Start(url)
            Else
                Dim errorMessage As String = "No URL found to click"
                ChangeStatus(errorMessage)
            End If
        End If
    End Sub

    Private Sub ScrapeNowToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ScrapeNowToolStripMenuItem.Click
        Dim clickedNode As TreeNode = GetClickedNode()
        If Not IsNothing(clickedNode) Then
            Dim url As String = clickedNode.Tag
            If url <> vbNullString Then
                ClearGrids()
                Application.DoEvents()
                LogState("Prepare scraping department '" + clickedNode.Text + "'" + " URL:" + clickedNode.Tag)
                Dim _stopWatcher As New Stopwatch
                _stopWatcher.Start()
                Dim _scraper As New Scraper
                _scraper.ScrapeDepartment(url, clickedNode.Text)
                _stopWatcher.Stop()
                If Me.InvokeRequired Then
                    Me.BeginInvoke(Sub()
                                       ChangeStatus("All Done in " + _stopWatcher.Elapsed.ToString("g"), True)
                                       ScrapeButton.Enabled = True
                                       SendToMagentoButton.Enabled = True
                                       StopButton.Enabled = False
                                   End Sub)
                Else
                    ChangeStatus("All Done in " + _stopWatcher.Elapsed.ToString("g"), True)
                    ScrapeButton.Enabled = True
                    SendToMagentoButton.Enabled = True
                    StopButton.Enabled = False
                End If
                SendToMagento()
            End If
        End If
    End Sub

    Private Sub MainForm_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        End
    End Sub

    Private Sub SendToMagentoButton_Click(sender As System.Object, e As System.EventArgs) Handles SendToMagentoButton.Click
        SendToMagento()
    End Sub

    Private Sub SettingsButton_Click(sender As System.Object, e As System.EventArgs) Handles SettingsButton.Click
        SettingsForm.ShowDialog()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub StopButton_Click(sender As System.Object, e As System.EventArgs) Handles StopButton.Click
        Dim res = MsgBox("Are you sure you want to restart?", vbYesNo)
        If res = MsgBoxResult.No Then Exit Sub
        IsCancelled = True
        BackgroundWorker1.CancelAsync()
        SaveLastCategoryScraped()
        My.Settings.LastNodeScrapedFullPath = vbNullString
        My.Settings.Save()
        SaveDataGridToExcelFile(iOfferOuputDataGridView, , GetAppFolderPath() + "offery.dat")
        SaveDataGridToExcelFile(DataGridView1, , GetAppFolderPath() + "magen.dat")
        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@ cancel async doesn't work
        SerilizeDataStructure()
        ' ScrapeButton_Click(Me, New EventArgs)
    End Sub

    Private Sub SaveButton_Click(sender As System.Object, e As System.EventArgs) Handles SaveButton.Click
        SaveDataGridToExcelFile(iOfferOuputDataGridView)
    End Sub

    Private Sub SaveBusinessGridButton_Click(sender As System.Object, e As System.EventArgs) Handles SaveBusinessGridButton.Click
        SaveDataGridToExcelFile(DataGridView1)
    End Sub

    Private Sub PauseButton_Click(sender As System.Object, e As System.EventArgs) Handles PauseButton.Click
        If PauseButton.Text = "Continue" Then
            If DataGridView1.Rows.Count > 1 Then
                SaveLastCategoryScraped()
                If Not BackgroundWorker1.IsBusy Then
                    BackgroundWorker1.RunWorkerAsync()
                End If
            End If
            IsPaused = False
            PauseButton.Text = "Pause"
        ElseIf PauseButton.Text = "Pause" Then
            IsPaused = True
            PauseButton.Text = "Continue"
            ScrapeButton.Enabled = True
        End If
    End Sub
#End Region

#Region "Helpers"
    Sub ScrapeAllNow(Optional workerThread As System.ComponentModel.BackgroundWorker = Nothing)
        If Me.InvokeRequired Then
            Me.BeginInvoke(Sub()
                               SettingsButton.Enabled = False
                               ScrapeButton.Enabled = False
                               SendToMagentoButton.Enabled = False
                               StopButton.Enabled = True
                           End Sub)
        Else
            SettingsButton.Enabled = False
            ScrapeButton.Enabled = False
            SendToMagentoButton.Enabled = False
            StopButton.Enabled = True
        End If

        Dim _stopWatcher As New Stopwatch
        _stopWatcher.Start()
        Dim _scraper As New Scraper
        Dim allChildDepartments = From kvp In DepartmentDict Where kvp.Key.Contains("/")
        
        Dim ignore As Boolean = False
        For Each kvp As KeyValuePair(Of String, Department) In allChildDepartments
            Try
                If IsCancelled Then
                    IsCancelled = False
                    Exit For
                End If

                Do While IsPaused
                    Application.DoEvents()
                Loop

                Dim lastCatPath As String = GetLastScrapedCategory()
                If lastCatPath <> vbNullString Then
                    If Not ignore Then
                        If kvp.Key <> lastCatPath Then
                            Continue For
                        Else
                            ignore = True
                            Continue For
                        End If
                    End If
                Else
                    ignore = True
                End If

                Dim departName As String = kvp.Value.DepartmentName
                Dim departUrl As String = kvp.Value.DepartmentUrl
                LogState("Prepare scraping department '" + departName + "'" + " URL:" + departUrl)
                ProductsPagesList = New List(Of String)
                _scraper.ScrapeDepartment(departUrl, departName, workerThread)
                '#work area 26/05/2013
                MarkDepartAsDoneByDepartmentPath(kvp.Key)

                SaveDataGridToExcelFile(iOfferOuputDataGridView, , GetAppFolderPath() + "offery.dat")
                SaveDataGridToExcelFile(DataGridView1, , GetAppFolderPath() + "magen.dat")
                SerilizeDataStructure()
            Catch ex As Exception
                LogState(ex.Message + vbNewLine + vbNewLine + ex.StackTrace)
            End Try
        Next
        allChildDepartments = Nothing



        _stopWatcher.Stop()
        'Dim filePath As String = GetAppFolderPath() + GetDateTimeFileName("txt")
        'My.Computer.FileSystem.WriteAllText(GetAppFolderPath() + filePath,
        '                                    ShownToolStripStatusLabel.Text + vbNewLine +
        '                                    NonUsaPricedToolStripStatusLabel.Text + vbNewLine +
        '                                    OutOfStockItemsToolStripStatusLabel.Text,
        '                                    False)
        If Me.InvokeRequired Then
            Me.BeginInvoke(Sub()
                               ChangeStatus("All Done in " + _stopWatcher.Elapsed.ToString("g"), True)
                               SettingsButton.Enabled = True
                               ScrapeButton.Enabled = True
                               SendToMagentoButton.Enabled = True
                               StopButton.Enabled = False
                           End Sub)
        Else
            ChangeStatus("All Done in " + _stopWatcher.Elapsed.ToString("g"), True)
            ScrapeButton.Enabled = True
            SettingsButton.Enabled = True
            SendToMagentoButton.Enabled = True
            StopButton.Enabled = False
        End If
    End Sub

    Sub DeSerilizeDataStructure()
        If IO.File.Exists(GetAppFolderPath() + "prodsList.dat") Then
            ProductsPagesList = DeserializeObject(GetAppFolderPath() + "prodsList.dat")
        End If
        If IO.File.Exists("ScraProds.dat") Then
            ScrapedProdcuts = DeserializeObject(GetAppFolderPath() + "ScraProds.dat")
        End If
    End Sub

    Sub SerilizeDataStructure()
        If IO.File.Exists("prodsList.dat") Then IO.File.Delete("prodsList.dat")
        If IO.File.Exists("ScraProds.dat") Then IO.File.Delete("ScraProds.dat")
        SerializeObject(ProductsPagesList, GetAppFolderPath() + "prodsList.dat")
        SerializeObject(ScrapedProdcuts, GetAppFolderPath() + "ScraProds.dat")
    End Sub

    Sub FillCatList()
        Using parser As New TextFieldParser("catlist.dat")
            parser.CommentTokens = New String() {"#"}
            parser.SetDelimiters(New String() {","})
            parser.HasFieldsEnclosedInQuotes = True
            parser.ReadLine()
            While Not parser.EndOfData
                Dim fields As String() = parser.ReadFields()
                CatList.Add(fields(0), fields(1))
            End While
        End Using
    End Sub

    Function CheckMagentoSettings() As Boolean
        '# Make sure DropBox is running...
        If Not IsDropboxRunning() Then
            MsgBox("Dropbox is not running, please run the Dropbox first.", MsgBoxStyle.Critical)
            LogState("Dropbox is not running, cannot send data to Magento.")
            Return False
        End If

        '# You must have the DropBox public folder Path as an option in the settings.
        If Not IsValidDropBoxPublicFolder(My.Settings.DropBoxPath) Then
            MsgBox("DropBox Public folder not valid, please choose a valid path", MsgBoxStyle.Critical, "Tmart Scraper")
            LogState("Sending to Magento quit, as DropBox folder path is not valid.")
            Return False
        End If

        '# Get public dropbox folder Ur.
        If Not IsValidDropBoxPublicUrl(My.Settings.DropBoxUrl) Then
            MsgBox("DropBox Public folder Url not valid, please choose a valid URL", MsgBoxStyle.Critical, "Tmart Scraper")
            LogState("Sending to Magento quit, as DropBox folder URL is not valid.")
            Return False
        End If
        Return True
    End Function

    Sub SendToMagento()
        If Not CheckMagentoSettings() Then Exit Sub

        'Get required DropBox data.
        If Not Strings.Right(My.Settings.DropBoxPath, 1) = "/" Then MagentoFileDropboxPath = My.Settings.DropBoxPath + "\"
        If Not Strings.Right(My.Settings.DropBoxUrl, 1) = "/" Then MagentoFileDropboxUrl = My.Settings.DropBoxUrl
        MagentoFileDropboxPath = MagentoFileDropboxPath + IO.Path.GetFileName(My.Settings.DropBoxUrl)

        'Delete current csv file.
        If IO.File.Exists(MagentoFileDropboxPath) Then IO.File.Delete(MagentoFileDropboxPath)

        'Looping till file got deleted
        Dim timeOutStopWatch As New Stopwatch
        LogState("Looping till file got deleted...")
        timeOutStopWatch.Start()
        Dim i As Integer = 0
        Dim isTimeOut As Boolean = True
        Do Until timeOutStopWatch.Elapsed.Seconds = 60
            i = i + 1
            If timeOutStopWatch.Elapsed.Seconds Mod 5 = 0 Then
                Application.DoEvents()
                If Not CheckIfDownloadableFile(MagentoFileDropboxUrl) Then
                    isTimeOut = False
                    Exit Do
                End If
            End If
        Loop
        timeOutStopWatch.Stop()
        'Make sure file is deleted.
        If isTimeOut Then
            LogState("Timeout and old csv file still not deleted from DropBox")
            MsgBox("Error, cannot complete the process", MsgBoxStyle.Critical)
            Exit Sub
        End If



        '# Create your new file.
        SaveDataGridToExcelFile(DataGridView1, MagentoFileDropboxPath)
        'Looping till file got deleted
        LogState("Looping till file got uploaded...")
        timeOutStopWatch.Reset()
        timeOutStopWatch.Start()
        i = 0
        isTimeOut = True
        Do Until timeOutStopWatch.Elapsed.Seconds = 60
            i = i + 1
            If timeOutStopWatch.Elapsed.Seconds Mod 5 = 0 Then
                Application.DoEvents()
                If CheckIfDownloadableFile(MagentoFileDropboxUrl) Then
                    isTimeOut = False
                    Exit Do
                End If
            End If
        Loop
        timeOutStopWatch.Stop()
        'Make sure file is uploaded.
        If isTimeOut Then
            LogState("Timeout and the new csv file still not uploaded to DropBox")
            MsgBox("Error, cannot complete the process", MsgBoxStyle.Critical)
            Exit Sub
        End If

        '# Send to Magento
        MagentoImporter.ShowDialog()
    End Sub

    Function GetClickedNode() As TreeNode
        Return DepartmentsTreeView.SelectedNode
    End Function

    Function GetClickedNode(x As Integer, y As Integer) As TreeNode
        Dim hitInfo As TreeViewHitTestInfo = DepartmentsTreeView.HitTest(x, y)
        Return hitInfo.Node
    End Function

    Sub ClearGrids()
        iOfferOutputBindingSource.Clear()
        MagentoOutputBindingSource.Clear()
        'OutOfStockItems = New Dictionary(Of String, String)
        'excludedWarehousesItems = New Dictionary(Of String, String)
        ProductsPagesList = New List(Of String)
        ScrapedProdcuts = New List(Of String)
    End Sub

    Sub MarkDepartAsUnDoneByDepartmentPath()
        Dim nodesQuery = From node As TreeNode In TreeNodesList

        For Each nodeObj As TreeNode In nodesQuery
            Dim treNode As TreeNode = nodeObj
            If Not IsNothing(nodesQuery) AndAlso nodesQuery.Count > 0 Then
                If Me.InvokeRequired Then
                    Me.BeginInvoke(Sub()
                                       treNode.BackColor = Color.White
                                   End Sub)
                Else
                    treNode.BackColor = Color.White
                End If
            End If
        Next

    End Sub

    Function GetDepartmentByURL(url As String) As Department
        Dim depQuery = From node As KeyValuePair(Of String, Department) In DepartmentDict
                       Where node.Value.DepartmentUrl = url
        Return depQuery.First.Value
    End Function

    Function GetDepartmentByPath(path As String) As Department
        Dim depQuery = From node As KeyValuePair(Of String, Department) In DepartmentDict
                       Where node.Key = path
        Return depQuery.First.Value
    End Function

    Sub SaveLastCategoryScraped()
        Dim lastNode As TreeNode = GetLastScrapedNode()
        If Not IsNothing(lastNode) Then
            My.Settings.LastNodeScrapedFullPath = lastNode.FullPath
            My.Settings.Save()
        End If
    End Sub

    Function GetLastScrapedNode() As TreeNode
        Dim nodesQuery = From node As TreeNode In TreeNodesList
                         Where node.BackColor = Color.LightGreen
        If nodesQuery.Count > 0 Then
            Return nodesQuery.Last
        Else
            Return Nothing
        End If
    End Function

    Function GetLastScrapedCategory() As String
        Dim lastNodeScrape As TreeNode = GetLastScrapedNode()
        If Not IsNothing(lastNodeScrape) Then
            Return lastNodeScrape.FullPath.Replace("\", "/")
        Else
            Return vbNullString
        End If
    End Function

    Sub MarkAllDoneNodes()
        If My.Settings.LastNodeScrapedFullPath = vbNullString Then Exit Sub
        Dim savedNode As String = My.Settings.LastNodeScrapedFullPath

        For Each node As TreeNode In TreeNodesList
            node.BackColor = Color.LightGreen
            If savedNode = node.FullPath Then
                DepartmentsTreeView.SelectedNode = node
                DepartmentsTreeView.Focus()
                Exit For
            End If
        Next
    End Sub

    Sub MarkDepartAsDoneByDepartmentPath(departmentPath As String)
        Dim depUrl = GetDepartmentByPath(departmentPath)

        If Not IsNothing(depUrl) Then
            Dim URL As String = depUrl.DepartmentUrl
            Dim nodesQuery = From node As TreeNode In TreeNodesList
                             Where node.Tag = URL
            If Not IsNothing(nodesQuery) AndAlso nodesQuery.Count > 0 Then
                If Me.InvokeRequired Then
                    Me.BeginInvoke(Sub()
                                       DepartmentsTreeView.SelectedNode = nodesQuery.First
                                       nodesQuery.First.BackColor = Color.LightGreen
                                       DepartmentsTreeView.Focus()
                                       SaveLastCategoryScraped()
                                   End Sub)
                Else
                    DepartmentsTreeView.SelectedNode = nodesQuery.First
                    nodesQuery.First.BackColor = Color.LightGreen
                    DepartmentsTreeView.Focus()
                    SaveLastCategoryScraped()
                End If
            End If
        End If
    End Sub
#End Region

    Private Sub DataGridView1_DataError(sender As System.Object, e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        My.Computer.FileSystem.WriteAllText("Magen-Errors.txt", e.Exception.Message + vbNewLine + vbNewLine + e.Exception.StackTrace, True)
    End Sub

    Private Sub iOfferOuputDataGridView_DataError(sender As System.Object, e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles iOfferOuputDataGridView.DataError
        My.Computer.FileSystem.WriteAllText("Offery-Errors.txt", e.Exception.Message + vbNewLine + vbNewLine + e.Exception.StackTrace, True)
    End Sub
End Class
