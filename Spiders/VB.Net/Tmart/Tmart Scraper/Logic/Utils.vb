Option Explicit On

Imports System.Net
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.IO

Module Utils

    Public DecriptionFooter As String = vbNullString

    Public ExcludedProducts As New List(Of String)

    Sub ReCreateLogFile()
        If IO.File.Exists(LogFilePath) Then IO.File.Delete(LogFilePath)
        LogState(DateTime.Now + ":")
    End Sub

    Sub ChangeStatus(status As String, Optional log As Boolean = True)
        'If status.Contains("object") Then Debugger.Break()
        If MainForm.InvokeRequired Then
            MainForm.BeginInvoke(Sub() MainForm.StatusToolStripStatusLabel.Text = status)
        Else
            MainForm.StatusToolStripStatusLabel.Text = status
        End If
        If log Then LogState(status)
    End Sub

    Function GetAppFolderPath() As String
        If Right(Application.StartupPath, 1) = "\" Then
            Return Application.StartupPath
        Else
            Return Application.StartupPath + "\"
        End If
    End Function

    Function GetDateTimeFileName(ext As String) As String
        Dim fileName As String = DateTime.Now.ToString("G").Replace(":", "-").Replace("/", "-") + "." + ext
        Return fileName
    End Function

    Sub LogState(state As String)
#If CONFIG = "Debug" Then
        Try
            'Disable till I figure out the threading error here.
            SyncLock state
                My.Computer.FileSystem.WriteAllText(GetAppFolderPath() + LogFilePath, state & vbNewLine, True)
            End SyncLock
        Catch ex As Exception
            My.Computer.FileSystem.WriteAllText(GetAppFolderPath() + "logging-errors.txt", state & vbNewLine & ex.Message & vbNewLine & ex.StackTrace, True)
        End Try
#End If
    End Sub

    Sub Navigate(browser As WebBrowser, url As String)
        If browser.InvokeRequired Then
            browser.BeginInvoke(Sub()
                                    If browser.Name = "MainWebBrowser" Then
                                        MainForm.IsMainComplete = False
                                    ElseIf browser.Name = "PaginationWebBrowser" Then
                                        MainForm.IsSubComplete = False
                                    End If
                                    browser.Navigate(url)
                                End Sub)
        Else
            If browser.Name = "MainWebBrowser" Then
                MainForm.IsMainComplete = False
            ElseIf browser.Name = "PaginationWebBrowser" Then
                MainForm.IsSubComplete = False
            End If
            browser.Navigate(url)
        End If
    End Sub

    Sub WaitForBrowserCompleted(browser As WebBrowser)
        If browser.Name = "MainWebBrowser" Then
            Do Until MainForm.IsMainComplete
                Application.DoEvents()
            Loop
            MainForm.IsMainComplete = False
        ElseIf browser.Name = "PaginationWebBrowser" Then
            Do Until MainForm.IsSubComplete
                Application.DoEvents()
            Loop
            MainForm.IsSubComplete = False
        End If
    End Sub

    Sub ReportBrowserComplete(browser As WebBrowser)
        If browser.ReadyState = WebBrowserReadyState.Complete Then
            If browser.Name = "MainWebBrowser" Then
                LogState("Main browser completes page:" + browser.Url.ToString)
                MainForm.IsMainComplete = True
            ElseIf browser.Name = "PaginationWebBrowser" Then
                LogState("Pagination browser completes page:" + browser.Url.ToString)
                MainForm.IsSubComplete = True
            End If
        End If
    End Sub

    Sub SaveDataGridToExcelFile(dataGrid As DataGridView, Optional sendToDropBox As String = vbNullString, Optional defaultFile As String = vbNullString)
        Dim SaveAsDialog As New SaveFileDialog

        Dim headers = (From header As DataGridViewColumn In dataGrid.Columns.Cast(Of DataGridViewColumn)() Select header.HeaderText).ToArray

        Dim rows = From row As DataGridViewRow In dataGrid.Rows.Cast(Of DataGridViewRow)() _
                   Where Not row.IsNewRow And Not row.Cells(0).Value Is "Column0"
                   Select Array.ConvertAll(row.Cells.Cast(Of DataGridViewCell).ToArray, Function(c) If(c.Value IsNot Nothing, c.Value.ToString, ""))

        Dim result As DialogResult
        Dim fileName As String = vbNullString
        If defaultFile <> vbNullString Then
            fileName = defaultFile
        ElseIf sendToDropBox = vbNullString Then
            SaveAsDialog.Filter = "CSV Files|*.csv"
            result = SaveAsDialog.ShowDialog()
            fileName = SaveAsDialog.FileName
        Else
            fileName = sendToDropBox
        End If

        If result = Windows.Forms.DialogResult.OK OrElse fileName <> vbNullString Then
            Using sw As New IO.StreamWriter(fileName)
                sw.WriteLine(String.Join(",", headers))
                For Each r In rows
                    For i = 0 To UBound(r)
                        r(i) = r(i).Replace(Chr(34), "'")
                        r(i) = Chr(34) & r(i) & Chr(34)
                    Next
                    sw.WriteLine(String.Join(",", r))
                    Application.DoEvents()
                Next
            End Using
        End If
    End Sub

    Function IsValidDropBoxPublicFolder(path As String) As Boolean
        If Strings.Mid(path, 2, 1) = ":" And path.Contains("Dropbox") AndAlso Strings.Right(path, 7) = "\Public" Then
            Return True
        End If
        Return False
    End Function

    Function IsValidDropBoxPublicUrl(url As String) As Boolean
        Dim regx As New System.Text.RegularExpressions.Regex("https://dl.dropboxusercontent.com/u/\d+")
        Dim match = regx.Match(url)
        If match.Success Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function CheckIfDownloadableFile(ByVal fileUrl As String) As Boolean
        Try
            Dim wClient As New Net.WebClient
            wClient.DownloadFile(New Uri(fileUrl), "temp.dat")
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Function IsDropboxRunning() As Boolean
        Dim processesQuery = From proc As Process In Process.GetProcesses()
                     Where proc.ProcessName.Contains("Dropbox")

        If processesQuery.Count = 0 Then
            Return False
        End If
        Return True
    End Function

    Public Sub AddChildrenTreeNode(Nodes As List(Of TreeNode), Node As TreeNode)
        For Each thisNode As TreeNode In Node.Nodes
            Nodes.Add(thisNode)
            AddChildrenTreeNode(Nodes, thisNode)
        Next
    End Sub

    Function GetTreeNodeByTag(tree As TreeView, tag As String) As TreeNode
        Dim Nodes As New List(Of TreeNode)
        For Each node As TreeNode In tree.Nodes
            AddChildrenTreeNode(Nodes, node)
            If node.Tag = tag Then
                Return node
            End If
        Next
        Return Nothing
    End Function

    Public Function SerializeObject(ByVal obj As Object, ByVal saveToPath As String) As Boolean
        Try
            Dim bFormatter As New BinaryFormatter
            Dim fStream As New FileStream(saveToPath, FileMode.Create)
            bFormatter.Serialize(fStream, obj)
            fStream.Close()
            fStream.Dispose()
            fStream = Nothing
            bFormatter = Nothing
            Return True
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            ' Alerter.REP("Error during serializing object.", ex, True)
            Return False
        End Try
    End Function

    Public Function DeserializeObject(ByVal readFromPath As String, Optional ByRef IsFailed As Boolean = False) As Object
        'ex:
        'CacaduErrorReporter = CType(Util.DeserializeObject("c:\ahmed.dat"), ErrorReporter)
        Dim obj As Object = Nothing
        Dim fStream = New FileStream(readFromPath, FileMode.Open)
        Dim BinFormatter As BinaryFormatter = New BinaryFormatter
        Try
            obj = BinFormatter.Deserialize(fStream)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            '   Alerter.REP("Object file is corrupted, error restoring a serialized object.", ex, True)
            IsFailed = True
        End Try
        fStream.Close()
        fStream.Dispose()
        fStream = Nothing
        BinFormatter = Nothing
        Return obj
    End Function
End Module