Imports Cacadore
Imports System.Data
Imports Cacadu.Connectors
Imports System.ComponentModel
Imports DevExpress.XtraBars.Ribbon
Imports DevExpress.XtraBars.Ribbon.ViewInfo
Imports Cacadu.Connectors.MainFormConnector

Namespace UI
    Friend Class MainForm
#Region "Fields"
        Dim WithEvents _watcherObj As Watcher
#End Region
        
#Region "Constructors"
        Public Sub New()
            Dim ahmedd As New Actions.FacebookActions.LastFacebookNotification
            Balora.Serializer.XMLSerialr.SerializeObject(ahmedd, "c:\asdfjhjkasdfhasdjkfhjkhasdfjkhasjkdfhjkasdhfjkashdfjkhasjkdfhjk")

            MsgBox("gotta me")
            InitializeComponent() 'This call is required by the designer
        End Sub
#End Region

#Region "MainForm Components Events Handlers"
#Region "Form Handlers"
        Private Sub MainForm_Load(sender As Object, e As System.EventArgs) Handles Me.Load
            LoadLayout()
            SetMainFormPlacement()
        End Sub

        Private Sub MainForm_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
            _watcherObj = New Watcher

            SendToTrayIfMinCommandLineIncluded()

            CheckIfSplashScreenEnabled()

            Skins.AddHandlerForStyleChangedEvent()

            PrepareCacadu()

            LoadMainForm()

            HideSomeButtonsIfRegistered()

            Dim trgsStarter As New TriggersStarter
            trgsStarter.TryStartTriggers()

            CheckForUpdates("https://dl.dropbox.com/u/5721699/cacadu-appcast/versioninfo.xml", Me.Icon)
            'Here is where Sparkle stores the key. If you delete it, you will force the update window to be shown. 
            '{HKEY_CURRENT_USER\Software\Conderella Software http://www.cacadu.conderella.com\Cacadu\AutoUpdate}
            'Move them to here to speed up openning
        End Sub

        Private Sub MainForm_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
            SaveMainFormPlacement()

            SaveLayout()

            If My.Settings.STST AndAlso Not e.CloseReason = CloseReason.ApplicationExitCall Then
                e.Cancel = True
                SendToSystray()
            Else
                RemoveHandler TecDAL.TypedDataSetCrud.TectonicDataSetObj.tasks.tasksRowChanged, AddressOf tasksRowChanged
                RemoveHandler Me.TasksListDataGridView.DataError, AddressOf MainFormConnector.TasksListDataGridView_DataError
            End If
        End Sub

        Private Sub MainForm_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
            'Not shutting down Quartz will prevent Cacadu from closing completely.
            ShutdownQuartz()
            'ShowWaitForm("Saving tasks...")
            'Disable because it slows down Cacadu closing.
            'Cacadore.Watcher.SaveWatchedTasks(cancelIfRunning:=True), and because every 2 minutes you have save them using updateTasks
            'CloseWaitForm()
        End Sub

        Private Sub MainTimer_Tick(sender As System.Object, e As System.EventArgs) Handles MainTimer.Tick
            Dim taskId As String = GetSelectedTaskId()
            Dim nodeType As String = GetSelectedNodeType()
            Dim isPaused As Boolean = Watcher.IsWatchedTaskPaused(taskId)
            Dim isRunning As Boolean = Watcher.IsWatchedTaskRunning(taskId)
            Dim isScheduled As Boolean = Watcher.IsWatchedTaskScheduled(taskId)

            If nodeType = "Task" AndAlso (isRunning OrElse isPaused OrElse isScheduled) Then
                TogglePauseButtonText(taskId)
                ToggleStopButtonEnable(True)
            Else
                DisablePauseButton()
                ToggleStopButtonEnable(False)
                ToggleQuickAccessButtons(False)
            End If

            If nodeType = "Task" Then ToggleQuickAccessButtons(True)

            ToggleStopAllButtonEnable()

            CheckIfNoNodes()

            'طريقة مفتكسة لإظهار فورم الألرت عشان مش هينفع اظهرها 
            'من ثريد تاني غير الثريد الرئيسي. تقوم الطريقة علي متغير عام 
            'هغيره لما تتبعت رسالة. وبعدين اعمل تشيك عليها من التايمر 
            'الرئيسي في الفورم.
            If AlertNow And AlertText <> "" Then
                ShowAlertForm("", AlertText, AlertLocation, AlertTag, ShowAlertButton, AlertImage, AlertState)
                AlertNow = False
                ShowAlertButton = False
                AlertText = ""
                AlertState = Nothing
                AlertTag = vbNullString
                AlertImage = Nothing
                AlertLocation = Cacadore.Actions.MessageBoxes.ShowAlert.ShowAlertsInputs.AlertLocationEnum.BottomRight
            End If
        End Sub

        Private Sub MainForm_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize
            HideMe(Me)
        End Sub
#End Region

#Region "Watcher Event Handler"
        Public Sub TaskState_Changed(taskName As String, historyEntry As Cacadore.HistoryEntry) Handles _watcherObj.TaskState_Changed
            Dim description As String = historyEntry.Description
            Dim entryDate As DateTime = historyEntry.EntryDateTime

            If Not IsNothing(HistoryDataGridView) Then AddNewHistoryRow(HistoryDataGridView.Rows.Count + 1, taskName, description, entryDate)

            If historyEntry.Status = ITaskState.ExecutionState.Failed AndAlso Not IsNothing(ErrorsDataGridView) Then
                AddNewErrorRow(ErrorsDataGridView.Rows.Count + 1, taskName, description, entryDate)
            End If

            If Cacadore.Settings.HistoryEnabled Then
                SyncLock pendingHistoryEntries
                    pendingHistoryEntries.Enqueue(historyEntry)
                End SyncLock
            End If
        End Sub

        Function CheckForRedRows(description As String) As Color
            Dim rowForeColor As Color
            If description.Contains("fail") OrElse description.Contains("Task has no actions") Then
                rowForeColor = Color.Red
            Else
                rowForeColor = Color.Black
            End If
            Return rowForeColor
        End Function

        Sub AddNewErrorRow(rowCount As Integer, taskName As String, description As String, entryDate As DateTime)
            If IsNothing(ErrorsDataGridView) Then Exit Sub

            Dim rowForeColor As Color = CheckForRedRows(description)

            If IsHandleCreated And Not IsDisposed Then
                If ErrorsDataGridView.InvokeRequired Then
                    ErrorsDataGridView.BeginInvoke(Function()
                                                       Dim rowIndex As Integer = ErrorsDataGridView.Rows.Add(ErrorsDataGridView.Rows.Count + 1, taskName, description, entryDate)
                                                       ErrorsDataGridView.Rows(rowIndex).DefaultCellStyle.ForeColor = Color.Red
                                                       Return True
                                                   End Function)
                Else
                    Dim rowIndex As Integer = ErrorsDataGridView.Rows.Add(rowCount, taskName, description, entryDate)
                    ErrorsDataGridView.Rows(rowIndex).DefaultCellStyle.ForeColor = rowForeColor
                End If
            End If
        End Sub

        Sub AddNewHistoryRow(rowCount As Integer, taskName As String, description As String, entryDate As DateTime)
            If IsNothing(HistoryDataGridView) Then Exit Sub

            Dim rowForeColor As Color = CheckForRedRows(description)

            If IsHandleCreated And Not IsDisposed Then
                If HistoryDataGridView.InvokeRequired Then
                    HistoryDataGridView.BeginInvoke(Function()
                                                        Dim rowIndex As Integer = HistoryDataGridView.Rows.Add(HistoryDataGridView.Rows.Count + 1, taskName, description, entryDate)
                                                        HistoryDataGridView.Rows(rowIndex).DefaultCellStyle.ForeColor = rowForeColor
                                                        Return True
                                                    End Function)
                Else
                    Dim rowIndex As Integer = HistoryDataGridView.Rows.Add(rowCount, taskName, description, entryDate)
                    HistoryDataGridView.Rows(rowIndex).DefaultCellStyle.ForeColor = rowForeColor
                End If
            End If
        End Sub

        Private Sub ClearToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ClearToolStripMenuItem.Click
            HistoryDataGridView.Rows.Clear()
        End Sub
#End Region
#Region "Ribbon Handlers"
#Region "Home Page Handlers"
        Private Sub TaskFlowHomePageBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles TaskFlowHomePageBarButton.ItemClick
            DisplayTaskFlow(GetSelectedNodeName(), GetSelectedNodeText(), GetSelectedNodeType())
        End Sub
        Private Sub TriggersHomePageBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles TriggersHomePageBarButton.ItemClick
            OnTriggersButtonClick()
        End Sub

        Private Sub SetGroupNameBarEditItem_EditValueChanged(sender As Object, e As System.EventArgs) Handles SetGroupNameBarEditItem.EditValueChanged
            AddNewGroup(SetGroupNameBarEditItem.EditValue.ToString)
        End Sub

        Private Sub QuickAddBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles QuickAddBarButton.ItemClick
            AddNewTaskForm.ShowDialog()
        End Sub

        Private Sub UsingWizardBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles UsingWizardBarButton.ItemClick
            MsgBox("Wizard Button...")
        End Sub

        Private Sub TasksListBarButtonItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles TasksListBarButtonItem.ItemClick
            ShowTasksList()
        End Sub

        Private Sub RibbonControlOfMainForm_MouseMove(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles RibbonControlOfMainForm.MouseMove
            'To fix bug where you cannot click action unless you click its text. i.e if you clicked action's icon it doesn't get dragged.
            'So you we use RibbonControlOfMainForm_MouseMove rather than RibbonControlOfMainForm_MouseDown where we make sure that left
            'button is clicked.
            If e.Button = Windows.Forms.MouseButtons.Left Then TryDragAction(e)
        End Sub

#Region "Run Button"
        Delegate Sub MethodToBeInjectedDlg(sender As Object, e As System.ComponentModel.DoWorkEventArgs)

        Private Sub RunHomePageBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles RunHomePageBarButton.ItemClick
            RunSelectedTask()
        End Sub

        Friend Sub backGroundWorkerValue_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs)
            Dim tsk As Cacadore.Task = CType(e.Argument, Task)
            Dim watchedTsk = Watcher.GetWatchedTaskById(tsk.TaskId)

            If IsNothing(watchedTsk) Then
                Watcher.Watch(tsk)
                watchedTsk = Watcher.GetWatchedTaskById(tsk.TaskId)
            End If

            If Not IsNothing(watchedTsk) Then
                If watchedTsk.Properties.IsEnabled Then AddNewHistoryRow(HistoryDataGridView.Rows.Count + 1, watchedTsk.Properties.TaskName, "Task started manually.", DateTime.Now)
                tsk.Execute()
            End If
        End Sub

        Friend Function RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) As Boolean
            Return True
        End Function
#End Region

        Private Sub PauseBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles PauseBarButton.ItemClick
            PauseOrResumeSelectedTask()
        End Sub

        Private Sub StopBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles StopBarButton.ItemClick
            StopSelectedTask()
        End Sub
        Private Sub StopAllBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles StopAllBarButton.ItemClick
            StopAllTasks()
        End Sub

        Private Sub BuyBarButtonItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BuyBarButtonItem.ItemClick
            OpenBuyNowLink()
        End Sub
#End Region

#Region "Groups Page Handlers"
        Private Sub NewGroupNameBarEditItem_EditValueChanged(sender As System.Object, e As System.EventArgs) Handles NewGroupNameBarEditItem.EditValueChanged
            If Not IsNothing(NewGroupNameBarEditItem.EditValue) Then RenameSelectedGroup(NewGroupNameBarEditItem.EditValue.ToString)
        End Sub

        Private Sub RenameTaskBarButtonBarEditItem_EditValueChanged(sender As System.Object, e As System.EventArgs) Handles RenameTaskBarButtonBarEditItem.EditValueChanged
            RenameSelectedTask()
        End Sub

        Private Sub DuplicateGroupBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles DuplicateGroupBarButton.ItemClick
            ShowWaitForm("Cloning group...")
            Dim selGroupId As String = GetSelectedGroupId()
            CacadoreToTectonic.CloneGroup(selGroupId)
            RefreshTasksTree()
            CloseWaitForm()
        End Sub

        Private Sub DeleteGroupBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles DeleteGroupBarButton.ItemClick
            Dim groupId As String = GetSelectedGroupId()
            ShowAlertForm("",
                          "Do you want to delete the group along with its tasks?",
                          AlertLocation,
                          "DeleteGroup",
                          True,
                          My.Resources.Cacadu_Image_46x46,
                          groupId)
        End Sub
#End Region

#Region "Tasks Page Handlers"
        Private Sub MyTasksBarButtonItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles MyTasksBarButtonItem.ItemClick
            TasksDockPanel.Show()
        End Sub

        Private Sub HistoryBarButtonItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles HistoryBarButtonItem.ItemClick
            HistoryDockPanel.Show()
        End Sub

        Private Sub ErrorsBarButtonItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles ErrorsBarButtonItem.ItemClick
            ErrorsParentDockPanel.Show()
        End Sub

        Private Sub TaskPropertiesBarButtonItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles TaskPropertiesBarButtonItem.ItemClick
            ShowTaskProperties()
        End Sub

        Private Sub DisableEnableTaskBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles DisableEnableTaskBarButton.ItemClick
            ToggleSelectedTaskEnableUponEnableRibbonButtonClick()
        End Sub

        Private Sub DeleteTaskBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles DeleteTaskBarButton.ItemClick
            ShowAlertForm("Cacadu",
                          "Confirm task delete?",
                          AlertLocation,
                          "DeleteTreeTask",
                          True, My.Resources.Cacadu_Image_46x46)
        End Sub

        Private Sub VariablesBarButtonItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles VariablesBarButtonItem.ItemClick
            VariablesAndConstantsConnector.SetSelectedTabPage(0)
            VariablesAndConstantsForm.ShowDialog()
        End Sub
#End Region

#Region "Tools Page Handlers"
        Private Sub EnableDisableCacadu_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles EnableDisableCacadu.ItemClick
            ToggleCacaduEnable()
        End Sub

        Private Sub GlobalVariablesBarButtonItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles GlobalVariablesBarButtonItem.ItemClick
            VariablesAndConstantsConnector.SetSelectedTabPage(1)
            VariablesAndConstantsForm.ShowDialog()
        End Sub

        Private Sub GlobalConstantsBarButtonItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles GlobalConstantsBarButtonItem.ItemClick
            VariablesAndConstantsConnector.SetSelectedTabPage(1)
            VariablesAndConstantsForm.ShowDialog()
        End Sub

        Private Sub OptionsBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles OptionsBarButton.ItemClick
            OptionsForm.ShowDialog()
        End Sub

        Private Sub ExportTasksBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles ExportTasksBarButton.ItemClick
            ExportTectonic()
        End Sub

        Private Sub ImportTasksBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles ImportTasksBarButton.ItemClick
            ImportTectonic()
        End Sub
#End Region

#Region "Help Page Handlers"
        Private Sub HelpContentsBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles HelpContentsBarButton.ItemClick
            ShowHelp()
        End Sub

        Private Sub RightHelpButtonBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles RightHelpButtonBarButton.ItemClick
            ShowHelp()
        End Sub

        Private Sub CheckUpdatesBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles CheckUpdatesBarButton.ItemClick
            CheckForUpdates("https://dl.dropbox.com/u/5721699/cacadu-appcast/versioninfo.xml", Me.Icon)
        End Sub

        Private Sub CacaduOnlineBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles CacaduOnlineBarButton.ItemClick
            Process.Start("http://cacadu.conderella.com")
        End Sub

        Private Sub FAQsBarButtons_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles FAQsBarButtons.ItemClick
            Process.Start("http://cacadu.conderella.com/support/faqs/")
        End Sub

        Private Sub VideoToursBarButtonItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles VideoToursBarButtonItem.ItemClick
            Process.Start("http://cacadu.conderella.com/video-tours/")
        End Sub
        Private Sub AboutBarButton_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles AboutBarButton.ItemClick
            AboutForm.ShowDialog()
        End Sub

        Private Sub RegisterBarButtonItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles RegisterBarButtonItem.ItemClick
            IntelliprotectorHelper.ShowRegistrationWindow()
        End Sub
#End Region

#Region "Caption Buttons Click Handlers"
        Private Sub StartGroupOfHomePage_CaptionButtonClick(sender As System.Object, e As DevExpress.XtraBars.Ribbon.RibbonPageGroupEventArgs) Handles StartGroupOfHomePage.CaptionButtonClick
            AddNewTaskForm.ShowDialog()
        End Sub

        Private Sub RunGroupOfHomePage_CaptionButtonClick(sender As System.Object, e As DevExpress.XtraBars.Ribbon.RibbonPageGroupEventArgs) Handles RunGroupOfHomePage.CaptionButtonClick
            RunSelectedTask()
        End Sub

        Private Sub EditTasksGroupOfHomePage_CaptionButtonClick(sender As System.Object, e As DevExpress.XtraBars.Ribbon.RibbonPageGroupEventArgs) Handles EditTasksGroupOfHomePage.CaptionButtonClick
            ShowTasksList()
        End Sub

        Private Sub BuyRibbonPageGroup_CaptionButtonClick(sender As System.Object, e As DevExpress.XtraBars.Ribbon.RibbonPageGroupEventArgs) Handles BuyRibbonPageGroup.CaptionButtonClick
            OpenBuyNowLink()
        End Sub

        Private Sub EditGroupsOfGroupsPage_CaptionButtonClick(sender As System.Object, e As DevExpress.XtraBars.Ribbon.RibbonPageGroupEventArgs) Handles EditGroupsOfGroupsPage.CaptionButtonClick
            AddNewGroupForm.ShowDialog()
        End Sub

        Private Sub TaskEditGroupOfTasksPage_CaptionButtonClick(sender As System.Object, e As DevExpress.XtraBars.Ribbon.RibbonPageGroupEventArgs) Handles TaskEditGroupOfTasksPage.CaptionButtonClick
            TasksDockPanel.Show()
        End Sub

        Private Sub TaskExecutionGroupOfTasksPage_CaptionButtonClick(sender As System.Object, e As DevExpress.XtraBars.Ribbon.RibbonPageGroupEventArgs) Handles TaskExecutionGroupOfTasksPage.CaptionButtonClick
            HistoryDockPanel.Show()
        End Sub

        Private Sub ToolsGroupOfToolsPage_CaptionButtonClick(sender As System.Object, e As DevExpress.XtraBars.Ribbon.RibbonPageGroupEventArgs) Handles ToolsGroupOfToolsPage.CaptionButtonClick
            OptionsForm.ShowDialog()
        End Sub

        Private Sub HelpGroupOfHelpPage_CaptionButtonClick(sender As System.Object, e As DevExpress.XtraBars.Ribbon.RibbonPageGroupEventArgs) Handles HelpGroupOfHelpPage.CaptionButtonClick
            ShowHelp()
        End Sub
#End Region
#End Region
#End Region
#Region "Treeview Handlers"
        Private Sub TasksTreeView_DoubleClick(sender As System.Object, e As System.EventArgs) Handles TasksTreeView.DoubleClick
            GetSelectedTaskFlow()
        End Sub

        Private Sub SearchTasksMruEdit_EditValueChanging(sender As System.Object, e As DevExpress.XtraEditors.Controls.ChangingEventArgs) Handles SearchTasksMruEdit.EditValueChanging
            ClearTasksTreeBackColor()
            SearchTasksTree(e.NewValue.ToString)
        End Sub

        Private Sub SearchTasksMruEdit_ButtonClick(sender As System.Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles SearchTasksMruEdit.ButtonClick
            If Not IsNothing(SearchTasksMruEdit.EditValue) Then
                ClearTasksTreeBackColor()
                SearchTasksTree(SearchTasksMruEdit.EditValue.ToString)
            End If
        End Sub

        Private Sub TasksTreeView_KeyDown(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles TasksTreeView.KeyDown
            If IsNothing(GetSelectedNode()) Then Exit Sub
            Select Case e.KeyData
                Case Keys.Delete
                    Dim nodeType As String = GetSelectedNodeType()
                    If nodeType = "Task" Then
                        ShowTaskDeleteAlert()
                    ElseIf nodeType = "Group" Then
                        ShowGroupDeleteAlert()
                    End If
                Case Keys.F2
                    GetSelectedNode().BeginEdit()
            End Select
        End Sub

        Private Sub TasksTreeView_MouseUp(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles TasksTreeView.MouseUp
            Dim _isNotFreeArea = IsFreeArea(e.X, e.Y)
            Select Case e.Button
                Case Windows.Forms.MouseButtons.Left
                    If IsNothing(_isNotFreeArea) Then FilterTasksList("")
                Case Windows.Forms.MouseButtons.Right
                    Try
                        If Not IsNothing(_isNotFreeArea) Then TasksTreeView.SelectedNode = _isNotFreeArea
                        TasksTreeContextMenuStrip.Show(TasksTreeView, e.X, e.Y)
                    Catch ex As Exception
                        Debug.Print("ASDFASDF")
                    End Try

            End Select
        End Sub

        Private Sub TasksTreeView_AfterLabelEdit(sender As System.Object, e As System.Windows.Forms.NodeLabelEditEventArgs) Handles TasksTreeView.AfterLabelEdit
            RenameTreeNode(e)
        End Sub

        Private Sub TasksDockPanel_Click(sender As System.Object, e As System.EventArgs) Handles TasksDockPanel.Click
            FilterTasksList("") 'Reset filter to display all nodes
        End Sub

        Private Sub TasksTreeView_AfterSelect(sender As Object, e As System.Windows.Forms.TreeViewEventArgs) Handles TasksTreeView.AfterSelect
            OnAfterNodeSelect(e.Node.Name, e.Node.Tag.ToString, e)
        End Sub
#End Region

#Region "TasksListDataGridView"
        Private Sub TasksListDataGridView_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles TasksListDataGridView.CellClick
            SetSelectedNode(GetSelectedTaskIdInTasksList())

            If e.ColumnIndex = 5 OrElse e.ColumnIndex = 4 Then DropComboColumn(e.ColumnIndex) 'To fix a bug where you must double click the combo cell to pull it down.
        End Sub

        Private Sub TasksListDataGridView_CellDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles TasksListDataGridView.CellDoubleClick
            If e.ColumnIndex <> 5 AndAlso e.ColumnIndex <> 4 Then
                Dim tsk = GetSelectedTaskInTasksList()
                If Not IsNothing(tsk) Then DisplayTaskFlow(tsk.TaskId, tsk.Properties.TaskName, "Task")
            End If
        End Sub
#End Region

#Region "XtraTabs"
        Private Sub XtraTabsOfMainForm_DragOver(sender As Object, e As System.Windows.Forms.DragEventArgs) Handles XtraTabsOfMainForm.DragOver
            If e.Data.GetDataPresent(GetType(GalleryItem)) Then e.Effect = DragDropEffects.Copy
        End Sub

        Private Sub XtraTabsOfMainForm_DragDrop(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles XtraTabsOfMainForm.DragDrop
            OnXtraTabsOfMainFormDragDrop(sender, e)
        End Sub

        Private Sub XtraTabsOfMainForm_SelectedPageChanged(sender As System.Object, e As DevExpress.XtraTab.TabPageChangedEventArgs) Handles XtraTabsOfMainForm.SelectedPageChanged
            If IsNothing(e.Page) Then Exit Sub
            Dim selectedTabPageName As String = e.Page.Name
            ChangeTaskProperitesBasedOnSelecetedTabPage(selectedTabPageName)
            ChangeRelatedTaskButtonsEnableOnTabPageChanged(selectedTabPageName)
        End Sub

        Private Sub XtraTabs_CloseButtonClick(sender As System.Object, e As System.EventArgs) Handles XtraTabsOfMainForm.CloseButtonClick
            Dim _closePageButtonEventArgs As DevExpress.XtraTab.ViewInfo.ClosePageButtonEventArgs =
                CType(e, DevExpress.XtraTab.ViewInfo.ClosePageButtonEventArgs)
            Dim _tapPage = CType(_closePageButtonEventArgs.Page, DevExpress.XtraTab.XtraTabPage)
            CloseTaskFlow(_tapPage)
        End Sub
#End Region

#Region "Property Grid(s)"
        Private Sub TaskPropertiesPropertyGrid_PropertyValueChanged(s As System.Object, e As System.Windows.Forms.PropertyValueChangedEventArgs) Handles TaskPropertiesPropertyGrid.PropertyValueChanged
            SaveTaskPropertiesPropertyGrid()
            OnTasksTreeUpdated()
        End Sub
#End Region

#Region "Actions Gallery"
        Private Sub ActionsRibbonGallery_GalleryItemClick(sender As System.Object, e As DevExpress.XtraBars.Ribbon.GalleryItemClickEventArgs) Handles ActionsRibbonGallery.GalleryItemClick
            ShowActionForm(CStr(e.Item.Tag))
        End Sub
#End Region

#Region "Alert Form"

#End Region

#Region "Context Menus"

#Region "TasksTreeContextMenuStrip"
        Shared Sub ToggleTreeContextMenuMode(mode As String)
            Dim taskVisiable As Boolean = False
            Dim groupVisiable As Boolean = False
            With MainForm
                Select Case mode
                    Case "Task"
                        taskVisiable = True
                        groupVisiable = False
                    Case "Group"
                        taskVisiable = False
                        groupVisiable = True
                End Select
                .RunTaskToolStripMenuItem.Visible = taskVisiable

                .RenameTaskToolStripMenuItem.Visible = taskVisiable
                .DeleteTaskToolStripMenuItem.Visible = taskVisiable
                .DashToolStripMenuItem2.Visible = taskVisiable
                .TaskFlowToolStripMenuItem.Visible = taskVisiable
                .TriggersToolStripMenuItem.Visible = taskVisiable
                .SaveToolStripMenuItem.Visible = taskVisiable
                .StopToolStripMenuItem.Visible = taskVisiable
                .TaskPropertiesToolStripMenuItem.Visible = taskVisiable
                .ToolStripSeparator1.Visible = taskVisiable

                '.AddGroupToolStripMenuItem.Visible = groupVisiable
                .RenameGroupToolStripMenuItem.Visible = groupVisiable
                .DeleteGroupToolStripMenuItem.Visible = groupVisiable
                .DashToolStripMenuItem3.Visible = groupVisiable
                .AddTaskToolStripMenuItem.Visible = groupVisiable
            End With
        End Sub

        Private Sub TasksTreeContextMenuStrip_Opening(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles TasksTreeContextMenuStrip.Opening
            ToggleTreeContextMenuMode(GetSelectedNodeType())
        End Sub

        Private Sub SaveToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SaveToolStripMenuItem.Click
            GetSelectedTask().Save()
        End Sub

        Private Sub TasksListToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles TasksListToolStripMenuItem.Click
            ShowTasksList()
        End Sub

        Private Sub RunTaskToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles RunTaskToolStripMenuItem.Click
            RunSelectedTask()
        End Sub

        Private Sub StopToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StopToolStripMenuItem.Click
            StopSelectedTask()
        End Sub

        Private Sub AddTaskToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AddTaskToolStripMenuItem.Click
            AddNewTaskForm.ShowDialog()
        End Sub

        Private Sub RenameTaskToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles RenameTaskToolStripMenuItem.Click
            BeginEditSelectedNode()
        End Sub

        Private Sub DeleteTaskToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DeleteTaskToolStripMenuItem.Click
            ShowTaskDeleteAlert()
        End Sub

        Private Sub AddGroupToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AddGroupToolStripMenuItem.Click
            AddNewGroupForm.ShowDialog()
        End Sub

        Private Sub RenameGroupToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles RenameGroupToolStripMenuItem.Click
            BeginEditSelectedNode()
        End Sub

        Private Sub DeleteGroupToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DeleteGroupToolStripMenuItem.Click
            ShowGroupDeleteAlert()
        End Sub

        Private Sub ExportTToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExportTToolStripMenuItem.Click
            ExportTectonic()
        End Sub

        Private Sub ImportTasksToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ImportTasksToolStripMenuItem.Click
            ImportTectonic()
        End Sub

        Private Sub TaskPropertiesToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles TaskPropertiesToolStripMenuItem.Click
            ShowTaskProperties()
        End Sub

        Private Sub TaskFlowToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles TaskFlowToolStripMenuItem.Click
            GetSelectedTaskFlow()
        End Sub

        Private Sub TriggersToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles TriggersToolStripMenuItem.Click
            OnTriggersButtonClick()
        End Sub
#End Region

#Region "SystrayContextMenuStrip"
        Private Sub SystrayContextMenuStrip_Opening(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles SystrayContextMenuStrip.Opening
            ToggleEnableDisableSystrayStrip(Cacadore.Settings.Enabled)
        End Sub

        Private Sub ShowCacaduToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ShowCacaduToolStripMenuItem.Click
            ToggleFormVisible()
        End Sub

        Private Sub EnableDisableCacaduToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EnableDisableCacaduToolStripMenuItem.Click
            ToggleCacaduEnable()
        End Sub

        Private Sub ExitToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExitToolStripMenuItem.Click
            'CloseCacadu()
            Application.Exit()
        End Sub
#End Region
#End Region

#Region "Systray Icon"
        Private Sub SystrayIcon_MouseDown(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles SystrayIcon.MouseDown
            ToggleSystrayMenuShownItemText()
        End Sub

        Private Sub SystrayIcon_MouseDoubleClick(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles SystrayIcon.MouseDoubleClick
            ShowMainForm(Me)
        End Sub
#End Region

#Region "Quick Access"
        Private Sub OpenTaskQuickAccess_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles OpenTaskQuickAccess.ItemClick
            GetSelectedTaskFlow()
        End Sub

        Private Sub SaveTaskQuickAccess_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles SaveTaskQuickAccess.ItemClick
            GetSelectedTask().Save()
        End Sub

        Private Sub OptionsQuickAccess_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles OptionsQuickAccess.ItemClick
            OptionsForm.ShowDialog()
        End Sub
#End Region

#Region "Main Menu"
        Private Sub NewGroupMainMenuBarButtonItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles NewGroupMainMenuBarButtonItem.ItemClick
            AddNewGroupForm.ShowDialog()
        End Sub

        Private Sub NewTaskBarButtonItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles NewTaskBarButtonItem.ItemClick
            AddNewTaskForm.ShowDialog()
        End Sub

        Private Sub SaveTaskApplicationMenu_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles SaveTaskApplicationMenu.ItemClick
            GetSelectedTask().Save()
        End Sub

        Private Sub ImportTasksBarButtonItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles ImportTasksBarButtonItem.ItemClick
            ImportTectonic()
        End Sub

        Private Sub ExportTasksBarButtonItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles ExportTasksBarButtonItem.ItemClick
            ExportTectonic()
        End Sub

        Private Sub CloseCacaduApplicationMenu_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles CloseCacaduApplicationMenu.ItemClick
            CloseCacadu()
        End Sub
#End Region

#Region "Status bar"
        Private Sub DefaultBarButtonItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles DefaultBarButtonItem.ItemClick
            Themes.Skins.OnThemeItemClick("Default")
        End Sub

        Private Sub RandomBarButtonItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles RandomBarButtonItem.ItemClick
            Themes.Skins.OnThemeItemClick("Random")
        End Sub

        Private Sub RightStatusBarLinkBarStaticItem_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs)
            Process.Start(My.Resources.CacaduWebSite)
        End Sub

        Private Sub RibbonStatusBar_MouseMove(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles RibbonStatusBar.MouseMove
            RibbonStatusBar.Cursor = Cursors.Default
        End Sub
#End Region
    End Class
End Namespace