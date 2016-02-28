Imports Cacadu.UI
Imports Cacadore.Commands
Imports Cacadu.UI.MainForm
Imports DevExpress.XtraBars.Alerter
Imports System.Runtime.InteropServices
Imports TecDAL.TectonicDataSet
Imports TecDAL
Imports DevExpress.XtraBars.Ribbon
Imports DevExpress.XtraBars.Ribbon.ViewInfo
Imports Cacadore
Imports System.ComponentModel

'الفايدة العامة من إستخدام كونيكتور هو الفصل بين مدخلات واجهة 
'الاستخدام وبين مكونات الواجهة نفسها. بمعني ان الكونيكتور بيتولي 
'تنفيذ مهام الواجهة وفي نفس الوقت بيتلقي المدخلات من المكونات من 
'غير ما يتعامل معاها بشكل مباشر. وبالتالي في أي وقت مش هيبقي 
'صعب تغيير الواجهة برمتها. اعمل اي واجهة جديدة وابعت مدخلاتها 
'للكونيتكتور. 

'مما سبق يتضح ان ماينفعش الكونيتكتور يعرف اي حاجة عن العناصر 
'المستخدمة في الواجهة وخصوصا في تمريرات الوظائف والا هتكسر 
'القاعدة السابقة.

Namespace Connectors
    Friend Class MainFormConnector
        Implements Balora.IListener

        Const ZeroGroupIndex As String = "0"

#Region "Events & Events Handlers"
        Private Shared Event TasksTreeUpdated()

        Private Sub MainFormConnector_TasksTreeUpdated() Handles Me.TasksTreeUpdated
            RefreshTasksTree()
        End Sub
#End Region

#Region "Fields"
        Friend Shared _typedDataSetCrud As TecDAL.TypedDataSetCrud = CType(Cacadore.Settings.CrudObject, TecDAL.TypedDataSetCrud)
#End Region

#Region "MainForm Helpers"
#Region "Main Form"
        ''' <summary>
        ''' Load event of MainForm
        ''' </summary>
        Shared Sub LoadMainForm()
            SetIconToCacaduIcon(MainForm)
            SetMainFormCaption()
            _typedDataSetCrud.AddListener(New MainFormConnector)
            OnTasksTreeUpdated()
            InitTasksList()
            DrawDuplicatedRibbonButtons()
            AddPausedTasksToWatcher()
            ToggleEnableDisableCacaduRibbonButton(Cacadore.Settings.Enabled)
        End Sub

        Shared Sub CheckIfSplashScreenEnabled()
            If My.Settings.SS Then CacaduGloby.CloseSplashScreen()
        End Sub

        Shared Sub SendToTrayIfMinCommandLineIncluded()
            If My.Application.CommandLineArgs.Count > 0 AndAlso My.Application.CommandLineArgs.Item(0) = "min" Then SendToSystray()
        End Sub

        Shared Sub HideSomeButtonsIfRegistered()
            With MainForm
                If IntelliprotectorHelper.IsCacaduRegistered() Then
                    .BuyRibbonPageGroup.Visible = False
                    .BuyBarButtonItem.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
                    .RegisterBarButtonItem.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
                End If
            End With
        End Sub

        ''' <summary>
        ''' Setting the caption of MainForm @ Mainform_Load
        ''' </summary>
        Shared Sub SetMainFormCaption()
            Dim appName As String = String.Format("{0} {1}", Application.ProductName, Application.ProductVersion)
#If DEBUG Then
            MainForm.Text = appName & "::Debug Mode"
#End If
        End Sub

        'So we can toggle pause button.
        Shared Sub AddPausedTasksToWatcher()
            Dim tskStatesDic = CacadoreToTectonic.GetAllRegisteredTasksStates()
            Dim queryResult = From sta In tskStatesDic.Values Where sta.TaskExecutionState = Cacadore.ITaskState.ExecutionState.Paused
            If queryResult.Count > 0 Then
                For Each one In queryResult
                    Cacadore.Watcher.Watch(one.taskId, True)
                Next
            End If
        End Sub

        Shared Sub ShowMainForm(mainformObj As Form)
            mainformObj.Show()
            mainformObj.WindowState = FormWindowState.Normal
        End Sub

        Shared Sub SendToSystray()
            MainForm.WindowState = FormWindowState.Minimized
        End Sub

        Shared Sub HideMe(mainformObj As Form)
            If mainformObj.WindowState = FormWindowState.Minimized Then
                mainformObj.Hide()
            End If
        End Sub

        Shared Sub ToggleSystrayMenuShownItemText()
            With MainForm
                If .WindowState = FormWindowState.Minimized Then
                    .ShowCacaduToolStripMenuItem.Text = "Show Cacadu"
                Else
                    .ShowCacaduToolStripMenuItem.Text = "Hide Cacadu"
                End If
            End With
        End Sub

        Shared Sub ToggleFormVisible()
            With MainForm
                If .ShowCacaduToolStripMenuItem.Text = "Show Cacadu" Then
                    ShowMainForm(MainForm)
                Else
                    SendToSystray()
                    HideMe(MainForm)
                End If
            End With
        End Sub

        Shared Sub SetMainFormPlacement()
            With MainForm
                .Width = My.Settings.CW
                .Height = My.Settings.CH
                If My.Settings.CT < 0 Then
                    .Top = 50
                Else
                    .Top = My.Settings.CT
                End If
                If My.Settings.CL < 0 Then
                    .Left = 0
                Else
                    .Left = My.Settings.CL
                End If
            End With
        End Sub

        Shared Sub SaveMainFormPlacement()
            With MainForm
                My.Settings.CW = .Width
                My.Settings.CH = .Height
                My.Settings.CT = .Top
                My.Settings.CL = .Left
                My.Settings.Save()
            End With
        End Sub

        Shared Sub LoadLayout()
            Dim layoutPath As String = Balora.PathsHelper.GetCurrentExecutingDirectory & "layout.xml"
            Dim quickAccessLayoutPath As String = Balora.PathsHelper.GetCurrentExecutingDirectory & "R-Layout.xml"
            If IO.File.Exists(layoutPath) Then
                MainForm.DockManager.RestoreLayoutFromXml(layoutPath)
            End If

            If IO.File.Exists(quickAccessLayoutPath) Then
                MainForm.RibbonControlOfMainForm.Toolbar.RestoreLayoutFromXml(quickAccessLayoutPath)
            End If
        End Sub

        Shared Sub SaveLayout()
            MainForm.DockManager.SaveLayoutToXml(Balora.PathsHelper.GetCurrentExecutingDirectory & "layout.xml")
            MainForm.RibbonControlOfMainForm.Toolbar.SaveLayoutToXml("R-Layout.xml")
        End Sub
#End Region

#Region "Ribbon Related"
        Shared Sub DrawDuplicatedRibbonButtons()
            With MainForm
                .EditGroupsOfGroupsPage.ItemLinks.Insert(.EditGroupsOfGroupsPage.ItemLinks(0), .AddGroupHomePageBarSubItem)
                .TaskEditGroupOfTasksPage.ItemLinks.Insert(.TaskEditGroupOfTasksPage.ItemLinks(1), .AddTaskHomePageBarSubItem)
                .TaskEditGroupOfTasksPage.ItemLinks.Insert(.TaskEditGroupOfTasksPage.ItemLinks(4), .TasksListBarButtonItem)

                .TaskExecutionGroupOfTasksPage.ItemLinks.Insert(.TaskExecutionGroupOfTasksPage.ItemLinks(0), .TaskFlowHomePageBarButton)
                .TaskExecutionGroupOfTasksPage.ItemLinks.Insert(.TaskExecutionGroupOfTasksPage.ItemLinks(1), .TriggersHomePageBarButton)
                .TaskExecutionGroupOfTasksPage.ItemLinks.Insert(.TaskExecutionGroupOfTasksPage.ItemLinks(2), .RunHomePageBarButton)
                .TaskExecutionGroupOfTasksPage.ItemLinks.Insert(.TaskExecutionGroupOfTasksPage.ItemLinks(3), .PauseBarButton)
                .TaskExecutionGroupOfTasksPage.ItemLinks.Insert(.TaskExecutionGroupOfTasksPage.ItemLinks(4), .StopBarButton)
                .TaskExecutionGroupOfTasksPage.ItemLinks.Insert(.TaskExecutionGroupOfTasksPage.ItemLinks(5), .StopAllBarButton)

                .HelpGroupOfHelpPage.ItemLinks.Insert(.TaskExecutionGroupOfTasksPage.ItemLinks(5), .BuyBarButtonItem)
            End With
        End Sub

        Shared Sub ChangeRelatedTaskButtonsEnableOnTabPageChanged(selectedTabPageName As String)
            'Changes the enabled for task related buttons. Related buttons are like Triggers And Task Flow ribbon buttons.
            Dim enabled As Boolean = False
            Dim tasksListTabPageName As String = GetTasksListTabPageName()
            'If tab page is not Tasks List (i.e Task Flow) then enable buttons...
            If Not selectedTabPageName.Contains(tasksListTabPageName) Then enabled = True
            ChangeRelatedTaskButtonEnable(enabled)
        End Sub

        Shared Sub ChangeRelatedTaskButtonsEnableOnNodeSelect(nodeType As String)
            Dim enabled As Boolean = False
            If nodeType = "Task" Then enabled = True
            ChangeRelatedTaskButtonEnable(enabled)
        End Sub

        Shared Sub ChangeRelatedTaskButtonEnable(enabled As Boolean)
            With MainForm
                .TriggersHomePageBarButton.Enabled = enabled
                .TaskFlowHomePageBarButton.Enabled = enabled
                Dim selTab = GetSelectedTabPageTaskId()
                If selTab.Contains("TasksListXtraTabPage") Then
                    .ActionsRibbonGallery.Enabled = False
                Else
                    .ActionsRibbonGallery.Enabled = enabled
                End If
            End With
        End Sub

        Shared Sub OnTriggersButtonClick()
            Dim taskId As String = GetSelectedTabPageTaskId()
            Dim isRegisteredTask = CacadoreToTectonic.IsRegisteredTaskId(taskId)
            If Not isRegisteredTask Then
                taskId = GetSelectedNodeTaskId()
            End If
            TriggersFormConnector.PrepareAndDisplayTriggersForm(taskId)
        End Sub

        Shared Sub ToggleTasksTabRelatedBarButtonEnable(enable As Boolean)
            With MainForm
                If .RenameTaskBarButton.Enabled <> enable Then .RenameTaskBarButton.Enabled = enable
                If .DuplicateTaskButtonBarButton.Enabled <> enable Then .DuplicateTaskButtonBarButton.Enabled = enable
                If .DeleteTaskBarButton.Enabled <> enable Then .DeleteTaskBarButton.Enabled = enable
                If .DisableEnableTaskBarButton.Enabled <> enable Then .DisableEnableTaskBarButton.Enabled = enable
                If .TaskPropertiesBarButtonItem.Enabled <> enable Then .TaskPropertiesBarButtonItem.Enabled = enable
                If .RunHomePageBarButton.Enabled <> enable Then .RunHomePageBarButton.Enabled = enable
            End With
        End Sub

        Shared Sub ToggleGroupTabRelatedBarButtonEnable(enable As Boolean)
            With MainForm
                If .RenameBarSubItem.Enabled <> enable Then .RenameBarSubItem.Enabled = enable
                If .DeleteGroupBarButton.Enabled <> enable Then .DeleteGroupBarButton.Enabled = enable
                If .DuplicateGroupBarButton.Enabled <> enable Then .DuplicateGroupBarButton.Enabled = enable
            End With
        End Sub

        Shared Sub ToggleDisableEnableRibbonButton(isEnabled As Boolean)
            With MainForm
                If isEnabled Then
                    .DisableEnableTaskBarButton.Caption = "Disable"
                Else
                    .DisableEnableTaskBarButton.Caption = "Enable"
                End If
            End With
        End Sub

        Shared Sub ToggleEnableDisableCacaduRibbonButton(enable As Boolean)

            With MainForm
                If enable Then
                    .EnableDisableCacadu.Caption = "Disable Cacadu"
                    .EnableDisableCacadu.LargeGlyph = My.Resources.switch_on
                Else
                    .EnableDisableCacadu.Caption = "Enable Cacadu"
                    WriteInStatusBar(My.Resources.CD, Color.Red)
                    MainForm.SystrayIcon.ShowBalloonTip(3000,
                                                        "Cacadu",
                                                        "Cacadu is disabled, please enable it in order to execute your tasks.",
                                                        ToolTipIcon.Info)
                    .EnableDisableCacadu.LargeGlyph = My.Resources.switch_off
                End If
            End With
        End Sub

        Shared Sub DisableCacadu()
            Cacadore.Settings.Save(ISettings.SettingsEnum.CacadoreEnabled, False)
            Helpers.QuartzHelper.QuartzUtils.PauseAllTriggers()
            Cacadore.Settings.Enabled = False
            WriteInStatusBar(My.Resources.CD, Color.Red)
        End Sub

        Shared Sub EnableCacadu()
            Dim trgsStarter As New TriggersStarter
            Cacadore.Settings.Save(ISettings.SettingsEnum.CacadoreEnabled, True)
            Cacadore.Settings.Enabled = True
            If trgsStarter.TryStartTriggers() Then
                Helpers.QuartzHelper.QuartzUtils.ResumeAllTriggers()
            End If
            WriteInStatusBar(My.Resources.CE, Color.Blue)
        End Sub

        Shared Sub ToggleCacaduEnable()
            Dim isEnabled As Boolean = Cacadore.Settings.Enabled

            If isEnabled Then
                'We will disable
                'To disable:-
                'Find how to disable all triggers QuartzUtils.PauseAllTriggers.
                'Pause all running tasks in Watcher.
                'Write in log that Cacadu is disabled.
                DisableCacadu()
            ElseIf Not isEnabled Then
                'We will enable
                'To Enable:-
                'Resume All Triggers.
                'Resume All paused tasks in Watcher
                EnableCacadu()
            End If

            ToggleEnableDisableCacaduRibbonButton(Cacadore.Settings.Enabled)
        End Sub

        Shared Sub StopAllTasks()
            If Watcher.WatchedTasks.Count <= 0 Then Exit Sub
            For Each tsk In Watcher.WatchedTasks.Values
                tsk.Cancel()
            Next
            ShowAlertForm("",
                          "All running tasks have stopped, do you want to pause their triggers?",
                          AlertLocation,
                          "PauseAllTriggers",
                          True,
                          My.Resources.Cacadu_Image_46x46)
        End Sub

        Shared Sub OpenBuyNowLink()
            Process.Start("http://cacadu.conderella.com/order/")
        End Sub

        Shared Sub ShowHelp()
            Process.Start("CacaduHelp.chm")
        End Sub
#End Region

#Region "XtraTabs"
        Shared Sub DisplayAllTasksInTasksList()
            FilterTasksList("") 'Reset filter to display all nodes
        End Sub

        Shared Sub OnAfterNodeSelect(nodeKey As String, nodeTag As String, Optional e As System.Windows.Forms.TreeViewEventArgs = Nothing)
            ChangeRelatedTaskButtonsEnableOnNodeSelect(nodeTag)
            Select Case nodeTag
                Case "Task"
                    Dim isEnabled As Boolean = False
                    ShowOneTaskInTasksList(nodeKey)
                    LoadTaskPropertiesInPropertyGrid(nodeKey, isEnabled)
                    ToggleGroupTabRelatedBarButtonEnable(False)
                    ToggleTasksTabRelatedBarButtonEnable(True)
                    ToggleDisableEnableRibbonButton(isEnabled)
                    If Not IsNothing(e) Then ToggleTaskNodeForeColor(e.Node, isEnabled)
                    MainForm.TasksTreeView.SelectedImageIndex = 1
                Case ("Group")
                    FilterTasksList(String.Format("group_id = '{0}'", nodeKey))
                    UnloadTaskPropertiesInPropertyGrid()
                    ToggleGroupTabRelatedBarButtonEnable(True)
                    ToggleTasksTabRelatedBarButtonEnable(False)
                    MainForm.TasksTreeView.SelectedImageIndex = 0
            End Select
        End Sub

        Shared Function GetTasksListTabPageName() As String
            Dim tasksListTabPageName As String = MainForm.XtraTabsOfMainForm.TabPages(0).Name
            Return tasksListTabPageName
        End Function

        Shared Function IsTasksListsTabPageSelected() As Boolean
            Return Not IsTaskTabSelected()
        End Function

        Shared Function IsTaskTabSelected() As Boolean
            Dim sel = GetSelectedTabPage.Name
            Dim tskLstsName As String = GetTasksListTabPageName()
            If tskLstsName.Contains(sel) Then
                Return False
            Else
                Return True
            End If
        End Function

        Shared Sub ShowTasksList()
            MainForm.XtraTabsOfMainForm.TabPages(0).PageVisible = True
            MainForm.XtraTabsOfMainForm.SelectedTabPageIndex = 0
        End Sub

        Shared Sub HideTasksList()
            MainForm.XtraTabsOfMainForm.TabPages(0).PageVisible = False
        End Sub

        Shared Function IsAnXtraPageDisplayed(ByVal taskId As String) As DevExpress.XtraTab.XtraTabPage
            Dim queryTasksTreeResult = From tabPage In MainForm.XtraTabsOfMainForm.TabPages
                                       Select tabPage
                                       Where tabPage.Name = taskId
            If queryTasksTreeResult.Count <> 0 Then
                Return queryTasksTreeResult(0)
            Else
                Return Nothing
            End If
        End Function

        Shared Sub AddTaskFlowToXtraTab(taskFlowXtraTabPage As DevExpress.XtraTab.XtraTabPage)
            Dim taskFlowCont As New Cacadu.TaskFlowControl
            taskFlowCont.Dock = System.Windows.Forms.DockStyle.Fill
            taskFlowCont.Location = New System.Drawing.Point(0, 0)
            taskFlowCont.Name = "TaskFlowControl1"
            taskFlowCont.Size = New System.Drawing.Size(571, 188)
            taskFlowCont.TabIndex = 0
            taskFlowXtraTabPage.Controls.Add(taskFlowCont)
            'This a bit complicated, after the user dragged an action and drop it in the task flow, we 
            'just wait for the label created in "OnXtraTabsOfMainFormDragDrop" to be added. Then
            'We display the target action form after that.
            AddHandler taskFlowXtraTabPage.ControlAdded, AddressOf DisplayDraggedAction
            AddHandler taskFlowXtraTabPage.Disposed, AddressOf taskFlowXtraTabPage_Disposed
        End Sub

        'This a bit complicated, after the user dragged an action and drop it in the task flow, we 
        'just wait for the label created in "OnXtraTabsOfMainFormDragDrop" to be added. Then
        'We display the target action form after that.
        Friend Shared Sub DisplayDraggedAction(sender As Object, e As ControlEventArgs)
            Dim selXtraTabPage As DevExpress.XtraTab.XtraTabPage = CType(sender, DevExpress.XtraTab.XtraTabPage)
            ShowActionForm(e.Control.Text)
            e.Control.Dispose()
        End Sub

        Private Shared Sub taskFlowXtraTabPage_Disposed(sender As Object, e As EventArgs)
            Dim selXtraTabPage As DevExpress.XtraTab.XtraTabPage = CType(sender, DevExpress.XtraTab.XtraTabPage)
            RemoveHandler selXtraTabPage.ControlAdded, AddressOf MainFormConnector.DisplayDraggedAction
            RemoveHandler selXtraTabPage.Disposed, AddressOf MainFormConnector.taskFlowXtraTabPage_Disposed
        End Sub

        Shared Sub DrawNewXtraTabPage(ByVal taskId As String, ByVal taskName As String)
            MainForm.TasksListXtraTabPage.SuspendLayout()
            'XtraTabPage1
            Dim taskFlowXtraTabPage = New DevExpress.XtraTab.XtraTabPage With {.Name = taskId,
                                                                               .Size = New System.Drawing.Size(653, 188),
                                                                               .Text = taskName & " Task Flow"}
            AddTaskFlowToXtraTab(taskFlowXtraTabPage)

            With MainForm.XtraTabsOfMainForm
                .TabPages.Add(taskFlowXtraTabPage)
                .SelectedTabPage = taskFlowXtraTabPage
            End With

            TaskFlowPresenter.FillTaskList(taskId)
            MainForm.TasksListXtraTabPage.ResumeLayout()
        End Sub

        Shared Sub DisplayTaskFlow(ByVal taskId As String, ByVal taskName As String, ByVal nodeType As String)
            Dim isXtraTabThatAlreadyDisplayed = IsAnXtraPageDisplayed(taskId)
            Select Case nodeType
                Case "Task"
                    'Task flow for this task isn't displayed.
                    If IsNothing(isXtraTabThatAlreadyDisplayed) Then
                        DrawNewXtraTabPage(taskId, taskName)
                    Else 'Task flow for this task is not visible but created or may be disposed.
                        If isXtraTabThatAlreadyDisplayed.IsHandleCreated Then
                            isXtraTabThatAlreadyDisplayed.PageVisible = True
                        Else
                            DrawNewXtraTabPage(taskId, taskName)
                        End If
                        MainForm.XtraTabsOfMainForm.SelectedTabPage = isXtraTabThatAlreadyDisplayed
                    End If
            End Select
        End Sub

        Shared Function GetSelectedTabPage() As DevExpress.XtraTab.XtraTabPage
            Return MainForm.XtraTabsOfMainForm.SelectedTabPage
        End Function

        Shared Function GetSelectedTabPageTaskId() As String
            If Not IsNothing(MainForm.XtraTabsOfMainForm.SelectedTabPage) Then
                Return MainForm.XtraTabsOfMainForm.SelectedTabPage.Name()
            Else
                Return ""
            End If
        End Function

        Shared Sub CloseTaskFlow(page As DevExpress.XtraTab.XtraTabPage)
            Dim _tabPage As DevExpress.XtraTab.XtraTabPage = CType(page, DevExpress.XtraTab.XtraTabPage)
            _tabPage.PageVisible = False
        End Sub

        Shared Sub OnXtraTabsOfMainFormDragDrop(sender As System.Object, e As System.Windows.Forms.DragEventArgs)
            Dim galleryAction As GalleryItem = TryCast(e.Data.GetData(GetType(GalleryItem)), GalleryItem)
            Dim selectedTab = GetSelectedTabPage()
            Dim newLabel As New Label
            With newLabel
                .Text = galleryAction.Tag.ToString
                .Location = TryCast(sender, Control).PointToClient(New Point(e.X, e.Y))
            End With
            selectedTab.Controls.Add(newLabel)
        End Sub

        Shared Function GetSelectedTabTaskFlowControl() As TaskFlowControl
            Dim tskFlowControl As TaskFlowControl = Nothing
            Dim selTabPage = Connectors.MainFormConnector.GetSelectedTabPage()
            For Each [control] In selTabPage.Controls
                If TypeOf ([control]) Is TaskFlowControl Then
                    tskFlowControl = CType([control], TaskFlowControl)
                    Return tskFlowControl
                End If
            Next
            Return Nothing
        End Function

        Shared Function GetTabTaskFlowControl(tab As DevExpress.XtraTab.XtraTabPage) As TaskFlowControl
            Dim tskFlowControl As TaskFlowControl = Nothing
            For Each [control] In tab.Controls
                If TypeOf ([control]) Is TaskFlowControl Then
                    tskFlowControl = CType([control], TaskFlowControl)
                    Return tskFlowControl
                End If
            Next
            Return Nothing
        End Function
#End Region

#Region "Task Tree"
        Shared Sub OnTasksTreeUpdated()
            RaiseEvent TasksTreeUpdated()
        End Sub

        Shared Sub AddNewGroup(groupName As String, Optional fromForm As Boolean = False)
            If fromForm Then If MainForm.SetGroupNameBarEditItem.EditValue.ToString = "" Then Exit Sub
            Dim isRegistered As Boolean = CacadoreToTectonic.IsRegisteredGroupName(groupName)
            'Only add the new group if its name isn't already taken by another group.
            'Keep in mind that Tectonic itself enable having the same name.
            If Not isRegistered Then
                Dim grp As Cacadore.Group = CacadoreToTectonic.CreateNewGroup(groupName)
                RaiseEvent TasksTreeUpdated()
                SetSelectedNode(grp.GroupId)
            Else
                ShowAlertForm("",
                              String.Format("Cannot add a new group, a group with name {0} is already added.", groupName),
                              AlertLocation,
                              "",
                              False,
                              My.Resources.Cacadu_Image_46x46)
            End If
            If fromForm Then MainForm.SetGroupNameBarEditItem.EditValue = ""
        End Sub

        Shared Sub DeleteSelectedTreeNode()
            With MainForm
                Dim nodeKey = GetSelectedNodeName()
                Select Case GetSelectedNodeType()
                    Case "Task"
                        MainFormConnector.DeleteTreeTask(nodeKey)
                    Case "Group"
                        MainFormConnector.DeleteTreeGroup(nodeKey)
                End Select
                MainForm.TasksTreeView.Nodes.Remove(GetSelectedNode())
            End With
            RaiseEvent TasksTreeUpdated()
        End Sub

        Shared Function RenameGroup(newName As String, currentName As String) As Boolean
            Dim isRegisteredName As Boolean = CacadoreToTectonic.IsRegisteredGroupName(newName)
            If Not isRegisteredName Then
                Dim grp = CacadoreToTectonic.GetGroupByName(currentName)
                CacadoreToTectonic.ChangeGroupName(grp, newName)
                Return True
            Else
                ShowAlertForm("",
                              String.Format("Error renaming a group, a group with name {0} is already added.", newName),
                               AlertLocation,
                              "",
                              False,
                              My.Resources.Cacadu_Image_46x46)
                Return False
            End If
        End Function

        Shared Sub RenameTreeNode(e As NodeLabelEditEventArgs)
            If e.Label = vbNullString Then
                e.CancelEdit = True
                Exit Sub
            End If
            With MainForm
                Select Case e.Node.Tag.ToString
                    Case "Task"
                        MainFormConnector.RenameTask(e.Node.Name, e.Label)
                    Case "Group"
                        Dim result = MainFormConnector.RenameGroup(e.Label, e.Node.Text)
                        e.CancelEdit = Not result
                End Select
            End With
            'To display only the node in the Tasks List.
            OnAfterNodeSelect(e.Node.Name, e.Node.Tag.ToString)
        End Sub

        Shared Function GetSelectedNode() As TreeNode
            Return MainForm.TasksTreeView.SelectedNode
        End Function

        ''' <summary>
        ''' Gets the name of the selected node regardless if it is a name for task or a group.
        ''' </summary><returns></returns>
        Shared Function GetSelectedNodeName() As String
            Dim selectedNode = GetSelectedNode()
            If Not IsNothing(selectedNode) Then
                Return selectedNode.Name
            Else
                Return vbNullString
            End If
        End Function

        ''' <summary>
        ''' Gets the name of the selected group (returns "Ungrouped" group if no group is selected.)
        ''' </summary><returns></returns>
        Shared Function GetSelectedGroupName() As String
            Dim selectedNodeType As String = MainFormConnector.GetSelectedNodeType
            Dim selectedNodeName As String = vbNullString
            If selectedNodeType = "Group" Then
                selectedNodeName = MainFormConnector.GetSelectedNodeText
            Else
                selectedNodeName = "Ungrouped"
            End If
            Return selectedNodeName
        End Function

        Shared Function GetSelectedGroupId() As String
            Dim selectedNodeType As String = MainFormConnector.GetSelectedNodeType
            Dim selectedNodeName As String = vbNullString
            If selectedNodeType = "Group" Then selectedNodeName = GetSelectedNodeName()
            Return selectedNodeName
        End Function

        Shared Function GetSelectedNodeType() As String
            Dim selectedNode = GetSelectedNode()
            If Not IsNothing(selectedNode) Then
                Return selectedNode.Tag.ToString
            Else
                Return vbNullString
            End If
        End Function

        Shared Function GetSelectedNodeText() As String
            Dim selectedNode = GetSelectedNode()
            If Not IsNothing(selectedNode) Then
                Return selectedNode.Text
            Else
                Return vbNullString
            End If
        End Function

        Shared Sub RefreshTasksTree()
            ClearTasksTree()
            FillTreeWithGroups()
        End Sub

        Shared Sub ClearTasksTree()
            MainForm.TasksTreeView.Nodes.Clear()
        End Sub

        Shared Sub FillTreeWithGroups()
            With MainForm
                .TasksTreeView.BeginUpdate()
                Dim registerdGroups = CacadoreToTectonic.GetGroups()
                For Each grp In registerdGroups
                    .TasksTreeView.Nodes.Add(grp.Key, grp.Value.Name, 0)
                    .TasksTreeView.Nodes(grp.Key).Tag = "Group"
                    Dim tasksCol = CacadoreToTectonic.GetTasksCollection(grp.Key)
                    FillGroupsWithTasks(tasksCol, grp.Key)
                Next
                .TasksTreeView.EndUpdate()
            End With
        End Sub

        Shared Sub FillGroupsWithTasks(tasksCol As Cacadore.TasksCollection, groupId As String)
            For Each tsk In tasksCol.Tasks
                If tsk.Value.Properties.TaskName <> "" Then
                    Dim newNode = MainForm.TasksTreeView.Nodes(groupId).Nodes.Add(tsk.Key, tsk.Value.Properties.TaskName, 1)
                    MainForm.TasksTreeView.Nodes(groupId).Nodes(tsk.Key).Tag = "Task"
                    ToggleTaskNodeForeColor(newNode, tsk.Value.Properties.IsEnabled)
                End If
            Next
        End Sub

        Shared Sub GetSelectedNodeData(ByRef id As String, ByRef nodeName As String, ByRef nodeType As String)
            id = MainForm.TasksTreeView.SelectedNode.Name
            nodeName = MainForm.TasksTreeView.SelectedNode.Text
            nodeType = MainForm.TasksTreeView.SelectedNode.Tag.ToString
        End Sub

        Shared Sub SearchTasksTree(text As String)
            Dim searchedWord As String = text
            For Each groupNode As TreeNode In MainForm.TasksTreeView.Nodes
                For Each taskNode As TreeNode In groupNode.Nodes
                    If taskNode.Text.StartsWith(searchedWord, True, My.Application.Culture) And searchedWord <> "" Then
                        MainForm.TasksTreeView.SelectedNode = taskNode
                        MainForm.TasksTreeView.SelectedNode.BackColor = Color.YellowGreen
                    End If
                Next
            Next
        End Sub

        Shared Sub ClearTasksTreeBackColor()
            Dim nodes As TreeNodeCollection = MainForm.TasksTreeView.Nodes
            For Each n As TreeNode In nodes
                ClearTasksTreeBackgroundRecursive(n)
            Next
        End Sub

        Shared Sub ClearTasksTreeBackgroundRecursive(treeNode As TreeNode)
            'Called by ClearBackColor function
            For Each tn As TreeNode In treeNode.Nodes
                tn.BackColor = Color.White
                ClearTasksTreeBackgroundRecursive(tn)
            Next
        End Sub

        ''' <summary>
        ''' Gets the selected node task id.
        ''' </summary>
        ''' <returns>return nothing if a group node.</returns>
        Shared Function GetSelectedNodeTaskId() As String
            Dim taskId As String
            Dim nodeType As String = GetSelectedNodeType()
            If nodeType = "Task" Then
                taskId = GetSelectedNodeName()
            Else
                taskId = vbNullString
            End If
            Return taskId
        End Function

        ''' <summary>
        ''' Determines whether the clicked point in a tree view is the free region or node region.
        ''' </summary>
        ''' <param name="x">The x.</param>
        ''' <param name="y">The y.</param><returns>
        '''   <c>true</c> if [is free area] [the specified x]; otherwise, <c>false</c>.
        ''' </returns>
        Shared Function IsFreeArea(x As Integer, y As Integer) As TreeNode
            Dim result = MainForm.TasksTreeView.GetNodeAt(x, y)
            Return result
        End Function

        Shared Sub SetSelectedNode(key As String)
            If MainForm.TasksTreeView.Nodes.Count <= 0 Then Exit Sub
            'key == TaskId or GroupId
            If key <> vbNullString Then
                Dim newlyAddedNode As TreeNode() = MainForm.TasksTreeView.Nodes.Find(key, True)
                If newlyAddedNode.Count > 0 Then
                    MainForm.TasksTreeView.SelectedNode = newlyAddedNode.First
                    ClearTasksTreeBackColor()
                    MainForm.TasksTreeView.SelectedNode.BackColor = Color.YellowGreen
                End If
            End If
        End Sub

        Shared Sub DropComboColumn(columnIndex As Integer)
            'For more info:-
            'http://stackoverflow.com/questions/241100/how-to-manually-drop-down-a-datagridviewcomboboxcolumn
            MainForm.TasksListDataGridView.BeginEdit(True)
            If columnIndex = 5 Or columnIndex = 4 Then
                Dim editingControl = MainForm.TasksListDataGridView.EditingControl
                If Not IsNothing(editingControl) Then DirectCast(editingControl, DataGridViewComboBoxEditingControl).DroppedDown = True
            End If
        End Sub

        Shared Sub RenameSelectedGroup(newName As String)
            Dim selectedGroupType As String = GetSelectedNodeType()
            If selectedGroupType = "Group" Then
                Dim selectedGroupName As String = GetSelectedGroupName()
                Dim selectedNodeId As String = GetSelectedGroupId()
                Dim result = MainFormConnector.RenameGroup(newName, selectedGroupName)
                RefreshTasksTree()
                SetSelectedNode(selectedNodeId)
            End If
        End Sub

        Shared Sub DeleteGroupAndItsTasks(groupId As String)
            Dim groupToBeDeleted = CacadoreToTectonic.GetGroupById(groupId)
            ShowWaitForm("Deleting group " & groupToBeDeleted.Name & " along with its tasks...")
            CacadoreToTectonic.DeleteGroupWithTasks(groupId)
            CloseWaitForm()
            RefreshTasksTree()
        End Sub

        Shared Sub ToggleTaskNodeForeColor(node As TreeNode, isEnabled As Boolean)
            If isEnabled Then
                node.ForeColor = Color.Black
            Else
                node.ForeColor = Color.Gray
            End If
        End Sub

        Shared Sub CheckIfNoNodes()
            'If tasks tree has no groups at all, then disable all related group buttons along with related tasks buttons.
            If MainForm.TasksTreeView.Nodes.Count = 0 Then
                ToggleGroupTabRelatedBarButtonEnable(False)
                ToggleTasksTabRelatedBarButtonEnable(False)
            End If
        End Sub

        Shared Sub BeginEditSelectedNode()
            GetSelectedNode.BeginEdit()
        End Sub
#End Region

#Region "General Helpers"
        Shared Sub RenameTask(taskId As String, newName As String)
            Dim tsk = CacadoreToTectonic.GetTaskById(taskId)
            tsk.Properties.TaskName = newName
            tsk.Save()
            tsk.Dispose()
            tsk = Nothing
        End Sub

        Shared Function DeleteTreeGroup(groupId As String) As Boolean
            Dim groupToBeDeleted = CacadoreToTectonic.GetGroupById(groupId)
            ShowWaitForm("Deleting group " & groupToBeDeleted.Name & "...")
            Dim res = CacadoreToTectonic.DeleteGroup(groupId)
            RefreshTasksTree()
            CloseWaitForm()
            Return res
        End Function

        Shared Function DeleteTreeTask(nodeKey As String) As Boolean
            Dim res = CacadoreToTectonic.DeleteTask(nodeKey)
            Return res
        End Function

        Shared Function AddNewTask(waitBetweenActions As Boolean,
                                   Optional groupId As String = "",
                                   Optional taskName As String = vbNullString,
                                   Optional isEnabled As Boolean = True) As String
            Dim tsk = CacadoreToTectonic.CreateNewTask(waitBetweenActions,
                                                       groupId,
                                                       taskName,
                                                       isEnabled)
            RaiseEvent TasksTreeUpdated()
            Return tsk.TaskId
        End Function

        Shared Function GetActionIconIndex(actionName As String) As Integer
            Dim iconIndex As Integer = 0
            iconIndex = UI.MainForm.ActionsIconsImageCollection.Images.Keys.IndexOf(actionName)
            If iconIndex = -1 Then iconIndex = 0
            Return iconIndex
        End Function

        Shared Sub ToggleStopButtonEnable(enable As Boolean)
            MainForm.StopBarButton.Enabled = enable
        End Sub

        Shared Sub ToggleQuickAccessButtons(enabled As Boolean)
            With MainForm
                .OpenTaskQuickAccess.Enabled = enabled
                .SaveTaskQuickAccess.Enabled = enabled
            End With
        End Sub

        Shared Sub ToggleStopAllButtonEnable()
            With MainForm
                If Cacadore.Watcher.WatchedTasks.Count <= 0 Then
                    .StopAllBarButton.Enabled = False
                Else
                    .StopAllBarButton.Enabled = True
                End If
            End With
        End Sub

        Shared Sub ShowTaskDeleteAlert()
            ShowAlertForm("",
                          "Confirm task delete?",
                          AlertLocation,
                          "DeleteTreeTask",
                          True,
                          My.Resources.Cacadu_Image_46x46)
        End Sub

        Shared Sub ShowGroupDeleteAlert()
            ShowAlertForm("",
                          "Confirm group delete?",
                          AlertLocation,
                          "DeleteTreeGroup",
                          True,
                          My.Resources.Cacadu_Image_46x46,
                          GetSelectedGroupId())
        End Sub

#Region "Selected Task Methods"
        Shared Sub RunSelectedTask()
            If Cacadore.Settings.Enabled Then
                With MainForm
                    Dim tsk As Cacadore.Task = GetSelectedTask()
                    If IsNothing(tsk) Then Exit Sub
                    Dim bckGrndWrkHelper As New Balora.BackgroundWorkerHelper
                    bckGrndWrkHelper.WorkToDo = New MethodToBeInjectedDlg(AddressOf .backGroundWorkerValue_DoWork)
                    bckGrndWrkHelper.RunWorkerCompletedDlg = New RunWorkerCompletedEventHandler(AddressOf .RunWorkerCompleted)
                    bckGrndWrkHelper.Start(tsk)
                End With
            Else
                ShowAlertForm("",
                              "You cannot run tasks while Cacadu is disabled.",
                              AlertLocation,
                              Nothing,
                              False,
                              My.Resources.Cacadu_Image_46x46)
            End If
        End Sub

        Shared Sub PauseOrResumeSelectedTask()
            Dim tsk = GetSelectedTask()
            If Not IsNothing(tsk) Then
                Dim watchedTsk = Watcher.GetWatchedTaskById(tsk.TaskId)
                If Not IsNothing(watchedTsk) Then TogglePause(watchedTsk)
            End If
        End Sub

        Shared Sub StopSelectedTask()
            Dim selTsk As Cacadore.Task = GetSelectedTask()
            If Not IsNothing(selTsk) Then
                Dim watchedTsk = Watcher.GetWatchedTaskById(selTsk.TaskId)
                Dim tskName As String = watchedTsk.Properties.TaskName
                tskName = Chr(34) & tskName & Chr(34)
                watchedTsk.Cancel()
                ShowAlertForm("",
                              "Task " & tskName & " has stopped" & ", do you want to pause all of its triggers?",
                              AlertLocation,
                              "PauseTaskTriggers",
                              True,
                              My.Resources.Cacadu_Image_46x46,
                              watchedTsk)
            End If
        End Sub

        Shared Sub ToggleSelectedTaskEnableUponEnableRibbonButtonClick()
            With MainForm
                Dim selTask = GetSelectedTask()
                If .DisableEnableTaskBarButton.Caption = "Disable" Then
                    selTask.Properties.IsEnabled = False : selTask.Save()
                    .DisableEnableTaskBarButton.Caption = "Enable"
                    ToggleTaskNodeForeColor(GetSelectedNode(), False)
                ElseIf .DisableEnableTaskBarButton.Caption = "Enable" Then
                    selTask.Properties.IsEnabled = True : selTask.Save()
                    .DisableEnableTaskBarButton.Caption = "Disable"
                    ToggleTaskNodeForeColor(GetSelectedNode(), True)
                End If
            End With
        End Sub

        Shared Function GetSelectedTaskIdInTasksList() As String
            If IsTasksListsTabPageSelected() Then
                If Not IsNothing(MainForm.TasksListDataGridView.SelectedRows) AndAlso
                   MainForm.TasksListDataGridView.SelectedRows.Count > 0 Then
                    Dim selectedRow = MainForm.TasksListDataGridView.SelectedRows.Item(0)
                    Dim taskId As String = CStr(selectedRow.Cells(0).Value)
                    Return taskId
                End If
            End If
            Return ""
        End Function

        Shared Function GetSelectedTaskInTasksList() As Cacadore.Task
            Dim taskId As String = GetSelectedTaskIdInTasksList()
            If taskId <> "" Then
                Return GetTaskById(taskId)
            End If
            Return Nothing
        End Function

        'Search for selected task either as a tab or as a row in tasks list.
        Shared Function GetSelectedTask() As Cacadore.Task
            Dim tsk As Cacadore.Task = GetSelectedTabPageTaskObject()
            If IsNothing(tsk) Then tsk = GetSelectedTaskInTasksList()
            Return tsk
        End Function

        'Search for selected task either as a tab or as a row in tasks list.
        Shared Function GetSelectedTaskId() As String
            Dim tskId As String = GetSelectedTabPageTaskId()
            If tskId.Contains("TasksListXtraTabPage") Then tskId = GetSelectedTaskIdInTasksList()
            Return tskId
        End Function

        Shared Function GetSelectedTabPageTaskObject() As Cacadore.Task
            Dim tskId As String = Connectors.MainFormConnector.GetSelectedTabPageTaskId()
            If tskId.Contains(GetTasksListTabPageName()) Then Return Nothing
            Dim tsk As Cacadore.Task = CacadoreToTectonic.GetTaskById(tskId)
            Return tsk
        End Function

        Shared Function GetSelectedTabPageTaskName() As String
            Return MainForm.XtraTabsOfMainForm.SelectedTabPage.Text
        End Function

        Shared Sub RenameSelectedTask()
            With MainForm
                Dim taskId As String = GetSelectedNodeTaskId()
                If Not IsNothing(.RenameTaskBarButtonBarEditItem.EditValue) Then
                    RenameTask(taskId, .RenameTaskBarButtonBarEditItem.EditValue.ToString)
                    RefreshTasksTree()
                End If
            End With
        End Sub

        ''' <summary>
        ''' Gets the name of the selected task by checking selected Tree node or selected tab page.
        ''' </summary>
        ''' <param name="taskId">The task id.</param>
        ''' <param name="nodeType">Type of the node.</param><returns></returns>
        Shared Function GetSelectedTaskName(taskId As String, nodeType As String) As String
            Dim selectedTaskName As String = vbNullString
            If nodeType = "Task" Then
                selectedTaskName = CacadoreToTectonic.GetTaskName(taskId)
            Else
                If MainFormConnector.GetSelectedTabPageTaskId() = taskId Then
                    selectedTaskName = CacadoreToTectonic.GetTaskName(taskId)
                End If
            End If
            Return selectedTaskName
        End Function

        Shared Sub GetSelectedTaskFlow()
            Dim nodeName = GetSelectedNodeName()
            Dim nodeText = GetSelectedNodeText()
            Dim nodeType = GetSelectedNodeType()
            DisplayTaskFlow(nodeName, nodeText, nodeType)
        End Sub
#End Region


#End Region

#Region "Playing with Tasks List"
        Friend Shared Sub TasksListDataGridView_DataError(sender As Object, e As DataGridViewDataErrorEventArgs)
            If Debugger.IsAttached Then Debugger.Break()
            Balora.Alerter.REP(e.Exception.Message, e.Exception, False, True)
        End Sub

        Friend Shared Sub tasksRowChanged(sender As Object, e As TecDAL.TectonicDataSet.tasksRowChangeEvent)
            'Commit is the equivalent for Delete,Change,Add
            If e.Action = Data.DataRowAction.Commit Then FillTasksList()
        End Sub

        Shared Sub ShowOneTaskInTasksList(taskId As String)
            FilterTasksList(String.Format("id = '{0}'", taskId))
        End Sub

        Shared Sub FilterTasksList(filter As String)
            If filter = vbNullString Then
                FillTasksList()
            Else
                Dim whereMark = filter.IndexOf("'")
                Dim id As String = Strings.Mid(filter, whereMark + 2, filter.Length - 3)
                id = id.Replace("'", "")
                If filter.Contains("group_id") Then
                    DisplayTasksGroup(id)
                Else
                    DispalyOneTask(id)
                End If
            End If
        End Sub

        Shared Sub InitTasksList()
            FillTasksList()
            AddHandler TecDAL.TypedDataSetCrud.TectonicDataSetObj.tasks.tasksRowChanged, AddressOf tasksRowChanged
            AddHandler MainForm.TasksListDataGridView.DataError, AddressOf TasksListDataGridView_DataError
        End Sub

        Shared Sub ClearTaskList()
            MainForm.TasksListDataGridView.Rows.Clear()
        End Sub

        Shared Function getTaskNextTimeRun(taskId As String) As System.Collections.Generic.IEnumerable(Of String)
            Dim times As New SortedSet(Of String)
            If Cacadore.Watcher.WatchedTasks.Count <= 0 Then Return Nothing
            If Not Cacadore.Watcher.IsTaskWatched(taskId) Then Return Nothing

            If Not IsNothing(Cacadore.Watcher.WatchedTasks(taskId).MyTriggers) Then
                Dim Tasktrgs = Cacadore.Watcher.WatchedTasks(taskId).MyTriggers.Triggers.Values

                For Each trg In Tasktrgs
                    Dim nextTime = trg.NextFireTimeUTC
                    If Not IsNothing(nextTime) AndAlso nextTime.HasValue Then times.Add(trg.NextFireTimeUTC.Value.ToLocalTime.ToString)
                Next
            End If

            Return times
        End Function

        Shared Function getTaskTriggersDescriptions(taskId As String) As System.Collections.Generic.IEnumerable(Of String)
            Dim triggersDescriptions = From trg In TecDAL.TypedDataSetCrud.TectonicDataSetObj.triggers
                                       Where trg.task_id = taskId
                                       Select trg.description
            Return triggersDescriptions
        End Function

        Shared Sub addNewTasksListRow(row As tasksRow,
                                      triggersDescriptions As System.Collections.Generic.IEnumerable(Of String),
                                      taskNextRunTimes As System.Collections.Generic.IEnumerable(Of String))
            Try
                Dim triggersComboCell As New DataGridViewComboBoxCell
                Dim nextRunTimeComboCell As New DataGridViewComboBoxCell
                triggersComboCell.Items.AddRange(triggersDescriptions.ToArray)
                If Not IsNothing(taskNextRunTimes) AndAlso taskNextRunTimes.Count > 0 Then nextRunTimeComboCell.Items.AddRange(taskNextRunTimes.ToArray)
                Dim newRowIndex As Integer = MainForm.TasksListDataGridView.Rows.Add()
                With MainForm.TasksListDataGridView
                    .Rows(newRowIndex).Cells(0).Value = row.id
                    .Rows(newRowIndex).Cells(1).Value = row.name
                    .Rows(newRowIndex).Cells(2).Value = row.date_last_time_run
                    .Rows(newRowIndex).Cells(3).Value = row.last_time_run_result
                    If Not IsNothing(nextRunTimeComboCell) AndAlso nextRunTimeComboCell.Items.Count > 0 Then
                        .Rows(newRowIndex).Cells(4) = nextRunTimeComboCell
                        .Rows(newRowIndex).Cells(4).Value = nextRunTimeComboCell.Items(0)
                    End If
                    If Not IsNothing(triggersComboCell) AndAlso triggersComboCell.Items.Count > 0 Then
                        .Rows(newRowIndex).Cells(5) = triggersComboCell
                        .Rows(newRowIndex).Cells(5).Value = triggersComboCell.Items(0)
                    End If
                    .Rows(newRowIndex).Cells(6).Value = row.group_id
                End With
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Balora.Alerter.REP(ex.Message, ex, False, True)
            End Try
        End Sub

        Shared Sub DisplayTasksGroup(groupId As String)
            Dim tasksRows = From row As TectonicDataSet.tasksRow
                            In TecDAL.TypedDataSetCrud.TectonicDataSetObj.tasks
                            Where row.group_id = groupId
            ClearTaskList()
            startFillingLoop(tasksRows)
        End Sub

        Shared Sub DispalyOneTask(taskId As String)
            Dim taskRow As TectonicDataSet.tasksRow = TecDAL.TypedDataSetCrud.TectonicDataSetObj.tasks.FindByid(taskId)
            ClearTaskList()
            startFillingLoop({taskRow})
        End Sub

        Shared Sub startFillingLoop(rows As IEnumerable)
            For Each row As tasksRow In rows
                Dim triggersDescriptions = getTaskTriggersDescriptions(row.id)
                Dim taskNextRunTimes = getTaskNextTimeRun(row.id)
                addNewTasksListRow(row, triggersDescriptions, taskNextRunTimes)
            Next
        End Sub

        Shared Sub FillTasksList()
            ClearTaskList()
            startFillingLoop(TecDAL.TypedDataSetCrud.TectonicDataSetObj.tasks.Rows)
        End Sub
#End Region

#Region "Task Properties"
        Shared Sub ChangeTaskProperitesBasedOnSelecetedTabPage(ByVal selectedTabPageName As String)
            Dim tasksListTabPageName As String = GetTasksListTabPageName()
            If Not selectedTabPageName.Contains(tasksListTabPageName) Then
                LoadTaskPropertiesInPropertyGrid(selectedTabPageName)
            Else
                UnloadTaskPropertiesInPropertyGrid()
            End If
        End Sub

        Shared Sub LoadTaskPropertiesInPropertyGrid(taskId As String, Optional ByRef isEnable As Boolean = False)
            Dim selectedTsk = CacadoreToTectonic.GetTaskById(taskId)
            MainForm.TaskPropertiesPropertyGrid.SelectedObject = selectedTsk
            isEnable = selectedTsk.Properties.IsEnabled
            MainForm.TaskPropertiesPropertyGrid.Refresh()
        End Sub

        Shared Sub UnloadTaskPropertiesInPropertyGrid()
            MainForm.TaskPropertiesPropertyGrid.SelectedObject = Nothing
        End Sub

        Shared Sub SaveTaskPropertiesPropertyGrid()
            Dim tsk = CType(MainForm.TaskPropertiesPropertyGrid.SelectedObject, Cacadore.Task)
            tsk.Save()
        End Sub

        Shared Sub ShowTaskProperties()
            MainForm.DockPanelProperties.Show()
        End Sub
#End Region

#Region "Actions Gallery"
        Shared Function CreateActionFormGalleryItem(galleryItemTag As String) As Cacadore.IAction
            Select Case galleryItemTag
                Case "StartApplication"
                    Return New Cacadore.Actions.System.StartApplication
                Case "ShutdownPC"
                    Return New Cacadore.Actions.System.ShutdownPC
                Case "ExitProcess"
                    Return New Cacadore.Actions.System.ExitProcess
                Case "TakeScreenShot"
                    Return New Cacadore.Actions.System.TakeScreenShot
                Case "ShowAlert"
                    Return New Cacadore.Actions.MessageBoxes.ShowAlert
                Case "ShowMessageBox"
                    Return New Cacadore.Actions.MessageBoxes.ShowMessageBox
                Case "LastFacebookNotification"
                    Return New Actions.FacebookActions.LastFacebookNotification
                Case Else
                    Return Nothing
            End Select
        End Function

        Shared Function GetActionForm(actionName As String,
                                      Optional targetTask As Cacadore.Task = Nothing,
                                      Optional isEditMode As Boolean = False,
                                      Optional action As Cacadore.Action = Nothing) As DevExpress.XtraEditors.XtraForm
            Select Case actionName
                Case "StartApplication"
                    If Not IsNothing(targetTask) Then
                        'اوعي ابدا تنسي في المستقبل، تمرر المهمة للنموذج زي ما عملت هنا 
                        UI.StartApplicationForm.TargetTask = targetTask
                        UI.StartApplicationForm.IsEditMode = isEditMode
                    End If
                    If Not IsNothing(action) Then UI.StartApplicationForm.CreatedAction = CType(action, Cacadore.Actions.System.StartApplication)
                    Return UI.StartApplicationForm

                Case "ExitProcess"
                    If Not IsNothing(targetTask) Then
                        'اوعي ابدا تنسي في المستقبل، تمرر المهمة للنموذج زي ما عملت هنا 
                        UI.ExitProcessActionForm.TargetTask = targetTask
                        UI.ExitProcessActionForm.IsEditMode = isEditMode
                    End If
                    If Not IsNothing(action) Then UI.ExitProcessActionForm.CreatedAction = CType(action, Cacadore.Actions.System.ExitProcess)
                    Return UI.ExitProcessActionForm

                Case "ShutdownPC"
                    If Not IsNothing(targetTask) Then
                        'اوعي ابدا تنسي في المستقبل، تمرر المهمة للنموذج زي ما عملت هنا 
                        UI.ShutdownPCForm.TargetTask = targetTask
                        UI.ShutdownPCForm.IsEditMode = isEditMode
                    End If
                    If Not IsNothing(action) Then UI.ShutdownPCForm.CreatedAction = CType(action, Cacadore.Actions.System.ShutdownPC)
                    Return UI.ShutdownPCForm

                Case "ShowMessageBox"
                    'رب ذنبٍ أدخل صاحبه الجنة، ورب طاعةٍ أدخلت صاحبها النار
                    If Not IsNothing(targetTask) Then
                        'اوعي ابدا تنسي في المستقبل، تمرر المهمة للنموذج زي ما عملت هنا 
                        UI.ShowMessageBoxForm.TargetTask = targetTask
                        UI.ShowMessageBoxForm.IsEditMode = isEditMode
                    End If
                    If Not IsNothing(action) Then UI.ShowMessageBoxForm.CreatedAction = CType(action, Cacadore.Actions.MessageBoxes.ShowMessageBox)
                    Return UI.ShowMessageBoxForm

                Case "TakeScreenShot"
                    Return Nothing

                Case "ShowAlert"
                    If Not IsNothing(targetTask) Then
                        'اوعي ابدا تنسي في المستقبل، تمرر المهمة للنموذج زي ما عملت هنا 
                        UI.ShowAlertActionForm.TargetTask = targetTask
                        UI.ShowAlertActionForm.IsEditMode = isEditMode
                    End If
                    If Not IsNothing(action) Then UI.ShowAlertActionForm.CreatedAction = CType(action, Cacadore.Actions.MessageBoxes.ShowAlert)
                    Return UI.ShowAlertActionForm

                Case "LastFacebookNotification"
                    If Not IsNothing(targetTask) Then
                        'اوعي ابدا تنسي في المستقبل، تمرر المهمة للنموذج زي ما عملت هنا 
                        UI.LastFacebookNotificationActionForm.TargetTask = targetTask
                        UI.LastFacebookNotificationActionForm.IsEditMode = isEditMode
                    End If
                    If Not IsNothing(action) Then UI.LastFacebookNotificationActionForm.CreatedAction = CType(action, Actions.FacebookActions.LastFacebookNotification)
                    Return UI.LastFacebookNotificationActionForm

                Case Else
                    Return Nothing
            End Select
        End Function

        Shared Sub TryDragAction(e As System.Windows.Forms.MouseEventArgs)
            With MainForm
                Dim hInfo As RibbonHitInfo = .RibbonControlOfMainForm.CalcHitInfo(New Point(e.X, e.Y))
                If hInfo.HitTest = RibbonHitTest.GalleryImage Or hInfo.HitTest = RibbonHitTest.GalleryItem Then
                    Dim item As GalleryItem = hInfo.GalleryItem
                    CType(e, DevExpress.Utils.DXMouseEventArgs).Handled = True
                    .DoDragDrop(item, DragDropEffects.Copy)
                End If
            End With
        End Sub

        Shared Sub ShowActionForm(actionName As String)
            Dim taskId As String = GetSelectedTabPageTaskId()
            Dim taskObj = CacadoreToTectonic.GetTaskById(taskId)

            Dim actionObject = CreateActionFormGalleryItem(actionName)
            Dim actionForm = GetActionForm(CStr(actionName), taskObj)

            actionForm.ShowDialog()
        End Sub

        Shared Sub DisablePauseButton()
            With MainForm
                .PauseBarButton.Caption = "Pause"
                .PauseBarButton.Enabled = False
                .RunHomePageBarButton.Enabled = True
            End With
        End Sub

        Shared Sub TogglePauseButtonText(taskId As String)
            With MainForm
                Dim buttonText As String = MainForm.PauseBarButton.Caption
                If Cacadore.Watcher.IsWatchedTaskRunning(taskId) Then
                    .PauseBarButton.Caption = "Pause"
                    .PauseBarButton.Enabled = True
                ElseIf Cacadore.Watcher.IsWatchedTaskPaused(taskId) Then
                    .PauseBarButton.Caption = "Resume"
                    .RunHomePageBarButton.Enabled = False
                    .PauseBarButton.Enabled = True
                End If
            End With
        End Sub

        Shared Sub TogglePause(tsk As Cacadore.Task)
            Dim buttonText As String = MainForm.PauseBarButton.Caption
            If Cacadore.Watcher.IsWatchedTaskRunning(tsk.TaskId) Then
                tsk.Pause()
            ElseIf Cacadore.Watcher.IsWatchedTaskPaused(tsk.TaskId) Then
                tsk.Resume()
            End If
            TogglePauseButtonText(tsk.TaskId)
        End Sub
#End Region
#End Region

#Region "Context Menu Related"
        Shared Sub ToggleEnableDisableSystrayStrip(enable As Boolean)
            With MainForm
                If enable Then
                    .EnableDisableCacaduToolStripMenuItem.Text = "Disable Cacadu"
                    .EnableDisableCacaduToolStripMenuItem.Image = My.Resources.switch_on
                Else
                    .EnableDisableCacaduToolStripMenuItem.Text = "Enable Cacadu"
                    .EnableDisableCacaduToolStripMenuItem.Image = My.Resources.switch_off
                End If
            End With
        End Sub
#End Region

#Region "IListner"
        Public Sub Notify(message As Object) Implements Balora.IListener.Notify
            Select Case CStr(message)
                Case Is = "GroupsChanged"
                    RaiseEvent TasksTreeUpdated()
            End Select
        End Sub
#End Region
    End Class
End Namespace