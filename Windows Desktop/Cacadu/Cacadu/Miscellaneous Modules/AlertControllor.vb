Imports Cacadore
Imports Cacadu.UI
Imports DevExpress.XtraBars.Alerter
Imports Cacadu.Connectors.MainFormConnector
Imports Cacadu.Connectors.TriggersFormConnector
Imports Cacadore.Actions.MessageBoxes.ShowAlert.ShowAlertsInputs

Module AlertControllor

#Region "Fields"
    'طريقة مفتكسة لإظهار فورم الألرت عشان مش هينفع اظهرها 
    'من ثريد تاني غير الثريد الرئيسي. تقوم الطريقة علي متغير عام 
    'هغيره لما تتبعت رسالة. وبعدين اعمل تشيك عليها من التايمر 
    'الرئيسي في الفورم.
    Friend AlertTag As String = ""
    Friend AlertText As String = ""
    Friend AlertNow As Boolean = False
    Friend AlertImage As Image = Nothing
    Friend AlertState As Object = Nothing
    Friend ShowAlertButton As Boolean = False
    Friend AlertLocation As AlertLocationEnum = CType(My.Settings.DPFCA, 
                                                      AlertLocationEnum)
#End Region

#Region "Delegates"
    Delegate Sub ShowAlertFormDlg(caption As String,
                                  message As String,
                                  tag As String,
                                  showButtons As Boolean,
                                  image As Image,
                                  state As Object)

    Delegate Sub InvokeAlertFormDlg(_alertInfo As AlertInfo)
#End Region

#Region "Alert Form Events"
    Friend Sub Alerter_AlertFormLoad(caption As String,
                                    message As String,
                                    tag As String,
                                    showButtons As Boolean,
                                    image As Object,
                                    state As Object)
        ShowAlertForm(caption, message, AlertLocation, tag, showButtons, CType(image, Drawing.Image), state)
    End Sub

    Private Sub AlertForm_AlertClick(sender As Object, e As AlertClickEventArgs)
        Dim regXMatcher As New Balora.RegexMatcher()
        Dim tagArr As Object() = CType(e.Info.Tag, Object())
        Dim url As String = CStr(tagArr(0))
        If url <> vbNullString Then
            Dim isUrl As Boolean = regXMatcher.MatchURL(url)
            If isUrl Then
                Process.Start(url)
            End If
        End If
    End Sub

    Private Sub AlertForm_ButtonClick(sender As System.Object, e As DevExpress.XtraBars.Alerter.AlertButtonClickEventArgs)
        Dim tagArr As Object() = CType(e.Info.Tag, Object())
        Dim tagString As String = CStr(tagArr(0))
        Dim tagObject As Object = tagArr(1)

        If e.ButtonName = "oKAlertButton" Then
            If tagString.Contains("DeleteTreeTask") Then

                Dim taskId As String = GetSelectedNodeTaskId()
                Dim isDisplayed = IsAnXtraPageDisplayed(taskId)
                If Not IsNothing(isDisplayed) Then CloseTaskFlow(CType(isDisplayed, DevExpress.XtraTab.XtraTabPage))
                DeleteSelectedTreeNode()
                MainForm.TasksTreeView.ExpandAll()

            ElseIf tagString.Contains("DeleteTreeGroup") Then

                DeleteGroupAndItsTasks(tagObject.ToString)

            ElseIf tagString.Contains("DeleteSelectedTriggers") Then

                DeleteSelectedTriggers()
                ShowOneTaskInTasksList(TaskObj.TaskId)
                Connectors.TriggersFormConnector.ShowTriggersCountInStatusBar()

            ElseIf tagString.Contains("AskToRunTask") Then

                Dim tsk As Cacadore.Task = CType(tagObject, Task)
                tsk.Execute()

            ElseIf tagString.Contains("DeleteSelectedActions") Then

                Dim tskFlwCont As TaskFlowControl = CType(tagObject, TaskFlowControl)
                If tskFlwCont.TaskFlowImageListBoxControl.SelectedItems.Count = 1 Then
                    TaskFlowPresenter.DeleteAction(tskFlwCont)
                ElseIf tskFlwCont.TaskFlowImageListBoxControl.SelectedItems.Count > 1 Then
                    TaskFlowPresenter.DeleteActions(tskFlwCont)
                End If

                TaskFlowPresenter.ShowTriggersCountInStatusBar(tskFlwCont)

            ElseIf tagString.Contains("PauseAllTriggers") Then

                Helpers.QuartzHelper.QuartzUtils.PauseAllTriggers()

            ElseIf tagString.Contains("PauseTaskTriggers") Then

                Dim tsk As Cacadore.Task = CType(tagObject, Task)
                tsk.PauseAllTriggers()

            ElseIf tagString.Contains("ShutdownNow") Then

                Dim acty As Cacadore.Actions.System.ShutdownPC = CType(tagObject, Cacadore.Actions.System.ShutdownPC)
                acty.PrepareForShutdown()

            ElseIf tagString.Contains("DeleteGroup") Then

                DeleteGroupAndItsTasks(tagObject.ToString)

            ElseIf tagString.Contains("AskToStartTriggers") Then

                Dim trgStarter As TriggersStarter = CType(tagObject, TriggersStarter)

                trgStarter.Start()

            End If
        End If
        e.AlertForm.Close()
    End Sub

    Private Sub AlertForm_FormClosing(sender As System.Object, e As DevExpress.XtraBars.Alerter.AlertFormClosingEventArgs)
        If Not My.Settings.ACA Then
            If e.CloseReason <> DevExpress.XtraBars.Alerter.AlertFormCloseReason.UserClosing Then e.Cancel = True
        End If
    End Sub
#End Region

#Region "Helpers"
    Friend Sub AddAlertFormEventsHandlers()
        AddHandler MainForm.AlertForm.AlertClick, AddressOf AlertControllor.AlertForm_AlertClick
        AddHandler Balora.Alerter.AlertForm_Load, AddressOf AlertControllor.Alerter_AlertFormLoad
        AddHandler MainForm.AlertForm.FormClosing, AddressOf AlertControllor.AlertForm_FormClosing
        AddHandler MainForm.AlertForm.ButtonClick, AddressOf AlertControllor.AlertForm_ButtonClick
        AddHandler Balora.Alerter.AlertForm_MessageSent, AddressOf AlertControllor.BaloraAlerter_MessageReceived
    End Sub

    Private Sub InvokeAlertForm(_alertInfo As AlertInfo)
        Try
            MainForm.AlertForm.Show(MainForm, _alertInfo)
        Catch ex As Exception
            MsgBox(_alertInfo.Text, MsgBoxStyle.Critical)
            ''للأسف السطر دة مش بيتنفذ، ومش بيطلع أخطاء ومش بقدر اكتب عن الخطأ بالاحمر في اللوج لسبب لا أعلمه حتي الآن
            'ابقي صلحه في المستقبل يا احمد
            'If Not IsNothing(MainForm.HistoryDataGridView) Then MainForm.AddNewHistoryRow(MainForm.HistoryDataGridView.Rows.Count + 1,
            '                                                                              "taskName", _alertInfo.Text, DateTime.Now)
        End Try
    End Sub

    Friend Sub ShowAlertForm(caption As String,
                             message As String,
                             alertLocation As AlertLocationEnum,
                             Optional tag As String = vbNullString,
                             Optional showButtons As Boolean = False,
                             Optional image As Image = Nothing,
                             Optional state As Object = Nothing)
        Dim tagArr As Object() = {tag, state}
        If Strings.Right(message, 1) <> "." Then message = message & "."
        Dim _alertInfo As New AlertInfo(caption, message)
        _alertInfo.Image = image
        _alertInfo.Tag = tagArr

        For Each butt As AlertButton In MainForm.AlertForm.Buttons
            butt.Visible = showButtons
        Next

        MainForm.AlertForm.FormLocation = CType(alertLocation, AlertFormLocation)

        If tag = "ShutdownPC" Then
            AlertText = message
            AlertNow = True
            ShowAlertButton = showButtons
            AlertTag = "ShutdownNow"
            If Not IsNothing(state) Then AlertState = state
            AlertImage = My.Resources.Cacadu_Image_46x46
            Exit Sub
        End If

        If MainForm.InvokeRequired Then
            MainForm.BeginInvoke(New InvokeAlertFormDlg(AddressOf InvokeAlertForm), _alertInfo)
        Else
            InvokeAlertForm(_alertInfo)
        End If
    End Sub

    Friend Sub BaloraAlerter_MessageReceived(message As String, state As Object)
        If message.Contains("AllTriggersPaused") Then
            WriteInStatusBar("All Triggers Paused.", Color.Red)
        ElseIf message.Contains("ClearStatusBarState") Then
            WriteInStatusBar("", Color.Black)
        ElseIf message.Contains("ReportTaskProgress") Then
            Dim args() As String = CType(state, String())
            Dim taskName As String = args(0)
            Dim runningProgress As Integer = CInt(args(1))
            WriteInStatusBar("Running Task " & taskName, Color.Black, runningProgress)
        ElseIf message.Contains("TaskCancelled") Then
            Dim taskName As String = CStr(state)
            WriteInStatusBar("Task " & Chr(34) & taskName & Chr(34) & " cancelled at " & DateTime.Now, Color.Black)
        ElseIf message.Contains("UpdateBottleneck") Then
            CacaduGloby.UpdateBottleneck(state)
        ElseIf message.Contains("AlertOnError") Then
            'طريقة مفتكسة لإظهار فورم الألرت عشان مش هينفع اظهرها 
            'من ثريد تاني غير الثريد الرئيسي. تقوم الطريقة علي متغير عام 
            'هغيره لما تتبعت رسالة. وبعدين اعمل تشيك عليها من التايمر 
            'الرئيسي في الفورم.
            AlertText = state.ToString
            AlertImage = My.Resources.Cacadu_Image_46x46
            AlertNow = True
        ElseIf message.Contains("DisplayThisMessage") Then
            'BTW, it is the same as "AlertOnError", however, I prefer to use "AlertOnError" on errors only
            AlertText = state.ToString
            AlertImage = My.Resources.Cacadu_Image_46x46
            AlertNow = True
        ElseIf message.Contains("AskUserForMisFire") Then
            Dim taskName As String = Chr(34) & CType(state, Cacadore.Task).Properties.TaskName & Chr(34)
            Dim messageBoxMessageFromResources As String = My.Resources.ATRT
            messageBoxMessageFromResources = messageBoxMessageFromResources.Replace("$TNH$", taskName)
            AlertState = state
            ShowAlertButton = True
            AlertTag = "AskToRunTask"
            AlertText = messageBoxMessageFromResources
            AlertImage = My.Resources.Cacadu_Image_46x46
            AlertNow = True
        ElseIf message.Contains("DoAlertAction") Then
            Dim alertMessageLocation = message.Split(CChar("#"))(1)
            AlertText = state.ToString
            AlertLocation = CType(alertMessageLocation, AlertLocationEnum)
            AlertNow = True
        ElseIf message.Contains("FacebookLoginSuccess") Then
            Dim sentState As Dictionary(Of String, Object) = CType(state, Global.System.Collections.Generic.Dictionary(Of String, Object))
            Dim proasdfasd As Image = CType(sentState("photourl"), Image)
            Dim userName As String = CType(sentState("user_name"), String)
            AlertText = userName & " has a successful Facebook login."
            AlertImage = proasdfasd
            AlertNow = True
        ElseIf message.Contains("FacebookLogOutSuccess") Then
            Dim fbActions As Actions.FacebookActions.FbHelper = CType(state, Actions.FacebookActions.FbHelper)
            AlertText = "You're now logged off Facebook."
            AlertImage = fbActions.UserProfilePicture
            AlertNow = True
        ElseIf message.Contains("lastNotificationTitleHtml") Then
            Dim fbActions As Actions.FacebookActions.FbHelper = CType(state, Actions.FacebookActions.FbHelper)
            AlertText = fbActions.LastNotificationTitle
            AlertImage = fbActions.LastNotificationIcon
            AlertTag = fbActions.lastNotificationURL
            AlertNow = True
        End If
    End Sub
#End Region
End Module
