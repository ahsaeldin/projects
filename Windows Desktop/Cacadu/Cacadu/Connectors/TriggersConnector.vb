Imports Cacadu.UI
Imports Helpers.QuartzHelper.QuartzUtils
Imports Cacadore

Namespace Connectors
    Friend Class TriggersFormConnector

#Region "Properties"
        Shared Property TaskObj As Cacadore.Task = Nothing
#End Region

#Region "Helpers"
        Shared Sub PrepareAndDisplayTriggersForm(taskId As String)
            TriggersForm.MyTaskId = taskId
            TaskObj = Cacadore.Commands.GetTaskById(taskId)
            Dim taskName As String = TaskObj.Properties.TaskName

            Dim newTriggersFormCaption As String = "Triggers list for" & "<i>" & taskName & "</i>"

            TriggersForm.HtmlText = newTriggersFormCaption

            FillTriggersList(TaskObj)

            TriggersForm.ShowDialog()
            If Not IsNothing(TaskObj) Then TaskObj.Dispose()
            If Not IsNothing(TaskObj) Then TaskObj = Nothing
        End Sub

#Region "Triggers Manage"
        Shared Sub FillTriggersList(tsk As Cacadore.Task)
            Dim itm As ListViewItem = Nothing
            Dim description As String = vbNullString

            TriggersForm.TriggersListView.Items.Clear()

            For Each trgKVP In tsk.MyTriggers.Triggers

                description = trgKVP.Value.Description
                description = Balora.CroXGen.FormatExpressionDescrption(description, Nothing, Nothing, False)

                'Fill Type Column
                If description.Contains("once") Then
                    itm = TriggersForm.TriggersListView.Items.Add("Run once")
                ElseIf description.Contains("Every") OrElse description.Contains("every") Then
                    Dim whereStarting As Integer = description.IndexOf("starting")
                    If whereStarting <> -1 Then
                        Dim typeString = Strings.Left(description, whereStarting)
                        typeString = typeString.Replace("every", "Every")
                        itm = TriggersForm.TriggersListView.Items.Add(typeString)
                    Else
                        itm = TriggersForm.TriggersListView.Items.Add(description)
                    End If
                End If

                If IsNothing(itm) Then
                    itm = TriggersForm.TriggersListView.Items.Add("N/A")
                End If
                itm.SubItems.Add(description)
                itm.ToolTipText = description
                itm.Group = TriggersForm.TriggersListView.Groups("scheduleTriggersListViewGroup")
                itm.Tag = trgKVP.Key

                Dim triggerState = GetTriggerState(trgKVP.Key)
                If IsTriggerUnScheduled(trgKVP.Key) Then
                    GrayTrigger(itm)
                    itm.SubItems.Add("UnScheduled")
                ElseIf IsTriggerPaused(trgKVP.Key) Then
                    GrayPausedTrigger(itm)
                    itm.SubItems.Add("Paused")
                ElseIf triggerState = Quartz.TriggerState.Normal Then
                    itm.SubItems.Add("Scheduled")
                End If
            Next
        End Sub

        Shared Sub EditSelectedTriggers()
            ScheduleTriggerFormConnector.ResetTabsSelectedProperty()

            Dim selectedTriggers = TriggersForm.TriggersListView.SelectedItems

            Dim itm = selectedTriggers.Item(0)
            Dim trgId As String = GetItemTrgId(itm)

            Dim trg = TriggersFormConnector.TaskObj.MyTriggers.Triggers(trgId)

            Select Case trg.TriggerType
                Case Cacadore.ITrigger.TriggerTypeEnum.SimpleTrigger

                    Dim parsedExpr As Balora.CroXGen.SimpleCron
                    Dim innerTrg As Quartz.ISimpleTrigger = CType(trg.InnerTrigger, Quartz.ISimpleTrigger)
                    Helpers.QuartzHelper.QuartzUtils.ParseSimpleTrigger(CType(trg, Helpers.QuartzHelper.QuartzTrigger), parsedExpr)
                    ScheduleTriggerFormConnector.FillSimpleControlsForEditing(innerTrg, parsedExpr)

                Case Cacadore.ITrigger.TriggerTypeEnum.CronTrigger

                    Dim innerTrg As Quartz.ICronTrigger = CType(trg.InnerTrigger, Quartz.ICronTrigger)
                    ScheduleTriggerFormConnector.FillAdvancedControlsForEditing(innerTrg.CronExpressionString)

            End Select

            ScheduleTriggerFormConnector.SetMisFireCommand(trg.MisFireCommand)
            If Not IsNothing(trg.Calendar) Then
                Dim croxExpr As Quartz.CronExpression = CType(trg.Calendar, Quartz.Impl.Calendar.CronCalendar).CronExpression
                ScheduleTriggerFormConnector.SetHolidays(CType(trg.Calendar, Quartz.Impl.Calendar.CronCalendar).CronExpression)
            End If

            With ScheduleTriggerForm
                .Mode = "Edit"
                .TriggerId = trgId
                .ShowDialog()
            End With

            MainFormConnector.ShowOneTaskInTasksList(Connectors.TriggersFormConnector.TaskObj.TaskId)
        End Sub

        Shared Sub DeleteSelectedTriggers()
            ShowWaitForm("Deleting...")
            Dim selectedTriggers = TriggersForm.TriggersListView.SelectedItems
            For Each itm As ListViewItem In selectedTriggers
                Dim trgId As String = GetItemTrgId(itm)
                itm.Remove()
                'Dim trg = TriggersFormConnector.TaskObj.MyTriggers.Triggers(trgId)
                Dim trg = Cacadore.Commands.GetTriggerById(trgId)
                TriggersFormConnector.TaskObj.UnAttachTrigger(trg)
                Watcher.RemoveTriggerFromWatchedScheduledTriggers(trgId)
            Next
            CloseWaitForm()
        End Sub
#End Region

#Region "General Helpers"
        Shared Function GetFirstSelectedItemTrgId() As String
            Dim selItems = GetSelectedItemsTrgIds()
            If Not IsNothing(selItems) AndAlso selItems.Count > 0 Then
                Return selItems.First
            Else
                Return vbNullString
            End If
        End Function

        Shared Function GetSelectedItemsTrgIds() As String()
            Dim lst As New List(Of String)
            For Each itm As ListViewItem In TriggersForm.TriggersListView.SelectedItems
                lst.Add(itm.Tag.ToString)
            Next
            If lst.Count > 0 Then
                Return lst.ToArray
            Else
                Return Nothing
            End If
        End Function

        Shared Function GetItemTrgId(itm As ListViewItem) As String
            Return itm.Tag.ToString
        End Function

        Shared Function GetItemTrgId(itmIndex As Integer) As String
            Return TriggersForm.TriggersListView.Items(itmIndex).Tag.ToString
        End Function

        Shared Sub SetButtonsEnabled()
            With TriggersForm
                Dim selectedItemsCount = .TriggersListView.SelectedItems.Count
                If selectedItemsCount > 0 Then
                    .DeleteTriggerSimpleButton.Enabled = True
                    If selectedItemsCount = 1 Then
                        .EditTriggerSimpleButton.Enabled = True
                    Else
                        .EditTriggerSimpleButton.Enabled = False
                    End If
                Else
                    .EditTriggerSimpleButton.Enabled = False
                    .DeleteTriggerSimpleButton.Enabled = False
                End If
            End With
        End Sub

        Shared Sub SelectAllTriggerItems(e As System.Windows.Forms.KeyEventArgs)
            If e.Control And e.KeyValue = Keys.A Then
                For Each itm As ListViewItem In TriggersForm.TriggersListView.Items
                    itm.Selected = True
                Next
            End If
        End Sub

        Shared Sub ToggeleResumeAllButton()
            With TriggersForm
                Dim isThereATriggerPaused As Boolean = IsThereTriggerPaused()
                If isThereATriggerPaused Then
                    .ResumeAllSimpleButton.Enabled = True
                Else
                    .ResumeAllSimpleButton.Enabled = False
                End If
            End With
        End Sub

        Shared Function GetTriggerItemFromList(triggerId As String) As ListViewItem
            Dim lstViewItem As ListViewItem = Nothing
            For Each trgItem As ListViewItem In TriggersForm.TriggersListView.Items
                If trgItem.Tag.ToString.Contains(triggerId) Then
                    Return trgItem
                End If
            Next
            Return lstViewItem
        End Function

        Shared Sub UnGrayResumedTrigger(trgListItem As ListViewItem)
            trgListItem.ForeColor = Color.Black
            trgListItem.Font = New Font(trgListItem.Font, FontStyle.Regular)
        End Sub

        Shared Sub GrayTrigger(trgListItem As ListViewItem)
            trgListItem.ForeColor = Color.Gray
        End Sub

        Shared Sub GrayPausedTrigger(trgListItem As ListViewItem)
            GrayTrigger(trgListItem)
            trgListItem.Font = New Font(trgListItem.Font, FontStyle.Italic)
        End Sub

        Shared Sub GrayPausedTrigger(triggerId As String)
            Dim isPause As Boolean = IsTriggerPaused(triggerId)
            Dim trgItem = GetTriggerItemFromList(triggerId)
            If Not IsNothing(trgItem) Then
                GrayPausedTrigger(trgItem)
            End If
        End Sub

        Shared Sub ToggleResumeAndPauseSingleTriggerButton()
            With TriggersForm
                If .TriggersListView.SelectedItems.Count > 1 OrElse .TriggersListView.SelectedItems.Count = 0 Then
                    .ResumeSimpleButton.Enabled = False
                    .PauseSimpleButton.Enabled = False
                ElseIf .TriggersListView.SelectedItems.Count = 1 Then
                    Dim triggerId As String = GetItemTrgId(.TriggersListView.SelectedItems.Item(0))
                    If triggerId <> "" Then
                        Dim isPaused As Boolean = IsTriggerPaused(triggerId)
                        Dim isScheduled As Boolean = IsTriggerScheduled(triggerId)
                        If isPaused Then
                            .PauseSimpleButton.Enabled = False
                            .ResumeSimpleButton.Enabled = True
                        ElseIf isScheduled Then
                            .PauseSimpleButton.Enabled = True
                            .ResumeSimpleButton.Enabled = False
                        End If
                    End If
                End If
            End With
        End Sub

        Shared Function GetTriggersRowsCount() As Integer
            Return TriggersForm.TriggersListView.Items.Count
        End Function

        Shared Sub ShowTriggersCountInStatusBar()
            TriggersForm.TriggerFormLeftToolStripStatusLabel.Text = GetTriggersRowsCount() & " Trigger(s)"
        End Sub
#End Region

#End Region
    End Class
End Namespace

