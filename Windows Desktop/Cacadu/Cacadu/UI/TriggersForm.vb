Imports Cacadu.Connectors
Imports Helpers.QuartzHelper.QuartzUtils
Imports Cacadu.Connectors.TriggersFormConnector

Namespace UI
    Friend Class TriggersForm
#Region "Properties"
        Property MyTaskId As String
        Property HitLocation As Point
#End Region

#Region "Form Handlers"
        Private Sub Triggers_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
            SetIconToCacaduIcon(Me)
            ToggeleResumeAllButton()
            ShowTriggersCountInStatusBar()
        End Sub

        Private Sub TriggersForm_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
            MyTaskId = vbNullString
            TaskObj.Dispose()
            TaskObj = Nothing
        End Sub
#End Region

#Region "Buttons"
        Private Sub AddScheduleTriggerSimpleButton_Click(sender As System.Object, e As System.EventArgs) Handles AddScheduleTriggerSimpleButton.Click
            ScheduleTriggerFormConnector.ResetTabsSelectedProperty()
            With ScheduleTriggerForm
                .Mode = ""
                .TriggerId = ""
                .ShowDialog()
            End With
            MainFormConnector.ShowOneTaskInTasksList(Connectors.TriggersFormConnector.TaskObj.TaskId)
            ShowTriggersCountInStatusBar()
        End Sub

        Private Sub EditTriggerSimpleButton_Click(sender As System.Object, e As System.EventArgs) Handles EditTriggerSimpleButton.Click
            EditSelectedTriggers()
        End Sub

        Private Sub DeleteTriggerSimpleButton_Click(sender As System.Object, e As System.EventArgs) Handles DeleteTriggerSimpleButton.Click
            ShowAlertForm("", "Confirm delete?", AlertLocation, "DeleteSelectedTriggers", True, My.Resources.Cacadu_Image_46x46)
        End Sub

        Private Sub ResumeSimpleButton_Click(sender As System.Object, e As System.EventArgs) Handles ResumeSimpleButton.Click
            Dim selItemTrgId As String = GetFirstSelectedItemTrgId()
            If selItemTrgId <> vbNullString Then
                ResumeTrigger(selItemTrgId)
                FillTriggersList(TaskObj)
            End If
        End Sub

        Private Sub PauseSimpleButton_Click(sender As System.Object, e As System.EventArgs) Handles PauseSimpleButton.Click
            Dim selItemTrgId As String = GetFirstSelectedItemTrgId()
            If selItemTrgId <> vbNullString Then
                PauseTrigger(selItemTrgId)
                FillTriggersList(TaskObj)
            End If
        End Sub

        Private Sub PauseAllSimpleButton_Click(sender As System.Object, e As System.EventArgs) Handles PauseAllSimpleButton.Click
            PauseAllTriggers()
            ToggeleResumeAllButton()
            FillTriggersList(TaskObj)
        End Sub

        Private Sub ResumeAllSimpleButton_Click(sender As System.Object, e As System.EventArgs) Handles ResumeAllSimpleButton.Click
            ResumeAllTriggers()
            ToggeleResumeAllButton()
            FillTriggersList(TaskObj)
        End Sub

        Private Sub CloseSimpleButton_Click(sender As System.Object, e As System.EventArgs) Handles CloseSimpleButton.Click
            OnCancel()
        End Sub

        Private Sub OnCancel()
            Me.Dispose()
        End Sub
#End Region

#Region "Triggers ListView"
        Private Sub TriggersListView_ItemSelectionChanged(sender As System.Object, e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles TriggersListView.ItemSelectionChanged
            SetButtonsEnabled()
        End Sub

        Private Sub TriggersListView_KeyDown(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles TriggersListView.KeyDown
            SelectAllTriggerItems(e)
            If e.KeyValue = Keys.Delete Then ShowAlertForm("", "Confirm delete?", AlertLocation, "DeleteSelectedTriggers", True, My.Resources.Cacadu_Image_46x46)
        End Sub

        Private Sub TriggersListView_DoubleClick(sender As System.Object, e As System.EventArgs) Handles TriggersListView.DoubleClick
            Dim clickedItemInfo = TriggersListView.HitTest(HitLocation)
            If Not IsNothing(clickedItemInfo.Item) Then
                TriggersListView_ItemDoubleClick(clickedItemInfo.Item)
            End If
        End Sub

        Private Sub TriggersListView_ItemDoubleClick(itm As ListViewItem)
            'My defined ItemDoubleClick event.
            'Ref:http://www.freevbcode.com/ShowCode.asp?ID=2472
            EditSelectedTriggers()
        End Sub

        Private Sub TriggersListView_ColumnClick(sender As System.Object, e As System.Windows.Forms.ColumnClickEventArgs) Handles TriggersListView.ColumnClick
            'If you wish to sort columns in the future.
            '    http://support.microsoft.com/kb/319401
            '    http://msdn.microsoft.com/en-us/library/ms229643(v=vs.80).aspx
            '    http://www.java2s.com/Code/CSharp/GUI-Windows-Form/SortaListViewbyAnyColumn.htm
        End Sub

        Private Sub TriggersListView_MouseUp(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles TriggersListView.MouseUp
            HitLocation = e.Location
            'We call it from MouseUp rather than MouseDown to make sure that selection 
            'has been done/unselected. so pause and resume button enabled will be detected rightly.
            ToggleResumeAndPauseSingleTriggerButton()
        End Sub
#End Region

    End Class
End Namespace
