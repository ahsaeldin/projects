Imports Cacadu.UI
Imports Cacadu.Connectors
Imports DevExpress.XtraBars.Ribbon
Imports DevExpress.XtraEditors.Controls
Imports Cacadu.Connectors.MainFormConnector
Imports DevExpress.XtraEditors.ViewInfo
Imports DevExpress.XtraEditors.Drawing
Imports System.Reflection
Imports DevExpress.XtraEditors

Friend Class TaskFlowPresenter
#Region "Flow Chart"
    Sub DrawFlowChartHeHeHeHe()
    End Sub
#End Region

#Region "Actions List"
    Shared Sub AddActionListItem([action] As Cacadore.Action, lstItemContent As String, imageIndex As Integer)
        Dim tskFlowControl As TaskFlowControl = Connectors.MainFormConnector.GetSelectedTabTaskFlowControl()
        With tskFlowControl
            .TaskFlowImageListBoxControl.BeginUpdate()
            Dim itmIndex As Integer = .TaskFlowImageListBoxControl.Items.Add(lstItemContent, imageIndex)
            .TaskFlowImageListBoxControl.SelectedIndex = itmIndex
            .TaskFlowImageListBoxControl.EndUpdate()

            'Because ImageListBoxItem has no tags, this is my way to store corresponding actions ids.
            Dim selItm As ImageListBoxItem = .TaskFlowImageListBoxControl.Items(itmIndex)
            If Not .ItemsDictionary.ContainsKey(selItm) Then
                .ItemsDictionary.Add(selItm, [action])
            Else
                .ItemsDictionary(selItm) = [action]
            End If
        End With
    End Sub

    Shared Sub EditActionListItem(actionIndex As Integer, lstItemContent As String, imageIndex As Integer)
        Dim tskFlowControl As TaskFlowControl = Connectors.MainFormConnector.GetSelectedTabTaskFlowControl()
        With tskFlowControl
            .TaskFlowImageListBoxControl.BeginUpdate()
            .TaskFlowImageListBoxControl.Items(actionIndex).Value = lstItemContent
            .TaskFlowImageListBoxControl.Items(actionIndex).ImageIndex = imageIndex
            .TaskFlowImageListBoxControl.EndUpdate()
        End With
    End Sub

    Shared Sub FillTaskList(taskId As String)
        Dim tskFlowControl As TaskFlowControl = Connectors.MainFormConnector.GetSelectedTabTaskFlowControl()
        Dim tsk = tskFlowControl.SelectedTask
        For Each [action] In tsk.MyActions.Actions.Values
            Dim icoIndex As Integer = GetActionIconIndex([action].GetType.Name)
            AddActionListItem([action], [action].GetFullDescription, icoIndex)
        Next
        ShowTriggersCountInStatusBar(tskFlowControl)
    End Sub
 
    Shared Sub DeleteAction(tskFlowCont As TaskFlowControl)
        Dim tsk As Cacadore.Task = tskFlowCont.SelectedTask
        Dim selIndex As Integer = tskFlowCont.TaskFlowImageListBoxControl.SelectedIndex
        Dim actionId As String = GetItemActionId(tskFlowCont.TaskFlowImageListBoxControl, selIndex)
        Dim acty = GetActionFromActionsDictionary(tskFlowCont, actionId)
        If Not IsNothing(acty) AndAlso tsk.UnAttachAction(acty) Then
            Dim itemToBeDeleted = tskFlowCont.TaskFlowImageListBoxControl.Items(tskFlowCont.TaskFlowImageListBoxControl.SelectedIndex)
            tskFlowCont.ItemsDictionary.Remove(itemToBeDeleted)
            tskFlowCont.TaskFlowImageListBoxControl.BeginUpdate()
            tskFlowCont.TaskFlowImageListBoxControl.Items.RemoveAt(tskFlowCont.TaskFlowImageListBoxControl.SelectedIndex)
            tskFlowCont.TaskFlowImageListBoxControl.EndUpdate()
        End If
    End Sub

    Shared Sub DeleteActionsObjects(tsk As Cacadore.Task, actionsToBeDeletedList As List(Of Cacadore.Action))
        For Each acty In actionsToBeDeletedList
            tsk.UnAttachAction(acty)
        Next
    End Sub

    Shared Sub GetDelationList(tskFlowCont As TaskFlowControl, actionsToBeDeletedList As List(Of Cacadore.Action), itmsToBeDeleted As SortedSet(Of Integer))
        Dim tsk As Cacadore.Task = tskFlowCont.SelectedTask
        For Each itm As ImageListBoxItem In tskFlowCont.TaskFlowImageListBoxControl.SelectedItems
            Dim itmIndex As Integer = tskFlowCont.TaskFlowImageListBoxControl.Items.IndexOf(itm)
            Dim actionId As String = GetItemActionId(tskFlowCont.TaskFlowImageListBoxControl, itmIndex)
            Dim acty = tsk.MyActions.Actions(actionId)
            If Not IsNothing(acty) Then actionsToBeDeletedList.Add(acty)
            RemoveItemActionId(tskFlowCont.TaskFlowImageListBoxControl, itmIndex)
            itmsToBeDeleted.Add(itmIndex)
        Next
    End Sub

    Shared Sub DeleteActionsItemsFromActionsList(tskFlowCont As TaskFlowControl, itmsToBeDeleted As SortedSet(Of Integer))
        'We reverse them in the orderedset to delete from botton to top.
        Dim reversedItmsToBeDeleted = itmsToBeDeleted.Reverse()
        Dim deletedItemsCount As Integer = itmsToBeDeleted.Count
        tskFlowCont.TaskFlowImageListBoxControl.BeginUpdate()
        For Each itm In reversedItmsToBeDeleted
            Dim itmToBeDeleted As Integer = 0
            If itmsToBeDeleted.Count > 0 Then
                tskFlowCont.TaskFlowImageListBoxControl.Items.RemoveAt(itm)
            End If
        Next
        tskFlowCont.TaskFlowImageListBoxControl.EndUpdate()
    End Sub

    Shared Sub DeleteActions(tskFlowCont As TaskFlowControl)
        ShowWaitForm("Deleting...")
        Dim itmsToBeDeleted As New SortedSet(Of Integer)
        Dim actionsToBeDeletedList As New List(Of Cacadore.Action)
        GetDelationList(tskFlowCont,
                        actionsToBeDeletedList,
                        itmsToBeDeleted)

        DeleteActionsObjects(tskFlowCont.SelectedTask,
                             actionsToBeDeletedList)

        DeleteActionsItemsFromActionsList(tskFlowCont,
                                          itmsToBeDeleted)
        CloseWaitForm()
    End Sub

    Shared Function GetClickedItem(_taskFlowImageListBoxControl As ImageListBoxControl,
                                   clickPoint As Point,
                                   ByRef clickedItemIndex As Integer) As ImageListBoxItem
        clickedItemIndex = _taskFlowImageListBoxControl.IndexFromPoint(clickPoint)
        If clickedItemIndex = -1 Then Return Nothing
        Dim clickedItem As ImageListBoxItem = _taskFlowImageListBoxControl.Items(clickedItemIndex)
        Return clickedItem
    End Function

    Shared Function GetClickedAction(_taskFlowImageListBoxControl As ImageListBoxControl,
                                     clickPoint As Point,
                                     tskFlowCont As TaskFlowControl) As Cacadore.Action
        Dim clickedItemIndex As Integer
        Dim tskFlowControl As TaskFlowControl = Connectors.MainFormConnector.GetSelectedTabTaskFlowControl()
        Dim clickedItem As ImageListBoxItem = GetClickedItem(_taskFlowImageListBoxControl, clickPoint, clickedItemIndex)
        If IsNothing(clickedItem) Then Return Nothing
        Dim tsk As Cacadore.Task = tskFlowCont.SelectedTask
        Dim [action] As Cacadore.Action = tskFlowControl.ItemsDictionary(clickedItem)
        Dim clickedActionKVP = GetActionFromActionsDictionary(tskFlowCont, [action].ActionId)
        clickedActionKVP.ParentTask = tsk
        Return clickedActionKVP
    End Function

    Shared Function GetItemActionId(_taskFlowImageListBoxControl As ImageListBoxControl, index As Integer) As String
        Dim tskFlowControl As TaskFlowControl = Connectors.MainFormConnector.GetSelectedTabTaskFlowControl()
        Dim lstItm As ImageListBoxItem = _taskFlowImageListBoxControl.Items(index)
        Dim [action] As Cacadore.Action = tskFlowControl.ItemsDictionary(lstItm)
        Return [action].ActionId
    End Function

    Shared Sub RemoveItemActionId(_taskFlowImageListBoxControl As ImageListBoxControl, index As Integer)
        Dim tskFlowControl As TaskFlowControl = Connectors.MainFormConnector.GetSelectedTabTaskFlowControl()
        Dim lstItm As ImageListBoxItem = _taskFlowImageListBoxControl.Items(index)
        tskFlowControl.ItemsDictionary.Remove(lstItm)
    End Sub

    Shared Function GetSelectedAction(_taskFlowImageListBoxControl As ImageListBoxControl, tskFlowCont As TaskFlowControl) As Cacadore.Action
        Dim clickedItemIndex As Integer = GetSelectedListItemIndex()
        Dim tsk As Cacadore.Task = tskFlowCont.SelectedTask
        Dim actionId As String = GetItemActionId(_taskFlowImageListBoxControl, clickedItemIndex)
        Dim clickedActionKVP = GetActionFromActionsDictionary(tskFlowCont, actionId)
        clickedActionKVP.ParentTask = tsk
        Return clickedActionKVP
    End Function

    Shared Function GetSelectedListItemIndex() As Integer
        Dim selItm As TaskFlowControl = Connectors.MainFormConnector.GetSelectedTabTaskFlowControl()
        Return selItm.TaskFlowImageListBoxControl.SelectedIndex
    End Function

    Shared Function GetClickLocation(taskFlowImageListBoxControl As ImageListBoxControl) As Point
        Dim localMousePosition As Point = taskFlowImageListBoxControl.PointToClient(Cursor.Position)
        Return localMousePosition
    End Function

    Shared Function GetClickedItemIndex(taskFlowImageListBoxControl As ImageListBoxControl) As Integer
        Dim localMousePosition As Point = GetClickLocation(taskFlowImageListBoxControl)
        Dim index = taskFlowImageListBoxControl.IndexFromPoint(localMousePosition)
        Return index
    End Function

    Shared Sub EditSelectedActionByLocation(location As Point,
                                            taskFlowImageListBoxControl As ImageListBoxControl,
                                            tskFlowCont As TaskFlowControl)
        Dim clickedAction = GetClickedAction(taskFlowImageListBoxControl, location, tskFlowCont)
        If IsNothing(clickedAction) Then Exit Sub
        Dim actionForm = GetActionForm(clickedAction.GetType.Name,
                                       clickedAction.ParentTask,
                                       isEditMode:=True,
                                       action:=clickedAction)
        actionForm.ShowDialog()
    End Sub

    Shared Sub EditSelectedAction(taskFlowImageListBoxControl As ImageListBoxControl, tskFlowCont As TaskFlowControl)
        Dim clickedAction = GetSelectedAction(taskFlowImageListBoxControl, tskFlowCont)
        Dim actionForm = GetActionForm(clickedAction.GetType.Name,
                                       clickedAction.ParentTask,
                                       isEditMode:=True,
                                       action:=clickedAction)
        actionForm.ShowDialog()
    End Sub

    Shared Sub DuplicateSelectedAction(_taskFlowImageListBoxControl As ImageListBoxControl, tskFlowCont As TaskFlowControl)
        Dim duplicatedAction As Cacadore.Action = GetSelectedAction(_taskFlowImageListBoxControl, tskFlowCont).Clone()

        Dim fullDescription As String = duplicatedAction.GetFullDescription
        Dim icoIndex As Integer = Connectors.MainFormConnector.GetActionIconIndex(duplicatedAction.GetType.Name)
        Dim actionId As String = duplicatedAction.ParentTask.AttachAction(duplicatedAction)

        If actionId <> "" Then
            TaskFlowPresenter.AddActionListItem(duplicatedAction, fullDescription, icoIndex)
            _taskFlowImageListBoxControl.SelectedItem = _taskFlowImageListBoxControl.Items(_taskFlowImageListBoxControl.ItemCount - 1)
        End If
    End Sub

    Shared Sub ToggleSelectedActionsEnable(_taskFlowImageListBoxControl As ImageListBoxControl, tskFlowCont As TaskFlowControl)
        For Each selIndex As Integer In _taskFlowImageListBoxControl.SelectedIndices
            Dim tsk As Cacadore.Task = tskFlowCont.SelectedTask
            Dim actionId As String = GetItemActionId(_taskFlowImageListBoxControl, selIndex)
            Dim selAct = GetActionFromActionsDictionary(tskFlowCont, actionId)
            If selAct.Enabled = False Then
                selAct.Enabled = True
            Else
                selAct.Enabled = False
            End If
            selAct.ParentTask.UpdateAction(selAct)
        Next
    End Sub

    Shared Sub ToggleActionItemEnable(taskFlowImageListBoxControl As ImageListBoxControl,
                                     _editActionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem)
        'Enable/Disable some buttons...
        If taskFlowImageListBoxControl.SelectedItems.Count > 1 Then
            _editActionToolStripMenuItem.Enabled = False
        ElseIf taskFlowImageListBoxControl.SelectedItems.Count = 1 Then
            _editActionToolStripMenuItem.Enabled = True
        End If
    End Sub

    Shared Sub ToggleActionItemText(taskFlowImageListBoxControl As ImageListBoxControl,
                                    _disableActionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem,
                                     tskFlowCont As TaskFlowControl)
        Dim selectedAction = GetSelectedAction(taskFlowImageListBoxControl, tskFlowCont)
        If selectedAction.Enabled = False Then
            _disableActionToolStripMenuItem.Text = "Enable Action"
        Else
            _disableActionToolStripMenuItem.Text = "Disable Action"
        End If
    End Sub

    Shared Sub SelectActionByRightClick(taskFlowImageListBoxControl As ImageListBoxControl)
        'Select item under mouse if not selected...
        Dim clickedItemIndex = GetClickedItemIndex(taskFlowImageListBoxControl)
        If clickedItemIndex <> -1 Then taskFlowImageListBoxControl.SelectedItem = taskFlowImageListBoxControl.Items(clickedItemIndex)
    End Sub

    Shared Sub InvertActionsListSelection(taskFlowImageListBoxControl As ImageListBoxControl)
        taskFlowImageListBoxControl.BeginUpdate()
        For i As Integer = 0 To taskFlowImageListBoxControl.ItemCount - 1
            taskFlowImageListBoxControl.SetSelected(i, (Not taskFlowImageListBoxControl.SelectedIndices.Contains(i)))
        Next i
        taskFlowImageListBoxControl.EndUpdate()
    End Sub

    Shared Sub MoveUpSelectedActionItem(_taskFlowImageListBoxControl As ImageListBoxControl, tskFlowCont As TaskFlowControl)
        With _taskFlowImageListBoxControl
            Dim precednce As Integer = 0
            Dim selAction = GetSelectedAction(_taskFlowImageListBoxControl,
                                              tskFlowCont)

            Dim selIndex As Integer = GetSelectedListItemIndex()

            If selIndex > 0 Then
                Dim selItem = .SelectedItem
                MoveActionItem(GetSelectedTabTaskFlowControl(),
                               _taskFlowImageListBoxControl,
                               selItem,
                               selIndex,
                               selIndex - 1)

                'If selIndex > 0 Then precednce = selIndex
                'Cacadore.Commands.MoveUpActionExecutionPrecedence(selAction.ActionId, precednce)

                '.BeginUpdate()
                '.Items.RemoveAt(selIndex)
                '.Items.Insert(precednce - 1, selItem)
                '.SelectedIndex = precednce - 1
                '.EndUpdate()
            End If
        End With
    End Sub

    Shared Sub MoveDownSelectedActionItem(_taskFlowImageListBoxControl As ImageListBoxControl, tskFlowCont As TaskFlowControl)
        With _taskFlowImageListBoxControl
            Dim precednce As Integer = 0
            Dim selAction = GetSelectedAction(_taskFlowImageListBoxControl, tskFlowCont)
            Dim selIndex As Integer = GetSelectedListItemIndex()
            If selIndex < .ItemCount - 1 Then
                Dim selItem = .SelectedItem
                MoveActionItem(GetSelectedTabTaskFlowControl(),
                              _taskFlowImageListBoxControl,
                              selItem,
                              selIndex,
                              selIndex + 2)

                'precednce = selIndex
                'Cacadore.Commands.MoveDownActionExecutionPrecedence(selAction.ActionId, precednce + 1)

                '.BeginUpdate()
                '.Items.RemoveAt(selIndex)
                '.Items.Insert(precednce + 1, selItem)
                '.SelectedIndex = precednce + 1
                '.EndUpdate()
            End If
        End With
    End Sub

    Shared Sub MoveActionItem(flowControl As TaskFlowControl,
                              taskFlowImageListBoxControl As ImageListBoxControl,
                              sourceItemObject As Object,
                              sourceIndex As Integer,
                              destinationIndex As Integer)
        With flowControl
            Dim selAction = GetSelectedAction(taskFlowImageListBoxControl, flowControl)
            If destinationIndex < sourceIndex Then
                Cacadore.Commands.MoveUpActionExecutionPrecedence(selAction.ActionId, destinationIndex + 1)
            ElseIf destinationIndex > sourceIndex Then
                Cacadore.Commands.MoveDownActionExecutionPrecedence(selAction.ActionId, destinationIndex - 1)
            End If

            If destinationIndex - 1 <= sourceIndex Then sourceIndex = sourceIndex + 1
            taskFlowImageListBoxControl.BeginUpdate()
            taskFlowImageListBoxControl.Items.Insert(destinationIndex, sourceItemObject)
            taskFlowImageListBoxControl.Items.RemoveAt(sourceIndex)
            taskFlowImageListBoxControl.EndUpdate()
            If destinationIndex < sourceIndex Then
                taskFlowImageListBoxControl.SelectedIndex = destinationIndex
            Else
                taskFlowImageListBoxControl.SelectedIndex = destinationIndex - 1
            End If
        End With
    End Sub

    Shared Function GetActionFromActionsDictionary(tskFlowCont As TaskFlowControl, actionId As String) As Cacadore.Action
        Dim queryResult = From acty In tskFlowCont.ItemsDictionary.Values
                          Where acty.ActionId = actionId
        If queryResult.Count > 0 Then Return queryResult.First
        Return Nothing
    End Function

    Shared Sub GrayItemIfActionDisable(_taskFlowImageListBoxControl As ImageListBoxControl,
                                       ByRef e As ListBoxDrawItemEventArgs,
                                       tskFlowCont As TaskFlowControl)
        Dim selTask = tskFlowCont.SelectedTask
        Dim actionId As String = GetItemActionId(_taskFlowImageListBoxControl, e.Index)
        Dim acty = GetActionFromActionsDictionary(tskFlowCont, actionId)
        If Not IsNothing(acty) Then If Not acty.Enabled Then e.Appearance.BackColor = Color.LightGray
    End Sub

    Shared Sub RunSelectedActions(_taskFlowImageListBoxControl As ImageListBoxControl, tskFlowCont As TaskFlowControl)
        For Each selIndex As Integer In _taskFlowImageListBoxControl.SelectedIndices
            Dim tsk As Cacadore.Task = tskFlowCont.SelectedTask
            Dim actionId As String = GetItemActionId(_taskFlowImageListBoxControl, selIndex)
            'Execute from a fresh copy rather than the object stored in the dictionary to avoid this error
            '"An async read operation has already been started on the stream."
            'which raised in calling MyProcess.BeginOutputReadLine()
            'Dim selAct = GetActionFromActionsDictionary(tskFlowCont, actionId)
            Dim freshCopyOfAction As Cacadore.Action = Cacadore.Commands.GetActionById(actionId)
            freshCopyOfAction.Execute()
        Next
    End Sub

    Shared Function GetActionsRowsCount(tskFlowCont As TaskFlowControl) As Integer
        Return tskFlowCont.TaskFlowImageListBoxControl.Items.Count
    End Function

    Shared Sub ShowTriggersCountInStatusBar(tskFlowCont As TaskFlowControl)
        tskFlowCont.TaskFlowToolStripStatusLabel.Text = GetActionsRowsCount(tskFlowCont) & " Actions(s)"
    End Sub

#Region "Drag Drop Helper Methods"
    Shared Sub PickItemForDrag(flowControl As TaskFlowControl, taskFlowImageListBoxControl As ImageListBoxControl)
        With flowControl
            .DragSourceLocation = GetClickLocation(taskFlowImageListBoxControl)
            .SourceItmIndex = GetClickedItemIndex(taskFlowImageListBoxControl)
            If .SourceItmIndex = -1 Then .DragSourceLocation = Point.Empty
        End With
    End Sub

    Shared Sub StartDragingDrop(flowControl As TaskFlowControl,
                                taskFlowImageListBoxControl As ImageListBoxControl,
                                e As System.Windows.Forms.MouseEventArgs)
        With flowControl
            If (.DragSourceLocation <> Point.Empty) AndAlso
               ((Math.Abs(e.X - .DragSourceLocation.X) > SystemInformation.DragSize.Width) OrElse
               (Math.Abs(e.Y - .DragSourceLocation.Y) > SystemInformation.DragSize.Height)) Then
                taskFlowImageListBoxControl.DoDragDrop(taskFlowImageListBoxControl, DragDropEffects.Move)
            End If
        End With
    End Sub

    Shared Sub StartDragOver(flowControl As TaskFlowControl,
                             taskFlowImageListBoxControl As ImageListBoxControl,
                             e As System.Windows.Forms.DragEventArgs)
        With flowControl
            .DragSourceLocation = taskFlowImageListBoxControl.PointToClient(New Point(e.X, e.Y))
            .DestinationItemIndex = taskFlowImageListBoxControl.IndexFromPoint(.DragSourceLocation)
            'Really the logic is complicated and impulsive; DestinationItemIndex is equal -1 only if you drag to
            'beyond the last bottom item, so simply you handle this using the next line by hand.
            If .DestinationItemIndex = -1 Then .DestinationItemIndex = taskFlowImageListBoxControl.Items.Count
            taskFlowImageListBoxControl.Refresh()
            e.Effect = DragDropEffects.Move
        End With
    End Sub

    Shared Sub PlaceDragedItem(flowControl As TaskFlowControl,
                               taskFlowImageListBoxControl As ImageListBoxControl,
                               e As System.Windows.Forms.DragEventArgs)

        With flowControl
            Dim newPoint As New Point(e.X, e.Y)
            newPoint = taskFlowImageListBoxControl.PointToClient(newPoint)
            Dim sourceItemObject As Object = taskFlowImageListBoxControl.Items(.SourceItmIndex)
            MoveActionItem(flowControl, taskFlowImageListBoxControl, sourceItemObject, flowControl.SourceItmIndex, flowControl.DestinationItemIndex)
        End With
    End Sub

    Shared Sub DrawDragedItem(flowControl As TaskFlowControl,
                              taskFlowImageListBoxControl As ImageListBoxControl,
                              e As ListBoxDrawItemEventArgs)
        With flowControl
            If e.Index <> .DestinationItemIndex Then Return
            If .InternalDraw Then Return

            .InternalDraw = True
            Dim b As ImageListBoxViewInfo = CType(taskFlowImageListBoxControl.GetViewInfo(), ImageListBoxViewInfo)
            Dim painter As PainterImageListBox = TryCast(b.Painter, PainterImageListBox)
            Dim info As MethodInfo = GetType(BaseListBoxPainter).GetMethod("DrawItem", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
            'info.Invoke(painter, 
            Dim args As New ControlGraphicsInfoArgs(b, e.Cache, b.Bounds)

            Dim itemInfo As BaseListBoxViewInfo.ItemInfo = b.GetItemByIndex(e.Index)
            info.Invoke(painter, New Object() {args, itemInfo})
            e.Graphics.DrawLine(Pens.Black, itemInfo.Bounds.Location, New Point(itemInfo.Bounds.Right, itemInfo.Bounds.Top))
            e.Handled = True
            .InternalDraw = False
        End With
    End Sub
#End Region
#End Region
End Class
