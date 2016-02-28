Imports Cacadu.UI
Imports System.Reflection
Imports Cacadu.TaskFlowPresenter
Imports DevExpress.XtraBars.Ribbon
Imports DevExpress.XtraEditors.Drawing
Imports DevExpress.XtraEditors.ViewInfo
Imports DevExpress.XtraEditors.Controls
Imports Cacadu.Connectors.MainFormConnector
Imports DevExpress.XtraEditors.BaseListBoxControl

Friend Class TaskFlowControl

#Region "Properties"
    Property SelectedTask As Cacadore.Task
    Property SourceItmIndex As Integer = -1
    Property InternalDraw As Boolean = False
    Property DestinationItemIndex As Integer = -1
    Property DragSourceLocation As Point = Point.Empty
    'Because ImageListBoxItem has no tags, this is my way to store corresponding actions ids.
    Property ItemsDictionary() As New Dictionary(Of ImageListBoxItem, Cacadore.Action)
#End Region

#Region "TaskFlowControl Events"
    Private Sub TaskFlowControl_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        SelectedTask = GetSelectedTabPageTaskObject()
        Me.TaskFlowImageListBoxControl.ImageList = MainForm.ActionsIconsImageCollection
    End Sub
#End Region

#Region "TaskFlowImageListBoxControl"
    Private Sub TaskFlowImageListBoxControl_KeyDown(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles TaskFlowImageListBoxControl.KeyDown
        If e.KeyData = Keys.Delete Then
            If e.KeyValue = Keys.Delete Then ShowAlertForm("",
                                                           "Confirm delete?",
                                                           AlertLocation,
                                                           "DeleteSelectedActions",
                                                           True,
                                                           My.Resources.Cacadu_Image_46x46,
                                                           Me)
        ElseIf e.KeyCode = Keys.A And e.Control Then
            TaskFlowImageListBoxControl.SelectAll()
        End If
    End Sub

    Private Sub TaskFlowImageListBoxControl_MouseDoubleClick(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles TaskFlowImageListBoxControl.MouseDoubleClick
        EditSelectedActionByLocation(e.Location, TaskFlowImageListBoxControl, Me)
    End Sub

    Private Sub TaskFlowImageListBoxControl_MouseDown(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles TaskFlowImageListBoxControl.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right And TaskFlowImageListBoxControl.ItemCount > 0 Then
            ToggleContextMenuVisibility(True)
            SelectActionByRightClick(TaskFlowImageListBoxControl)
            ToggleActionItemEnable(TaskFlowImageListBoxControl,
                                   EditActionToolStripMenuItem)
            ToggleActionItemText(TaskFlowImageListBoxControl,
                                 DisableActionToolStripMenuItem,
                                 Me)
        ElseIf e.Button = Windows.Forms.MouseButtons.Right And TaskFlowImageListBoxControl.ItemCount = 0 Then
            ToggleContextMenuVisibility(False)
            AddActionToolStripMenuItem.Visible = True
        ElseIf e.Button = Windows.Forms.MouseButtons.Left Then
            PickItemForDrag(Me, TaskFlowImageListBoxControl)
        End If
    End Sub

    Private Sub TaskFlowImageListBoxControl_MouseMove(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles TaskFlowImageListBoxControl.MouseMove
        If e.Button = MouseButtons.Left Then StartDragingDrop(Me, TaskFlowImageListBoxControl, e)
    End Sub

    Private Sub TaskFlowImageListBoxControl_DragOver(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles TaskFlowImageListBoxControl.DragOver
        If e.Data.GetDataPresent(GetType(GalleryItem)) Then
            e.Effect = DragDropEffects.Copy
        Else
            StartDragOver(Me, TaskFlowImageListBoxControl, e)
        End If
    End Sub

    Private Sub TaskFlowImageListBoxControl_DragDrop(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles TaskFlowImageListBoxControl.DragDrop
        If e.Data.GetDataPresent(GetType(GalleryItem)) Then
            Connectors.MainFormConnector.OnXtraTabsOfMainFormDragDrop(sender, e)
        Else
            PlaceDragedItem(Me, TaskFlowImageListBoxControl, e)
        End If
    End Sub

    Private Sub TaskFlowImageListBoxControl_DrawItem(sender As System.Object, e As DevExpress.XtraEditors.ListBoxDrawItemEventArgs) Handles TaskFlowImageListBoxControl.DrawItem
        GrayItemIfActionDisable(_TaskFlowImageListBoxControl, e, Me)
        DrawDragedItem(Me, TaskFlowImageListBoxControl, e)
    End Sub

    Private Sub TaskFlowImageListBoxControl_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles TaskFlowImageListBoxControl.SelectedIndexChanged
        If TaskFlowImageListBoxControl.Items.Count > 0 Then
            TaskFlowToolStrip.Enabled = True
        Else
            TaskFlowToolStrip.Enabled = False
        End If
    End Sub

    Private Sub ToggleContextMenuVisibility(visiable As Boolean)
        If TaskFlowImageListBoxControl.SelectedItems.Count <= 1 Then
            AddActionToolStripMenuItem.Visible = visiable
            EditActionToolStripMenuItem.Visible = visiable
            RunMeToolStripMenuItem.Visible = visiable
            DisableActionToolStripMenuItem.Visible = visiable
            DuplicateActionToolStripMenuItem.Visible = visiable
            SelectAllToolStripMenuItem.Visible = visiable
            InvertSelectionToolStripMenuItem.Visible = visiable
            DeleteSelectionToolStripMenuItem.Visible = visiable
            ClearSelectionToolStripMenuItem.Visible = visiable
            MoveUpToolStripMenuItem.Visible = visiable
            MoveDownToolStripMenuItem.Visible = visiable
            DashToolStripMenuItem1.Visible = visiable
            DashToolStripMenuItem2.Visible = visiable
            DashToolStripMenuItem3.Visible = visiable
            'DashToolStripMenuItem4.Visible = visiable
            'CopyToolStripMenuItem.Visible = visiable
            'CutToolStripMenuItem.Visible = visiable
            'PasteToolStripMenuItem.Visible = visiable
        ElseIf TaskFlowImageListBoxControl.SelectedItems.Count >= 1 Then
            RunMeToolStripMenuItem.Visible = visiable
            DisableActionToolStripMenuItem.Visible = visiable
            DeleteSelectionToolStripMenuItem.Visible = visiable
            ClearSelectionToolStripMenuItem.Visible = visiable
            DuplicateActionToolStripMenuItem.Visible = False
            SelectAllToolStripMenuItem.Visible = False
            InvertSelectionToolStripMenuItem.Visible = False
            AddActionToolStripMenuItem.Visible = False
            EditActionToolStripMenuItem.Visible = False
            MoveUpToolStripMenuItem.Visible = False
            MoveDownToolStripMenuItem.Visible = False
            DashToolStripMenuItem1.Visible = False
            DashToolStripMenuItem2.Visible = False
            DashToolStripMenuItem3.Visible = False
            'DashToolStripMenuItem4.Visible = visiable
            'CopyToolStripMenuItem.Visible = visiable
            'CutToolStripMenuItem.Visible = visiable
            'PasteToolStripMenuItem.Visible = visiable
        End If
    End Sub
#End Region

#Region "Context Menu & Toolbar"
    Private Sub EditActionToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EditActionToolStripMenuItem.Click, EditActionToolStripButton.Click
        EditSelectedAction(TaskFlowImageListBoxControl, Me)
    End Sub

    Private Sub RunMeToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles RunMeToolStripMenuItem.Click, RunMeToolStripButton.Click
        RunSelectedActions(TaskFlowImageListBoxControl, Me)
    End Sub

    Private Sub DisableActionToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DisableActionToolStripMenuItem.Click, DisableActionToolStripButton.Click
        ToggleSelectedActionsEnable(_TaskFlowImageListBoxControl, Me)
    End Sub

    Private Sub DuplicateActionToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DuplicateActionToolStripMenuItem.Click, DuplicateActionToolStripButton.Click
        DuplicateSelectedAction(TaskFlowImageListBoxControl, Me)
        ShowTriggersCountInStatusBar(Me)
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click, SelectAllToolStripButton.Click
        TaskFlowImageListBoxControl.SelectAll()
    End Sub

    Private Sub InvertSelectionToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles InvertSelectionToolStripMenuItem.Click, InvertSelectionToolStripButton.Click
        InvertActionsListSelection(TaskFlowImageListBoxControl)
    End Sub

    Private Sub DeleteSelectionToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DeleteSelectionToolStripMenuItem.Click, DeleteSelectionToolStripButton.Click
        ShowAlertForm("",
                      "Confirm delete?",
                      AlertLocation,
                      "DeleteSelectedActions",
                      True,
                      My.Resources.Cacadu_Image_46x46,
                      Me)
    End Sub

    Private Sub ClearSelectionToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ClearSelectionToolStripMenuItem.Click, ClearSelectionToolStripButton.Click
        TaskFlowImageListBoxControl.UnSelectAll()
    End Sub

    Private Sub MoveUpToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles MoveUpToolStripMenuItem.Click, MoveUpToolStripButton.Click
        MoveUpSelectedActionItem(TaskFlowImageListBoxControl, Me)
    End Sub

    Private Sub MoveDownToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles MoveDownToolStripMenuItem.Click, MoveDownToolStripButton.Click
        MoveDownSelectedActionItem(TaskFlowImageListBoxControl, Me)
    End Sub

    Private Sub AddAction_Click(sender As System.Object, e As System.EventArgs) Handles StartApplicationToolStripMenuItem.Click, ShutdownPCToolStripMenuItem.Click, ExitProcessToolStripMenuItem.Click, TaskScreenShotToolStripMenuItem.Click, ShowAlertToolStripMenuItem.Click, MessageBoxToolStripMenuItem.Click, LastNotificationToolStripMenuItem.Click
        Dim menuItem As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        ShowActionForm(CStr(menuItem.Tag))
        ShowTriggersCountInStatusBar(Me)
    End Sub
#End Region
End Class
