Imports Cacadu.UI

Namespace Connectors
    Friend Class AddNewTaskFormConnector
#Region "Events Handlers"
        Shared Sub OnFormLoad()
            SetIconToCacaduIcon(AddNewTaskForm)
            Dim groupsNames As SortedSet(Of String) = CacadoreToTectonic.GetGroupsNames()
            fillGroupsComboBoxEdit(groupsNames)
        End Sub

        Shared Sub OnFormOK(waitBetweenActions As Boolean,
                            Optional groupName As String = "",
                            Optional taskName As String = vbNullString,
                            Optional isEnabled As Boolean = True)
            ShowWaitForm()
            Dim groupId As String = CacadoreToTectonic.GetGroupIdByName(groupName)
            Dim tskId As String = MainFormConnector.AddNewTask(waitBetweenActions, groupId, taskName, isEnabled)
            MainFormConnector.SetSelectedNode(tskId)
            CloseWaitForm()
        End Sub
#End Region
#Region "Events Helpers"
        Shared Sub fillGroupsComboBoxEdit(groupsNames As SortedSet(Of String))
            Dim selectedGroupName As String = MainFormConnector.GetSelectedGroupName()
            With AddNewTaskForm.AddToGroupComboBoxEdit
                .Properties.Items.BeginUpdate()
                .Properties.Items.AddRange(groupsNames)
                .SelectedIndex = groupsNames.ToList.IndexOf(selectedGroupName)
                .Properties.Items.EndUpdate()
            End With
        End Sub
#End Region
    End Class
End Namespace
