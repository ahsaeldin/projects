'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   CacadoreToTectonic Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
'An incredible class that connects between Crud layers and  Cacadore classes, it creates Cacadore classes then save to tectonic using Crud
'+Just need to be abstracted by inject depends (TypedDataSetCrud or any other crud layer)
Imports Cacadore
Imports TecDAL

'This file added as a link to Cacadu And integration tests project, so you can hide it from reflectors.
Friend Class CacadoreToTectonic

#Region "Query Area"
    ''' <summary>
    ''' Gets full history of tasks only.
    ''' </summary><returns></returns>
    Public Shared Function GetTasksHistory() As List(Of HistoryEntry)
        Return History.GetFullTasksHistory()
    End Function

    ''' <summary>
    ''' Gets full history of actions only.
    ''' </summary><returns></returns>
    Public Shared Function GetActionsHistory() As List(Of HistoryEntry)
        Return History.GetFullActionsHistory
    End Function

    ''' <summary>
    ''' Gets full history of triggers only.
    ''' </summary><returns></returns>
    Public Shared Function GetTriggersHistory() As List(Of HistoryEntry)
        Return History.GetFullTriggersHistory
    End Function

    ''' <summary>
    ''' Gets the single task history using its id
    ''' </summary>
    ''' <param name="taskId">The task id.</param><returns></returns>
    Public Shared Function GetSingleTaskHistory(taskId As String) As List(Of HistoryEntry)
        Return History.GetTaskHistory(taskId)
    End Function

    ''' <summary>
    ''' Gets a dictionary of all members in the Globals 
    ''' by choosing members type(either variables or constants).
    ''' </summary>
    ''' <param name="globalMemberType">Type of the global member.</param><returns></returns>
    Public Shared Function GetGlobals(ByVal globalMemberType As IGlobals.GlobalType) As Dictionary(Of String, String)
        Dim glob As New Cacadore.Globals
        If globalMemberType = IGlobals.GlobalType.Constant Then
            Return glob.Constants
        ElseIf globalMemberType = IGlobals.GlobalType.Variable Then
            Return glob.Variables
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Gets a dictionary of all Tectonic groups.
    ''' </summary><returns></returns>
    Public Shared Function GetGroups() As Dictionary(Of String, Group)
        Return Commands.GetGroups()
    End Function

    ''' <summary>
    ''' Gets an ordered list of all Tectonic groups names.
    ''' </summary><returns></returns>
    Public Shared Function GetGroupsNames() As SortedSet(Of String)
        Return Commands.GetGroupsNames()
    End Function

    Public Shared Function GetGroupIdByName(ByVal groupName As String) As String
        Return Commands.GetGroupIdByName(groupName)
    End Function

    ''' <summary>
    ''' Example in how to get settings from Settings table.
    ''' </summary><returns></returns>
    Public Shared Function GetSettings() As Dictionary(Of String, Boolean)
        Dim settingsList As New Dictionary(Of String, Boolean)
        settingsList.Add("CacadoreEnabled", Settings.Enabled)
        settingsList.Add("HistoryEnabled", Settings.HistoryEnabled)
        settingsList.Add("TriggersEnabled", Settings.TriggersEnabled)
        Return settingsList
    End Function

    ''' <summary>
    ''' Give me a task, give you its variables in return.
    ''' </summary>
    ''' <param name="tsk">The TSK.</param><returns></returns>
    Public Shared Function GetTaskVariables(tsk As Task) As VariablesCollection
        Return tsk.Variables.Variables
    End Function

    ''' <summary>
    ''' Give me a task and a variable name, 
    ''' give you variable value in return.
    ''' </summary>
    ''' <param name="tsk">The TSK.</param>
    ''' <param name="varName">Name of the var.</param><returns></returns>
    Public Shared Function GetTaskVariableValue(ByVal tsk As Task, varName As String) As String
        Dim varValue = tsk.Variables.GetVariableValue(varName)
        Return varValue
    End Function

    ''' <summary>
    ''' Queries the name of the task.
    ''' </summary>
    ''' <param name="tsk">The TSK.</param><returns></returns>
    Public Shared Function QueryTaskName(ByVal tsk As Task) As String
        Dim tskName As String = CStr(Commands.GetTaskProperty(tsk.TaskId, TaskProperties.PropertiesNames.TaskName))
        Return tskName
    End Function

    ''' <summary>
    ''' Queries the state of the task enabled.
    ''' </summary>
    ''' <param name="tsk">The TSK.</param><returns></returns>
    Public Shared Function QueryTaskEnabledState(ByVal tsk As Task) As Boolean
        Dim isEnabled = CBool(Commands.GetTaskProperty(tsk.TaskId, TaskProperties.PropertiesNames.IsTaskEnabled))
        Return isEnabled
    End Function

    ''' <summary>
    ''' Queries the task last time run.
    ''' </summary>
    ''' <param name="tsk">The TSK.</param><returns></returns>
    Public Shared Function QueryTaskLastTimeRun(ByVal tsk As Task) As Date
        Dim lastTimeRunDate = CDate(Commands.GetTaskProperty(tsk.TaskId, TaskProperties.PropertiesNames.LastTimeRun))
        Return lastTimeRunDate
    End Function

    ''' <summary>
    ''' Queries the task last time run result.
    ''' </summary>
    ''' <param name="tsk">The TSK.</param><returns></returns>
    Public Shared Function QueryTaskLastTimeRunResult(ByVal tsk As Task) As String
        Dim lastTimeRunResult = CStr(Commands.GetTaskProperty(tsk.TaskId, TaskProperties.PropertiesNames.LastTimeRunResult))
        Return lastTimeRunResult
    End Function

    ''' <summary>
    ''' Queries the task next time to be run.
    ''' </summary>
    ''' <param name="tsk">The TSK.</param><returns></returns>
    Public Shared Function QueryTaskNextTimeToBeRun(ByVal tsk As Task) As Date
        Dim nextTimeToBeRun = CDate(Commands.GetTaskProperty(tsk.TaskId, TaskProperties.PropertiesNames.NextTimeToBeRun))
        Return nextTimeToBeRun
    End Function

    ''' <summary>
    ''' Queries the task triggers when.
    ''' </summary>
    ''' <param name="tsk">The TSK.</param><returns></returns>
    ''' <remarks>
    ''' For both, they will be blank for now until I finish the 
    ''' logic of triggers Tasks>=Triggers_When ... TaskProperties.Triggers_When
    ''' Tasks_States>=Next_Trigger_Id ... TaskState.NextTriggerId
    ''' </remarks>
    Public Shared Function QueryTaskTriggersWhen(ByVal tsk As Task) As String
        Dim triggersWhen = CStr(Commands.GetTaskProperty(tsk.TaskId, TaskProperties.PropertiesNames.TriggerWhen))
        Return triggersWhen
    End Function

    ''' <summary>
    ''' Gets the task triggers.
    ''' </summary>
    ''' <param name="tsk">The TSK.</param><returns></returns>
    Public Shared Function GetTaskTriggers(tsk As Task) As Dictionary(Of String, Trigger)
        Return tsk.MyTriggers.Triggers
    End Function

    ''' <summary>
    ''' Gets the task triggers.
    ''' </summary>
    Public Shared Function GetTaskTriggers(taskId As String) As Dictionary(Of String, Trigger)
        Dim tsk = GetTaskById(taskId)
        Dim trgs = tsk.MyTriggers.Triggers
        Return trgs
    End Function

    ''' <summary>
    ''' Gets an action by its id.
    ''' </summary>
    ''' <param name="actionId">The action id.</param><returns></returns>
    Public Shared Function GetTaskActionById(actionId As String) As Action
        Return Commands.GetActionById(actionId)
    End Function

    ''' <summary>
    ''' Gets a trigger by its id.
    ''' </summary>
    ''' <param name="triggerId">The trigger id.</param><returns></returns>
    Public Shared Function GetTaskTriggerById(triggerId As String) As Trigger
        Return CType(Commands.GetTriggerById(triggerId), Trigger)
    End Function

    ''' <summary>
    ''' Give me a group id, give you all tasks 
    ''' associated with it in a TasksCollection in return.
    ''' </summary>
    ''' <param name="groupId">The group id.</param><returns></returns>
    Public Shared Function GetTasksCollection(groupId As String) As TasksCollection
        Dim tempGroup As New Group(groupId)
        Return tempGroup.TasksCollection
    End Function

    ''' <summary>
    ''' Gets a task by it id.
    ''' </summary>
    ''' <param name="taskId">The task id.</param><returns></returns>
    Public Shared Function GetTaskById(taskId As String) As Task
        Return Commands.GetTaskById(taskId)
    End Function

    Public Shared Function GetTaskPropertiesById(taskId As String) As TaskProperties
        Return GetTaskById(taskId).Properties
    End Function

    ''' <summary>
    ''' Gets a group by its id.
    ''' </summary>
    ''' <param name="groupId">The group id.</param><returns></returns>
    Public Shared Function GetGroupById(groupId As String) As Group
        Return Commands.GetGroupById(groupId)
    End Function

    ''' <summary>
    ''' Gets a group by its id.
    ''' </summary>
    Public Shared Function GetGroupByName(name As String) As Group
        Return Commands.GetGroupById(GetGroupIdByName(name))
    End Function

    ''' <summary>
    ''' Gets all tasks registered in Tectonic.
    ''' </summary><returns></returns>
    Public Shared Function GetAllRegisteredTasks() As Dictionary(Of String, Task)
        Return Commands.GetAllTasks()
    End Function

    ''' <summary>
    ''' Gets all registered tasks states.
    ''' </summary><returns></returns>
    Public Shared Function GetAllRegisteredTasksStates() As Dictionary(Of String, TaskState)
        Return Commands.GetAllTasksStates
    End Function

    Public Shared Function GetAllTriggers() As Dictionary(Of String, Trigger)
        Return Commands.GetAllTriggers
    End Function

    Public Shared Function IsRegisteredGroupName(name As String) As Boolean
        Dim groupsDic = GetGroups()
        Dim result = From grp In groupsDic Where grp.Value.Name = name Select grp.Value
        If result.Count <= 0 Then Return False
        Return True
    End Function

    Public Shared Function IsRegisteredTaskId(taskId As String) As Boolean
        Return Commands.IsRegisterdTask(taskId)
    End Function

    Public Shared Function GetTaskName(taskId As String) As String
        Return Commands.GetTaskProperty(taskId, TaskProperties.PropertiesNames.TaskName).ToString
    End Function
#End Region

#Region "Insertion Area"
    ''' <summary>
    ''' Creates a new group.
    ''' </summary>
    ''' <param name="name">The name.</param><returns></returns>
    Public Shared Function CreateNewGroup(name As String) As Group
        Dim newGroup As New Group(TypedDataSetCrud.gngi)
        With newGroup
            newGroup.Name = name
        End With
        If newGroup.Save() Then
            Return newGroup
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Creates a new task.
    ''' </summary>
    ''' <param name="waitBetweenActions">if set to <c>true</c> [wait between actions].</param>
    ''' <param name="groupid">The groupid.</param><returns></returns>
    Public Shared Function CreateNewTask(waitBetweenActions As Boolean,
                                         Optional groupId As String = "",
                                         Optional taskName As String = vbNullString,
                                         Optional isEnabled As Boolean = True) As Task
        Dim taskId As String = TypedDataSetCrud.gnti
        Dim newTask As New Task(taskId)
        'Dim taskVar As New Cacadore.TaskVariables(taskId)
        With (newTask)
            If taskName <> vbNullString Then .Properties.TaskName = taskName
            .Properties.IsEnabled = isEnabled
            .ParentGroupId = groupId
            .WaitBetweenActions = waitBetweenActions
        End With
        If newTask.Save() Then
            Return newTask
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Adds a new global variable.
    ''' </summary>
    ''' <param name="name">The name.</param>
    ''' <param name="value">The value.</param><returns></returns>
    Public Shared Function AddGlobalVariable(name As String, value As String) As Boolean
        Dim globalsObj As New Cacadore.Globals
        Return globalsObj.AddVariable(name, value)
    End Function

    ''' <summary>
    ''' Adds a new global constant.
    ''' </summary>
    ''' <param name="name">The name.</param>
    ''' <param name="value">The value.</param><returns></returns>
    Public Shared Function AddNewGlobalConstant(name As String, value As String) As Boolean
        Dim globalsObj As New Cacadore.Globals
        Return globalsObj.AddConstant(name, value)
    End Function

    ''' <summary>
    ''' Adds a new task variable.
    ''' </summary>
    ''' <param name="taskId">The task id.</param>
    ''' <param name="name">The name.</param>
    ''' <param name="value">The value.</param>
    Public Sub AddNewTaskVariable(taskId As String, name As String, value As String)
        Dim taskVar As New Cacadore.TaskVariables(taskId)
        taskVar.AddVariable(name, value)
    End Sub

    ''' <summary>
    ''' Attaches a dump trigger to task.
    ''' </summary>
    ''' <param name="tsk">The TSK.</param><returns></returns>
    Public Shared Function AttachDumpTriggerToTask(ByRef tsk As Task) As Trigger
        Dim trg As New Triggers.DumpTrigger
        tsk.AttachTrigger(trg)
        Return trg
    End Function

    Public Shared Function AttachScheduleTriggerToTask(ByRef tsk As Task) As Trigger
        Dim trg As New Triggers.ScheduleTrigger
        tsk.AttachTrigger(trg)
        Return trg
    End Function

    ''' <summary>
    ''' Attaches an action to a task.
    ''' </summary>
    ''' <param name="newTask">The new task.</param>
    ''' <param name="actionType">Type of the action.</param>
    ''' <param name="description">The description.</param>
    ''' <param name="IsInactive">if set to <c>true</c> [is inactive].</param>
    ''' <param name="actionInput">The action input.</param>
    ''' <param name="actionOutput">The action output.</param>
    ''' <param name="actionShape">The action shape.</param><returns></returns>
    Public Shared Function AttachActionToTask(ByVal newTask As Task,
                                              actionType As IAction.ActionTypeEnum,
                                              description As String,
                                              IsInactive As Boolean,
                                              actionInput As IAction.IInputs,
                                              actionOutput As IAction.IOutputs,
                                              actionShape As Object) As Action
        Dim newAction As Action = Nothing
        Select Case actionType
            Case IAction.ActionTypeEnum.StartApplication
                newAction = New Cacadore.Actions.System.StartApplication
            Case IAction.ActionTypeEnum.ShowAlert
                newAction = New Cacadore.Actions.MessageBoxes.ShowAlert
            Case IAction.ActionTypeEnum.ShowMessageBox
                newAction = New Cacadore.Actions.MessageBoxes.ShowMessageBox
            Case IAction.ActionTypeEnum.TakeScreenShot
                newAction = New Cacadore.Actions.System.TakeScreenShot
            Case IAction.ActionTypeEnum.ExitProcess
                newAction = New Cacadore.Actions.System.ExitProcess
            Case IAction.ActionTypeEnum.ShutdownPC
                newAction = New Cacadore.Actions.System.ShutdownPC
            Case IAction.ActionTypeEnum.LastFacebookNotification
                newAction = New Actions.FacebookActions.LastFacebookNotification
        End Select
    
        newAction.Comments = description
        newAction.Enabled = IsInactive
        newAction.Inputs = actionInput
        newAction.Outputs = actionOutput
        newAction.Result = ITaskState.ExecutionState.NotScheduled
        newAction.Shape = New Shape
        Dim actionId As String = newTask.AttachAction(newAction)
        Return newAction
    End Function

    ''' <summary>
    ''' Inserts a new history entry.
    ''' </summary>
    ''' <param name="hisEntry">The new entry.</param><returns></returns>
    Public Shared Function InsertNewHistoryEntry(hisEntry As HistoryEntry) As Boolean
        Return History.AddNewEntry(hisEntry)
    End Function

    Public Shared Function CloneGroup(groupId As String) As Group
        Dim groupObject As Group = GetGroupById(groupId)

        Dim actionMemStream As New Global.System.IO.MemoryStream
        Balora.Serializer.BinarySerializer.SerializeObject(groupObject, actionMemStream)
        Dim duplicatedGroup As Group = CType(Balora.Serializer.BinarySerializer.DeserializeObject(actionMemStream.ToArray), Group)

        Dim clonedTasksList As New List(Of Task)
        Dim newGroupId As String = TypedDataSetCrud.gngi
        For Each tsk As Task In groupObject.TasksCollection().Tasks.Values
            Dim newTaskId As String = TypedDataSetCrud.gnti
            Balora.Util.SetPrivateField(tsk, "_taskIdValue", newTaskId)
            tsk.ParentGroupId = newGroupId
            tsk.SetTaskState(New TaskState(newTaskId))
            clonedTasksList.Add(tsk.Clone)
            MsgBox("Save has bugs, develop in the future")
            tsk.Save()
            tsk.Dispose()
            tsk = Nothing
        Next

        Balora.Util.SetPrivateField(duplicatedGroup, "_groupIdValue", newGroupId)
        duplicatedGroup.Save()
        actionMemStream.Dispose()
        actionMemStream = Nothing
        Return duplicatedGroup
    End Function
#End Region

#Region "Update Area"
    ''' <summary>
    ''' Updates a task.
    ''' </summary>
    ''' <param name="task">The task.</param><returns></returns>
    Public Shared Function UpdateTask(ByRef task As Task) As Boolean
        Return task.Save()
    End Function

    ''' <summary>
    ''' Updates Cacadore settings.
    ''' </summary>
    Public Shared Function UpdateAllSettings(settEntry As Settings.SettingEntry) As Boolean
        Return Settings.Save(settEntry)
    End Function

    ''' <summary>
    ''' Updates Cacadore settings.
    ''' </summary>
    Public Shared Function UpdateSetting(settingName As ISettings.SettingsEnum, settingValue As Boolean) As Boolean
        Return Settings.Save(settingName, settingValue)
    End Function

    ''' <summary>
    ''' Updates a global variable.
    ''' </summary>
    ''' <param name="name">The name.</param>
    ''' <param name="value">The value.</param><returns></returns>
    Public Shared Function UpdateGlobalVariable(name As String, value As String) As Boolean
        Dim globalsObj As New Cacadore.Globals
        Return globalsObj.EditVariable(name, value)
    End Function

    ''' <summary>
    ''' Updates a group.
    ''' </summary>
    ''' <param name="group">The group.</param>
    ''' <param name="newName">The new name.</param><returns></returns>
    Public Shared Function ChangeGroupName([group] As Group, newName As String) As Boolean
        [group].Name = newName
        Return [group].Save()
    End Function

    ''' <summary>
    ''' Updates a task variable.
    ''' </summary>
    ''' <param name="taskId">The task id.</param>
    ''' <param name="name">The name.</param>
    ''' <param name="value">The value.</param><returns></returns>
    Public Shared Function UpdateTaskVariable(taskId As String, name As String, value As String) As Boolean
        Dim taskVarObj As New Cacadore.TaskVariables(taskId)
        Return (taskVarObj.UpdateVariable(name, value))
    End Function

    ''' <summary>
    ''' Updates a task action.
    ''' </summary>
    ''' <param name="action">The action.</param>
    ''' <param name="task">The task.</param><returns></returns>
    Public Shared Function UpdateTaskAction([action] As Action, ByVal [task] As Task) As Boolean
        Dim res = [task].UpdateAction([action])
        Return res
    End Function
#End Region
#Region "Delete Area"
    ''' <summary>
    ''' Deletes a task.
    ''' </summary>
    ''' <param name="task">The task.</param><returns></returns>
    Public Shared Function DeleteTask(ByRef task As Task) As Boolean
        Dim result As Boolean = task.Delete()
        If result Then task.Dispose()
        Return result
    End Function

    ''' <summary>
    ''' Deletes a task.
    ''' </summary>
    Public Shared Function DeleteTask(taskId As String) As Boolean
        Dim tsk = CacadoreToTectonic.GetTaskById(taskId)
        Dim res = CacadoreToTectonic.DeleteTask(tsk)
        tsk.Dispose()
        tsk = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Deletes all tectonic tasks.
    ''' </summary><returns></returns>
    Public Shared Function DeleteAllTectonicTasks() As Boolean
        Return Commands.DeleteAllTasks()
    End Function

    ''' <summary>
    ''' Deletes a group.
    ''' </summary>
    ''' <param name="group">The group.</param><returns></returns>
    Public Shared Function DeleteGroup(group As Group) As Boolean
        Return group.Delete
    End Function

    ''' <summary>
    ''' Deletes a group.
    ''' </summary>
    Public Shared Function DeleteGroup(groupId As String) As Boolean
        Dim grp = CacadoreToTectonic.GetGroupById(groupId)
        Dim res = CacadoreToTectonic.DeleteGroup(grp)
        Return res
    End Function

    Public Shared Function DeleteGroupWithTasks(groupId As String) As Boolean
        Dim groupToBeDeleted = CacadoreToTectonic.GetGroupById(groupId)
        Dim groupTasks = CacadoreToTectonic.GetTasksCollection(groupId)
        For Each tsk In groupTasks.Tasks.Values
            If Not tsk.Delete() Then
                Return False
            End If
        Next
        If groupToBeDeleted.Delete() Then Return True
        Return False
    End Function

    ''' <summary>
    ''' Deletes all groups.
    ''' </summary><returns></returns>
    Public Shared Function DeleteAllGroups() As Boolean
        Return Commands.DeleteAllGroups()
    End Function

    ''' <summary>
    ''' Deletes a global variable.
    ''' </summary>
    ''' <param name="globalVariableName">Name of the global variable.</param><returns></returns>
    Public Shared Function DeleteGlobalVariable(globalVariableName As String) As Boolean
        Dim globVar As New Cacadore.Globals
        Return globVar.RemoveMemberByName(globalVariableName)
    End Function

    ''' <summary>
    ''' Deletes all globals. (variables and constants).
    ''' </summary><returns></returns>
    Public Shared Function DeleteAllGlobals() As Boolean
        Return Commands.DeleteAllGlobals()
    End Function

    ''' <summary>
    ''' Deletes a single task variable.
    ''' </summary>
    ''' <param name="variables">The variables.</param>
    ''' <param name="variableName">Name of the variable.</param><returns></returns>
    Public Shared Function DeleteSingleTaskVariable(variables As TaskVariables, variableName As String) As Boolean
        Return variables.DeleteVariable(variableName)
    End Function

    ''' <summary>
    ''' Deletes all tasks variables registered in Tectonic.
    ''' </summary><returns></returns>
    Public Shared Function DeleteAllTasksVariables() As Boolean
        Return Commands.DeleteAllVariables()
    End Function

    ''' <summary>
    ''' Deletes the state of a task.
    ''' </summary>
    ''' <param name="tskState">State of the TSK.</param><returns></returns>
    Public Shared Function DeleteTaskState(tskState As TaskState) As Boolean
        Return tskState.Clear()
    End Function

    ''' <summary>
    ''' Deletes all tasks states registered in Tectonic.
    ''' </summary>
    ''' <param name="tskState">State of the TSK.</param><returns></returns>
    Public Shared Function DeleteAllTasksStates(tskState As TaskState) As Boolean
        Return Commands.DeleteAllTaskStates
    End Function

    ''' <summary>
    ''' Deletes a single trigger.
    ''' </summary>
    ''' <param name="trigger">The trigger.</param><returns></returns>
    Public Shared Function DeleteSingleTrigger(trigger As Trigger) As Boolean
        Return (trigger.Delete())
    End Function

    ''' <summary>
    ''' Deletes all triggers registered in Tectonic.
    ''' </summary><returns></returns>
    Public Shared Function DeleteAllTriggers() As Boolean
        Return Commands.DeleteAllTriggers()
    End Function

    ''' <summary>
    ''' Deletes the single action.
    ''' </summary>
    ''' <param name="action">The action.</param>
    ''' <param name="tsk">Action's Task</param>
    Public Shared Function DeleteSingleAction([action] As Action, Optional ByVal tsk As Task = Nothing) As Boolean
        Dim res = tsk.UnAttachAction([action])
        Return res
    End Function

    ''' <summary>
    ''' Deletes all actions. registered in Tectonic.
    ''' </summary><returns></returns>
    Public Shared Function DeleteAllActions() As Boolean
        Return Commands.DeleteAllActions
    End Function

    ''' <summary>
    ''' Clears the full history.
    ''' </summary><returns></returns>
    Public Shared Function ClearFullHistory() As Boolean
        Return History.ClearHistory()
    End Function

    ''' <summary>
    ''' Clears the type of the history by entry type.
    ''' </summary>
    ''' <param name="byEntryType">Type of the by entry.</param><returns></returns>
    Public Shared Function ClearHistoryByEntryType(byEntryType As IHistoryEntry.HistoryEntryType) As Boolean
        Return History.ClearHistory(byEntryType)
    End Function

    ''' <summary>
    ''' Clears the history by time range like FF.
    ''' </summary>
    ''' <param name="clearTimeRange">The clear time range.</param>
    ''' <param name="customDate">The custom date.</param><returns></returns>
    Public Shared Function ClearHistoryByTimeRange(ByVal clearTimeRange As IHistory.TimeRange, Optional ByVal customDate As Date = #12:00:00 AM#) As Boolean
        Dim hist As New History
        Return hist.ClearHistory(clearTimeRange, customDate)
    End Function
#End Region

#Region "Helper Methods"

#End Region
End Class