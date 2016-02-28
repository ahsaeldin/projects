'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   Commands Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
Option Strict Off

Public Class Commands
    Inherits CacadoreBase
    Implements ICrudable
    Implements IDisposable

    Enum CommandsEnum
        DeleteAllTasks
        DeleteAllGroups
        DeleteAllGlobals
        DeleteAllVariables
        DeleteAllTasksStates
        DeleteAllActions
        DeleteAllTriggers
        GetTaskProperty
        GetTaskTriggersCount
        GetTaskActionsCount
        GetActionById
        GetTriggerById
        GetTaskById
        GetGroupById
        GetAllTasks
        GetTaskExecutionState
        GetAllTasksStates
        GetUnfullfilledTasks
        GetAllTriggers
        GetTimeOutTriggers
        IsRegisterdTask
        IsRegisterdAction
        MoveUpActionExecutionPrecedence
        MoveDownActionExecutionPrecedence
        IsRegisterdTaskVariable
    End Enum

#Region "Queries"
    ''' <summary>
    ''' Query task property from Cacadore database by taskId and property name passed via args()
    ''' </summary>
    '''<example>
    '''GetTaskProperty(Me.taskId, PropertiesNames.TaskName.ToString)
    '''GetTaskProperty(Me.taskId, PropertiesNames.IsTaskEnabled.ToString)
    '''GetTaskProperty(Me.taskId, PropertiesNames.LastTimeRun.ToString)
    '''GetTaskProperty(Me.taskId, PropertiesNames.LastTimeRunResult.ToString)   
    '''GetTaskProperty(Me.taskId, PropertiesNames.NextTimeToBeRun.ToString)
    '''GetTaskProperty(Me.taskId, PropertiesNames.TriggerWhen.ToString))
    ''' </example>
    Public Shared Function GetTaskProperty(taskId As String, taskProperty As TaskProperties.PropertiesNames) As Object
        Dim args() As String = {CStr(CommandsEnum.GetTaskProperty), taskId, taskProperty.ToString}
        Return executeCommand(args, QueryOperationDlg)
    End Function

    ''' <summary>
    ''' Gets the action from Tectonic by id.
    ''' </summary>
    ''' <param name="actionId">The action id.</param><returns></returns>
    Public Shared Function GetActionById(actionId As String) As Action
        Dim args() As String = {CStr(CommandsEnum.GetActionById), actionId}
        Return CType(executeCommand(args, QueryOperationDlg), Action)
    End Function

    ''' <summary>
    ''' Gets the trigger from Tectonic by id.
    ''' </summary>
    ''' <param name="triggerId">The trigger id.</param><returns></returns>
    Public Shared Function GetTriggerById(triggerId As String) As ITrigger
        Dim args() As String = {CStr(CommandsEnum.GetTriggerById), triggerId}
        Return CType(executeCommand(args, QueryOperationDlg), ITrigger)
    End Function

    Public Shared Function GetTriggerTask(triggerId As String) As Task
        Return GetTriggerById(triggerId).Task
    End Function

    Public Shared Function GetTriggerInnerTrigger(triggerId As String) As Object
        Return GetTriggerById(triggerId).InnerTrigger
    End Function

    Public Shared Function GetTriggerCalendar(triggerId As String) As Object
        Return GetTriggerById(triggerId).Calendar
    End Function

    Public Shared Function GetTaskById(taskId As String) As Task
        Dim args() As String = {CStr(CommandsEnum.GetTaskById), taskId}
        Dim queryResult As Object = executeCommand(args, QueryOperationDlg)
        Dim tsk As Task = CType(queryResult, Task)
        Return tsk
    End Function

    Public Shared Function GetTaskTriggersCount(taskId As String) As Integer
        Dim args() As String = {CStr(CommandsEnum.GetTaskTriggersCount), taskId}
        Return CType(executeCommand(args, QueryOperationDlg), Integer)
    End Function

    Public Shared Function GetTaskActionsCount(taskId As String) As Integer
        Dim args() As String = {CStr(CommandsEnum.GetTaskActionsCount), taskId}
        Return CType(executeCommand(args, QueryOperationDlg), Integer)
    End Function

    Public Shared Function GetGroupById(groupId As String) As Group
        Dim args() As String = {CStr(CommandsEnum.GetGroupById), groupId}
        Return CType(executeCommand(args, QueryOperationDlg), Group)
    End Function

    Public Shared Function GetGroupName(groupId As String) As String
        Return GetGroupById(groupId).Name
    End Function

    ''' <summary>
    ''' Gets a dictionary of all Tectonic groups.
    ''' </summary><returns></returns>
    Public Shared Function GetGroups() As Dictionary(Of String, Group)
        Dim tectonicGroups As New GroupsCollection
        Return tectonicGroups.Groups
    End Function

    Private Shared Function GetGroupId(ByVal groupName As String,
                                       ByVal groupsDic As Dictionary(Of String, Group)) As IEnumerable(Of String)
        Return From grp In groupsDic
                    Where grp.Value.Name = groupName
                    Select grp.Key
    End Function

    Public Shared Function GetTaskIdByName(ByVal taskName As String) As String
        Dim tasksDic = GetAllTasks()
        Dim queryResult = From tsks In tasksDic Where tsks.Value.Properties.TaskName = taskName Select tsks.Key
        If queryResult.Count > 0 Then Return queryResult.First
        tasksDic = Nothing
        Return ""
    End Function

    Public Shared Function GetGroupIdByName(ByVal groupName As String) As String
        Dim result As String
        Dim groupsDic = GetGroups()
        If GetGroupId(groupName, groupsDic).Count = 0 AndAlso groupName <> "Ungrouped" Then
            result = "0"
        ElseIf GetGroupId(groupName, groupsDic).Count = 0 AndAlso groupName = "Ungrouped" Then
            Return ""
        Else
            result = GetGroupId(groupName, groupsDic)(0)
        End If
        groupsDic = Nothing
        Return result
    End Function

    ''' <summary>
    ''' Gets an ordered list of all Tectonic groups names.
    ''' </summary><returns></returns>
    Public Shared Function GetGroupsNames() As SortedSet(Of String)
        Dim tectonicGroups As New GroupsCollection()
        Dim _list = tectonicGroups.Groups.Values.ToList()
        Dim sortedSet As New SortedSet(Of String)
        For Each grp In _list
            sortedSet.Add(grp.Name)
        Next
        tectonicGroups.Dispose()
        tectonicGroups = Nothing
        Return sortedSet
    End Function

    ''' <summary>
    ''' Gets all tasks registered in Tectonic.
    ''' </summary><returns></returns>
    Public Shared Function GetAllTasks() As Dictionary(Of String, Task)
        Dim args() As String = {CStr(CommandsEnum.GetAllTasks)}
        Return CType(executeCommand(args, QueryOperationDlg), Dictionary(Of String, Task))
    End Function

    ''' <summary>
    ''' Gets all registered tasks states.
    ''' </summary><returns></returns>
    Public Shared Function GetAllTasksStates() As Dictionary(Of String, TaskState)
        Dim args() As String = {CStr(CommandsEnum.GetAllTasksStates)}
        Return CType(executeCommand(args, QueryOperationDlg), Dictionary(Of String, TaskState))
    End Function

    ''' <summary>
    ''' Gets the state of the task previous execution.
    ''' </summary>
    ''' <param name="taskId">The task id.</param><returns></returns>
    Public Shared Function GetTaskPreviousExecutionState(taskId As String) As ITaskState.ExecutionState
        Dim histObj As New History
        Dim fullTskHist = History.GetTaskHistory(taskId)
        Dim taskStat = From stat In fullTskHist Where stat.TaskId = taskId Select stat.Status
        Dim statusList As List(Of ITaskState.ExecutionState) = taskStat.ToList
        If statusList.Count >= 2 Then
            'statusList.Count - 2 is the previous state index in the list.
            Return CType(statusList.Item(statusList.Count - 2), ITaskState.ExecutionState)
        Else
            Return ITaskState.ExecutionState.NotScheduled
        End If
    End Function

    Public Shared Function GetTaskExecutionState(taskId As String) As ITaskState.ExecutionState
        Dim args() As String = {CStr(CommandsEnum.GetTaskExecutionState), taskId}
        Return CType(executeCommand(args, QueryOperationDlg), ITaskState.ExecutionState)
    End Function

    Public Shared Function GetUnfullfilledTasks() As Dictionary(Of String, Task)
        Dim args() As String = {CStr(CommandsEnum.GetUnfullfilledTasks)}
        Return CType(executeCommand(args, QueryOperationDlg), Dictionary(Of String, Task))
    End Function

    Public Shared Function GetAllTriggers() As Dictionary(Of String, Trigger)
        Dim args() As String = {CStr(CommandsEnum.GetAllTriggers)}
        Return CType(executeCommand(args, QueryOperationDlg), Dictionary(Of String, Trigger))
    End Function

    Public Shared Function GetTimeOutTriggers() As Dictionary(Of String, ITrigger)
        Dim args() As String = {CStr(CommandsEnum.GetTimeOutTriggers)}
        Return CType(executeCommand(args, QueryOperationDlg), Dictionary(Of String, ITrigger))
    End Function

    Public Shared Function IsRegisterdTask(taskId As String) As Boolean
        Dim args() As String = {CStr(CommandsEnum.IsRegisterdTask), taskId}
        Return DirectCast(executeCommand(args, QueryOperationDlg), Boolean)
    End Function

    Public Shared Function GetNextRunTime(ByVal tsk As Task) As Date?
        Dim taskTriggers As TriggersCollection = tsk.MyTriggers
        Dim result = GetNextRunTime(taskTriggers)
        taskTriggers.Dispose()
        taskTriggers = Nothing
        Return result
    End Function

    Public Shared Function GetNextRunTime(trgs As TriggersCollection) As Date?
        Dim queryResult = From kvp As KeyValuePair(Of String, Cacadore.Trigger)
                          In trgs.Triggers
                          Select kvp.Value.NextFireTimeUTC
        queryResult.ToList.Sort()
        Return CDate(queryResult.First)
    End Function

    Public Shared Function IsTriggerAttachedByCronExpression(task As Cacadore.Task, cronExpression As String) As ITrigger
        Dim queryResult = From trg As Cacadore.ITrigger In task.MyTriggers.Triggers.Values
                          Where Not IsNothing(trg.InnerTrigger) AndAlso trg.TriggerType = ITrigger.TriggerTypeEnum.CronTrigger AndAlso trg.InnerTrigger.CronExpressionString = cronExpression
                          Select trg
        If queryResult.Count > 0 Then
            Dim result = queryResult.First
            queryResult = Nothing
            Return result
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function ExtractTaskId(Id As String) As String
        Dim whereIsDash = Id.IndexOf("-")
        Dim number As String = Strings.Left(Id, whereIsDash)
        Return number
    End Function

    Public Shared Function ExtractActionNumber(actionId As String) As Integer
        Dim whereIsDash = actionId.IndexOf("-")
        Dim number As String = Strings.Mid(actionId, whereIsDash + 1)
        Return CInt(number.Replace("-a", ""))
    End Function

    Public Shared Function ExtractTriggerNumber(triggerId As String) As Integer
        Dim whereIsDash = triggerId.IndexOf("-")
        Dim number As String = Strings.Mid(triggerId, whereIsDash + 1)
        Return CInt(number.Replace("-t", ""))
    End Function

    Public Shared Function IsRegisteredAction(actionId As String) As Boolean
        Dim args() As String = {CStr(CommandsEnum.IsRegisterdAction), actionId}
        Return DirectCast(executeCommand(args, QueryOperationDlg), Boolean)
    End Function

#Region "Variables & Constants"
    Public Shared Function IsRegisteredGlobalConstant(constantName As String) As Boolean
        Dim _globals As New Globals
        Return _globals.Constants.ContainsKey(constantName)
    End Function

    Public Shared Function IsRegisteredGlobalVariable(varName As String) As Boolean
        Dim _globals As New Globals
        Return _globals.Variables.ContainsKey(varName)
    End Function

    Public Shared Function IsRegisteredTaskVariable(taskId As String, varName As String) As Boolean
        Dim args() As String = {CStr(CommandsEnum.IsRegisterdTaskVariable), taskId, varName}
        Return DirectCast(executeCommand(args, QueryOperationDlg), Boolean)
    End Function

    Public Shared Function GetGlobalConstant(constantName As String) As String
        Dim value As String = vbNullString
        If IsRegisteredGlobalConstant(constantName) Then
            Dim _globals As New Globals
            _globals.Constants.TryGetValue(constantName, value)
        End If
        Return value
    End Function

    Public Shared Function GetGlobalVariable(varName As String) As String
        Dim value As String = vbNullString
        If IsRegisteredGlobalVariable(varName) Then
            Dim _globals As New Globals
            _globals.Variables.TryGetValue(varName, value)
        End If
        Return value
    End Function

    Public Shared Function GetTaskVariable(ByVal task As Task, varName As String) As String
        If IsNothing(task) Then Return False
        Dim value As String = vbNullString
        If IsRegisteredTaskVariable(task.TaskId, varName) Then
            value = task.Variables.GetVariableValue(varName)
        End If
        Return value
    End Function

    Public Shared Function GetTaskVariable(taskId As String, varName As String) As String
        If taskId = "" Then Return False
        Dim tsk As Task = GetTaskById(taskId)
        Dim value = GetTaskVariable(tsk, varName)
        Return value
    End Function

    Public Shared Function GetTaskVariables(taskId As String) As VariablesCollection
        If taskId = "" Then Return Nothing
        Using tskVars As New TaskVariables(taskId)
            Return tskVars.Variables
        End Using
    End Function

    Public Shared Function GetMemberValue(taskId As String, member As String) As String
        Dim memberValue As String = vbNullString
        If Commands.IsRegisteredGlobalConstant(member) Then
            memberValue = Commands.GetGlobalConstant(member)
        ElseIf Commands.IsRegisteredGlobalVariable(member) Then
            memberValue = Commands.GetGlobalVariable(member)
        ElseIf Commands.IsRegisteredTaskVariable(taskId, member) Then
            memberValue = Commands.GetTaskVariable(taskId, member)
        End If
        Return memberValue
    End Function
#End Region
#End Region

#Region "Delete(s)"
    Public Shared Function DeleteAllTasks() As Boolean
        Dim args() As String = {CStr(CommandsEnum.DeleteAllTasks)}
        Return CType(executeCommand(args, DeleteOperationDlg), Boolean)
    End Function

    Public Shared Function DeleteAllGroups() As Boolean
        Dim args() As String = {CStr(CommandsEnum.DeleteAllGroups)}
        Return CType(executeCommand(args, DeleteOperationDlg), Boolean)
    End Function

    ''' <summary>
    ''' Deletes all globals.
    ''' </summary><returns></returns>
    Public Shared Function DeleteAllGlobals() As Boolean
        Dim args() As String = {CStr(CommandsEnum.DeleteAllGlobals)}
        Return CType(executeCommand(args, DeleteOperationDlg), Boolean)
    End Function

    ''' <summary>
    ''' Deletes all variables.
    ''' </summary><returns></returns>
    Public Shared Function DeleteAllVariables() As Boolean
        Dim args() As String = {CStr(CommandsEnum.DeleteAllVariables)}
        Return CType(executeCommand(args, DeleteOperationDlg), Boolean)
    End Function

    ''' <summary>
    ''' Deletes all task states.
    ''' </summary><returns></returns>
    Public Shared Function DeleteAllTaskStates() As Boolean
        Dim args() As String = {CStr(CommandsEnum.DeleteAllTasksStates)}
        Return CType(executeCommand(args, DeleteOperationDlg), Boolean)
    End Function

    ''' <summary>
    ''' Deletes all actions.
    ''' </summary><returns></returns>
    Public Shared Function DeleteAllActions() As Boolean
        Dim args() As String = {CStr(CommandsEnum.DeleteAllActions)}
        Return CType(executeCommand(args, DeleteOperationDlg), Boolean)
    End Function

    ''' <summary>
    ''' Deletes all triggers.
    ''' </summary><returns></returns>
    Public Shared Function DeleteAllTriggers() As Boolean
        Dim args() As String = {CStr(CommandsEnum.DeleteAllTriggers)}
        Return CType(executeCommand(args, DeleteOperationDlg), Boolean)
    End Function

    ''' <summary>
    ''' Deletes the time out triggers.
    ''' </summary>
    ''' <returns>Return false if one/all triggers cannot be deleted.</returns>
    Public Shared Function DeleteTimeOutTriggers() As Boolean
        Dim timeOutTriggers = Watcher.TimeOutTriggers
        For Each trg In timeOutTriggers
            If Not trg.Value.Delete() Then Return False
        Next
        Return True
    End Function

    Public Shared Function DeleteTaskVariables(taskId As String) As Boolean
        Using tskVars As New TaskVariables(taskId)
            For Each var As Variable In tskVars.Variables
                If Not tskVars.DeleteVariable(var.Name) Then Return False
            Next
        End Using
        Return True
    End Function
#End Region

#Region "Reusable"
#Region "Execution Commands"
    Private Shared Function executeCommand(args() As String, commandType As [Delegate]) As Object
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Dim res As Object = InvokeTectonicDlg(New Commands, args, commandType, My.Resources.CmdNotPresisted)
            Return res
        End If
    End Function
#End Region

    ' ''' <summary>
    ' ''' Creates new type of the action by IAction.ActionTypeEnum
    ' ''' </summary>
    ' ''' <param name="actionType">Type of the action.</param><returns></returns>
    'Public Shared Function CreateActionByType(actionType As IAction.ActionTypeEnum) As Action
    '    Select Case actionType
    '        Case IAction.ActionTypeEnum.StartApplication
    '            Return New Actions.System.StartApplication
    '        Case IAction.ActionTypeEnum.ShowAlert
    '            Return New Actions.MessageBoxes.ShowAlert
    '        Case IAction.ActionTypeEnum.ShowMessageBox
    '            Return New Actions.MessageBoxes.ShowMessageBox
    '        Case IAction.ActionTypeEnum.TakeScreenShot
    '            Return New Actions.System.TakeScreenShot
    '        Case IAction.ActionTypeEnum.ExitProcess
    '            Return New Actions.System.ExitProcess
    '        Case IAction.ActionTypeEnum.ShutdownPC
    '            Return New Actions.System.ShutdownPC()
    '        Case IAction.ActionTypeEnum.LastFacebookNotification
    '            '  Why not just to get type using reflection from fbactions.dll then return it here
    '            '  Return
    '    End Select
    '    Return Nothing
    'End Function

    ''' <summary>
    ''' Determines whether [id] is a trigger id.
    ''' </summary>
    ''' <param name="id">The id.</param><returns>
    '''   <c>true</c> if [is trigger id] [the specified id]; otherwise, <c>false</c>.
    ''' </returns>
    Public Shared Function IsValidTriggerId(id As String) As Boolean
        If IsNumeric(id.Last) Then
            Dim last3Chars As String = Right(id, 3)
            If Left(last3Chars, 2) = "-t" Then
                Return True
            Else
                Return False
            End If
        End If
        Return False
    End Function

    Public Shared Function IsValidActionId(id As String) As Boolean
        If IsNumeric(id.Last) Then
            Dim last3Chars As String = Right(id, 3)
            If Left(last3Chars, 2) = "-a" Then
                Return True
            Else
                Return False
            End If
        End If
        Return False
    End Function

    Public Shared Function MoveUpActionExecutionPrecedence(actionId As String, newPrecedence As Integer) As Boolean
        Dim args() As String = {CStr(CommandsEnum.MoveUpActionExecutionPrecedence), actionId, newPrecedence}
        Return CType(executeCommand(args, UpdateOperationDlg), Boolean)
    End Function

    Public Shared Function MoveDownActionExecutionPrecedence(actionId As String, newPrecedence As Integer) As Boolean
        Dim args() As String = {CStr(CommandsEnum.MoveDownActionExecutionPrecedence), actionId, newPrecedence}
        Return CType(executeCommand(args, UpdateOperationDlg), Boolean)
    End Function
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
