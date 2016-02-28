'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   Watcher Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
''' <summary>
''' This Class watch "Task, Action and Trigger" classes.
''' 1.Tells you what tasks are running now by what triggers.
''' 2.What all the Scheduled/upcoming tasks by what triggers.
''' 3.What is unfulfilled tasks and @ what actions have been stopped.
''' 
''' The philosophy of Watcher is any task added to watcher will not be unwatched unless:
''' 1.System restart/shutdown
''' 2.Task.Delete
''' 3.Task.DispoFse
''' </summary>
Public Class Watcher
    Inherits CacadoreBase
    Implements ICrudable
    Implements IDisposable

#Region "Fields"
    Dim _interval As Integer = 1000 * 60 * 15 '15 minute
    Dim _due As Integer = 1000 * 60 * 15 '15 minute
    Private endedTriggersCheckerXTimer As Balora.xTimer = New Balora.xTimer(_due,
                                                                            _interval,
                                                                            True,
                                                                            Sub() updateTasks(New Object))
    Private Shared _watchedTasksValue As New Dictionary(Of String, Task)
    Private Shared _watchedTasksStatesValue As New Dictionary(Of String, ITaskState.ExecutionState)
    Friend Shared _isBusy As Boolean = False
#End Region

#Region "Properties"
    ''' <summary>
    ''' Gets a value indicating whether Cacadore is busy in running tasks.
    ''' </summary>
    ''' <value>
    ''' Returns <c>true</c> if set to true from other members, 
    ''' Returns <c>true</c> if any task is running,
    ''' Returns <c>false</c> if no task is running and not set
    ''' to <c>false</c> from other Cacadore constructs.
    ''' </value>
    Public Shared ReadOnly Property IsBusy() As Boolean
        Get
            'If _isBusy is set to true from other Cacadore
            'constructs then فبها ونعمت
            'else if _isBusy is false then, make a check
            'no task is running...
            checkIfBusy()
            'Return true if set to true from other members, 
            'Return true if any task is running,
            'Return false if no task is running and not set
            'to false from other Cacadore constructs
            Return _isBusy
        End Get
    End Property

    Public Shared ReadOnly Property IsWatching As Boolean
        Get
            If WatchedTasks.Count > 0 Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    Public Shared ReadOnly Property IsExecuting As Boolean
        Get
            Dim queryResult = From _stateKVP
                              In WatchedTasksStates
                              Where _stateKVP.Value = ITaskState.ExecutionState.Running
            If queryResult.Count > 0 Then
                queryResult = Nothing
                Return True
            Else
                queryResult = Nothing
                Return False
            End If
        End Get
    End Property

    Public Shared ReadOnly Property WatchedTasks As Dictionary(Of String, Task)
        Get
            Return _watchedTasksValue
        End Get
    End Property

    Public Shared ReadOnly Property WatchedTasksStates As Dictionary(Of String, ITaskState.ExecutionState)
        Get
            Return _watchedTasksStatesValue
        End Get
    End Property

    'Refreshed every 50 seconds by Successive Checks.
    Public Shared ReadOnly Property UnFullFilledTasks As Dictionary(Of String, Task)
        Get
            Return Commands.GetUnfullfilledTasks()
        End Get
    End Property

    Public Shared ReadOnly Property TimeOutTriggers As Dictionary(Of String, ITrigger)
        Get
            Return Commands.GetTimeOutTriggers
        End Get
    End Property

    Public Shared Property ScheduledTriggers As New Dictionary(Of String, Trigger)

    Public Shared Property ToBeCancelledTasks As New List(Of String)
#End Region

#Region "Methods"
    Public Shared Sub Watch(taskId As String, Optional bybassSave As Boolean = False)
        Watch(Commands.GetTaskById(taskId), bybassSave)
    End Sub

    Public Shared Sub Watch(ByVal taskToBeWatched As Task, Optional bybassSave As Boolean = False)
        '1.Check if taskId is already registered in tasksExecutionStatesValue or not and add if not
        'or update tasksExecutionStatesValue if registered.
        '2.Raise up the appropriate events.
        If Not _watchedTasksValue.ContainsKey(taskToBeWatched.TaskId) Then 'Task isn't in Watcher.
            addNewTask(taskToBeWatched)
            If Not bybassSave Then taskToBeWatched.Save()
        Else 'Task already added to Watcher.
            udpateExistingTask(taskToBeWatched)
        End If
    End Sub

    Public Shared Sub AddNewEntry(taskId As String,
                                  taskName As String,
                                  entryType As IHistoryEntry.HistoryEntryType,
                                  description As String,
                                  entryDateTime As DateTime,
                                  status As ITaskState.ExecutionState)
        description = getFullStateDescription(taskName, description)
        Dim histEntry = createHistoryEntry(taskId, entryType, description, entryDateTime, status)
        RaiseEvent TaskState_Changed(taskName, histEntry)
    End Sub

    Public Shared Function UnWatch(ByVal taskToBeWatched As Task) As Boolean
        Try
            If _watchedTasksValue.ContainsKey(taskToBeWatched.TaskId) Then
                removeTask(taskToBeWatched)
            End If
            taskToBeWatched.Dispose()
            taskToBeWatched = Nothing
            Return True
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            _privateRaiseErrorUp("", ex, True)
            Return False
        End Try
    End Function

    Public Shared Sub AddScheduledTrigger(ByVal trigger As ITrigger)
        SyncLock Watcher.ScheduledTriggers
            If Not ScheduledTriggers.ContainsKey(trigger.TriggerId) Then
                ScheduledTriggers.Add(trigger.TriggerId, CType(trigger, Cacadore.Trigger))
            Else
                ScheduledTriggers(trigger.TriggerId) = CType(trigger, Cacadore.Trigger)
            End If
        End SyncLock
        OnTriggerScheduled(trigger)
        trigger = Nothing
    End Sub

    Public Shared Sub SaveWatchedTasks(Optional cancelIfRunning As Boolean = False)
        If IsWatching Then
            If WatchedTasks.Values.Count > 0 Then
                For i = 0 To WatchedTasks.Values.Count - 1
                    Dim tsk As Task = WatchedTasks.Values(i)
                    If Not IsNothing(tsk) Then
                        'I may uncomment in the future.
                        'If tsk.State.TaskExecutionState = ITaskState.ExecutionState.Running Then
                        '    If cancelIfRunning Then tsk.Cancel()
                        'End If
                        tsk.Save()
                    End If
                Next
            End If
        End If
    End Sub

    Public Shared Function IsTaskWatched(taskId As String) As Boolean
        If WatchedTasks.ContainsKey(taskId) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Function GetWatchedTaskState(taskId As String) As ITaskState.ExecutionState
        If IsTaskWatched(taskId) Then
            Dim tskState As ITaskState.ExecutionState
            Dim tsk = GetWatchedTaskById(taskId)
            If Not IsNothing(tsk) Then
                If Not IsNothing(tsk.State) Then
                    tskState = GetWatchedTaskById(taskId).State.TaskExecutionState
                End If
            End If
            Return tskState
        End If
        Return ITaskState.ExecutionState.NotScheduled
    End Function

    Public Shared Function IsWatchedTaskScheduled(taskId As String) As Boolean
        If IsTaskWatched(taskId) Then
            Dim tskState = GetWatchedTaskState(taskId)
            If tskState = ITaskState.ExecutionState.Scheduled Then
                Return True
            End If
        End If
        Return False
    End Function

    Public Shared Function IsWatchedTaskEnabled(taskId As String) As Boolean
        If IsTaskWatched(taskId) Then
            Dim tskState = GetWatchedTaskState(taskId)
            If tskState = ITaskState.ExecutionState.Scheduled Then
                Return True
            End If
        End If
        Return False
    End Function

    Public Shared Function IsWatchedTaskPaused(taskId As String) As Boolean
        If IsTaskWatched(taskId) Then
            Dim tskState = GetWatchedTaskState(taskId)
            If tskState = ITaskState.ExecutionState.Paused Then
                Return True
            End If
        End If
        Return False
    End Function

    Public Shared Function IsWatchedTaskRunning(taskId As String) As Boolean
        If IsTaskWatched(taskId) Then
            Dim tskState = GetWatchedTaskState(taskId)
            If tskState = ITaskState.ExecutionState.Running Then
                Return True
            End If
        End If
        Return False
    End Function

    Public Shared Function GetWatchedTaskById(taskId As String) As Task
        If WatchedTasks.ContainsKey(taskId) Then
            Return WatchedTasks(taskId)
        Else
            Return Nothing
        End If
    End Function

    Public Shared Sub AddCancelledTask(taskId As String)
        If Not ToBeCancelledTasks.Contains(taskId) Then
            ToBeCancelledTasks.Add(taskId)
        End If
    End Sub

    Public Shared Function IsTaskCancelled(taskId As String) As Boolean
        If ToBeCancelledTasks.Contains(taskId) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Sub RemoveCancelledTask(taskId As String)
        If ToBeCancelledTasks.Contains(taskId) Then
            ToBeCancelledTasks.Remove(taskId)
        End If
    End Sub

    Public Shared Sub PauseAllWatchedTask()
        For Each tsk In WatchedTasks.Values
            tsk.Pause()
        Next
    End Sub

    Public Shared Sub StopAllWatchedTask()
        For Each tsk In WatchedTasks.Values
            If tsk.State.TaskExecutionState = ITaskState.ExecutionState.Running Then
                tsk.Cancel()
            End If
        Next
    End Sub

    Public Shared Sub ResumeAllWatchedTask()
        For Each tsk In WatchedTasks.Values
            tsk.Resume()
        Next
    End Sub

    Private Sub updateTasks(stateInfo As Object)
        SaveWatchedTasks()
        DeleteEndedTriggers()
    End Sub

    Shared Sub RemoveTriggerFromWatchedScheduledTriggers(trgId As String)
        If Watcher.ScheduledTriggers.ContainsKey(trgId) Then
            Watcher.ScheduledTriggers.Remove(trgId)
        End If
    End Sub

    Sub DeleteEndedTriggers()
        SyncLock Watcher.ScheduledTriggers
            Dim toBeClearedFromList As New List(Of String)
            'Pause all triggers attached to this task.
            For Each trg In ScheduledTriggers
                Dim willFireAgain = trg.Value.MayFireAgain
                If Not willFireAgain Then
                    trg.Value.Task.UnAttachTrigger(trg.Value)
                    toBeClearedFromList.Add(trg.Key)
                End If
            Next
            For Each id In toBeClearedFromList
                RemoveTriggerFromWatchedScheduledTriggers(id)
            Next
            RaiseEvent Trigger_Deleted()
        End SyncLock
    End Sub

    Private Shared Function getFullStateDescription(taskName As String,
                                                    stateDescription As String,
                                                    Optional nextFireTime As String = vbNullString) As String
        If stateDescription = ITaskState.ExecutionState.Running.ToString Then
            stateDescription = String.Format("Task ({0}) is running.", taskName)
        ElseIf stateDescription = ITaskState.ExecutionState.Failed.ToString Then
            stateDescription = String.Format("Task ({0}) failed.", taskName)
        ElseIf stateDescription = ITaskState.ExecutionState.Done.ToString Then
            stateDescription = String.Format("Task ({0}) has been executed successfully.", taskName)
        ElseIf stateDescription = ITaskState.ExecutionState.NotScheduled.ToString Then
            stateDescription = String.Format("Task ({0}) is not scheduled to run anymore.", taskName)
        ElseIf stateDescription = ITaskState.ExecutionState.Unfulfilled.ToString Then
            stateDescription = String.Format("Task ({0}) is unfulfilled.", taskName)
        ElseIf stateDescription = ITaskState.ExecutionState.Paused.ToString Then
            stateDescription = String.Format("Task ({0}) paused.", taskName)
        ElseIf stateDescription = ITaskState.ExecutionState.DoneWithSomeErrors.ToString Then
            stateDescription = String.Format("Task ({0}) done with some errors.", taskName)
        ElseIf stateDescription = ITaskState.ExecutionState.Resumed.ToString Then
            stateDescription = String.Format("Task ({0}) resumed.", taskName)
        ElseIf stateDescription = ITaskState.ExecutionState.CacadoreDisabled.ToString Then
            stateDescription = String.Format("Cacadu is disabled right now.", taskName)
        End If
        Return stateDescription
    End Function

    Private Shared Sub addNewTask(ByVal taskToBeWatched As Task)
        _watchedTasksValue.Add(taskToBeWatched.TaskId, taskToBeWatched)
        _watchedTasksStatesValue.Add(taskToBeWatched.TaskId, taskToBeWatched.State.TaskExecutionState)
    End Sub

    Private Shared Sub udpateExistingTask(ByVal taskToBeWatched As Task)
        _watchedTasksValue(taskToBeWatched.TaskId) = taskToBeWatched
        _watchedTasksStatesValue(taskToBeWatched.TaskId) = taskToBeWatched.State.TaskExecutionState
    End Sub

    Private Shared Sub removeTask(ByVal taskToBeWatched As Task)
        _watchedTasksValue.Remove(taskToBeWatched.TaskId)
        _watchedTasksStatesValue.Remove(taskToBeWatched.TaskId)
        taskToBeWatched.Dispose()
        taskToBeWatched = Nothing
    End Sub

    Private Shared Function createHistoryEntry(taskId As String,
                                               entryType As IHistoryEntry.HistoryEntryType,
                                               description As String,
                                               entryDateTime As DateTime,
                                               status As ITaskState.ExecutionState) As HistoryEntry
        Dim historyEntry As New HistoryEntry
        historyEntry.TaskId = taskId
        historyEntry.EntryDateTime = entryDateTime
        historyEntry.EntryType = entryType
        historyEntry.Description = description
        historyEntry.Status = status
        Return historyEntry
    End Function

    Private Shared Sub checkIfBusy()
        If Not _isBusy Then
            SyncLock _watchedTasksStatesValue
                For Each taskState In _watchedTasksStatesValue
                    If taskState.Value = ITaskState.ExecutionState.Running Then
                        _isBusy = True
                    End If
                Next
                _isBusy = False
            End SyncLock
        End If
    End Sub

    Public Shared Sub OnTriggerScheduled(trigger As ITrigger)
        RaiseEvent Trigger_Scheduled(trigger)
    End Sub

    Private Shared Sub New_Trigger_Scheduled(trigger As ITrigger) Handles Me.Trigger_Scheduled
        Debug.Print(trigger.TriggerId)
    End Sub
#End Region

#Region "Events"
    Public Shared Event Trigger_Deleted()
    Public Shared Event Trigger_Scheduled(trigger As ITrigger)
    'NotScheduled 'Task is added to Cacadore but not scheduled or attached to a trigger.
    'Scheduled 'Will be running in the future @ a specified trigger.
    'Running  'Hoppa, is running right now.
    'Unfulfilled 'Oops, actually it started but interpreted.
    'Failed 'Started but ends by failing.
    'Paused 'Task is paused by user.
    Public Shared Event TaskState_Changed(taskName As String, historyEntry As HistoryEntry)
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                endedTriggersCheckerXTimer.Dispose()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
            endedTriggersCheckerXTimer = Nothing
            _watchedTasksValue = Nothing
            _watchedTasksStatesValue = Nothing
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