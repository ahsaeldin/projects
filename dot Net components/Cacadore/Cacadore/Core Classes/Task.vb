'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   Task Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
Imports System.ComponentModel

<Serializable()> Public Class Task
    Inherits CacadoreBase
    Implements ITask
    Implements ICrudable
    Implements Balora.IListener
    Implements IDisposable

#Region "Fields"
#Region "Serializable Fields"
    Private _taskIdValue As String
    Private _triggersCount As Integer = 0
    Private _actionCount As Integer = 0
    Private _taskListeners As ArrayList
    Private _parentGroupValue As String
    Private _taskStateValue As TaskState
    Private _variablesValue As TaskVariables
    Private _propertiesValue As TaskProperties
    Private _waitBetweenActionsValue As Boolean
    Private _triggersValue As TriggersCollection
    Private _actionsCollectionValue As ActionsCollection
#End Region
#End Region

#Region "Non Serializable Fields"
    <NonSerialized()> Private _executionGate As New Threading.ManualResetEvent(False)
#End Region

#Region "Execution Background workder delegates"
    Private Delegate Sub _doWorkDlg(sender As Object, e As System.ComponentModel.DoWorkEventArgs)
    Private Delegate Sub _progressChangedDlg(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs)
    Private Delegate Sub _runWorkerCompletedDlg(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs)
#End Region

#Region "Constructors"
    ''' <summary>
    ''' Initializes a new instance of the <see cref="Task" /> class.
    ''' </summary>
    ''' <param name="taskId">The task id.</param>
    ''' <param name="store">
    ''' if set to <c>false</c> task will not be sent to Tectonic
    ''' from the constructor (i.e: SetTaskProperties and ChangeExecutionState won't get called from constructor)
    ''' </param>
    ''' <remarks>
    ''' *.You never can add a task without setting the its id
    ''' because we will generate actions and triggers ids based
    ''' on that id.
    ''' *.store parameter are useful when you want to create a new task that
    ''' will be filled by a query of Tectonic whereas so you don't need ever
    ''' to send the just created task to Tectonic
    ''' </remarks>
    Public Sub New(ByVal taskId As String, Optional store As Boolean = True)
        initClassFields(taskId)
        UpdateProperties(New Nullable(Of DateTime),
                         returnNotScheduledTaskState.ToString,
                         New Nullable(Of DateTime))
        initTaskState(taskId)
        initTaskVariables(taskId)
        AddHandler Settings.CacadoreEnabledChanged, AddressOf Settings_CacadoreEnabledChanged
    End Sub
#End Region

#Region "Properties"
    'To decrease trips to Tectonic.
    <System.ComponentModel.ReadOnly(True)>
    <System.ComponentModel.DisplayName("Actions Count")>
    <System.ComponentModel.Category("Task Properties")>
    Public Property ActionsCount() As Integer Implements ITask.ActionsCount
        Get
            If _actionCount <= 0 Then _actionCount = Cacadore.Commands.GetTaskActionsCount(TaskId)
            Return _actionCount
        End Get
        Set(ByVal value As Integer)
            _actionCount = value
        End Set
    End Property

    <System.ComponentModel.ReadOnly(True)>
    <System.ComponentModel.DisplayName("Triggers Count")>
    <System.ComponentModel.Category("Task Properties")>
    Public Property TriggersCount As Integer Implements ITask.TriggersCount
        Get
            _triggersCount = Commands.GetTaskTriggersCount(TaskId)
            Return _triggersCount
        End Get
        Set(ByVal value As Integer)
            _triggersCount = value
        End Set
    End Property

    <System.ComponentModel.Browsable(False)>
    Public Property Errors As New TaskErrors Implements ITask.Errors

    <System.ComponentModel.Browsable(False)>
    Public ReadOnly Property History As List(Of HistoryEntry) Implements ITask.History
        Get
            Return Cacadore.History.GetTaskHistory(Me.TaskId)
        End Get
    End Property

    <System.ComponentModel.Category("Task Properties"), TypeConverter(GetType(GroupsNames))>
    <System.ComponentModel.DisplayName("Parent Group")>
    <System.ComponentModel.Description("The group name in which the task belongs to.")>
    Public Property ParentGroup() As String Implements ITask.ParentGroupName
        Get
            Return Commands.GetGroupName(_parentGroupValue)
        End Get
        Set(value As String)
            ParentGroupId = Commands.GetGroupIdByName(value)
        End Set
    End Property

    <System.ComponentModel.Browsable(False)>
    Public Property ParentGroupId As String Implements ITask.ParentGroupId
        Get
            Return _parentGroupValue
        End Get
        Set(ByVal value As String)
            _parentGroupValue = value
        End Set
    End Property

    <System.ComponentModel.Category("Task Properties")>
    <System.ComponentModel.Description("Enables you to modify selected task properties.")>
    Public Property Properties As TaskProperties Implements ITask.Properties
        Get
            Return _propertiesValue
        End Get
        Set(ByVal value As TaskProperties)
            If Me.TaskId <> value.taskId Then
                _propertiesValue = New TaskProperties(Me.TaskId)
            Else
                _propertiesValue = value
            End If
        End Set
    End Property

    <System.ComponentModel.Browsable(False)>
    Public ReadOnly Property TaskId As String Implements ITask.TaskId
        Get
            Return _taskIdValue
        End Get
    End Property

    '<System.ComponentModel.Category("Behavior")>
    '<System.ComponentModel.DisplayName("Task Actions")>
    '<System.ComponentModel.Description("The attached actions to the task.")>
    <System.ComponentModel.Browsable(False)>
    Public ReadOnly Property MyActions As ActionsCollection Implements ITask.ActionsCollection
        Get
            Return _actionsCollectionValue
        End Get
    End Property

    '<System.ComponentModel.Category("Behavior")>
    '<System.ComponentModel.DisplayName("Task Triggers")>
    '<System.ComponentModel.Description("Task's triggers to be fired at.")>
    <System.ComponentModel.Browsable(False)>
    Public ReadOnly Property MyTriggers As TriggersCollection Implements ITask.Triggers
        Get
            Return _triggersValue
        End Get
    End Property

    'This is very important
    'Direct assignment of Task.TaskVariables property to any
    'other TaskVariable class is acceptable only if both has
    'the same id, otherwise it won't set. Check TaskVariables 
    'Property in Task Class.
    <System.ComponentModel.Category("Task Properties")>
    <System.ComponentModel.DisplayName("Task Variables")>
    <System.ComponentModel.Description("View defined variables for the task.")>
    Public Property Variables As TaskVariables Implements ITask.Variables
        Get
            Return _variablesValue  '+ Use LINQ or anything to look up tectonic to fill this property.
        End Get
        Set(ByVal value As TaskVariables)
            'Direct assignment of Task.TaskVariables property to any
            'other TaskVariable class is acceptable only if both has
            'the same id.
            If value.TaskId = Me.TaskId Then
                _variablesValue = value
            End If
        End Set
    End Property

    'Terminate execution and report failed if one of task's actions failed. default if false
    <System.ComponentModel.Category("Task Properties")>
    <System.ComponentModel.DisplayName("Break On Failed")>
    <System.ComponentModel.Description("Break task's execution, if any of its actions failed.")>
    Public Property BreakOnFailed() As Boolean Implements ITask.BreakOnFailed

    <System.ComponentModel.Category("Task Properties")>
    <System.ComponentModel.DisplayName("Wait Between Actions")>
    <System.ComponentModel.Description("If true, task will wait until each action completes its execution.")>
    Public Property WaitBetweenActions As Boolean Implements ITask.WaitBetweenActions
        Get
            Return _waitBetweenActionsValue
        End Get
        Set(ByVal value As Boolean)
            _waitBetweenActionsValue = value
        End Set
    End Property

    <System.ComponentModel.Browsable(False)>
    Public ReadOnly Property State As TaskState Implements ITask.State
        Get
            Return _taskStateValue
        End Get
    End Property

    <System.ComponentModel.Category("Task Properties")>
    <System.ComponentModel.DisplayName("May Fired Again")>
    <System.ComponentModel.Description("Indicates if the task may fire again, depending on task's triggers settings, false if no trigger attached.")>
    Public ReadOnly Property MayFiredAgain As Boolean Implements ITask.MayFiredAgain
        Get
            For Each trg In MyTriggers.Triggers
                If trg.Value.MayFireAgain Then
                    Return True
                End If
            Next
            Return False
        End Get
    End Property
#End Region

#Region "Methods"
#Region "Public Methods"
    'Returns the attached ActionId or false if failed
    Public Function AttachAction(ByVal actionObject As Action) As String Implements ITask.AttachAction
        Try
            Dim newActionId As String = getNewActionId()
            setActionId(actionObject, newActionId)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            _privateRaiseErrorUp("", ex, True)
            Return ""
        End Try
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Dim result As Boolean = CBool(InvokeTectonicDlg(actionObject, {}, InsertOperationDlg, My.Resources.ActNotPresisted))
            Return actionObject.ActionId
        End If
    End Function

    Public Function AttachTrigger(ByVal trg As Trigger) As String Implements ITask.AttachTrigger
        TriggersCount = TriggersCount + 1
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            trg.Save(Me)
            Return trg.TriggerId
        End If
    End Function

    Public Function UpdateAction(ByVal [action] As Action) As Boolean Implements ITask.UpdateAction
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return CBool(InvokeTectonicDlg([action], {}, UpdateOperationDlg, My.Resources.ActNotPresisted))
        End If
    End Function

    Public Function UnAttachAction(ByVal actionObj As Action) As Boolean Implements ITask.UnAttachAction
        If ActionsCount > 0 Then ActionsCount = ActionsCount - 1
        Return deleteOneOfMyActions(actionObj)
    End Function

    Public Function UnAttachTrigger(ByVal trgId As String) As Boolean Implements ITask.UnAttachTrigger
        Dim trg As ITrigger = MyTriggers.Triggers(trgId)
        If Not IsNothing(trg) Then Return UnAttachTrigger(trg)
        Return False
    End Function

    ''' <summary>
    ''' Delete the trigger from the store, and change task state if required.
    ''' </summary>
    Public Function UnAttachTrigger(ByVal trg As ITrigger) As Boolean Implements ITask.UnAttachTrigger
        If TriggersCount > 0 Then TriggersCount = TriggersCount - 1
        Dim result As Boolean = stopListeningToTriggers(trg)
        unScheduleMeIfNoTriggerIsAttached()
        Return result
    End Function

    Public Function CheckIfNoActionsAttached() As ITaskState.ExecutionState
        If ActionsCount <= 0 Then
            logState(returnTaskHasNoActions())
            Return returnTaskHasNoActions()
        End If
        Return ITaskState.ExecutionState.NotScheduled
    End Function

    Public Function Execute() As ITaskState.ExecutionState Implements ITask.Execute
        If CheckIfNoActionsAttached() = ITaskState.ExecutionState.HasNoActions Then
            returnTaskHasNoActions()
        End If

        If Not Properties.IsEnabled Then
            logState(returnTaskDisabled())
            Return returnTaskDisabled()
        End If

        If Not isCacadoreEnabled() Then
            logState(returnCacadoreIsDisableState())
            Return returnCacadoreIsDisableState()
        End If

        If isTaskPaused() Then
            Return returnPausedTaskState()
        End If

        If IsNothing(_executionGate) Then
            _executionGate = New Threading.ManualResetEvent(False)
        End If

        doBeforeExecutionWork() 'Generates the great hang, but I needs it for Pause/Resume functionally.

        Dim results = startExecuteLoop()

        doAfterExecutionWork(results)

        Return returnRunningTaskState()
    End Function


    ''' <summary>
    ''' Pauses task execution if it is running.
    ''' </summary>
    ''' <remarks>
    ''' Pause has no effect if task isn't running.
    ''' </remarks>
    Public Sub Pause() Implements ITask.Pause
        Watcher.WatchedTasks(TaskId).State.ChangeExecutionState(ITaskState.ExecutionState.Paused)
        changeMyState(returnPausedTaskState())
        logState(returnPausedTaskState())
        raiseTaskEvents(returnPausedTaskState())
    End Sub

    Public Sub [Resume]() Implements ITask.Resume
        If isTaskPaused() Then
            onResume()
        Else
            'Balora.Hodhod.STOW("Cannot resume; task is not paused.")
        End If
    End Sub

    Public Sub Cancel() Implements ITask.Cancel
        Try
            Dim currentTaskState = Me.State.GetExecutionState()
            If currentTaskState = returnRunningTaskState() Or isTaskPaused() Then
                _executionGate.Set()
                Watcher.AddCancelledTask(TaskId)
            Else
                'Balora.Hodhod.STOW("Task isn't running or paused to be canceled.")
            End If
        Catch ex As InvalidOperationException
            _privateRaiseErrorUp("Cancellation is not supported.", ex, False)
        End Try
    End Sub

    Public Sub SetTaskState(tskState As TaskState) Implements ITask.SetTaskState
        _taskStateValue = tskState
    End Sub

    Public Sub UpdateProperties(lastFireTime As Nullable(Of DateTime),
                                 lastRunResult As String,
                                 nextFireTime As Nullable(Of DateTime)) Implements ITask.SetTaskProperties
        _updateProperties(lastFireTime, lastRunResult, nextFireTime)
    End Sub

    'Until now, Save do both Update and Insert into Tectonic.
    Public Function Save() As Boolean Implements ITask.Save
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return CBool(InvokeTectonicDlg(Me, {}, InsertOperationDlg, My.Resources.TskNotPresisted))
        End If
    End Function

    Public Sub PauseAllTriggers() Implements ITask.PauseAllTriggers
        For Each trg In MyTriggers.Triggers.Values
            trg.PauseTrigger()
        Next
    End Sub

    Public Function Clone() As Task Implements ITask.Clone
        Dim actionMemStream As New Global.System.IO.MemoryStream
        Balora.Serializer.BinarySerializer.SerializeObject(Me, actionMemStream)
        Dim duplicatedTask As Task = CType(Balora.Serializer.BinarySerializer.DeserializeObject(actionMemStream.ToArray), Task)
        Return duplicatedTask
    End Function

    Public Function Delete() As Boolean Implements ITask.Delete
        ActionsCount = 0
        TriggersCount = 0
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Dim result As Boolean = CBool(InvokeTectonicDlg(Me, {}, DeleteOperationDlg, My.Resources.TskNotPresisted))
            removeFromWatcher()
            Return result
        End If
    End Function
#End Region

#Region "Private Methods"
#Region "Private calls"
    Private Sub _updateProperties(ByVal lastFireTime As Nullable(Of DateTime),
                                   ByVal lastRunResult As String,
                                   ByVal nextFireTime As Nullable(Of DateTime))
        'For every newly created task (@ constructor), task properties set by SetTaskProperties method.
        'For every newly created QuartzTrigger(@ QuartzTrigger_Scheduled), task properties set by SetTaskProperties method.
        'For every QuartzTrigger(@ QuartzTrigger_Executed), task properties set by SetTaskProperties method.
        If Not IsNothing(lastFireTime) Then _propertiesValue.LastTimeRun = CDate(lastFireTime).ToLocalTime
        _propertiesValue.SetLastTimeRunResult(lastRunResult)
        If Not IsNothing(nextFireTime) Then _propertiesValue.NextTimeToBeRun = CDate(nextFireTime).ToLocalTime
    End Sub
#End Region
#Region "Initlization Methods"
    Private Sub initTaskState(ByVal taskId As String)
        'Just set TaskState to taskStateValue class
        Me._taskStateValue = New TaskState(taskId)
        Me.State.MyTask = Me
        Me.State.NextTriggerId = ""
    End Sub

    Private Sub initTaskVariables(ByVal taskId As String)
        'Attach Variable to this task.
        _variablesValue = New TaskVariables(taskId)
    End Sub

    Private Sub initClassFields(ByRef taskId As String)
        'You can pass any unique id here,
        'but it will check first if it is inserted
        'before in the database
        _taskIdValue = taskId
        _taskListeners = New ArrayList
        '0 == root group
        _parentGroupValue = CStr(0)
        _triggersValue = New TriggersCollection
        _actionsCollectionValue = New ActionsCollection
        _triggersValue.TaskId = Me.TaskId
        _actionsCollectionValue.TaskId = Me.TaskId
        'Attach properties to this class
        _propertiesValue = New TaskProperties(taskId)
        _actionCount = Commands.GetTaskActionsCount(taskId)
    End Sub
#End Region

#Region "do On methods"
    Private Sub doAfterExecution(ByVal actionsExecutionResults As List(Of ITaskState.ExecutionState),
                                           ByVal _action As Action,
                                           ByVal result As ITaskState.ExecutionState)
        'Execute action then EnList the result.
        actionsExecutionResults.Add(result)
        'Save this action as the current executing action in tectonic.
        'Note:- saving an action to "currentExecutingActionIdValue" means
        'it has been executed and has a result.
        Me.State.CurrentExecutingActionId = _action.ActionId
    End Sub

    Private Sub onResume()
        Watcher.WatchedTasks(TaskId).State.ChangeExecutionState(ITaskState.ExecutionState.Running)
        changeMyState(ITaskState.ExecutionState.Running)
        'Dodge the lock in backGroudWorkerThread_DoWork
        _executionGate.Set()
        raiseTaskEvents(returnResumedTaskState())
        logState(returnResumedTaskState())
    End Sub

    Private Sub checkIfPaused()
        If isTaskPaused() Then
            ''AhSaElDin 20120103: الترتيب هنا مهم.
            logState(returnPausedTaskState())
            If Not isTaskPaused() Then changeMyState(returnPausedTaskState())
            raiseTaskEvents(returnPausedTaskState())
            _executionGate.Reset()
            _executionGate.WaitOne()
            'Resume after waiting...
            changeMyState(returnResumedTaskState)
        End If
    End Sub

    Private Function doOnActionFail(actionNo As Integer, actionResult As ITaskState.ExecutionState) As Boolean
        'If action report fail and "FailedOnError" is set to true...
        'then the whole task will be reported as failed.
        If actionResult = returnFailedTaskState() Then
            If BreakOnFailed Then
                Cacadore.Watcher.AddNewEntry(TaskId,
                                             Me.Properties.TaskName,
                                             IHistoryEntry.HistoryEntryType.Action,
                                             String.Format("Action #{0} failed, task cancelled because (Break On Fail) set to true.", actionNo),
                                             DateTime.Now,
                                             returnFailedTaskState())
            Else
                Cacadore.Watcher.AddNewEntry(TaskId,
                                             Me.Properties.TaskName,
                                             IHistoryEntry.HistoryEntryType.Action,
                                             String.Format("Action #{0} failed.", actionNo),
                                             DateTime.Now,
                                             returnFailedTaskState())
            End If
            Return True
        End If
        Return False
    End Function

    Private Sub doIfCancelled(ByRef _executionEventArgs As ITask.ExecutionEventArgs)
        'Balora.Hodhod.STOW("Task has been canceled.")
        _executionEventArgs.state = returnCanceledTaskState()
        changeMyState(returnCanceledTaskState())
        raiseTaskEvents(returnCanceledTaskState())
        logState(returnCanceledTaskState())
        'Balora.Alerter.SendMessageToCacadu("TaskCancelled", Me.Properties.TaskName)
    End Sub

    Private Sub doIfCompeleted(actionsExecutionResults As List(Of ITaskState.ExecutionState), ByVal _executionEventArgs As ITask.ExecutionEventArgs)
        Dim isAllDone As IEnumerable(Of ITaskState.ExecutionState) = getAllDoneActions(actionsExecutionResults)
        Dim isAllFailed As IEnumerable(Of ITaskState.ExecutionState) = getIsAllFailedActions(actionsExecutionResults)
        'In order to pass actions result to Executed event.
        _executionEventArgs.result = actionsExecutionResults

        setTaskStateAfterCompliation(_executionEventArgs, actionsExecutionResults, isAllDone, isAllFailed)
        raiseTaskEvents(returnDoneTaskState(), _executionEventArgs) 'Memory Leak 'koko  
        logState(returnDoneTaskState())
        actionsExecutionResults = Nothing
        isAllDone = Nothing
        isAllFailed = Nothing
    End Sub
#End Region

#Region "Return state methods"
    Private Function returnFailedTaskState() As ITaskState.ExecutionState
        Return ITaskState.ExecutionState.Failed
    End Function

    Private Function returnPausedTaskState() As ITaskState.ExecutionState
        Return ITaskState.ExecutionState.Paused
    End Function

    Private Function returnDoneTaskState() As ITaskState.ExecutionState
        Return ITaskState.ExecutionState.Done
    End Function

    Private Function returnDoneWithSomeErrorsTaskState() As ITaskState.ExecutionState
        Return ITaskState.ExecutionState.DoneWithSomeErrors
    End Function

    Private Function returnCanceledTaskState() As ITaskState.ExecutionState
        Return ITaskState.ExecutionState.Canceled
    End Function

    Private Function returnResumedTaskState() As ITaskState.ExecutionState
        Return ITaskState.ExecutionState.Resumed
    End Function

    Private Function returnBeforeRunningTaskState() As ITaskState.ExecutionState
        Return ITaskState.ExecutionState.BeforeRunning
    End Function

    Private Function returnCacadoreIsDisableState() As ITaskState.ExecutionState
        Return ITaskState.ExecutionState.CacadoreDisabled
    End Function

    Private Function returnTaskHasNoActions() As ITaskState.ExecutionState
        Return ITaskState.ExecutionState.HasNoActions
    End Function

    Private Function returnTaskDisabled() As ITaskState.ExecutionState
        Return ITaskState.ExecutionState.Disabled
    End Function

    Private Function returnNotScheduledTaskState() As ITaskState.ExecutionState
        Return ITaskState.ExecutionState.NotScheduled
    End Function

    Private Function returnFiredTaskState() As ITaskState.ExecutionState
        Return ITaskState.ExecutionState.Fired
    End Function

    Private Function returnRunningTaskState() As ITaskState.ExecutionState
        Return ITaskState.ExecutionState.Running
    End Function

    Private Function returnScheduledTaskState() As ITaskState.ExecutionState
        Return ITaskState.ExecutionState.Scheduled
    End Function
#End Region

#Region "Execution Methods"
    Private Function executeAction(ByVal _action As Action) As ITaskState.ExecutionState
        'Now Execute -- Fire...
        Dim result As ITaskState.ExecutionState = _action.Execute()
        Return result
    End Function

    Private Sub setTaskStateAfterCompliation(ByRef _executionEventArgs As ITask.ExecutionEventArgs,
                                             ByVal actionsExecutionResults As List(Of ITaskState.ExecutionState),
                                             ByVal isAllDone As IEnumerable(Of ITaskState.ExecutionState),
                                             ByVal isAllFailed As IEnumerable(Of ITaskState.ExecutionState))
        If isAllDone.Count = actionsExecutionResults.Count Then 'If all done...
            _executionEventArgs.state = returnDoneTaskState()
        ElseIf isAllFailed.Count = actionsExecutionResults.Count Then 'If all failed...
            _executionEventArgs.state = returnFailedTaskState()
            changeMyState(returnFailedTaskState())
        Else 'mixed "Done, with some errors"...
            _executionEventArgs.state = returnDoneWithSomeErrorsTaskState()
            changeMyState(returnDoneWithSomeErrorsTaskState())
        End If
    End Sub

    Function startExecuteLoop() As List(Of ITaskState.ExecutionState)
        Dim numberofProcessedActions As Integer = 0
        Dim actionsExecutionResults As New List(Of ITaskState.ExecutionState)
        For Each kvp As KeyValuePair(Of String, Action) In MyActions.Actions
            If kvp.Value.Enabled Then
                If Watcher.IsTaskCancelled(TaskId) Then
                    Watcher.RemoveCancelledTask(TaskId)
                    Return actionsExecutionResults
                End If
                checkIfPaused()
                setActionParentTask(kvp.Value)
                Dim result As ITaskState.ExecutionState = executeAction(kvp.Value)
                doAfterExecution(actionsExecutionResults, kvp.Value, result)
                If doOnActionFail(kvp.Value.GetActionNumber(), result) And BreakOnFailed Then Exit For
            End If

            Dim progressPercentage As Integer = 0 'تحسبًا اني اعملها في يوم من الايام.
            reportProgress(progressPercentage)
        Next
        'Used later in complete event
        Return actionsExecutionResults
    End Function

    Sub reportProgress(progressPercentage As Integer)
        raiseTaskEvents(returnRunningTaskState, New EventArgs)
        logState(returnRunningTaskState)
        'Balora.Alerter.SendMessageToCacadu("reporttaskprogress",
        '                                   {Me.Properties.TaskName,
        '                                   progressPercentage.ToString})
    End Sub

    Sub doAfterExecutionWork(res As List(Of ITaskState.ExecutionState))
        setCurrentExecutingActionIdToNull()
        'Will be passed to Esxecuted event.
        Dim _executionEventArgs As New ITask.ExecutionEventArgs
        If Watcher.IsTaskCancelled(TaskId) Then
            doIfCancelled(_executionEventArgs) 'Memory Leak 'koko 
        Else
            doIfCompeleted(res, _executionEventArgs) 'Memory Leak 'koko 
        End If
        changeStateByCheckingTriggersCount() 'Memory Leak 'koko 
    End Sub

    Private Function isTaskPaused() As Boolean
        If Me.State.GetExecutionState() = returnPausedTaskState() Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function isCacadoreEnabled() As Boolean
        If Not Cacadore.Settings.Enabled Then
            Return False
        Else
            Return True
        End If
    End Function

    Sub doBeforeExecutionWork()
        raiseTaskEvents(returnBeforeRunningTaskState())
        logState(returnBeforeRunningTaskState())
        changeMyState(returnRunningTaskState())
    End Sub
#End Region

#Region "Tectonicing Methods"
    Private Sub saveMe()
        Me.Save()
    End Sub
#End Region

#Region "Helper Methods"
    Private Sub setActionParentTask(ByRef _action As Action)
        'AhSaElDin 20120126: Action will use its own task to
        'get some property like WaitBetweenAction in order to
        'check for delaying in the overridable WaitForExit method.
        If IsNothing(_action.ParentTask) Then _action.ParentTask = Me
    End Sub

    Private Function getAllDoneActions(ByVal actionsExecutionResults As List(Of ITaskState.ExecutionState)) As IEnumerable(Of ITaskState.ExecutionState)
        'Query if all actions are done.
        Dim isAllDone = From result In actionsExecutionResults Where result = returnDoneTaskState()
        Return isAllDone
    End Function

    Private Function getIsAllFailedActions(ByVal actionsExecutionResults As List(Of ITaskState.ExecutionState)) As IEnumerable(Of ITaskState.ExecutionState)
        'Query if all actions are failed.
        Dim isAllFailed = From result In actionsExecutionResults Where result = returnFailedTaskState()
        Return isAllFailed
    End Function

    Private Sub setCurrentExecutingActionIdToNull()
        'First of all set this value to nothing and save in database
        Me.State.CurrentExecutingActionId = vbNullString
    End Sub

    Private Sub raiseTaskEvents(_state As ITaskState.ExecutionState, Optional evtArgs As Object = Nothing)
        If IsNothing(evtArgs) Then evtArgs = New EventArgs()
        Select Case _state
            Case returnDoneTaskState()
                RaiseEvent Executed(Me, CType(evtArgs, ITask.ExecutionEventArgs))
            Case returnRunningTaskState()
                RaiseEvent Executing(Me, CType(evtArgs, EventArgs))
            Case returnResumedTaskState()
                RaiseEvent Resumed(Me, CType(evtArgs, EventArgs))
            Case returnBeforeRunningTaskState()
                RaiseEvent Before_Executing(Me, CType(evtArgs, EventArgs))
            Case returnPausedTaskState()
                RaiseEvent Paused(Me, CType(evtArgs, EventArgs))
            Case returnCanceledTaskState()
                RaiseEvent Canceled(Me, CType(evtArgs, EventArgs))
            Case returnFiredTaskState()
                RaiseEvent Fired(Me, CType(evtArgs, ITask.ExecutionEventArgs))
            Case returnScheduledTaskState()
                RaiseEvent Scheduled(Me, CType(evtArgs, ITask.ScheduleEventArgs))
        End Select
    End Sub

    Private Function getNewActionId(Optional pos As Integer = 0) As String
        Dim lastActionId As String
        Dim actionNo As Integer
        Dim tskActionsCount As Integer = Commands.GetTaskActionsCount(TaskId)
        If pos > 0 Then
            ActionsCount = pos + 1
        ElseIf tskActionsCount > 0 Then
            lastActionId = MyActions.Actions.Last.Key
            actionNo = Commands.ExtractActionNumber(lastActionId)
            ActionsCount = actionNo + 1
        Else
            ActionsCount = 1
        End If
        'Increment the old Id to create the new one.
        Dim newActionId As String = String.Format("{0}-a{1}", TaskId, ActionsCount)
        If Commands.IsRegisteredAction(newActionId) Then Return getNewActionId(ActionsCount)
        Return newActionId
    End Function

    Private Sub removeFromWatcher()
        Watcher.UnWatch(Me)
    End Sub

    Private Sub setActionId(ByRef actionObject As Action, ByVal newActionId As String)
        'Set the action id (attach action to the task)
        actionObject.SetActionId(newActionId)
        actionObject.SetParentTaskId(Me.TaskId)
    End Sub

    Private Sub unScheduleMeIfNoTriggerIsAttached()
        If TriggersCount = 0 Then
            changeMyState(returnNotScheduledTaskState())
        End If
    End Sub

    Private Function deleteOneOfMyActions(ByVal _action As Action) As Boolean
        ActionsCount = ActionsCount - 1
        Return _action.Delete
    End Function

    Private Function stopListeningToTriggers(ByVal trg As ITrigger) As Boolean
        Dim result As Boolean = trg.StopListening()
        Return result
    End Function

    Private Sub logState(notificationEvent As ITaskState.ExecutionState)
        Dim message As String = vbNullString
        Dim tskName As String = Properties.TaskName
        Select Case notificationEvent
            Case Cacadore.ITaskState.ExecutionState.BeforeRunning
                message = "Task is running"
            Case Cacadore.ITaskState.ExecutionState.Canceled
                message = "Task cancelled."
            Case Cacadore.ITaskState.ExecutionState.Done
                message = "Task done."
            Case Cacadore.ITaskState.ExecutionState.Running
                message = "" 'we move this message to "Before Running" and prevent this message even from getting displayed,
                'because with every action executed it will be displayed and we need only one indication about task is running.
            Case Cacadore.ITaskState.ExecutionState.Paused
                message = "Task paused."
            Case Cacadore.ITaskState.ExecutionState.Resumed
                message = "Task resumed."
            Case Cacadore.ITaskState.ExecutionState.Scheduled
                message = "Task scheduled successfully."
            Case Cacadore.ITaskState.ExecutionState.Disabled
                message = "Task is disabled, please enable it in order to execute."
            Case ITaskState.ExecutionState.CacadoreDisabled
                message = "Cacadu is disabled, please enable it in order to execute your tasks."
            Case ITaskState.ExecutionState.HasNoActions
                message = "Task has no actions."
        End Select
        If notificationEvent <> ITaskState.ExecutionState.Scheduled AndAlso notificationEvent <> ITaskState.ExecutionState.Running Then
            'what will happen if I could not delete this? TriggerFired وهل ليها لزمة وهي اصلا موجودة في
            Watcher.AddNewEntry(Me.TaskId,
                                Me.Properties.TaskName,
                                IHistoryEntry.HistoryEntryType.Trigger,
                                message,
                                DateTime.Now,
                                notificationEvent)
        End If
    End Sub

    Private Sub raiseFiredOrScheduledEvent(ByVal message As ArrayList, ByVal _state As ITaskState.ExecutionState)
        Dim trgId As String = CStr(message.Item(0))
        If Not Commands.IsValidTriggerId(trgId) Then Exit Sub
        Dim trg As Trigger = Me.MyTriggers.Triggers(trgId)
        If IsNothing(trg) Then Exit Sub
        Select Case _state
            Case returnFiredTaskState()
                raiseTaskEvents(ITaskState.ExecutionState.Fired, New ITask.ExecutionEventArgs With {.Trigger = trg})
            Case returnScheduledTaskState()
                raiseTaskEvents(returnScheduledTaskState, New ITask.ScheduleEventArgs With {.Trigger = trg})
        End Select
    End Sub

    Private Sub changeMyState(_state As ITaskState.ExecutionState)
        State.ChangeExecutionState(_state)
    End Sub

    Private Sub updateTaskState()
        setCurrentTriggerId()
        Watcher.Watch(Me)
    End Sub

    Private Sub setCurrentTriggerId()
        'If State.GetExecutionState() = ITaskState.ExecutionState.Running Then
        '    Me.Task.State.CurrentTriggerId = Me.InnerTrigger.Name
        'Else
        '    Me.Task.State.CurrentTriggerId = ""
        'End If
    End Sub

    Private Sub changeStateByCheckingTriggersCount()
        'If TriggersCount = 0 Then 
        '    changeMyState(returnNotScheduledTaskState())
        'Else
        '    changeMyState(returnScheduledTaskState)
        'End If
    End Sub

    'Protected Overrides Sub dispose(ByVal disposing As Boolean)
    '    diposeAllPrivateFields()
    'End Sub
#End Region
#End Region
#End Region

#Region "Implemented Interfaces"
#Region "IListener Interface"
    Public Sub Notify(message As Object) Implements Balora.IListener.Notify 'This method is a Dangerous Tectonic Trip.
        Dim msg As ArrayList = CType(message, ArrayList)
        If msg.Contains("Scheduled") Then
            raiseFiredOrScheduledEvent(msg, returnScheduledTaskState)
        ElseIf msg.Contains("Fired") Then
            raiseFiredOrScheduledEvent(msg, returnFiredTaskState)
        End If
        msg = Nothing
    End Sub
#End Region
#End Region

#Region "Events Handlers"
    Private Sub Settings_CacadoreEnabledChanged(enabled As Boolean)
        If enabled Then
            'Release gate, enable/resume execution...
            _executionGate.Set()
            changeMyState(ITaskState.ExecutionState.Paused)

            'Very important: "search and delete this comment if you intended to save it to database."
            'PreviousTaskExecutionState property is not added yet to the database and it will not be saved there in state table.
            If Me.State.TaskPreviousExecutionState = returnRunningTaskState() Then
                changeMyState(returnRunningTaskState)
            End If
        Else
            Pause()
        End If
    End Sub

    Sub Task_Paused(sender As Task, e As EventArgs) Handles Me.Paused
        Dim myFiredTrgs = getMyJustFiredTriggers()
        'Pause all triggers attached to this task.
        For Each trg In myFiredTrgs
            trg.PauseTrigger()
        Next
        updateTaskState()
        myFiredTrgs = Nothing
    End Sub

    Sub Task_Resumed(sender As Task, e As EventArgs) Handles Me.Resumed
        Dim myFiredTrgs = getMyJustFiredTriggers()
        'Reusme all triggers attached to this task.
        For Each trg In myFiredTrgs
            trg.ResumeTrigger()
        Next
        updateTaskState()
        myFiredTrgs = Nothing
    End Sub

    Sub Task_Canceled(sender As Task, e As EventArgs) Handles Me.Canceled
        updateTaskState()
    End Sub

    Sub Task_Fired(sender As Task, e As ITask.ExecutionEventArgs) Handles Me.Fired
        updateTaskState()
    End Sub

    Sub Before_TaskExecuting(sender As Task, e As EventArgs) Handles Me.Before_Executing
        updateTaskState()
    End Sub

#Region "Executed Event handler"
    Function getJustExecutedTriggers() As System.Collections.Generic.IEnumerable(Of Trigger)
        SyncLock Watcher.ScheduledTriggers
            'Get all just executed triggers, because all "False" triggers in FiredTriggers means it was just executed.
            Dim queryResult = From trg As Trigger
                              In Watcher.ScheduledTriggers.Values
                              Where trg.IsExecuting = False And trg.Task.TaskId = Me.TaskId
            Return queryResult
        End SyncLock
    End Function

    Function getMyJustFiredTriggers() As System.Collections.Generic.IEnumerable(Of Trigger)
        SyncLock Watcher.ScheduledTriggers
            'Get all just executed triggers, because all "False" triggers in FiredTriggers means it was just executed.
            Dim queryResult = From trg As Trigger
                              In Watcher.ScheduledTriggers.Values
                              Where trg.IsExecuting = True And trg.Task.TaskId = Me.TaskId
            Return queryResult
        End SyncLock
    End Function

    Sub sortDates(ByRef JustExecutedTriggers As System.Collections.Generic.IEnumerable(Of Trigger),
                  ByRef nextFireTimeUTCList As SortedSet(Of Date),
                  ByRef previousFireTimeUTCList As SortedSet(Of Date))
        'Get and sort dates
        For Each trg In JustExecutedTriggers
            If trg.NextFireTimeUTC.HasValue Then nextFireTimeUTCList.Add(trg.NextFireTimeUTC.Value)
            If trg.PreviousFireTimeUTC.HasValue Then previousFireTimeUTCList.Add(trg.PreviousFireTimeUTC.Value)
        Next
    End Sub

    Sub Task_Executed(sender As Task, e As ITask.ExecutionEventArgs) Handles Me.Executed
        updateTaskState()

        'Get all just executed triggers, because all "False" triggers in FiredTriggers means it was just executed.
        Dim queryResult = getJustExecutedTriggers()

        'Get and sort dates
        Dim nextFireTimeUTCList As New SortedSet(Of Date) : Dim previousFireTimeUTCList As New SortedSet(Of Date)
        sortDates(queryResult, nextFireTimeUTCList, previousFireTimeUTCList)
        
        SyncLock Watcher.WatchedTasksStates
            'Update Task properties
            If nextFireTimeUTCList.Count <> 0 AndAlso previousFireTimeUTCList.Count <> 0 Then UpdateProperties(previousFireTimeUTCList.Last,
                                                                                                               Watcher.WatchedTasksStates(TaskId).ToString,
                                                                                                               nextFireTimeUTCList.First)
        End SyncLock
        'Save task.
        Save()

        queryResult = Nothing
        previousFireTimeUTCList = Nothing
        nextFireTimeUTCList = Nothing
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                If Not IsNothing(Errors) Then Errors.Dispose()
                If Not IsNothing(_triggersValue) Then _triggersValue.Dispose()
                If Not IsNothing(_variablesValue) Then _variablesValue.Dispose()
                If Not IsNothing(_taskStateValue) Then _taskStateValue.Dispose()
                If Not IsNothing(_propertiesValue) Then _propertiesValue.Dispose()
                If Not IsNothing(_actionsCollectionValue) Then _actionsCollectionValue.Dispose()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
            Errors = Nothing
            _triggersValue = Nothing
            _variablesValue = Nothing
            _taskStateValue = Nothing
            _propertiesValue = Nothing
            _actionsCollectionValue = Nothing
            RemoveHandler Settings.CacadoreEnabledChanged, AddressOf Settings_CacadoreEnabledChanged
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region
#End Region
    Sub Task_Scheduled(sender As Task, e As ITask.ScheduleEventArgs) Handles Me.Scheduled

    End Sub

    Sub Task_Executing(sender As Task, e As EventArgs) Handles Me.Executing

    End Sub
#End Region
#Region "Events Declarations"
    Event Paused(sender As Task, e As EventArgs) Implements ITask.Paused
    Event Resumed(sender As Task, e As EventArgs) Implements ITask.Resumed
    Event Canceled(sender As Task, e As EventArgs) Implements ITask.Canceled
    Event Fired(sender As Task, e As ITask.ExecutionEventArgs) Implements ITask.Fired
    Event Executed(sender As Task, e As ITask.ExecutionEventArgs) Implements ITask.Executed
    Event Scheduled(sender As Task, e As ITask.ScheduleEventArgs) Implements ITask.Scheduled
    Event Before_Executing(sender As Task, e As EventArgs) Implements ITask.Before_Executing
    Event Executing(sender As Task, e As EventArgs) Implements ITask.Executing
#End Region
End Class

#Region "Property Grid"
Friend Class GroupsNames
    Inherits StringConverter

    Public Overloads Overrides Function GetStandardValuesSupported(ByVal context As ITypeDescriptorContext) As Boolean
        Return True
    End Function

    Public Overloads Overrides Function GetStandardValues(ByVal context As ITypeDescriptorContext) As StandardValuesCollection
        Return New StandardValuesCollection(Commands.GetGroupsNames())
    End Function

End Class
#End Region