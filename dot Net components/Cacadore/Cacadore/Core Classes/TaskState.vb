'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   TaskState Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
'The only way for change "Task.State.TaskExecutionState property" 
'is to call "TaskState.ChangeExecutionState" 
'Ex:- .State.ChangeExecutionState(ITaskState.ExecutionState.Running)
Imports System.Windows.Forms

<Serializable()> Public Class TaskState
    Inherits CacadoreBase
    Implements ITaskState
    Implements ICrudable
    Implements IDisposable

#Region "Fields"
    Private _myTaskValue As Task
    Friend _taskIdValue As String
    Private _nextTriggerIdValue As String
    Private _currentTriggerIdValue As String
    Private _previoudTriggerIdValue As String
    Private _currentExecutingActionIdValue As String
    Private _taskPreviousExecutionState As ITaskState.ExecutionState
    Private _taskExecutionState As ITaskState.ExecutionState = ITaskState.ExecutionState.NotScheduled
#End Region

#Region "Constructors"
    Public Sub New(taskId As String)
        _taskIdValue = taskId
    End Sub
#End Region

#Region "Properties"
    Public Property MyTask As Task Implements ITaskState.MyTask
        Get
            If Not IsNothing(_myTaskValue) Then
                Return _myTaskValue
            Else
                _myTaskValue = Commands.GetTaskById(_taskIdValue)
                Return _myTaskValue
            End If
        End Get
        Set(value As Task)
            _myTaskValue = value
        End Set
    End Property

    Public ReadOnly Property taskId As String Implements ITaskState.taskId
        Get
            Return _taskIdValue
        End Get
    End Property

    Public Property CurrentExecutingActionId As String Implements ITaskState.CurrentExecutingActionId
        Get
            Return _currentExecutingActionIdValue
        End Get
        Set(value As String)
            _currentExecutingActionIdValue = value
        End Set
    End Property

    ''' <summary>
    ''' Gets the next trigger id. when task get scheduled/done/canceled
    ''' </summary>
    Public Property PreviousTriggerId As String Implements ITaskState.PreviousTriggerId
        Get
            Return _previoudTriggerIdValue
        End Get
        Set(value As String)
            _previoudTriggerIdValue = value
        End Set
    End Property

    ''' <summary>
    ''' Gets the current trigger id, only if task is currently running.
    ''' </summary>
    Public Property CurrentTriggerId As String Implements ITaskState.CurrentTriggerId
        Get
            Return _currentTriggerIdValue
        End Get
        Set(value As String)
            _currentTriggerIdValue = value
        End Set
    End Property

    ''' <summary>
    ''' Gets the next trigger id. when task get scheduled/done/canceled
    ''' </summary>
    Public Property NextTriggerId As String Implements ITaskState.NextTriggerId
        Get
            Return _nextTriggerIdValue
        End Get
        Set(value As String)
            _nextTriggerIdValue = value
        End Set
    End Property

    Public Property TaskExecutionState As ITaskState.ExecutionState Implements ITaskState.TaskExecutionState
        Get
            'Use GetExecutionState method to get a fresh value.
            Return _taskExecutionState
        End Get
        Set(value As ITaskState.ExecutionState)
            _taskExecutionState = value
        End Set
    End Property

    'Very important:  "search and delete this comment if you intended to save it to database."
    'PreviousTaskExecutionState property is not added yet to the database and it will not be saved there in state table.
    Public Property TaskPreviousExecutionState As ITaskState.ExecutionState Implements ITaskState.PreviousTaskExecutionState
        Get
            If _taskPreviousExecutionState <> Nothing Then
                Return _taskPreviousExecutionState
            Else
                Return getTaskPreviousExecutionState()
            End If
        End Get
        Set(value As ITaskState.ExecutionState)
            _taskPreviousExecutionState = value
        End Set
    End Property
#End Region

#Region "Methods"
    'The only way for change "Task.State.TaskExecutionState property" is to call "Task.State.ChangeExecutionState"
    '.TaskState.ChangeExecutionState(ITaskState.ExecutionState.Running)
    Public Sub ChangeExecutionState(tskState As ITaskState.ExecutionState) Implements ITaskState.ChangeExecutionState
        saveTaskProperties(MyTask, tskState)
        'AhSaElDin 20111230: TaskPreviousExecutionState is writable now
        'Very important: "search and delete this comment if you intended to save it to database."
        'PreviousTaskExecutionState property is not added yet to the database and it will not be saved there in state table.
        'Balora.Util.SetPrivateField(tsk.State, "_taskPreviousExecutionState", tsk.State.TaskExecutionState)
        _taskPreviousExecutionState = TaskExecutionState
        'First step is to update current instance field, to avoid making trips to database.
        _taskExecutionState = tskState
        MyTask.State.TaskExecutionState = tskState
        'Disable saving the task to avoid calling tectonic while executing the task.
        'MyTask.Save() '*.What might happen if uncommented 'MyTask.Save() @ ChangeExecutionState hehehehehe (I may try someday.)
        addToWather()
    End Sub

    'Use this method to get a fresh value of the TaskExectionState
    Public Function GetExecutionState() As ITaskState.ExecutionState
        If Watcher.WatchedTasks.ContainsKey(taskId) Then
            _taskExecutionState = Watcher.GetWatchedTaskById(taskId).State.TaskExecutionState
        End If
        Return _taskExecutionState
    End Function

    'Set me to nothing and delete its entires from database.
    Public Function Clear() As Boolean Implements ITaskState.Clear
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Dim result As Boolean = CBool(InvokeTectonicDlg(Me, {Me.taskId}, DeleteOperationDlg, My.Resources.TskNotPresisted))
            clearFields()
            Return result
        End If
    End Function

    Private Function getTaskPreviousExecutionState() As ITaskState.ExecutionState
        Return Commands.GetTaskPreviousExecutionState(taskId)
    End Function

    Private Sub getCurrenAndNextAndPreivousTriggersIds(ByVal tsk As Task, ByRef tskNextTriggerId As String, ByRef tskCurTriggerId As String, ByRef tskPreTriggerId As String)
        tskNextTriggerId = tsk.State.NextTriggerId
        tskCurTriggerId = tsk.State.CurrentTriggerId
        tskPreTriggerId = tsk.State.PreviousTriggerId
    End Sub

    Private Sub getPreTimeRunUTC(ByVal tsk As Task, ByRef preTimeRunUTC As Nullable(Of DateTime), ByVal tskPreTriggerId As String)
        If tskPreTriggerId <> "" Then preTimeRunUTC = tsk.MyTriggers.Triggers(tskPreTriggerId).PreviousFireTimeUTC
    End Sub

    Private Sub getNextTimeRunUTC(ByVal tsk As Task, ByRef nextTimeRunUTC As Nullable(Of DateTime), ByVal tskNextTriggerId As String)
        If tskNextTriggerId <> "" Then nextTimeRunUTC = tsk.MyTriggers.Triggers(tskNextTriggerId).NextFireTimeUTC
    End Sub

    Private Sub getLastTimeRunResult(ByVal tsk As Task, ByRef lastTimeRunResult As String)
        lastTimeRunResult = CStr(tsk.State.TaskPreviousExecutionState)
    End Sub

    Private Sub addToWather()
        Watcher.Watch(MyTask)
    End Sub

    Private Sub saveTaskProperties(tsk As Task, tskState As ITaskState.ExecutionState)
        Dim preTimeRunUTC As Nullable(Of DateTime)
        Dim nextTimeRunUTC As Nullable(Of DateTime)
        Dim tskCurTriggerId As String = vbNullString
        Dim tskPreTriggerId As String = vbNullString
        Dim tskNextTriggerId As String = vbNullString
        Dim lastTimeRunResult As String = vbNullString
        getCurrenAndNextAndPreivousTriggersIds(tsk, tskNextTriggerId, tskCurTriggerId, tskPreTriggerId)
        getPreTimeRunUTC(tsk, preTimeRunUTC, tskPreTriggerId)
        getNextTimeRunUTC(tsk, nextTimeRunUTC, tskNextTriggerId)
        getLastTimeRunResult(tsk, lastTimeRunResult)
        tsk.UpdateProperties(preTimeRunUTC, lastTimeRunResult, nextTimeRunUTC)
    End Sub

    Private Sub clearFields()
        _taskIdValue = ""
        _currentExecutingActionIdValue = ""
        NextTriggerId = ""
        _taskExecutionState = Nothing
    End Sub
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                If Not IsNothing(_myTaskValue) Then _myTaskValue.Dispose()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
            _myTaskValue = Nothing
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

End Class