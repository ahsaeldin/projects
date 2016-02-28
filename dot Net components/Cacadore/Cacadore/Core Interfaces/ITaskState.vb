'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   ITaskState Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
Public Interface ITaskState
    'Any watched task, will be one of a three types
    '1.It is running right now.
    '2.It is scheduled task in the future.
    '3.It failed and unfulfilled.
    Enum ExecutionState
        NotScheduled '0.A task not scheduled or attached to a trigger.
        Scheduled '1.A task that has a one or more trigger.
        Running  '2.Task is running right now.
        Unfulfilled '3.Started and running but interpreted.
        Failed '4.If all task's actions failed or one action failed (@ FailedOnError is true).
        Paused '5.Task is paused by user.
        Done '6.Task running successfully.
        DoneWithSomeErrors '7.Task done with some errors.
        CacadoreDisabled '8.Returned if Cacadore is disabled.
        Canceled '9.Task canceled.
        Resumed '10
        BeforeRunning '11
        Deleted '12 'Just a reference in database that task has been deleted.
        Fired '13
        Disabled '14
        HasNoActions
        WaitForAction
        WaitForConfirm
    End Enum

    ReadOnly Property taskId As String
    Property PreviousTriggerId As String
    Property CurrentTriggerId As String
    Property NextTriggerId As String
    Property CurrentExecutingActionId As String
    Property TaskExecutionState As ExecutionState
    Property PreviousTaskExecutionState As ExecutionState
    Property MyTask As Task

    Function Clear() As Boolean
    Sub ChangeExecutionState(tskState As ITaskState.ExecutionState)
End Interface