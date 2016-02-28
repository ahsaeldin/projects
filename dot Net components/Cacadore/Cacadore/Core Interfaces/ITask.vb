'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   ITask Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
''' <summary>
''' Base interface for Task class.
''' </summary>
Public Interface ITask
#Region "Structures"
    Structure ExecutionEventArgs
        Public result As Object
        Public trigger As ITrigger
        Public state As ITaskState.ExecutionState
    End Structure

    Structure ScheduleEventArgs
        Public Trigger As Trigger
    End Structure
#End Region

#Region "Properties"
    Property Errors As TaskErrors
    Property ActionsCount As Integer
    Property ParentGroupId As String
    Property BreakOnFailed As Boolean
    Property TriggersCount As Integer
    ReadOnly Property TaskId As String
    Property ParentGroupName As String
    Property Variables As TaskVariables
    ReadOnly Property State As TaskState
    Property Properties As TaskProperties
    Property WaitBetweenActions As Boolean
    ReadOnly Property MayFiredAgain As Boolean
    ReadOnly Property Triggers As TriggersCollection
    ReadOnly Property ActionsCollection As ActionsCollection
    ReadOnly Property History As System.Collections.Generic.List(Of HistoryEntry)
#End Region

#Region "Methods"
    Sub Pause()
    Sub Cancel()
    Sub [Resume]()
    Sub PauseAllTriggers()
    Function Clone() As Task
    Function Save() As Boolean
    Function Delete() As Boolean
    Sub SetTaskState(tskState As TaskState)
    Function Execute() As ITaskState.ExecutionState
    Function UpdateAction(action As Action) As Boolean
    Function UnAttachTrigger(trg As ITrigger) As Boolean
    Function UnAttachTrigger(ByVal trgId As String) As Boolean
    Function UnAttachAction(action As Action) As Boolean
    Function AttachAction(ByVal [action] As Action) As String
    Function AttachTrigger(ByVal [trigger] As Trigger) As String
    Sub SetTaskProperties(lastFireTime As Date?, lastRunResult As String, nextFireTime As Date?)
#End Region

#Region "Events"
    Event Paused(Sender As Task, e As EventArgs)
    Event Resumed(Sender As Task, e As EventArgs)
    Event Canceled(Sender As Task, e As EventArgs)
    Event Before_Executing(Sender As Task, e As EventArgs)
    Event Executed(Sender As Task, e As ExecutionEventArgs)
    Event Fired(sender As Task, e As ITask.ExecutionEventArgs)
    Event Scheduled(Sender As Task, e As ITask.ScheduleEventArgs)
    Event Executing(sender As Task, e As EventArgs)
#End Region

End Interface
