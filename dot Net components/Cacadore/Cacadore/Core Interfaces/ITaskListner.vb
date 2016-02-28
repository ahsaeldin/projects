'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   ITaskListner Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
''' <summary>
''' Base interface for any class want to be a listener for Task class.
''' </summary>
Public Interface ITaskListner
    Inherits Balora.IListener

#Region "Methods"
    Overloads Sub Notify(tsk As Task, notificationEvent As ITaskState.ExecutionState)
#End Region

#Region "Events"
    Event Scheduled(Sender As Task, e As EventArgs)
    Event BeforeExecuting(Sender As Task, e As EventArgs)
    Event Executing(sender As Task, e As EventArgs)
    Event Executed(Sender As Task, e As EventArgs)
    Event Paused(Sender As Task, e As EventArgs)
    Event Canceled(Sender As Task, e As EventArgs)
    Event Resumed(Sender As Task, e As EventArgs)
#End Region
End Interface
