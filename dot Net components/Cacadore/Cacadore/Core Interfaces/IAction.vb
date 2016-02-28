'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   IAction Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
''' Base interface for Action class.
''' </summary>

Public Interface IAction
#Region "Enums & Structure"
    Enum ActionTypeEnum
        'Every time you inherit a new action, add its type here and...
        'Set this type in the Overridable ReadOnly Property ActionType
        'in the new derived action class, just as you do in Actions.vb
        StartApplication
        TakeScreenShot
        ShowMessageBox
        ExitProcess
        ShutdownPC
        ShowAlert
        LastFacebookNotification
        'Todo Note: Very important Note:-
        '*.For each new action class, add it to 
        '       *.Commands.CreateActionByType()
    End Enum

    Interface IInputs
    End Interface

    Interface IOutputs
    End Interface

    Interface IErrors
    End Interface
#End Region

#Region "Properties"
    Property Shape As Shape
    Property Skip As Boolean
    Property Input As IInputs
    Property Errors As IErrors
    Property ParentTask As Task
    Property Enabled As Boolean
    Property Comments As String
    Property Outputs As IOutputs
    Property RetryTime As Integer
    ReadOnly Property TaskId As String
    ReadOnly Property ActionId As String
    Property Result As ITaskState.ExecutionState
    ReadOnly Property ActionType As ActionTypeEnum
    Property ExceptionsManager As ActionExceptionsManager
#End Region

#Region "Methods"
    Sub WaitForExit(Optional ByRef state As Object = Nothing)
    Sub SetPossibleExceptions()
    Function Delete() As Boolean
    Function Highlight() As Boolean
    Function Execute(Optional retryTimes As Integer = 1) As ITaskState.ExecutionState
#End Region

#Region "Events"
    Event Completed As EventHandler
#End Region
End Interface