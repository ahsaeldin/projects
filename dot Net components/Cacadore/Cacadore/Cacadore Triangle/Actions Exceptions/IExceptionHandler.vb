Public Interface IExceptionHandler

    Structure RaisedExceptionArgs
        Public SelectedErrorAction As SelectErrorActionEnum
        Public AlertnativeTaskIdValue As String
        Public ParentTask As Cacadore.Task
        Public Action As Cacadore.Action
        Public RetryTimes As Integer
    End Structure

    Enum SelectErrorActionEnum
        AlertOnError
        ContinueOnError
        RetryAction
        QuitOnError
        RunAnotherTaskOnError
    End Enum

    Property QuitOnError As Boolean
    Property RunAnotherTaskOnError As Boolean
    Property AlternativeTaskId As String
    Property RetryTimes As Integer
    Property AlertOnError As Boolean
    Property ContinueOnError As Boolean
    Property RetryAction As Boolean
    Property AttachedExceptionType As Type

End Interface
