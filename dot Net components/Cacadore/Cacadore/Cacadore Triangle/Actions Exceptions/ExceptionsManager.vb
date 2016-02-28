'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   ActionExceptionsManager Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
<Serializable()> Public Class ActionExceptionsManager
    Inherits CacadoreBase
    Implements IActionExceptionsManager
    Implements IDisposable

#Region "Fields"
    Private _raisedExceptionValue As Exception
#End Region

#Region "Properties"
    Public Property PossibleExceptions As New Dictionary(Of String, ActionExceptionHandler) Implements IActionExceptionsManager.PossibleExceptions

    Public ReadOnly Property PossibleExceptionsDescriptions As List(Of String) Implements IActionExceptionsManager.PossibleExceptionsDescriptions
        Get
            Return PossibleExceptions.Keys.ToList
        End Get
    End Property
#End Region

#Region "Methods"
    ''' <summary>
    ''' Adds the possible exception that can be managed.
    ''' </summary>
    ''' <remarks>
    ''' Use this method to add all types of Exception 
    ''' that may be raised upon executing this action.
    ''' </remarks>
    Public Sub AddPossibleException(exceptionType As Type, message As String) Implements IActionExceptionsManager.AddPossibleException
        Dim _actionExceptionHandler As New ActionExceptionHandler
        _actionExceptionHandler.AttachedExceptionType = exceptionType
        PossibleExceptions.Add(message, _actionExceptionHandler)
    End Sub

    ''' <summary>
    ''' Gets the selected exception handler using the message specified by the user.
    ''' </summary>
    ''' <param name="message">The message.</param><returns></returns>
    ''' <remarks>
    ''' Use this method to get the ActionExceptionHandler used to handle the message
    ''' selected by the user. It searches all PossibleExceptions dictionary of the 
    ''' corresponding ActionExceptionHandler to a given message.
    ''' </remarks>
    Public Function GetSelectedExceptionHandler(message As String) As ActionExceptionHandler Implements IActionExceptionsManager.GetSelectedExceptionHandler
        Dim kvp = From possEx In PossibleExceptions Where possEx.Key.Contains(message)
        If IsNothing(kvp.First.Value) Then Return New ActionExceptionHandler
        Return kvp.First.Value
    End Function

    ''' <summary>
    ''' Gets the raised exception handler by comparing the type of the raised exception
    ''' with the type of the stored one in PossibleExceptions.
    ''' </summary>
    ''' <param name="ex">The ex.</param><returns></returns>
    Public Function GetRaisedExceptionHandler(ex As Exception) As ActionExceptionHandler Implements IActionExceptionsManager.GetRaisedExceptionHandler
        For Each kvp In PossibleExceptions
            Dim exType = ex.GetType
            If exType Is kvp.Value.AttachedExceptionType Then
                Return kvp.Value
            End If
        Next
        Return New ActionExceptionHandler
    End Function

    ''' <summary>
    ''' Handles the exception and called directly by the action whenever an exception occured.
    ''' </summary>
    ''' <param name="ex">The ex.</param>
    Public Sub HandleException(ex As Exception,
                               _task As Cacadore.Task,
                               _action As Cacadore.Action) Implements IActionExceptionsManager.HandleException
        Dim actToTake = GetRaisedExceptionHandler(ex)
        Dim actToTakeArgs = actToTake.GetRaisedExceptionArgs
        actToTakeArgs.ParentTask = _task
        actToTakeArgs.Action = _action
        RaiseEvent Exception_Raised(ex, actToTakeArgs)
    End Sub
#End Region

#Region "Events"
    Private Event Exception_Raised(exception As System.Exception, actionToTake As IExceptionHandler.RaisedExceptionArgs) Implements IActionExceptionsManager.Exception_Raised
    Public Event After_Exception_Raised(exception As System.Exception, actionToTake As IExceptionHandler.RaisedExceptionArgs) Implements IActionExceptionsManager.After_Exception_Raised


#Region "Event Handlers"
    'Internal implementation of Exception_Raised
    Private Sub ActionExceptionsManager_Exception_Raised(exception As System.Exception, actionToTake As IExceptionHandler.RaisedExceptionArgs) Handles Me.Exception_Raised
        'Todo : add Actions to be taken here
        Select Case actionToTake.SelectedErrorAction
            Case IExceptionHandler.SelectErrorActionEnum.AlertOnError
                Dim taskName As String = actionToTake.ParentTask.Properties.TaskName
                Dim actionNo As Integer = Commands.ExtractActionNumber(actionToTake.Action.ActionId)
                Dim alertMessage As String = "Error executing action #" & actionNo & " of " & taskName & " task. "
                alertMessage = alertMessage & exception.Message
                Balora.Alerter.SendMessageToCacadu("AlertOnError", alertMessage)
            Case IExceptionHandler.SelectErrorActionEnum.RetryAction
                actionToTake.Action.RetryTime = actionToTake.Action.RetryTime + 1
                If actionToTake.Action.RetryTime <= actionToTake.RetryTimes Then actionToTake.Action.Execute()
                'Right now I don't now how to retry an action.
            Case IExceptionHandler.SelectErrorActionEnum.ContinueOnError
                'Nothing to to here.
                Debug.Print("asdfasdf")
            Case IExceptionHandler.SelectErrorActionEnum.RunAnotherTaskOnError
                actionToTake.ParentTask.Cancel()
                Dim otherTask = Commands.GetTaskById(actionToTake.AlertnativeTaskIdValue)
                otherTask.Execute()
                If Not IsNothing(otherTask) Then otherTask.Dispose()
                If Not IsNothing(otherTask) Then otherTask = Nothing
            Case IExceptionHandler.SelectErrorActionEnum.QuitOnError
                actionToTake.ParentTask.Cancel()
        End Select
        RaiseEvent After_Exception_Raised(exception, actionToTake)
    End Sub
#End Region
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
