Public Interface IActionExceptionsManager

    ReadOnly Property PossibleExceptionsDescriptions As List(Of String)

    Property PossibleExceptions As Dictionary(Of String, ActionExceptionHandler)

    ''' <summary>
    ''' Adds the possible exception that can be managed.
    ''' </summary>
    ''' <remarks>
    ''' Use this method to add all types of Exception 
    ''' that may be raised upon executing this action.
    ''' </remarks>
    Sub AddPossibleException(exceptionType As Type, message As String)

    ''' <summary>
    ''' Gets the selected exception handler using the message specified by the user.
    ''' </summary>
    ''' <param name="message">The message.</param><returns></returns>
    ''' <remarks>
    ''' Use this method to get the ActionExceptionHandler used to handle the message
    ''' selected by the user. It searches all PossibleExceptions dictionary of the 
    ''' corresponding ActionExceptionHandler to a given message.
    ''' </remarks>
    Function GetSelectedExceptionHandler(message As String) As ActionExceptionHandler

    ''' <summary>
    ''' Gets the raised exception handler by comparing the type of the raised exception
    ''' with the type of the stored one in PossibleExceptions.
    ''' </summary>
    ''' <param name="ex">The ex.</param><returns></returns>
    Function GetRaisedExceptionHandler(ex As Exception) As ActionExceptionHandler

    ''' <summary>
    ''' Handles the exception and called directly by the action whenever an exception occurred.
    ''' </summary>
    ''' <param name="ex">The ex.</param>
    Sub HandleException(ex As Exception,
                        _task As Cacadore.Task,
                        _action As Cacadore.Action)

    Event Exception_Raised(exception As System.Exception, actionToTake As IExceptionHandler.RaisedExceptionArgs)

    Event After_Exception_Raised(exception As Exception, actionToTake As IExceptionHandler.RaisedExceptionArgs)

End Interface