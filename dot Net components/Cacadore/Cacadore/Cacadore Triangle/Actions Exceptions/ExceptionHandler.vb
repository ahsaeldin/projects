'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   ActionExceptionHandler Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
<Serializable()> Public Class ActionExceptionHandler
    Inherits CacadoreBase
    Implements IExceptionHandler
    Implements IDisposable

    'Important: Default action to take is defined in two places
    '1.@Private _alertOnErrorValue As Boolean = True
    '2.@GetSelectedActionToTake

#Region "Fields"
    Private _quitOnErrorValue As Boolean
    Private _alertOnErrorValue As Boolean = True
    Private _continueOnErrorValue As Boolean
    Private _alertnativeTaskIdValue As String
    Private _retryActionValue As Boolean
    Private _retryTimesValue As Integer
    Private _raisedExceptionValue As Exception
    Private _runAnotherTaskOnErrorValue As Boolean
#End Region

#Region "Properties"
#Region "Radio Properties"
    Public Property AlternativeTaskId As String Implements IExceptionHandler.AlternativeTaskId
        Get
            Return _alertnativeTaskIdValue
        End Get
        Set(value As String)
            SwitchOfRelatedProperties()
            _alertnativeTaskIdValue = value
        End Set
    End Property

    Public Property RetryAction As Boolean Implements IExceptionHandler.RetryAction
        Get
            Return _retryActionValue
        End Get
        Set(value As Boolean)
            SwitchOfRelatedProperties()
            _retryActionValue = value
        End Set
    End Property

    Public Property RetryTimes As Integer Implements IExceptionHandler.RetryTimes
        Get
            Return _retryTimesValue
        End Get
        Set(value As Integer)
            _retryTimesValue = value
        End Set
    End Property

    Public Property AlertOnError As Boolean Implements IExceptionHandler.AlertOnError
        Get
            Return _alertOnErrorValue
        End Get
        Set(value As Boolean)
            SwitchOfRelatedProperties()
            _alertOnErrorValue = value
        End Set
    End Property

    Public Property ContinueOnError As Boolean Implements IExceptionHandler.ContinueOnError
        Get
            Return _continueOnErrorValue
        End Get
        Set(value As Boolean)
            SwitchOfRelatedProperties()
            _continueOnErrorValue = value
        End Set
    End Property

    Public Property QuitOnError As Boolean Implements IExceptionHandler.QuitOnError
        Get
            Return _quitOnErrorValue
        End Get
        Set(value As Boolean)
            SwitchOfRelatedProperties()
            _quitOnErrorValue = value
        End Set
    End Property

    Public Property RunAnotherTaskOnError As Boolean Implements IExceptionHandler.RunAnotherTaskOnError
        Get
            Return _runAnotherTaskOnErrorValue
        End Get
        Set(value As Boolean)
            SwitchOfRelatedProperties()
            _runAnotherTaskOnErrorValue = value
        End Set
    End Property

    Public Property AttachedExceptionType As Type Implements IExceptionHandler.AttachedExceptionType

#End Region
#End Region

#Region "Methods"
    Public Function GetRaisedExceptionArgs() As IExceptionHandler.RaisedExceptionArgs
        Dim actToTakeArgs As New IExceptionHandler.RaisedExceptionArgs
        actToTakeArgs.SelectedErrorAction = GetSelectedActionToTake()
        actToTakeArgs.AlertnativeTaskIdValue = AlternativeTaskId
        actToTakeArgs.RetryTimes = RetryTimes
        Return actToTakeArgs
    End Function

    Private Function returnAlertOnError() As IExceptionHandler.SelectErrorActionEnum
        Return IExceptionHandler.SelectErrorActionEnum.AlertOnError
    End Function

    Private Function returnContinueOnError() As IExceptionHandler.SelectErrorActionEnum
        Return IExceptionHandler.SelectErrorActionEnum.ContinueOnError
    End Function

    Private Function returnQuitOnError() As IExceptionHandler.SelectErrorActionEnum
        Return IExceptionHandler.SelectErrorActionEnum.QuitOnError
    End Function

    Private Function returnRunAnotherTaskOnError() As IExceptionHandler.SelectErrorActionEnum
        Return IExceptionHandler.SelectErrorActionEnum.RunAnotherTaskOnError
    End Function

    Private Function returnRetryAction() As IExceptionHandler.SelectErrorActionEnum
        Return IExceptionHandler.SelectErrorActionEnum.RetryAction
    End Function

    Private Function GetSelectedActionToTake() As IExceptionHandler.SelectErrorActionEnum
        If _alertOnErrorValue Then
            Return returnAlertOnError()
        ElseIf _continueOnErrorValue Then
            Return returnContinueOnError()
        ElseIf _runAnotherTaskOnErrorValue Then
            Return returnRunAnotherTaskOnError()
        ElseIf _retryActionValue Then
            Return returnRetryAction()
        ElseIf _quitOnErrorValue Then
            Return returnQuitOnError()
        Else 'Very important man, this is the default action to take if an exception arise.
            Return returnAlertOnError()
        End If
    End Function

    Private Sub SwitchOfRelatedProperties()
        _retryActionValue = False
        _alertOnErrorValue = False
        _continueOnErrorValue = False
        _quitOnErrorValue = False
        _runAnotherTaskOnErrorValue = False
    End Sub
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
