Imports Cacadore
Imports System.Runtime.Serialization

<Serializable()> Public Class LastFacebookNotification
    Inherits Action
    Implements IDisposable

#Region "Fields"
    Private fb As New Actions.FacebookActions.FbHelper
    Private _inputsValue As LastFacebookNotificationInputs
    Private _errorsValue As LastFacebookNotificationErrors
    Private _outputsValue As LastFacebookNotificationOutputs
#End Region

#Region "Constructors"
    Public Sub New()
        'To use inputs internally.
        _inputsValue = CType(Inputs, LastFacebookNotificationInputs)
        _outputsValue = New LastFacebookNotificationOutputs
        _errorsValue = New LastFacebookNotificationErrors
        Errors = _errorsValue
        SetPossibleExceptions()
    End Sub
#End Region

#Region "Properties"
    Public Overrides Property Errors As IAction.IErrors
        Get
            Return _errorsValue
        End Get
        Set(value As IAction.IErrors)
            _errorsValue = CType(value, LastFacebookNotificationErrors)
        End Set
    End Property

    Public Overrides Property Inputs As IAction.IInputs
        Get
            Return _inputsValue
        End Get
        Set(value As IAction.IInputs)
            _inputsValue = CType(value, LastFacebookNotificationInputs)
        End Set
    End Property

    Public Overrides Property Outputs As IAction.IOutputs
        Get
            Return _outputsValue
        End Get
        Set(value As IAction.IOutputs)
            _outputsValue = CType(value, LastFacebookNotificationOutputs)
        End Set
    End Property

    Public Overrides ReadOnly Property ActionType As Cacadore.IAction.ActionTypeEnum
        Get
            Return IAction.ActionTypeEnum.LastFacebookNotification
        End Get
    End Property
#End Region

#Region "Methods"
    Public Overrides Function GetFullDescription() As String
        Dim [action] = Me
        Dim inputs As LastFacebookNotificationInputs = CType(Me.Inputs, LastFacebookNotificationInputs)

        Dim listItem As String
        listItem = "Last Facebook Notification" & " Action"

        Dim formatedComments As String = GetFormattedComment()
        If formatedComments <> vbNullString Then listItem = listItem & ", Comments: " & formatedComments
        Return listItem
    End Function

    'How to:-'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'TODO : contains the logic of execution for any new action.
    'Set Input Property here
    'Set Output Property @ return
    'Set Result Property @ return
    'And after execution, you will have a result with input and output
    'And @ finish return "ITaskState.ExecutionState" that indicates the result
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Input and Output are optional because some tasks are parameter less.
    Public Overrides Function Execute(Optional retryTimes As Integer = 1) As Cacadore.ITaskState.ExecutionState
        'TODO : You must add the next line to any new action.
        If Me.Skip Then Return ITaskState.ExecutionState.Canceled

        Dim _inputs As LastFacebookNotificationInputs = CType(Inputs, LastFacebookNotificationInputs)

        If _inputs.LoginInAnotherAccount Then fb.LogOut()

        Dim _outputs As LastFacebookNotificationOutputs = CType(Outputs, LastFacebookNotificationOutputs)

        Dim waitHandle As New Threading.ManualResetEvent(False)

        StartExecutionLoop(_inputs, _outputs, retryTimes, waitHandle)

        WaitForExit(waitHandle)

        waitHandle.Dispose()

        waitHandle = Nothing

        'TODO : You must add the next line for any new action.
        Me.RaiseCompletedEvent(Me, New EventArgs)
        Return ITaskState.ExecutionState.Done
    End Function

    Private Sub StartExecutionLoop(_inputs As LastFacebookNotificationInputs,
                           _outputs As LastFacebookNotificationOutputs,
                           retryTimes As Integer,
                           waitHandle As Threading.ManualResetEvent)
        Const waitForLoginOnly As Integer = -2
        Const waitForNotificationReceived As Integer = 0
        For i = 1 To retryTimes
            'Unfortunately, we must use a manual reset event inside login.
            'because fbLoginDlg opened in another thread, so 
            'if you called it from the background worker, the 
            'thread won't wait till getting the access_token 
            'therefore nothing will work after that like getting 
            'last(notification.

            Try
                If _inputs.WaitForExitPeriod = waitForLoginOnly Then
                    fb.Login(waitHandle)
                Else
                    fb.Login()
                End If
            Catch ex As Exception
                GetErrors(ex)
            End Try


            Try
                If _inputs.WaitForExitPeriod = waitForNotificationReceived Then
                    fb.DisplayLastNotification(_outputs, waitHandle)
                Else
                    fb.DisplayLastNotification(_outputs)
                End If
            Catch ex As Exception
                GetErrors(ex)
            End Try

        Next
    End Sub

    Protected Overrides Sub WaitForExit(Optional ByRef state As Object = Nothing)
        Const waitForLoginOnly As Integer = -2
        Const waitForNotificationReceived As Integer = 0

        If Not IsNothing(Me.ParentTask) AndAlso Not Me.ParentTask.WaitBetweenActions Then Exit Sub
        'So WaitBetweenActions is true or Me.ParentTask is nothing.
        Dim _inputs As LastFacebookNotificationInputs = CType(Inputs, LastFacebookNotificationInputs)
        If Not IsNothing(_inputs) Then
            LogWait(_inputs.WaitForExitPeriod)
            If _inputs.WaitForExitPeriod > 0 Then
                Threading.Thread.Sleep(CInt(_inputs.WaitForExitPeriod))
            ElseIf _inputs.WaitForExitPeriod = waitForNotificationReceived Then
                If Not IsNothing(state) Then
                    Dim waitHandle As Threading.ManualResetEvent = CType(state, Threading.ManualResetEvent)
                    waitHandle.WaitOne(10000)
                End If
            ElseIf _inputs.WaitForExitPeriod = -1 Then
                'Continue Immediately.
            ElseIf _inputs.WaitForExitPeriod = waitForLoginOnly Then
                If Not FbHelper.IsLogon Then
                    Dim waitHandle As Threading.ManualResetEvent = CType(state, Threading.ManualResetEvent)
                    waitHandle.WaitOne(10000)
                End If
            End If
        End If
    End Sub

    Public Overrides Sub SetPossibleExceptions()
        ExceptionsManager.AddPossibleException(GetType(FbHelper.FacebookLogonCancelled),
                                               "Facebook logon cancelled.")
        ExceptionsManager.AddPossibleException(GetType(FbHelper.FacebookLogonFailed),
                                               "Facebook logon failed.")
        ExceptionsManager.AddPossibleException(GetType(FbHelper.FacebookGenericError),
                                               "Facebook generic error.")
    End Sub

    Protected Sub GetErrors(ByVal ex As Global.System.Exception)
        _errorsValue.ErrorList.Add(ex.Message)
        Me.ParentTask.Errors.ListError(ActionId, ex.Message)
        Me.ExceptionsManager.HandleException(ex, ParentTask, Me)
    End Sub

    Public Overrides Function Clone() As Cacadore.Action
        Dim actionMemStream As New Global.System.IO.MemoryStream
        Balora.Serializer.BinarySerializer.SerializeObject(Me, actionMemStream)
        Dim duplicatedAction As LastFacebookNotification = CType(Balora.Serializer.BinarySerializer.DeserializeObject(actionMemStream.ToArray), LastFacebookNotification)
        Return duplicatedAction
    End Function
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                fb.Dispose()
                _inputsValue = Nothing
                _errorsValue = Nothing
                _outputsValue = Nothing
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

#Region "Inner Classes"
#Region "Inputs Class"
    <Serializable()>
    Public Class LastFacebookNotificationInputs
        Implements IAction.IInputs
        Property LoginInAnotherAccount As Boolean = False
        Property WaitForExitPeriod As Decimal = 1
    End Class
#End Region

#Region "Output Class"
    <Serializable()>
    Public Class LastFacebookNotificationOutputs
        Implements IAction.IOutputs
        Property SaveToMember As String
        Property ParentAction As Cacadore.Action
    End Class
#End Region

#Region "Errors Class"
    <Serializable()>
    Public Class LastFacebookNotificationErrors
        Implements IAction.IErrors
        Property ErrorList As New List(Of String)
    End Class
#End Region
#End Region
End Class