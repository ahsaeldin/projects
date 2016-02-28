'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   ShutdownPC Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
Imports Balora.SysManage

Namespace Actions
    Namespace System
        <Serializable()> Public Class ShutdownPC
            Inherits Action
            Implements IDisposable

            Structure ShutdownEventArgs
                Dim PeriodBeforeExecute As Integer
                Dim IsConfirmBeforeExecute As Boolean
            End Structure

#Region "Fields"
            Private _inputsValue As ShutdownPCInputs
            Private _errorsValue As ShutdownPCErrors
            Private _outputsValue As ShutdownPCOutputs
            Private _sysManage As New Balora.SysManage
            Private shutdownAfterPeriodTimer As New Timers.Timer()
#End Region

#Region "Constructors"
            Public Sub New()
                'To use inputs internally.
                _inputsValue = CType(Inputs, ShutdownPCInputs)
            End Sub

            Public Sub New(input As ShutdownPCInputs)
                'To access inputs externally.
                Inputs = input
                'To use inputs internally.
                _inputsValue = CType(Inputs, ShutdownPCInputs)
                _errorsValue = New ShutdownPCErrors
                Errors = _errorsValue
            End Sub
#End Region

#Region "Properties"
#Region "Overrided Properties"
            Public Overrides Property Errors As IAction.IErrors
                Get
                    Return _errorsValue
                End Get
                Set(value As IAction.IErrors)
                    _errorsValue = CType(value, ShutdownPCErrors)
                End Set
            End Property

            Public Overrides Property Inputs As IAction.IInputs
                Get
                    Return _inputsValue
                End Get
                Set(value As IAction.IInputs)
                    _inputsValue = CType(value, ShutdownPCInputs)
                End Set
            End Property

            Public Overrides Property Outputs As IAction.IOutputs
                Get
                    Return _outputsValue
                End Get
                Set(value As IAction.IOutputs)
                    _outputsValue = CType(value, ShutdownPCOutputs)
                End Set
            End Property

            Public Overrides ReadOnly Property ActionType As IAction.ActionTypeEnum
                Get
                    Return IAction.ActionTypeEnum.ShutdownPC
                End Get
            End Property
#End Region

#Region "Added Properties"
            Private Property Canceled As Boolean = False
#End Region
#End Region

#Region "Methods"
#Region "Overrided Methods"
#Region "Execution Methods"
            Public Overrides Function Execute(Optional retryTimes As Integer = 1) As ITaskState.ExecutionState
                If Me.Skip Then Return ITaskState.ExecutionState.Canceled
                If IsNothing(_inputsValue) Then _inputsValue = CType(Inputs, ShutdownPCInputs)
                _checkIfForced()
                Dim executionResult As Boolean = False
                PrepareAndRaiseEvent()
                'Dim _outputs As ShutdownPCOutputs = CType(Outputs, ShutdownPCOutputs)
                'SetOutput(_outputs.SaveToMember, "new value here")
                If Not Canceled Then Return ShowShutdownAlert()
                Return ITaskState.ExecutionState.Canceled
            End Function

            Private Function ShowShutdownAlert() As ITaskState.ExecutionState
                Dim shutDownType As String = _inputsValue.ShutdownType.ToString.Replace("Forced", "")
                Dim shutdownMessage As String = "Your computer will " & shutDownType
                If _inputsValue.PeriodBeforeShutdown > 0 Then
                    shutdownMessage = shutdownMessage & " after " & _inputsValue.PeriodBeforeShutdown & " seconds."
                Else
                    shutdownMessage = shutdownMessage & " now."
                End If

                If _inputsValue.ConfirmBeforeExecute OrElse _inputsValue.AllowCancel Then
                    shutdownMessage = shutdownMessage & vbNewLine & "Confirm?"
                    'AhSaElDin 29/09/2012: Disable alert form and use a normal message box, 
                    'because alert hangs up the entire application after 5 minutes from not
                    'chooseing yes/no. May be I use it again in the future. Just enable the next line again.
                    'Balora.Alerter.LoadAlertForm("", shutdownMessage, "ShutdownPC", True, Nothing, Me)
                    Dim msgResult As MsgBoxResult = MsgBox(shutdownMessage, CType(MsgBoxStyle.YesNo + MsgBoxStyle.Question, MsgBoxStyle), "Cacadu")
                    If msgResult = MsgBoxResult.Yes Then Me.PrepareForShutdown()
                    Return ITaskState.ExecutionState.WaitForConfirm
                ElseIf _inputsValue.PeriodBeforeShutdown > 0 Then
                    'AhSaElDin 29/09/2012: Disable alert form and use a normal message box, 
                    'because alert hangs up the entire application after 5 minutes from not
                    'chooseing yes/no. May be I use it again in the future. Just enable the next line again.
                    Balora.Alerter.LoadAlertForm("", shutdownMessage, "ShutdownPC", False, Nothing, Me)
                    Return PrepareForShutdown()
                Else
                    Return PrepareForShutdown()
                End If
            End Function

            Public Function PrepareForShutdown() As ITaskState.ExecutionState
                Dim executionResult As Boolean = False
                If Not Canceled Then
                    Dim isThereWaitPeriod As Boolean = WaitPeriodBeforeShutdown()
                    'للأسف مش لاقي حل اني اخد النتيجة الحقيقة من التايمر.
                    If isThereWaitPeriod Then Return ITaskState.ExecutionState.Done
                    executionResult = PerformShutdown()
                Else
                    RaiseEvent Shutdown_Canceled(_inputsValue.ShutdownType)
                End If
                Return AfterShutdown(executionResult)
            End Function

            Private Function PerformShutdown() As Boolean
                Dim executionResult As Boolean = False
                executionResult = _sysManage.Shutdown(_inputsValue.ShutdownType)
                Return executionResult
            End Function

            Private Function AfterShutdown(executionResult As Boolean) As ITaskState.ExecutionState
                If Not executionResult Then
                    GetErrors()
                    If Canceled Then
                        Return ITaskState.ExecutionState.Canceled
                    Else
                        Return ITaskState.ExecutionState.Failed
                    End If
                End If
                Me.RaiseCompletedEvent(Me, New EventArgs)
                Return ITaskState.ExecutionState.Done
            End Function
#End Region

            Protected Sub GetErrors()
                Dim errorMessage As String = "Cannot execute shutdown action."
                _errorsValue.ErrorList.Add(errorMessage)
                Me.ParentTask.Errors.ListError(Me.ActionId, errorMessage)
            End Sub

            Public Overrides Function GetFullDescription() As String
                Dim [action] = Me
                Dim inputs As ShutdownPCInputs = CType(Me.Inputs, ShutdownPCInputs)
                Dim shutdownType As String = inputs.ShutdownType.ToString

                Dim listItem As String
                listItem = "Shutdown PC" & " Action,"
                listItem = listItem & " Shutdown Type: " & shutdownType

                Dim formatedComments As String = GetFormattedComment()
                If formatedComments <> vbNullString Then listItem = listItem & ", Comments: " & formatedComments
                Return listItem
            End Function

            Public Overrides Sub SetPossibleExceptions()
                'ExceptionsManager.AddPossibleException(GetType(Global.System.ComponentModel.Win32Exception),
                '                       "The associated process could not be terminated")
                'ExceptionsManager.AddPossibleException(GetType(NotSupportedException),
                '                                       "You are attempting to call Kill for a process that is running on a remote computer")
                'ExceptionsManager.AddPossibleException(GetType(InvalidOperationException),
                '                                       "The process has already exited")
            End Sub

            Public Overrides Function Clone() As Action
                Dim actionMemStream As New Global.System.IO.MemoryStream
                Balora.Serializer.BinarySerializer.SerializeObject(Me, actionMemStream)
                Dim duplicatedAction As ShutdownPC = CType(Balora.Serializer.BinarySerializer.DeserializeObject(actionMemStream.ToArray), ShutdownPC)
                Return duplicatedAction
            End Function
#End Region

#Region "Added Methods"
            Public Function CancelShutDown() As Boolean
                If Not _inputsValue.AllowCancel Then Return False
                Canceled = True
                Return _sysManage.CancelShutdown()
            End Function

            Private Sub _checkIfForced()
                If _inputsValue.IsForced Then
                    Select Case _inputsValue.ShutdownType
                        Case Balora.ShutdownTypeEnum.Hibernate
                            _inputsValue.ShutdownType = Balora.ShutdownTypeEnum.ForcedHibernate
                        Case Balora.ShutdownTypeEnum.LogOff
                            _inputsValue.ShutdownType = Balora.ShutdownTypeEnum.ForcedLogOff
                        Case Balora.ShutdownTypeEnum.Restart
                            _inputsValue.ShutdownType = Balora.ShutdownTypeEnum.ForcedRestart
                        Case Balora.ShutdownTypeEnum.Shutdown
                            _inputsValue.ShutdownType = Balora.ShutdownTypeEnum.ForcedShutdown
                        Case Balora.ShutdownTypeEnum.Sleep
                            _inputsValue.ShutdownType = Balora.ShutdownTypeEnum.ForcedSleep
                    End Select
                End If
            End Sub

            Private Function PrepareEventArgs() As ShutdownEventArgs
                Dim _shutdownEventArgs As New ShutdownEventArgs
                _shutdownEventArgs.IsConfirmBeforeExecute = _inputsValue.ConfirmBeforeExecute
                _shutdownEventArgs.PeriodBeforeExecute = _inputsValue.PeriodBeforeShutdown
                Return _shutdownEventArgs
            End Function

            Private Sub PrepareAndRaiseEvent()
                If Canceled Then Exit Sub
                Dim _shutdownEventArgs As ShutdownEventArgs = PrepareEventArgs()
                RaiseEvent Before_Shutdown(_inputsValue.ShutdownType, _shutdownEventArgs)
            End Sub

            Private Function WaitPeriodBeforeShutdown() As Boolean
                If Canceled Then Return False
                Dim _periodBeforeShutdownMilleSeconds As Integer = _inputsValue.PeriodBeforeShutdown * 1000
                If _periodBeforeShutdownMilleSeconds = 0 Then Return False
                shutdownAfterPeriodTimer.Enabled = True
                shutdownAfterPeriodTimer.Interval = _periodBeforeShutdownMilleSeconds
                AddHandler shutdownAfterPeriodTimer.Elapsed, AddressOf shutdownAfterPeriodTimer_Elapsed
                Return True
            End Function

            Private Sub shutdownAfterPeriodTimer_Elapsed(sender As Object, e As Global.System.Timers.ElapsedEventArgs)
                Dim executionResult As Boolean = PerformShutdown()
                'للأسف مش لاقي حل اني اخد النتيجة الحقيقة من التايمر.
                AfterShutdown(executionResult)
                shutdownAfterPeriodTimer.Enabled = False
            End Sub
#End Region

#Region "Events"
            Event Before_Shutdown(shutdownType As Balora.ShutdownTypeEnum, shutdownEventArgs As ShutdownEventArgs)
            Event Shutdown_Canceled(shutdownType As Balora.ShutdownTypeEnum)
#End Region
#End Region

#Region "Inner Classes"
            ''' <summary>
            ''' Holds all required inputs for this action
            ''' and constructor forces to define the mandatory
            ''' inputs.
            ''' </summary>
            <Serializable()>
            Public Class ShutdownPCInputs
                Implements IAction.IInputs

#Region "Properties"
                Property ShutdownType As Balora.ShutdownTypeEnum
                'Is forced action.
                Public Property IsForced() As Boolean
                'Period in seconds just before shutdown really executes. Default is 30 seconds.
                Property PeriodBeforeShutdown As Integer = 0
                'To indicate if you want to get user confirmation before execute.
                Property ConfirmBeforeExecute As Boolean = False
                'A property for allowing the cancelation of Shutdown/Restart/LogOff.
                Property AllowCancel As Boolean
#End Region
            End Class

            <Serializable()>
            Public Class ShutdownPCOutputs
                Implements IAction.IOutputs
            End Class

            <Serializable()>
            Public Class ShutdownPCErrors
                Implements IAction.IErrors
                Property ErrorList As New List(Of String)
            End Class
#End Region

#Region "IDisposable Support"
            Private disposedValue As Boolean ' To detect redundant calls

            ' IDisposable
            Protected Overridable Sub Dispose(disposing As Boolean)
                If Not Me.disposedValue Then
                    If disposing Then
                        ' TODO: dispose managed state (managed objects).
                        shutdownAfterPeriodTimer.Dispose()
                    End If

                    _sysManage = Nothing
                    _inputsValue = Nothing
                    _errorsValue = Nothing
                    _outputsValue = Nothing
                    shutdownAfterPeriodTimer = Nothing
                    ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                    ' TODO: set large fields to null.
                    RemoveHandler shutdownAfterPeriodTimer.Elapsed, AddressOf shutdownAfterPeriodTimer_Elapsed
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
    End Namespace
End Namespace

#Region "References"
'Check Balora.SysManage:-
#End Region