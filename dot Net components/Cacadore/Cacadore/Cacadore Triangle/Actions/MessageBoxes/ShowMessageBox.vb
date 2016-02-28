'++CacadoreToTectonic is your shepherd here.

Namespace Actions
    'Every action related group is defined under a namespace.
    Namespace MessageBoxes
        <Serializable()> Public Class ShowMessageBox
            Inherits Action
            Implements IDisposable

#Region "Fields"
            Private _inputsValue As ShowMessageBoxInputs
            Private _outputsValue As ShowMessageBoxOutputs
            Private _errorsValue As ShowMessageBoxErrors
#End Region

#Region "Constructor"
            Public Sub New()
                _inputsValue = New ShowMessageBoxInputs("")
                _outputsValue = New ShowMessageBoxOutputs
                _errorsValue = New ShowMessageBoxErrors
            End Sub
#End Region

#Region "Properties"
            Public Overrides ReadOnly Property ActionType As IAction.ActionTypeEnum
                Get
                    Return IAction.ActionTypeEnum.ShowMessageBox
                End Get
            End Property

            Public Overrides Property Inputs As IAction.IInputs
                Get
                    Return _inputsValue
                End Get
                Set(value As IAction.IInputs)
                    _inputsValue = CType(value, ShowMessageBoxInputs)
                End Set
            End Property

            Public Overrides Property Outputs As IAction.IOutputs
                Get
                    Return _outputsValue
                End Get
                Set(value As IAction.IOutputs)
                    _outputsValue = CType(value, ShowMessageBoxOutputs)
                End Set
            End Property

            Public Overrides Property Errors As IAction.IErrors
                Get
                    Return _errorsValue
                End Get
                Set(value As IAction.IErrors)
                    _errorsValue = CType(value, ShowMessageBoxErrors)
                End Set
            End Property
#End Region

            'How to:-'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'TODO : contains the logic of execution for any new action.
            'Set Input Property here
            'Set Output Property @ return
            'Set Result Property @ return
            'And after execution, you will have a result with input and output
            'And @ finish return "ITaskState.ExecutionState" that indicates the result
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Public Overrides Function Execute(Optional retryTimes As Integer = 1) As ITaskState.ExecutionState
                If Me.Skip Then Return ITaskState.ExecutionState.Canceled
                Dim _title As String = vbNullString
                Dim _message As String = vbNullString

                Dim startingResult As ITaskState.ExecutionState
                Dim _inputs As ShowMessageBoxInputs = CType(Inputs, ShowMessageBoxInputs)

                Dim _messageType As MsgBoxStyle = _inputs.MessageBoxType

                _message = Commands.GetMemberValue(TaskId, _inputs.Message)
                If _message = vbNullString Then _message = If(IsNothing(_inputs), "", _inputs.Message)

                If Not IsNothing(_inputs.Title) Then _title = Commands.GetMemberValue(TaskId, _inputs.Title)
                If _title = vbNullString Then _title = If(IsNothing(_inputs), "", _inputs.Title)
                If _title = vbNullString Then _title = "Message Box action of task " & Chr(34) & ParentTask.Properties.TaskName & Chr(34)


                Dim _messageBoxIcon As Windows.Forms.MessageBoxIcon = Windows.Forms.MessageBoxIcon.Information
                Select Case _messageType
                    Case MsgBoxStyle.Exclamation
                        _messageBoxIcon = Windows.Forms.MessageBoxIcon.Exclamation
                    Case MsgBoxStyle.Critical
                        _messageBoxIcon = Windows.Forms.MessageBoxIcon.Error
                    Case MsgBoxStyle.Information
                        _messageBoxIcon = Windows.Forms.MessageBoxIcon.Information
                End Select

                For i = 1 To retryTimes
                    Global.System.Windows.Forms.MessageBox.Show(_message, _title, Windows.Forms.MessageBoxButtons.OK, _messageBoxIcon)
                Next

                'Dim _outputs As ShowMessageBoxOutputs = CType(Outputs, ShowMessageBoxOutputs)
                'SetOutput(_outputs.SaveToMember, "new value here")

                If startingResult = ITaskState.ExecutionState.Failed Then Return ITaskState.ExecutionState.Failed
                WaitForExit()
                Me.RaiseCompletedEvent(Me, New EventArgs)

                Return ITaskState.ExecutionState.Done
            End Function

            Protected Overrides Sub WaitForExit(Optional ByRef state As Object = Nothing)
                If Not IsNothing(Me.ParentTask) AndAlso Not Me.ParentTask.WaitBetweenActions Then Exit Sub
                'So WaitBetweenActions is true or Me.ParentTask is nothing.
                Dim _inputs As ShowMessageBoxInputs = CType(Me.Inputs, ShowMessageBoxInputs)
                If Not IsNothing(_inputs) Then
                    LogWait(_inputs.WaitForExitPeriod)
                    If _inputs.WaitForExitPeriod > 0 Then
                        Threading.Thread.Sleep(CInt(_inputs.WaitForExitPeriod))
                    Else
                        'This option is hidden from user.
                        'WaitAppExitCheckEdit is hidden.
                    End If
                End If
            End Sub

            Public Overrides Sub SetPossibleExceptions()
                'لو عايز تطبقها في المستقبل
                '*.Get all possible exception by checking the new action calling points.
                '       (calling points are places where you really execute or prepare to execute)

                '*.Overrides Sub SetPossibleExceptions() in the new action and fill ActionExceptionsManager.AddPossibleException with all possible exceptions.

                '*.In the new action, make a call to Me.ExceptionsManager.HandleException(ex)
                '	 just like we did in StartApplication.GetErrors method and a call to 
                '              Me.ParentTask.Errors.ListError(Me.ActionId, ex.Message)
            End Sub

            Public Overrides Function GetFullDescription() As String
                Dim [action] = Me
                Dim inputs As ShowMessageBoxInputs = CType(Me.Inputs, ShowMessageBoxInputs)
                Dim message As String = Strings.Left(inputs.Message, 30)
                If inputs.Message.Length > 30 Then message = message & "... "

                Dim listItem As String
                listItem = "Show Message Box" & " Action,"
                listItem = listItem & " Message: " & message

                Dim formatedComments As String = GetFormattedComment()
                If formatedComments <> vbNullString Then listItem = listItem & ", Comments: " & formatedComments
                Return listItem
            End Function

            Public Overrides Function Clone() As Action
                Dim actionMemStream As New Global.System.IO.MemoryStream
                Balora.Serializer.BinarySerializer.SerializeObject(Me, actionMemStream)
                Dim duplicatedAction As ShowMessageBox = CType(Balora.Serializer.BinarySerializer.DeserializeObject(actionMemStream.ToArray), ShowMessageBox)
                Return duplicatedAction
            End Function

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

#Region "Inner Classes"
            <Serializable()>
            Public Class ShowMessageBoxInputs
                Implements IAction.IInputs

#Region "Constructors"
                Public Sub New(msg As String)
                    Message = msg
                End Sub
#End Region

#Region "Properties"
                Property Title As String
                Property Message As String
                Property MessageBoxType As MsgBoxStyle
                Property WaitForExitPeriod As Decimal = 1
#End Region
            End Class
#End Region

#Region "Output Class"
            <Serializable()>
            Public Class ShowMessageBoxOutputs
                Implements IAction.IOutputs
#Region "Properties"
                Property Output As New List(Of String)
#End Region
            End Class
#End Region

#Region "Errors Class"
            <Serializable()>
            Public Class ShowMessageBoxErrors
                Implements IAction.IErrors
#Region "Properties"
                Property ErrorList As New List(Of String)
#End Region
            End Class
#End Region
        End Class
    End Namespace
End Namespace