'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   ExitProcess Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
'Check #Region "How to inherit a new action class" in Action class.
'Never inherit a new action class without of using the instructions
'In the mentioned #Region.
Namespace Actions
    Namespace System
        <Serializable()> Public Class ExitProcess
            Inherits StartApplication

#Region "Fields"
            Private _inputsValue As ExitProcessInputs
            Private _outputsValue As ExitProcessOutputs
            Private _errosValue As ExitProcessErrors
#End Region

#Region "Constructor"
            Public Sub New()
                'Just to prevent calling base class constructor.
                _errosValue = New ExitProcessErrors
                SetPossibleExceptions()
            End Sub
#End Region
#Region "Overidable Action Properties"
            Public Overrides ReadOnly Property ActionType As IAction.ActionTypeEnum
                Get
                    Return IAction.ActionTypeEnum.ExitProcess
                End Get
            End Property

            Public Overrides Property Errors As IAction.IErrors
                Get
                    Return _errosValue
                End Get
                Set(value As IAction.IErrors)
                    _errosValue = CType(value, ExitProcessErrors)
                End Set
            End Property

            Public Overrides Property Inputs As IAction.IInputs
                Get
                    Return _inputsValue
                End Get
                Set(value As IAction.IInputs)
                    _inputsValue = CType(value, ExitProcessInputs)
                End Set
            End Property

            Public Overrides Property Outputs As IAction.IOutputs
                Get
                    Return _outputsValue
                End Get
                Set(value As IAction.IOutputs)
                    _outputsValue = CType(value, ExitProcessOutputs)
                End Set
            End Property

            'Mark as non-serializable so when serialize for clone/duplicate it.
            <NonSerialized()> Private processInfoValue As New ProcessStartInfo
            Public Property ProcessInfo() As ProcessStartInfo
                Get
                    Return processInfoValue
                End Get
                Set(ByVal value As ProcessStartInfo)
                    processInfoValue = value
                End Set
            End Property

            Property ProcessName As String
#End Region

#Region "Methods"
            'How to:-'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Set Input Property here
            'Set Output Property @ return
            'Set Result Property @ return
            'And after execution, you will have a result with input and output
            'And @ finish return "ITaskState.ExecutionState" that indicates the result
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Input and Output are optional because some tasks are parameter less.
            Public Overrides Function Execute(Optional retryTimes As Integer = 1) As ITaskState.ExecutionState
                If Me.Skip Then Return ITaskState.ExecutionState.Canceled
                ProcessInfo = SetProcessInfo()
                Dim endingResult As ITaskState.ExecutionState = EndProcess()
                'Dim _outputs As ExitProcessOutputs = CType(Outputs, ExitProcessOutputs)
                'SetOutput(_outputs.SaveToMember, "new value here")
                If endingResult = ITaskState.ExecutionState.Failed Then Return ITaskState.ExecutionState.Failed
                Me.RaiseCompletedEvent(Me, New EventArgs)
                MyProcess.Dispose()
                MyProcess = Nothing
                Return ITaskState.ExecutionState.Done
            End Function

            Public Shared Function GetRunningProcessesNames() As List(Of String)
                Dim processesList As New List(Of String)
                For Each _process In Process.GetProcesses
                    processesList.Add(_process.ProcessName)
                Next
                Return processesList
            End Function

            Private Function EndProcess() As ITaskState.ExecutionState
                Try

                    Dim processFound As Boolean = False
                    For Each _process In Process.GetProcessesByName(ProcessName)
                        processFound = True
                        MyProcess = _process
                        MyProcess.CloseMainWindow()
                        If Balora.Util.IsProcessRunning(ProcessName) Then MyProcess.Kill()
                        WaitForExit()
                        MyProcess.Dispose()
                        MyProcess = Nothing
                    Next
                    'Causes error alert to be displayed because it thorws exception,
                    'And it get thorwed because MyProcess == Nothing above.
                    'If processFound Then
                    '    If MyProcess.HasExited Then
                    '        OnProcessExited(Me, New EventArgs)
                    '    End If
                    'End If
                    Return ITaskState.ExecutionState.Done
                Catch ex As ObjectDisposedException
                    GetErrors(ex)
                    Return ITaskState.ExecutionState.Failed
                Catch ex As InvalidOperationException
                    GetErrors(ex)
                    Return ITaskState.ExecutionState.Failed
                Catch ex As Global.System.ComponentModel.Win32Exception
                    GetErrors(ex)
                    Return ITaskState.ExecutionState.Failed
                End Try
            End Function

            Private Function SetProcessInfo() As ProcessStartInfo
                Dim _inputs As ExitProcessInputs = CType(Me.Inputs, ExitProcessInputs)
                If IsNothing(_inputs) Then Return New ProcessStartInfo
                
                _processName = Commands.GetMemberValue(TaskId, _inputs.ProcessName)
                If _processName = vbNullString Then _processName = If(IsNothing(_inputs), "", _inputs.ProcessName)

                'Prepare process info.
                Dim _processInfo As New ProcessStartInfo(_processName)
                ProcessName = _processName
                Return _processInfo
            End Function

            Protected Overrides Sub WaitForExit(Optional ByRef state As Object = Nothing)
                'Avoids a deadlock condition by calling p.StandardOutput.ReadToEnd before p.WaitForExit. A deadlock condition can
                'result if the parent process calls p.WaitForExit before p.StandardOutput.ReadToEnd and the child process writes enough
                'text to fill the redirected stream. The parent process would wait indefinitely for the child process to exit. The child
                'process would wait indefinitely for the parent to read from the full StandardOutput stream.
                'http://msdn.microsoft.com/en-us/library/system.diagnostics.process.standardoutput.aspx

                'Exit sub if ParentTask isn't nothing AndAlso WaitBetweenActions <> False
                If Not IsNothing(Me.ParentTask) AndAlso Not Me.ParentTask.WaitBetweenActions Then Exit Sub
                'So WaitBetweenActions is true or Me.ParentTask is nothing.
                Dim _inputs As ExitProcessInputs = CType(Me.Inputs, ExitProcessInputs)
                If Not IsNothing(_inputs) Then
                    If _inputs.WaitForExitPeriod > 0 Then
                        MyProcess.WaitForExit(CInt(_inputs.WaitForExitPeriod))
                    Else
                        MyProcess.WaitForExit()
                    End If
                End If
            End Sub

            Private Shadows Sub SetPossibleExceptions()
                ExceptionsManager.AddPossibleException(GetType(Global.System.ComponentModel.Win32Exception),
                                                       "The associated process could not be terminated")
                ExceptionsManager.AddPossibleException(GetType(NotSupportedException),
                                                       "You are attempting to call Kill for a process that is running on a remote computer")
                ExceptionsManager.AddPossibleException(GetType(InvalidOperationException),
                                                       "The process has already exited")
            End Sub

            Public Overrides Function GetFullDescription() As String
                Dim [action] = Me
                Dim inputs As ExitProcessInputs = CType(Me.Inputs, ExitProcessInputs)
                Dim appName As String = inputs.ProcessName

                Dim listItem As String
                listItem = "Exit Process" & " Action,"
                listItem = listItem & " Process Name: " & appName

                Dim formatedComments As String = GetFormattedComment()
                If formatedComments <> vbNullString Then listItem = listItem & ", Comments: " & formatedComments
                Return listItem
            End Function

            Public Overrides Function Clone() As Action
                Dim actionMemStream As New Global.System.IO.MemoryStream
                Balora.Serializer.BinarySerializer.SerializeObject(Me, actionMemStream)
                Dim duplicatedAction As ExitProcess = CType(Balora.Serializer.BinarySerializer.DeserializeObject(actionMemStream.ToArray), ExitProcess)
                duplicatedAction.MyProcess = MyProcess
                duplicatedAction.ProcessInfo = ProcessInfo
                Return duplicatedAction
            End Function
#End Region

#Region "IDisposable Support"
            Private disposedValue As Boolean ' To detect redundant calls
            ' IDisposable
            Protected Overrides Sub Dispose(disposing As Boolean)
                If Not Me.disposedValue Then
                    If disposing Then
                        MyProcess.Dispose()
                    End If

                    ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                    ' TODO: set large fields to null.
                    MyProcess = Nothing
                    _inputsValue = Nothing
                    _outputsValue = Nothing
                    _errosValue = Nothing
                End If
                Me.disposedValue = True
            End Sub

            ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
            Protected Overrides Sub Finalize()
                ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
                Dispose(False)
                MyBase.Finalize()
            End Sub

            ' This code added by Visual Basic to correctly implement the disposable pattern.
            Public Overloads Sub Dispose()
                ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
                Dispose(True)
                GC.SuppressFinalize(Me)
            End Sub
#End Region

#Region "Inner Classes"
            ''' <summary>
            ''' Holds all required inputs for this action
            ''' and constructor forces to define the mandatory
            ''' inputs.
            ''' </summary>
            <Serializable()>
            Public Class ExitProcessInputs
                Implements IAction.IInputs

#Region "Constructors"
                Public Sub New(Name As String)
                    ProcessName = Name
                End Sub
#End Region

#Region "Props"
                Property ProcessName As String
                Property WaitForExitPeriod As Decimal = 1
#End Region
            End Class

            <Serializable()>
            Public Class ExitProcessOutputs
                Implements IAction.IOutputs
                Property Output As New List(Of String)
            End Class

            <Serializable()>
            Public Class ExitProcessErrors
                Implements IAction.IErrors
                Property ErrorList As New List(Of String)
            End Class
#End Region
        End Class
    End Namespace
End Namespace

#Region "References"
'Process.Kill Method 
'	http://msdn.microsoft.com/en-us/library/system.diagnostics.process.kill%28v=vs.100%29.aspx

'Process.CloseMainWindow Method 
'	http://msdn.microsoft.com/en-us/library/system.diagnostics.process.closemainwindow.aspx

'How do I kill a process using Vb.NET or C#?
'	http://stackoverflow.com/questions/116090/how-do-i-kill-a-process-using-vb-net-or-c
#End Region