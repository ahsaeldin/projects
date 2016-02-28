'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   StartApplication Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
'Check #Region "How to inherit a new action class" in Action class.
'Never inherit a new action class without of using the instructions
'In the mentioned #Region.
Namespace Actions

    Namespace System
        <Serializable()> Public Class StartApplication
            Inherits Action
            Implements IDisposable

#Region "Fields"
            Private _verbValue As String
            'Mark as non-serializable so when serialize for clone/duplicate it.
            <NonSerialized()> Private _processValue As Process
            Private _inputsValue As StartApplicationInputs
            Private _outputsValue As StartApplicationOutputs
            Private _errorsValue As StartApplicationErrors
#End Region

#Region "Constructors"
            Public Sub New()
                MyProcess = New Process()
                _errorsValue = New StartApplicationErrors
                _outputsValue = New StartApplicationOutputs
                SetPossibleExceptions()
            End Sub
#End Region

#Region "Properties"

#Region "Overidable Action Properties"
            Public Overrides ReadOnly Property ActionType As IAction.ActionTypeEnum
                Get
                    Return IAction.ActionTypeEnum.StartApplication
                End Get
            End Property

            Public Overrides Property Inputs As IAction.IInputs
                Get
                    Return _inputsValue
                End Get
                Set(value As IAction.IInputs)
                    _inputsValue = CType(value, StartApplicationInputs)
                    'To enable getting all verbs before execution.
                    MyProcess.StartInfo = SetProcessInfo()
                End Set
            End Property

            Public Overrides Property Outputs As IAction.IOutputs
                Get
                    Return _outputsValue
                End Get
                Set(value As IAction.IOutputs)
                    _outputsValue = CType(value, StartApplicationOutputs)
                End Set
            End Property

            Public Overrides Property Errors As IAction.IErrors
                Get
                    Return _errorsValue
                End Get
                Set(value As IAction.IErrors)
                    _errorsValue = CType(value, StartApplicationErrors)
                End Set
            End Property
#End Region

#Region "New Added Properties"
            Public Property MyProcess As Process
                Get
                    If IsNothing(_processValue) Then _processValue = New Process()
                    Return _processValue
                End Get
                Set(ByVal value As Process)
                    _processValue = value
                End Set
            End Property

            Public ReadOnly Property Handle() As IntPtr
                Get
                    Return MyProcess.Handle
                End Get
            End Property

            Public Property Verb() As String
                Get
                    Me.Dispose()
                    Return MyProcess.StartInfo.Verb
                End Get
                Set(value As String)
                    MyProcess.StartInfo.Verb = value
                End Set
            End Property

            Public ReadOnly Property Verbs() As List(Of String)
                Get
                    Return MyProcess.StartInfo.Verbs.ToList
                End Get
            End Property
#End Region
#End Region

#Region "Methods"
            'How to:-'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'TODO : contains the logic of execution for any new action.
            'Set Input Property here
            'Set Output Property @ return
            'Set Result Property @ return
            'And after execution, you will have a result with input and output
            'And @ finish return "ITaskState.ExecutionState" that indicates the result
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Input and Output are optional because some tasks are parameter less.
            Public Overrides Function Execute(Optional retryTimes As Integer = 1) As ITaskState.ExecutionState
                'TODO : You must add the next line to any new action.
                If Me.Skip Then Return ITaskState.ExecutionState.Canceled
                ' Inputs = Inputs 'This funny line will translate variables/globals in process name if not yet translated.

                SetEventHandler()
                Dim startingResult As ITaskState.ExecutionState
                For i = 1 To retryTimes
                    startingResult = StartProcess()
                Next

                'Dim _outputs As StartApplicationOutputs = CType(Outputs, StartApplicationOutputs)
                'SetOutput(_outputs.SaveToMember, "new value here")

                If startingResult = ITaskState.ExecutionState.Failed Then Return ITaskState.ExecutionState.Failed
                'BeginRead() 'Memory Leak 'koko 'Uncomment when you develop reading the output to display to the user.
                WaitForExit()
                'TODO : You must add the next line for any new action.
                Me.RaiseCompletedEvent(Me, New EventArgs)
                MyProcess.Dispose()
                MyProcess = Nothing
                Return ITaskState.ExecutionState.Done
            End Function

            Private Function StartProcess() As ITaskState.ExecutionState
                Try
                    MyProcess.Start()
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

            Private Sub CheckIfFileNameIsVariable(ByRef filePath As String)
                Dim _inputs As StartApplicationInputs = CType(Me.Inputs, StartApplicationInputs)
                filePath = Commands.GetMemberValue(TaskId, _inputs.FileName)
                If filePath = vbNullString Then filePath = If(IsNothing(_inputs), "", _inputs.FileName)
            End Sub

            Private Function SetProcessInfo() As ProcessStartInfo
                Dim filePath As String = vbNullString
                Dim _inputs As StartApplicationInputs = CType(Me.Inputs, StartApplicationInputs)
                If IsNothing(_inputs) Then Return New ProcessStartInfo
     
                CheckIfFileNameIsVariable(filePath)

                'This is a fix for when you set a portable firefox as the default browser, then
                'WaitForExit got break if you set WaitBetweenAction to true.
                'Check bug https://conderella.com/wi/client/index.php?issue=488
                CheckIfURL(filePath, _inputs.Arguments)

                'Prepare process info.
                Dim _processInfo As New ProcessStartInfo(filePath)
                _processInfo.Arguments = _inputs.Arguments
                _processInfo.CreateNoWindow = _inputs.CreateNoWindow
                If _inputs.WindowStyle <> -1 Then _processInfo.WindowStyle = _inputs.WindowStyle
                _processInfo.WorkingDirectory = _inputs.WorkingDirectory
                TryDisableShellExecute(filePath, _processInfo)
                Return _processInfo
            End Function

            Private Sub CheckIfURL(ByRef filePath As String, ByRef arguments As String)
                'This is a fix for when you set a portable firefox as the default browser, then
                'WaitForExit got break if you set WaitBetweenAction to true.
                'Check bug https://conderella.com/wi/client/index.php?issue=488&page=1
                Dim regxMatch As New Balora.RegexMatcher
                Dim isUrl = regxMatch.MatchURL(filePath)
                If isUrl Then
                    Dim defaultBrowserPath As String = Balora.PathsHelper.GetDefaultBrowserPath()
                    If arguments = vbNullString Then
                        arguments = filePath
                    Else
                        arguments = filePath & " " & arguments
                    End If
                    filePath = defaultBrowserPath
                End If
            End Sub

            Private Sub TryDisableShellExecute(ByVal filePath As String, ByRef _processInfo As ProcessStartInfo)
                Dim windowStyle = _processInfo.WindowStyle
                Dim fileExtension As String = Global.System.IO.Path.GetExtension(filePath)
                If windowStyle = ProcessWindowStyle.Maximized OrElse windowStyle = ProcessWindowStyle.Minimized Then Exit Sub
                If fileExtension.Contains(".exe") OrElse fileExtension = "" OrElse windowStyle = ProcessWindowStyle.Hidden Then
                    DisableShellExecute(_processInfo)
                End If
            End Sub

            Public Sub DisableShellExecute()
                DisableShellExecute(MyProcess.StartInfo)
            End Sub

            Public Sub DisableShellExecute(ByRef _processInfo As ProcessStartInfo)
                'When the ProcessStartInfo.UseShellExecute property is set to its default value, true, you can start applications and documents in a way that is similar to using
                'the Run dialog box of the Windows Start menu. When ProcessStartInfo.UseShellExecute is false, you can start only executables.
                'http://msdn.microsoft.com/en-us/library/system.diagnostics.process.process.aspx
                _processInfo.UseShellExecute = False
                _processInfo.RedirectStandardOutput = True
                _processInfo.RedirectStandardError = True
            End Sub

            Protected Overrides Sub WaitForExit(Optional ByRef state As Object = Nothing)
                'Avoids a deadlock condition by calling p.StandardOutput.ReadToEnd before p.WaitForExit. A deadlock condition can
                'result if the parent process calls p.WaitForExit before p.StandardOutput.ReadToEnd and the child process writes enough
                'text to fill the redirected stream. The parent process would wait indefinitely for the child process to exit. The child
                'process would wait indefinitely for the parent to read from the full StandardOutput stream.
                'http://msdn.microsoft.com/en-us/library/system.diagnostics.process.standardoutput.aspx
                'Exit sub if ParentTask isn't nothing AndAlso WaitBetweenActions <> False
                If Not IsNothing(Me.ParentTask) AndAlso Not Me.ParentTask.WaitBetweenActions Then Exit Sub
                'So WaitBetweenActions is true or Me.ParentTask is nothing.
                Dim _inputs As StartApplicationInputs = CType(Me.Inputs, StartApplicationInputs)
                If Not IsNothing(_inputs) Then
                    LogWait(_inputs.WaitForExitPeriod)
                    If _inputs.WaitForExitPeriod > 0 Then
                        MyProcess.WaitForExit(CInt(_inputs.WaitForExitPeriod))
                    Else
                        MyProcess.WaitForExit()
                    End If
                End If
            End Sub

            Private Sub BeginRead()
                If MyProcess.StartInfo.UseShellExecute Then Exit Sub
                If MyProcess.StartInfo.RedirectStandardOutput Then
                    If Not IsNothing(MyProcess) Then
                        MyProcess.BeginOutputReadLine()
                    End If
                End If
                If MyProcess.StartInfo.RedirectStandardError Then
                    If Not IsNothing(MyProcess) Then MyProcess.BeginErrorReadLine()
                End If
            End Sub

            Protected Sub SetEventHandler()
                AddHandler MyProcess.Exited, AddressOf Exited
                MyProcess.EnableRaisingEvents = True
                If MyProcess.StartInfo.UseShellExecute Then Exit Sub
                If MyProcess.StartInfo.RedirectStandardOutput Then
                    AddHandler MyProcess.OutputDataReceived, AddressOf OutputDataReceived
                End If
                If MyProcess.StartInfo.RedirectStandardError Then
                    AddHandler MyProcess.ErrorDataReceived, AddressOf ErrorDataReceived
                End If
            End Sub

            Protected Sub GetErrors(ByVal ex As Global.System.Exception)
                _errorsValue.ErrorList.Add(ex.Message)
                Me.ParentTask.Errors.ListError(ActionId, ex.Message)
                Me.ExceptionsManager.HandleException(ex, ParentTask, Me)
            End Sub

            Protected Overridable Sub OnProcessExited(sender As Object, e As EventArgs)
                RaiseEvent Process_Exited(sender, e)
            End Sub

            Public Overrides Sub SetPossibleExceptions()
                ExceptionsManager.AddPossibleException(GetType(Global.System.ComponentModel.Win32Exception),
                                       "There was an error in opening the associated file")
                ExceptionsManager.AddPossibleException(GetType(InvalidOperationException),
                                       "No file name was specified")
            End Sub

            Public Overrides Function GetFullDescription() As String
                Dim [action] = Me
                Dim inputs As StartApplicationInputs = CType(Me.Inputs, StartApplicationInputs)
                Dim appName As String = inputs.FileName

                Dim listItem As String
                listItem = "Start Application" & " Action,"
                listItem = listItem & " File Name: " & appName
                If inputs.WaitForExitPeriod = 0 Then
                    listItem = listItem & ", Delay Settings: Wait For Exit"
                ElseIf inputs.WaitForExitPeriod > 0 Then
                    listItem = listItem & ", Delay Settings: Wait For " & inputs.WaitForExitPeriod / 1000 & " second(s)"
                End If

                Dim formatedComments As String = GetFormattedComment()
                If formatedComments <> vbNullString Then listItem = listItem & ", Comments: " & formatedComments
                Return listItem
            End Function

            Public Overrides Function Clone() As Action
                Dim actionMemStream As New Global.System.IO.MemoryStream
                Balora.Serializer.BinarySerializer.SerializeObject(Me, actionMemStream)
                Dim duplicatedAction As StartApplication = CType(Balora.Serializer.BinarySerializer.DeserializeObject(actionMemStream.ToArray), StartApplication)
                duplicatedAction.MyProcess = MyProcess
                Return duplicatedAction
            End Function


#Region "IDisposable Support"
            Private disposedValue As Boolean ' To detect redundant calls

            ' IDisposable
            Protected Overridable Sub Dispose(disposing As Boolean)
                If Not Me.disposedValue Then
                    If disposing Then
                        MyProcess.Dispose()
                    End If

                    ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                    ' TODO: set large fields to null.
                    MyProcess = Nothing
                    RemoveHandler MyProcess.Exited, AddressOf Exited
                    RemoveHandler MyProcess.OutputDataReceived, AddressOf OutputDataReceived
                    RemoveHandler MyProcess.ErrorDataReceived, AddressOf ErrorDataReceived
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
            Public Sub Dispose() Implements IDisposable.Dispose
                ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
                Dispose(True)
                GC.SuppressFinalize(Me)
            End Sub
#End Region
#End Region

#Region "Events"
            Event Process_Exited As EventHandler
#End Region

#Region "Events Handlers"
            Private Sub Exited(sender As Object,
                               e As EventArgs)
                RaiseEvent Process_Exited(sender, New EventArgs)
            End Sub

            Private Sub OutputDataReceived(sender As Object,
                                           e As Global.System.Diagnostics.DataReceivedEventArgs)
                If Not IsNothing(e.Data) Then _outputsValue.Output.Add(e.Data)
            End Sub

            Private Sub ErrorDataReceived(sender As Object,
                                          e As Global.System.Diagnostics.DataReceivedEventArgs)
                If Not IsNothing(e.Data) Then _errorsValue.ErrorList.Add(e.Data)
            End Sub
#End Region

#Region "Inner Classes"

#Region "Inputs Class"
            ''' <summary>
            ''' Holds all required inputs for this action
            ''' and constructor forces to define the mandatory
            ''' inputs.
            ''' </summary>
            <Serializable()>
            Public Class StartApplicationInputs
                Implements IAction.IInputs

#Region "Constructors"
                Public Sub New(filePath As String)
                    FileName = filePath
                End Sub
#End Region

#Region "Properties"
                Property Verb As String
                Property FileName As String
                Property Arguments As String
                Property CreateNoWindow As Boolean
                Property WorkingDirectory As String
                Property WaitForExitPeriod As Decimal = 1
                Property WaitForInputHandlerPeriod As Integer = 1
                Property WindowStyle As ProcessWindowStyle = ProcessWindowStyle.Normal
#End Region

            End Class
#End Region

#Region "Output Class"
            <Serializable()>
            Public Class StartApplicationOutputs
                Implements IAction.IOutputs
                Property Output As New List(Of String)
            End Class
#End Region

#Region "Errors Class"
            <Serializable()>
            Public Class StartApplicationErrors
                Implements IAction.IErrors
                Property ErrorList As New List(Of String)
            End Class
#End Region
#End Region
        End Class
    End Namespace
End Namespace

#Region "References"
'Process Class:-
'	http://msdn.microsoft.com/en-us/library/system.diagnostics.process.aspx

'   -Process Constructor
'	 http://msdn.microsoft.com/en-us/library/system.diagnostics.process.process.aspx

'   -Process.StandardError Property 
'	 http://msdn.microsoft.com/en-us/library/system.diagnostics.process.standarderror.aspx

'   -Process.StandardInput Property
'	 http://msdn.microsoft.com/en-us/library/system.diagnostics.process.standardinput.aspx

'   -Process.StandardOutput Property
'	 http://msdn.microsoft.com/en-us/library/system.diagnostics.process.standardoutput.aspx

'   -Process.WaitForExit Method
'	http://msdn.microsoft.com/en-us/library/fb4aw7b8.aspx

'   -Process.OutputDataReceived Event
'	http://msdn.microsoft.com/en-us/library/system.diagnostics.process.outputdatareceived.aspx

'   -Process.ErrorDataReceived Event
'	http://msdn.microsoft.com/en-us/library/system.diagnostics.process.errordatareceived.aspx

'   -Process.Exited Event
'	http://msdn.microsoft.com/en-us/library/system.diagnostics.process.exited.aspx

'   -Process.EnableRaisingEvents Property 
'   http://msdn.microsoft.com/en-us/library/system.diagnostics.process.enableraisingevents.aspx

'ProcessStartInfo Class:-
'	http://msdn.microsoft.com/en-us/library/system.diagnostics.processstartinfo.aspx

'   -ProcessStartInfo.Verb Property 
'    http://www.youtube.com/watch?feature=player_embedded&v=vHQU86MuC28#t=11s
'   -ProcessStartInfo.Verbs Property
'    http://msdn.microsoft.com/en-us/library/system.diagnostics.processstartinfo.verbs.aspx

'   -ProcessStartInfo.RedirectStandardError Property 
'	 http://msdn.microsoft.com/en-us/library/system.diagnostics.processstartinfo.redirectstandarderror.aspx

'   -ProcessStartInfo.RedirectStandardInput Property 
'	 http://msdn.microsoft.com/en-us/library/system.diagnostics.processstartinfo.redirectstandardinput.aspx

'   -ProcessStartInfo.RedirectStandardOutput Property 
'	 http://msdn.microsoft.com/en-us/library/system.diagnostics.processstartinfo.redirectstandardoutput.aspx

'   -How to send arguments
'   http://msdn.microsoft.com/en-us/library/system.diagnostics.processstartinfo.arguments.aspx

'Other References:-
'   -Process.Start() Quips
'    https://msmvps.com/blogs/manoj/archive/2004/10/23/16505.aspx

'   -Dos command in c# - RedirectStandardError/Output problem
'	 http://stackoverflow.com/questions/2974946/dos-command-in-c-sharp-redirectstandarderror-output-problem

'   -Redirect Standard Output Efficiently in .NET
'	 http://stackoverflow.com/questions/164736/redirect-standard-output-efficiently-in-net

'   -ProcessInfo and RedirectStandardOutput
'	 http://stackoverflow.com/questions/1145969/processinfo-and-redirectstandardoutput

'   -C# RedirectStandardOutput Example
'	 http://www.dotnetperls.com/redirectstandardoutput
#End Region