'Based on this example
'http://www.eggheadcafe.com/tutorials/aspnet/478ca475-2fd9-45d6-9090-5acdaeb00162/make-your-apps-talk-to-each-other-asynchronous-named-pipes-library.aspx

Imports System.IO.Pipes
Imports System.Threading

Namespace Ipc.NamedPipes
    '/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \ 
    '                                                   Client Class                                                                 
    '\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /
    Friend Class PipeClient
        Inherits NamedPipesBase
        Implements IPipeClient

#Region "Fields"
        Private _connectGate As ManualResetEvent
        Private ReadOnly _instanceLock As Object
        Private _connectingTimeout As Integer = 5000
        Private _connectionCheckerTimer As Threading.Timer
        Private _timerCallBack As TimerCallback = AddressOf TryConnect
#End Region

#Region "Constructors"
        Public Sub New(pipeName As String, pipingModeOption As PipeOptions)
            MyBase.New(pipeName, pipingModeOption)
            Me._instanceLock = New Object()
            Me._connectGate = New ManualResetEvent(False)
        End Sub
#End Region

#Region "Properties"
        ''' <summary>
        ''' Gets or sets the connecting time out.
        ''' </summary>
        ''' <value>
        ''' The connecting time out.
        ''' </value>
        Public Property ConnectingTimeout() As Integer Implements IPipeClient.ConnectionTimeout
            Get
                Return _connectingTimeout
            End Get
            Set(ByVal value As Integer)
                _connectingTimeout = value
            End Set
        End Property
#End Region

#Region "Methods"
        ''' <summary>
        ''' Connects this instance.
        ''' </summary>
        ''' <remarks>
        ''' Because CPUD usage decrease @ calling "_client.Connect(ConnectingTimeOut)"...
        ''' the default is to timeout after 5 seconds and reconnect after 10 from start 
        ''' connecting (or 5 seconds after timeout) and this will decrease the CPU usage 
        ''' for 5 seconds until get connected and this is my initial approach that may 
        ''' change in the future.
        ''' </remarks>
        Public Sub Connect() Implements IPipeClient.Connect
            _connect()
        End Sub

        ''' <summary>
        ''' Tries the connect. 
        ''' </summary>
        ''' <param name="stateInfo">The state info.</param><returns></returns>
        Public Function TryConnect(stateInfo As Object) As Boolean Implements IPipeClient.TryConnect
            Return _tryConnect(stateInfo)
        End Function

        ''' <summary>
        ''' Sends the message.
        ''' </summary>
        ''' <param name="message">The message.</param><returns></returns>
        Public Overrides Function SendMessage(message As Byte()) As Boolean Implements IPipeClient.SendMessage
            Return _sendMessage(message)
        End Function

        ''' <summary>
        ''' Ends the read.
        ''' </summary>
        ''' <param name="result">The result.</param>
        Public Sub EndRead(result As IAsyncResult) Implements IPipeClient.EndRead
            _endRead(result)
        End Sub

#Region "Privates"
        ''' <summary>
        ''' _connects this instance. (conceal from reflectors)
        ''' </summary>
        Private Sub _connect()
            'reset connection gate state to close now...
            _connectGate.Reset()
            _connectionCheckerTimer = New Threading.Timer(_timerCallBack, "Connecting...", 0, 10000)
        End Sub

        ''' <summary>
        ''' _tries the connect. (conceal from reflectors)
        ''' </summary>
        ''' <param name="stateInfo">The state info.</param><returns></returns>
        Private Function _tryConnect(stateInfo As Object) As Boolean
            SyncLock _instanceLock
                Select Case PipingMode
                    Case PipeOptions.None
                        'ToDo: Add synchronous named pipe code here. "Synchronous isn't supported for now."
                    Case PipeOptions.Asynchronous
                        Try
                            If Not _isConnected() Then
                                'Connecting to a pipe that doesn't exist... 
                                'will raise "operation has timed out" exception.
                                _client.Connect(ConnectingTimeout)
                                If _isConnected() Then
                                    _fireAfterConnectionWork()
                                End If
                            End If
                            Return True
                        Catch ex As TimeoutException
                            'Connecting to a pipe that doesn't exist... 
                            'will raise "operation has timed out" exception.
                            _privateRaiseErrorUp("Connection attempt has timed out.", ex, False, True)
                            Return False
                        Catch ex As ArgumentOutOfRangeException
                            _privateRaiseErrorUp("", ex, True)
                        End Try
                End Select
                Return False
            End SyncLock
        End Function

        Private Function _sendMessage(ByVal message As Byte()) As Boolean
            Try
                'Close connection gate before sending a message.
                If Me._connectGate.WaitOne(100) Then
                    SyncLock Me._instanceLock
                        If _isConnected() Then
                            Return _doSendMessage(message)
                        Else
                            _enqueMessage(message)
                            _connect()
                        End If
                    End SyncLock
                Else
                    _enqueMessage(message)
                End If
                Return False
            Catch ex As Exception
                _privateRaiseErrorUp("", ex, True)
                Return False
            End Try
        End Function

        Private Sub _endRead(ByVal result As IAsyncResult)
            Dim pipeObject As PipeStream = _client
            Try
                Dim length As Integer = _invokeEndRead(result, pipeObject)
                Dim asyncState As Byte() = DirectCast(result.AsyncState, Byte())
                If length > 0 Then
                    _fireOnMessageReceived(length, asyncState)
                End If
                SyncLock _instanceLock
                    If _isConnected() Then
                        _invokeBeginRead(pipeObject, asyncState)
                    End If
                End SyncLock
            Catch ex As Exception
                _privateRaiseErrorUp("", ex, True)
            End Try
        End Sub

        Private Function _doSendMessage(message As Byte()) As Boolean
            Return DoSendMessage(message)
        End Function

        Private Sub _enqueMessage(message As Byte())
            EnqueueMessage(message)
        End Sub

        ''' <summary>
        ''' _ises the connected. (conceal from reflectors)
        ''' </summary><returns></returns>
        Private Function _isConnected() As Boolean
            Return _client.IsConnected
        End Function

        ''' <summary>
        ''' _fires the after connection work. (conceal from reflectors)
        ''' </summary>
        Private Sub _fireAfterConnectionWork()
            OnConnection()
            _connectionCheckerTimer.Dispose()
            _client.ReadMode = PipeTransmissionMode.Message
            Dim buffer As Byte() = New Byte(BufferLength - 1) {}
            _invokeBeginRead(_client, buffer)
            'Now open connection gate if it has been closed...
            _connectGate.[Set]()
            SendQueuedMessages()
        End Sub

        Private Sub _fireOnMessageReceived(ByVal length As Integer, ByVal asyncState As Byte())
            Dim destinationArray As Byte() = New Byte(length - 1) {}
            Array.Copy(asyncState, 0, destinationArray, 0, length)
            OnMessageReceived(New MessageEveArgs(destinationArray))
        End Sub

        Private Sub _invokeBeginRead(ByVal pipeObject As PipeStream, ByVal asyncState As Byte())
            pipeObject.BeginRead(asyncState, 0, BufferLength, New AsyncCallback(AddressOf Me.EndRead), asyncState)
        End Sub

        Private Function _invokeEndRead(ByVal result As IAsyncResult, ByVal pipeObject As PipeStream) As Integer
            Return pipeObject.EndRead(result)
        End Function
#End Region
#End Region
    End Class
End Namespace