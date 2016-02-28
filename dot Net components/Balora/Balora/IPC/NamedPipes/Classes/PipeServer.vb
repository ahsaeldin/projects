'Based on this example
'http://www.eggheadcafe.com/tutorials/aspnet/478ca475-2fd9-45d6-9090-5acdaeb00162/make-your-apps-talk-to-each-other-asynchronous-named-pipes-library.aspx

Imports System.IO.Pipes
Imports System.Threading

Namespace Ipc.NamedPipes
    '/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \
    '                                                   Server Class
    '\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /
    Friend Class PipeServer
        Inherits NamedPipesBase
        Implements IPipeServer

#Region "Fields"
        Private _connections As List(Of PipeStream)
#Region "Check Connection Timer"
        Private _timerCallBack As TimerCallback = AddressOf CheckConnections
        Private _connectionCheckerTimer As New Timer(_timerCallBack, "tick...", 5000, 1000)
#End Region
#End Region

#Region "Constructors"
        Public Sub New(pipeName As String, pipingModeOption As PipeOptions)
            MyBase.New(pipeName, pipingModeOption)
            Me._connections = New List(Of PipeStream)
            Listen()
        End Sub
#End Region

#Region "Methods"
        ''' <summary>
        ''' Checks the connections in a separate thread
        ''' and remove disconnected clients from _connections
        ''' </summary>
        ''' <param name="stateInfo">The state info.</param>
        Public Sub CheckConnections(stateInfo As Object)
            _checkConnectionsExtracted()
        End Sub

        ''' <summary>
        ''' Listens for incoming connections...
        ''' </summary><returns></returns>
        Public Function Listen() As Boolean Implements IPipeServer.Listen
            Return _listen()
        End Function

        ''' <summary>
        ''' Connected callback of "BeginWaitForConnection"
        ''' </summary>
        ''' <param name="result">The result.</param>
        Private Sub ClientConnected(result As IAsyncResult) Implements IPipeServer.ClientConnected
            _clientConnected(result)
        End Sub

        ''' <summary>
        ''' Ends the read.
        ''' </summary>
        ''' <param name="result">The result.</param>
        Public Sub EndRead(result As IAsyncResult) Implements IPipeServer.EndRead
            _endRead(result)
        End Sub

        ''' <summary>
        ''' Sends the message.
        ''' </summary>
        ''' <param name="message">The message.</param><returns></returns>
        Public Overrides Function SendMessage(message As Byte()) As Boolean Implements IPipeServer.SendMessage
            Return _sendMessage(message)
        End Function

        ''' <summary>
        ''' Disconnects this instance.
        ''' </summary><returns></returns>
        Public Overrides Function Disconnect() As Boolean Implements IPipeServer.Disconnect
            Try
                Return _disconnect()
            Catch ex As Exception
                _privateRaiseErrorUp("", ex, True)
                Return False
            End Try
        End Function

#Region "Privates"
        Private Function _listen() As Boolean
            Select Case Me.PipingMode
                Case PipeOptions.None
                    'ToDo: Add synchronous named pipe code here. "Synchronous isn't supported for now."
                Case PipeOptions.Asynchronous
                    Try
                        _invokeBeginWaitForConnection(_server)
                        Return True
                    Catch ex As Exception
                        _privateRaiseErrorUp("", ex, True)
                        Return False
                    End Try
            End Select
            Return False
        End Function

        Private Sub _clientConnected(ByVal result As IAsyncResult)
            _server = TryCast(result.AsyncState, NamedPipeServerStream)
            _invokeEndWaitForConnection(result)
            If _isConnected() Then
                _fireOnConnectionWork()
            End If
            Dim serverObj As NamedPipeServerStream = _returnNewServerObject()
            Try
                _invokeBeginWaitForConnection(serverObj)
            Catch ex As Exception
                _privateRaiseErrorUp("", ex, True)
            End Try
        End Sub

        Private Sub _endRead(ByVal result As IAsyncResult)
            Dim pipeObject As PipeStream = _server
            Try
                Dim length As Integer = _getLengthByInvokeingEndRead(result, pipeObject)
                Dim asyncState As Byte() = DirectCast(result.AsyncState, Byte())
                If length > 0 Then
                    _invokeOnMessageReceivedWork(length, asyncState)
                End If
                SyncLock _connections
                    If _isConnected() Then
                        _invokeBeginRead(asyncState)
                    End If
                End SyncLock
            Catch ex As Exception
                _privateRaiseErrorUp("", ex, True)
            End Try
        End Sub

        Private Function _sendMessage(ByVal message As Byte()) As Boolean
            Dim list As List(Of PipeStream) = Nothing
            Dim flag As Boolean = False
            Dim sendingResultsList As New List(Of Boolean)
            SyncLock Me._connections
                _sendMessageToAllClients(message, list, flag, sendingResultsList)
                'Remove disconnected clients from _connections list.
                If list IsNot Nothing Then
                    For Each connection As PipeStream In list
                        _connections.Remove(connection)
                    Next
                End If
                If Not _isConnected() Then
                    EnqueueMessage(message)
                End If
            End SyncLock
            'Check for results, if all messages sent then true, otherwise true.
            For Each Result In sendingResultsList
                If Not Result Then
                    Return False
                End If
            Next
            Return True
        End Function

        Private Sub _checkConnectionsExtracted()
            SyncLock _connections
                For Each conn As NamedPipeServerStream In _connections
                    If Not conn.IsConnected Then
                        _connections.Remove(conn)
                    End If
                    If _connections.Count = 0 Then Exit For 'Exit if no connections to avoid error if "Next" called.
                Next
            End SyncLock
        End Sub

        Private Sub _invokeBeginWaitForConnection(_serverObj As Object)
            Dim _serverObject As NamedPipeServerStream = CType(_serverObj, NamedPipeServerStream)
            _serverObject.BeginWaitForConnection(New AsyncCallback(AddressOf Me.ClientConnected), _serverObject)
        End Sub

        Private Sub _invokeBeginRead(ByVal buffer As Byte())
            _server.BeginRead(buffer, 0, BufferLength, New AsyncCallback(AddressOf Me.EndRead), buffer)
        End Sub

        Private Function _isConnected() As Boolean
            Return _server.IsConnected
        End Function

        Private Sub _fireOnConnectionWork()
            OnConnection()
            SyncLock _connections
                _connections.Add(_server)
            End SyncLock
            'Server must be @ listen mode in order to send queued messages.
            SendQueuedMessages()
            Dim buffer As Byte() = New Byte(BufferLength - 1) {}
            _invokeBeginRead(buffer)
        End Sub

        Private Sub _invokeEndWaitForConnection(ByVal result As IAsyncResult)
            _server.EndWaitForConnection(result)
        End Sub

        Private Function _returnNewServerObject() As NamedPipeServerStream
            Dim serverObj As New NamedPipeServerStream(MyBase.PipeName, PipeDirection.InOut, -1, PipeTransmissionMode.Message, PipeOptions.Asynchronous)
            Return serverObj
        End Function

        Private Sub _sendMessageToAllClients(ByVal message As Byte(), ByRef list As List(Of PipeStream), ByVal flag As Boolean, ByVal sendingResultsList As List(Of Boolean))
            For Each connection As PipeStream In Me._connections
                Try
                    flag = Not connection.IsConnected
                    If Not flag Then
                        If _isConnected() Then
                            'The magic line that sends to all clients.
                            Me._server = CType(connection, NamedPipeServerStream)
                            'Send message and add result to a list
                            sendingResultsList.Add(DoSendMessage(message))
                        End If
                    End If
                Catch
                    flag = True
                End Try
                'Make a list of client's that are disconnected now...
                'in order to remove them from _connections list.
                If flag Then
                    connection.Close()
                    If list Is Nothing Then
                        list = New List(Of PipeStream)()
                    End If
                    list.Add(connection)

                End If
            Next
        End Sub

        Private Sub _invokeOnMessageReceivedWork(ByVal length As Integer, ByVal asyncState As Byte())
            Dim destinationArray As Byte() = New Byte(length - 1) {}
            Array.Copy(asyncState, 0, destinationArray, 0, length)
            OnMessageReceived(New MessageEveArgs(destinationArray))
        End Sub

        Private Shared Function _getLengthByInvokeingEndRead(ByVal result As IAsyncResult, ByVal pipeObject As PipeStream) As Integer
            Dim length As Integer = pipeObject.EndRead(result)
            Return length
        End Function

        Private Function _disconnect() As Boolean
            SyncLock Me._connections
                For Each connection As PipeStream In Me._connections
                    Try
                        connection.Close()
                    Catch
                    End Try
                    Me._connections.Clear()
                Next
            End SyncLock
            MyBase.Disconnect()
            Return True
        End Function
#End Region

#End Region
    End Class
End Namespace