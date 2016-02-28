'Based on this example
'http://www.eggheadcafe.com/tutorials/aspnet/478ca475-2fd9-45d6-9090-5acdaeb00162/make-your-apps-talk-to-each-other-asynchronous-named-pipes-library.aspx

Imports System.IO.Pipes


Namespace Ipc.NamedPipes
    '/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \ 
    '                                                   NamedPipesBase Class                                                         '       
    '\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /
    Friend MustInherit Class NamedPipesBase
        Inherits BaloraBase
        Implements INamedPipesBase

#Region "Fields"
        Private _pipeName As String
        Private _pipingModeValue As PipeOptions
        Protected _client As NamedPipeClientStream
        Protected _server As NamedPipeServerStream
        Private ReadOnly _queueLock As New Object
        Private ReadOnly _eventLock As New Object
        Private ReadOnly _instanceLock As New Object
        Private _pendingMessages As Queue(Of Byte())
        Public Shared ReadOnly BufferLength As Integer = 65536
        '    Public Delegate Sub MessageEventHandler(sender As Object, e As MessageEveArgs)
        Private _messageReceived As EventHandler(Of MessageEveArgs)
#End Region

#Region "Constructors"
        ''' <summary>
        ''' Initializes a new instance of the <see cref="NamedPipesBase" /> class.
        ''' </summary>
        ''' <param name="pipeName">Name of the pipe.</param>
        ''' <param name="pipingModeOption">The piping mode option.</param>
        ''' <remarks>
        ''' Based on pipeName, this constructor detect if it is a remote calling.
        ''' Based on pipingModeOption, this constructor creates the suitable objects.
        ''' </remarks>
        Protected Sub New(pipeName As String, pipingModeOption As PipeOptions)
            _new(pipeName, pipingModeOption)
        End Sub

        Private Sub _new(ByVal pipeName As String, ByVal pipingModeOption As PipeOptions)
            Dim serverName As String = vbNullString
            Dim pipeNameOnly As String = vbNullString
            ExtractPipeName(pipeName, serverName, pipeNameOnly)
            PipingMode = pipingModeOption
            Select Case PipingMode
                Case PipeOptions.None 'Synchronous named pipe. "Synchronous isn't supported for now."
                    Select Case Me.GetType.Name
                        Case GetType(PipeServer).Name
                            '_server = New NamedPipeServerStream(pipeNameOnly, PipeDirection.InOut) "Synchronous isn't supported for now."
                        Case GetType(PipeClient).Name
                            '_client = New NamedPipeClientStream(".", pipeNameOnly, PipeDirection.InOut) "Synchronous isn't supported for now."
                    End Select
                Case PipeOptions.Asynchronous 'Asynchronous named pipes.
                    _defineConnectionObject(serverName, pipeNameOnly)
            End Select
        End Sub

#End Region

#Region "Properties"
        ''' <summary>
        ''' Gets or sets the piping mode. (None, asynchronous or write through.)
        ''' </summary>
        ''' <value>
        ''' The piping mode.
        ''' </value>
        Public Property PipingMode() As PipeOptions Implements INamedPipesBase.PipingMode
            Get
                Return _pipingModeValue
            End Get
            Set(ByVal value As PipeOptions)
                _pipingModeValue = value
            End Set
        End Property

        ''' <summary>
        ''' Gets the name of the pipe.
        ''' </summary>
        ''' <value>
        ''' The name of the pipe.
        ''' </value>
        Public ReadOnly Property PipeName() As String Implements INamedPipesBase.PipeName
            Get
                Return _pipeName
            End Get
        End Property

        ''' <summary>
        ''' Gets a value indicating whether this instance is connected.
        ''' </summary>
        ''' <value>
        ''' 
        ''' <c>true</c> if this instance is connected; otherwise, <c>false</c>.
        ''' 
        ''' </value>
        Public ReadOnly Property IsConnected() As Boolean Implements INamedPipesBase.IsConnected
            Get
                Dim _stream As PipeStream = _getStream()
                SyncLock _instanceLock
                    Return _stream.IsConnected
                End SyncLock
            End Get
        End Property
#End Region

#Region "Methods"
        Private Sub ExtractPipeName(ByVal pipeName As String, ByRef serverName As String, ByRef pipeNameOnly As String)
            pipeNameOnly = pipeName
            _pipeName = pipeName
            'Check if a remote call
            If pipeName.Contains("\") Then
                serverName = pipeNameOnly.Substring(pipeNameOnly.IndexOf("\") + 1, pipeNameOnly.LastIndexOf("\") - 1)
                pipeNameOnly = pipeNameOnly.Substring(_pipeName.LastIndexOf("\") + 1)
            Else 'Else a local call
                serverName = "."
            End If
        End Sub

        ''' <summary>
        ''' Sends the message.
        ''' </summary>
        ''' <param name="message">The message.</param><returns></returns>
        Public MustOverride Function SendMessage(message As Byte()) As Boolean Implements INamedPipesBase.SendMessage

        ''' <summary>
        ''' Does the send message work.
        ''' </summary>
        ''' <param name="message">The message.</param><returns></returns>
        Public Function DoSendMessage(message As Byte()) As Boolean Implements INamedPipesBase.DoSendMessage
            Return _doSendMessage(message)
        End Function

        ''' <summary>
        ''' Ends the send message.
        ''' </summary>
        ''' <param name="result">The result.</param>
        Public Sub EndSendMessage(result As IAsyncResult) Implements INamedPipesBase.EndSendMessage
            _endSendMessage(result)
        End Sub

        ''' <summary>
        ''' Raises the <see cref="E:MessageReceived" /> event.
        ''' </summary>
        ''' <param name="args">The <see cref="Balora.IPC.NamedPipes.MessageEveArgs" /> instance containing the event data.</param>
        ''' <remarks>
        ''' Added because derived classes cannot raise base class events, 
        ''' so Client and Server classes will use this method to raise 
        ''' MessageReceived events up.
        ''' </remarks>
        Public Overridable Sub OnMessageReceived(args As MessageEveArgs) Implements INamedPipesBase.OnMessageReceived
            _onMessageReceived(args)
        End Sub

        ''' <summary>
        ''' Called when [connection].
        ''' </summary>
        ''' <remarks>
        ''' Added because derived classes cannot raise base class events, 
        ''' so Client and Server classes will use this method to raise 
        ''' OnConnection events up.
        ''' </remarks>
        Public Overridable Sub OnConnection() Implements INamedPipesBase.OnConnection
            _onConnection()
        End Sub

        ''' <summary>
        ''' Sends the queued messages.
        ''' </summary>
        Protected Sub SendQueuedMessages() Implements INamedPipesBase.SendQueuedMessages
            _sendQueuedMessages()
        End Sub

        ''' <summary>
        ''' Enques the message to send later.
        ''' </summary>
        ''' <param name="message">The message.</param>
        Protected Sub EnqueueMessage(message As Byte()) Implements INamedPipesBase.EnqueueMessage
            _EnqueueMessage(message)
        End Sub

        ''' <summary>
        ''' Disconnects this instance.
        ''' </summary><returns></returns>
        Public Overridable Function Disconnect() As Boolean Implements INamedPipesBase.Disconnect
            Return _disconnect()
        End Function

#Region "Privates"
        Private Sub _defineConnectionObject(ByVal serverName As String, ByVal pipeNameOnly As String)
            Select Case Me.GetType.Name
                Case GetType(PipeServer).Name
                    _server = New NamedPipeServerStream(pipeNameOnly, PipeDirection.InOut, -1, PipeTransmissionMode.Message, PipeOptions.Asynchronous)
                Case GetType(PipeClient).Name
                    _client = New NamedPipeClientStream(serverName, pipeNameOnly, PipeDirection.InOut, PipeOptions.Asynchronous)
            End Select
        End Sub

        Private Sub _invokeBeginWrite(ByVal message As Byte(), ByVal _stream As PipeStream)
            _stream.BeginWrite(message, 0, message.Length, New AsyncCallback(AddressOf EndSendMessage), Nothing)
            _stream.Flush()
        End Sub

        Private Function _getStream() As PipeStream
            Dim _stream As PipeStream = Nothing
            Select Case Me.GetType.Name
                Case GetType(PipeServer).Name
                    _stream = _server
                Case GetType(PipeClient).Name
                    _stream = _client
            End Select
            Return _stream
        End Function

        Private Function _doSendMessage(ByVal message As Byte()) As Boolean
            Dim _stream As PipeStream = _getStream()
            Try
                message = If(message, New Byte(-1) {})
                _invokeBeginWrite(message, _stream)
                Return True
            Catch ex As Exception
                _privateRaiseErrorUp("", ex, True)
                Return False
            End Try
        End Function

        Private Sub _InvokeEndSendMessage(ByVal result As IAsyncResult, ByVal _stream As PipeStream)
            SyncLock Me._instanceLock
                _stream.EndWrite(result)
                _stream.Flush()
            End SyncLock
        End Sub

        Private Sub _endSendMessage(ByVal result As IAsyncResult)
            Dim _stream As PipeStream = _getStream()
            _InvokeEndSendMessage(result, _stream)
        End Sub

        Private Sub _onConnection()
            RaiseEvent Connected(Me, New EventArgs)
        End Sub

        Private Sub _EnqueueMessage(ByVal message As Byte())
            SyncLock _queueLock
                If _pendingMessages Is Nothing Then
                    _pendingMessages = New Queue(Of Byte())()
                End If
                Try
                    _pendingMessages.Enqueue(message)
                Catch ex As Exception
                    _privateRaiseErrorUp("", ex, True)
                End Try
            End SyncLock
        End Sub

        Private Sub _onMessageReceived(ByVal args As MessageEveArgs)
            SyncLock Me._eventLock
                RaiseEvent MessageReceived(Me, args)
            End SyncLock
        End Sub

        Private Sub _sendQueuedMessages()
            SyncLock _queueLock
                If _pendingMessages IsNot Nothing Then
                    While _pendingMessages.Count > 0
                        SendMessage(_pendingMessages.Dequeue())
                    End While
                    _pendingMessages = Nothing
                End If
            End SyncLock
        End Sub

        Private Function _disconnect() As Boolean
            SyncLock Me._eventLock
                Me._messageReceived = Nothing
                Dim _stream As PipeStream = _getStream()
                _stream.Close()
                Return True
            End SyncLock
        End Function
#End Region
#End Region

#Region "Events"
        ''' <summary>
        ''' Occurs when [connected].
        ''' </summary>
        Public Event Connected(sender As Object, e As System.EventArgs) Implements INamedPipesBase.Connected
        ''' <summary>
        ''' Occurs when [message received].
        ''' </summary>
        Public Custom Event MessageReceived As EventHandler(Of MessageEveArgs) Implements INamedPipesBase.MessageReceived
            AddHandler(ByVal value As EventHandler(Of MessageEveArgs))
                SyncLock Me._eventLock
                    'I think this code to combine clients that connected to a server.
                    _messageReceived = DirectCast([Delegate].Combine(Me._messageReceived, value), EventHandler(Of MessageEveArgs))
                End SyncLock
            End AddHandler
            RemoveHandler(ByVal value As EventHandler(Of MessageEveArgs))
                SyncLock Me._eventLock
                    _messageReceived = DirectCast([Delegate].Remove(Me._messageReceived, value), EventHandler(Of MessageEveArgs))
                End SyncLock
            End RemoveHandler
            RaiseEvent(sender As Object, args As MessageEveArgs)
                SyncLock Me._eventLock
                    _messageReceived.Invoke(sender, args)
                End SyncLock
            End RaiseEvent
        End Event
#End Region

#Region "IDisposable Support"
        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Private disposedValue As Boolean ' To detect redundant calls
        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    If Not _server Is Nothing Then _server.Dispose()
                    If Not _client Is Nothing Then _client.Dispose()
                    _pendingMessages = Nothing
                End If
                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        Protected Overrides Sub Finalize()
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(False)
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace