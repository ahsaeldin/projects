'Based on this example
'http://www.eggheadcafe.com/tutorials/aspnet/478ca475-2fd9-45d6-9090-5acdaeb00162/make-your-apps-talk-to-each-other-asynchronous-named-pipes-library.aspx

'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \ 
'                                                   INamedPipesBase Interface                                                    '       
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /
Imports System.IO.Pipes

Namespace Ipc.NamedPipes
    Friend Interface INamedPipesBase
        Inherits IDisposable
        Property PipingMode() As PipeOptions
        ReadOnly Property PipeName() As String
        ReadOnly Property IsConnected() As Boolean
        Function SendMessage(message As Byte()) As Boolean
        Function DoSendMessage(message As Byte()) As Boolean
        Sub EndSendMessage(result As IAsyncResult)
        Sub OnMessageReceived(args As MessageEveArgs)
        Sub OnConnection()
        Sub SendQueuedMessages()
        Sub EnqueueMessage(message As Byte())
        Function Disconnect() As Boolean
        Event Connected As EventHandler
        Event MessageReceived As EventHandler(Of MessageEveArgs)
    End Interface
End Namespace