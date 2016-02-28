'Based on this example
'http://www.eggheadcafe.com/tutorials/aspnet/478ca475-2fd9-45d6-9090-5acdaeb00162/make-your-apps-talk-to-each-other-asynchronous-named-pipes-library.aspx

'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \ 
'                                                   IServer Interface                                                            '       
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /
Namespace Ipc.NamedPipes
    Friend Interface IPipeServer
        Function Listen() As Boolean
        Sub ClientConnected(result As IAsyncResult)
        Sub EndRead(result As IAsyncResult)
        Function SendMessage(message As Byte()) As Boolean
        Function Disconnect() As Boolean
    End Interface
End Namespace