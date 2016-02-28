'Based on this example
'http://www.eggheadcafe.com/tutorials/aspnet/478ca475-2fd9-45d6-9090-5acdaeb00162/make-your-apps-talk-to-each-other-asynchronous-named-pipes-library.aspx

'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \ 
'                                                   IMessage Interface                                                           '       
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /
Namespace Ipc.NamedPipes
    ''' <summary>
    ''' Interface for Async Pipes Message objects.
    ''' </summary>
    Friend Interface IMessage
        Property MessageId() As Guid
        Property Originator() As [String]
        Property Recipient() As [String]
        Property MessageDateTime() As DateTime
        Property MessageType() As Type
        Property Payload() As Object
        Property Description As String

    End Interface
End Namespace
