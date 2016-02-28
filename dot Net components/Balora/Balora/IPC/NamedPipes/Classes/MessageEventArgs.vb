'Based on this example
'http://www.eggheadcafe.com/tutorials/aspnet/478ca475-2fd9-45d6-9090-5acdaeb00162/make-your-apps-talk-to-each-other-asynchronous-named-pipes-library.aspx

'/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \ 
'                                                   MessageEveArgs Class                                                        '       
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /
Namespace Ipc.NamedPipes
    <Serializable()> _
    Friend Class MessageEveArgs
        Inherits EventArgs
        Implements IMessageEveArgs
        Private ReadOnly _message As Object
        ''' <summary>
        ''' Initializes a new instance of the <see cref="MessageEveArgs"/> class.
        ''' </summary>
        ''' <param name="message">The message.</param>
        Public Sub New(message As Object)
            Me._message = message
        End Sub
        ''' <summary>
        ''' Gets the message.
        ''' </summary>
        ''' <value>The message.</value>
        Public ReadOnly Property Message() As Object Implements IMessageEveArgs.Message
            Get
                Return Me._message
            End Get
        End Property
    End Class
End Namespace

