Namespace Ipc.NamedPipes
    '\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /
    '                                                   Message Class
    '\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /	
    <Serializable()> _
    Friend Class Message
        Implements Ipc.NamedPipes.IMessage

#Region "IMessage Members"
        Private _messageId As Guid
        Private _originator As String
        Private _recipient As String
        Private _messageDateTime As DateTime
        Private _messageType As Type
        Private _payload As Object
        Private _descriptionValue As String

        ''' <summary>
        ''' Gets or sets the message guid.
        ''' </summary>
        ''' <value>
        ''' The message id.
        ''' </value>
        Public Property MessageId() As Guid Implements Ipc.NamedPipes.IMessage.MessageId
            Get
                Return _messageId
            End Get
            Set(value As Guid)
                _messageId = value
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the originator. "The sender app."
        ''' </summary>
        ''' <value>
        ''' The originator.
        ''' </value>
        Public Property Originator() As String Implements Ipc.NamedPipes.IMessage.Originator
            Get
                Return _originator
            End Get
            Set(value As String)
                _originator = value
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the recipient.
        ''' </summary>
        ''' <value>
        ''' The recipient.
        ''' </value>
        Public Property Recipient() As String Implements Ipc.NamedPipes.IMessage.Recipient
            Get
                Return _recipient
            End Get
            Set(value As String)
                _recipient = value
            End Set
        End Property

        ''' <summary>
        ''' Gets or sets the message date time.
        ''' </summary>
        ''' <value>
        ''' The message date time.
        ''' </value>
        Public Property MessageDateTime() As DateTime Implements Ipc.NamedPipes.IMessage.MessageDateTime
            Get
                Return _messageDateTime
            End Get
            Set(value As DateTime)
                _messageDateTime = value
            End Set
        End Property

        ''' <summary>
        ''' Tell the recipient here about message data type.
        ''' </summary>
        ''' <value>
        ''' The type of the message.
        ''' </value>
        Public Property MessageType() As Type Implements Ipc.NamedPipes.IMessage.MessageType
            Get
                Return _messageType
            End Get
            Set(value As Type)
                _messageType = value
            End Set
        End Property

        ''' <summary>
        ''' The message itself (The part of message that carries the data.)
        ''' </summary>
        ''' <value>
        ''' The payload.
        ''' </value>
        Public Property Payload() As Object Implements Ipc.NamedPipes.IMessage.Payload
            Get
                Return _payload
            End Get
            Set(value As Object)
                _payload = value
            End Set
        End Property

        Public Property Description() As String Implements Ipc.NamedPipes.IMessage.Description
            Get
                Return _descriptionValue
            End Get
            Set(value As String)
                _descriptionValue = value
            End Set
        End Property

#End Region

        Public Sub New()

        End Sub

        Public Sub New(messageId As Guid, originator As String, recipient As String, messageDateTime As DateTime, messageType As Type, payload As Byte())
            Me._messageId = messageId
            Me._originator = originator
            Me._recipient = recipient
            Me._messageDateTime = messageDateTime
            Me._messageType = messageType
            Me._payload = payload
        End Sub

    End Class
End Namespace