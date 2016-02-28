Imports Balora
Imports NUnit.Framework

Namespace BaloraUT
    <Ignore()>
    <TestFixture()> _
    Public Class NamedPipesTest

        Private receivedString As String
        Private autoEvent As New Threading.AutoResetEvent(False)
        Private WithEvents _server As Balora.Ipc.NamedPipes.PipeServer
        Private WithEvents _client As Balora.Ipc.NamedPipes.PipeClient

        <TestFixtureSetUp()> _
        Public Sub SetupMethods()
            PrepareCacaduComponents()
            Console.WriteLine("Get authorization to use Balora.")
            Console.WriteLine("Create server and client...")
            _server = New Balora.Ipc.NamedPipes.PipeServer("HiMan", IO.Pipes.PipeOptions.Asynchronous)
            _client = New Balora.Ipc.NamedPipes.PipeClient("HiMan", IO.Pipes.PipeOptions.Asynchronous)
            Console.WriteLine("Client::Connecting to server...")
            _client.Connect()
            Console.WriteLine("Client::Wait for server to accept connection...")
            autoEvent.WaitOne()

        End Sub

        <TestFixtureTearDown()> _
        Public Sub TearDownMethods()
        End Sub

        <SetUp()> _
        Public Sub SetupTest()
        End Sub

        <TearDown()> _
        Public Sub TearDownTest()
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub TestSendMessageFromClientToServer()

            Console.WriteLine("Client::Test sending a message from client to server.")
            Dim actual As Boolean = False
            Dim expected As Boolean = True
            Console.WriteLine("Client::Sending a message...")
            sendMessageFromClientToServer()
            autoEvent.Reset()
            Console.WriteLine("Client::Waiting for the message to be sent...")
            autoEvent.WaitOne()
            If receivedString = "This is message man" Then
                Console.WriteLine("Client::Message sent.")
                actual = True
            End If
            Assert.AreEqual(expected, actual)

        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub TestSendMessageFromServerToClient()

            Dim actual As Boolean = False
            Dim expected As Boolean = True
            receivedString = vbNullString
            Console.WriteLine("Server::Test sending  a message from server to client...")
            sendMessageFromServerToClient()
            autoEvent.Reset()
            autoEvent.WaitOne()
            Console.WriteLine("Server::Waiting for the message to be sent...")
            If receivedString = "This is message man" Then
                actual = True
                Console.WriteLine("Server::Message sent.")
            End If
            Assert.AreEqual(expected, actual)

        End Sub

        Private Sub _server_Connected(sender As Object, e As System.EventArgs) Handles _server.Connected

            Console.WriteLine("Server::Server connected.")
            autoEvent.Set()

        End Sub

        Private Sub _server_MessageReceived(sender As Object, args As Ipc.NamedPipes.MessageEveArgs) Handles _server.MessageReceived

            Dim message As Ipc.NamedPipes.Message = CType(Balora.Util.ByteArrayToObject(CType(args.Message, Byte())), Ipc.NamedPipes.Message)
            receivedString = System.Text.Encoding.UTF8.GetString(CType(message.Payload, Byte()))
            Console.WriteLine("Server::A new message received from client.")
            autoEvent.Set()

        End Sub

        Private Sub _client_MessageReceived(sender As Object, args As Ipc.NamedPipes.MessageEveArgs) Handles _client.MessageReceived

            Dim message As Ipc.NamedPipes.Message = CType(Balora.Util.ByteArrayToObject(CType(args.Message, Byte())), Ipc.NamedPipes.Message)
            receivedString = System.Text.Encoding.UTF8.GetString(CType(message.Payload, Byte()))
            Console.WriteLine("Client::A new message received from server.")
            autoEvent.Set()

        End Sub

        Private Sub sendMessageFromClientToServer()

            Dim m As New Ipc.NamedPipes.Message()
            m.MessageId = Guid.NewGuid()
            m.Recipient = "AppTwo"
            m.Originator = "AppOne"
            m.MessageType = GetType([String])
            m.MessageDateTime = DateTime.Now
            m.Payload = System.Text.Encoding.UTF8.GetBytes("This is message man")
            Dim b As Byte() = Balora.Util.ObjectToByteArray(m)
            _client.SendMessage(b)

        End Sub

        Private Sub sendMessageFromServerToClient()

            Dim m As New Ipc.NamedPipes.Message()
            m.MessageId = Guid.NewGuid()
            m.Recipient = "AppTwo"
            m.Originator = "AppOne"
            m.MessageType = GetType([String])
            m.MessageDateTime = DateTime.Now
            m.Payload = System.Text.Encoding.UTF8.GetBytes("This is message man")
            Dim b As Byte() = Balora.Util.ObjectToByteArray(m)
            _server.SendMessage(b)

        End Sub

    End Class

End Namespace
