Namespace Ipc.NamedPipes
    Friend Module NamedPipeHelper

        Public Function ProcessReceivedMessage(e As MessageEveArgs) As Object

            Dim message As Ipc.NamedPipes.Message = CType(Balora.Util.ByteArrayToObject(CType(e.Message, Byte())), Ipc.NamedPipes.Message)

            Select Case message.MessageType

                Case GetType(String)

                    Dim receivedString As String = System.Text.Encoding.UTF8.GetString(CType(message.Payload, Byte()))

                    Return receivedString

                Case Else

                    Return message.Payload

            End Select

        End Function

        ''' <summary>
        ''' Sends the message helper.
        ''' </summary>
        ''' <param name="pipeObject">The pipe object.</param>
        ''' <param name="recipient">The recipient.</param>
        ''' <param name="originator">The originator.</param>
        ''' <param name="message">The message.</param><returns></returns>
        ''' <remarks>
        ''' You can call this method directly to send a message, however,
        ''' I prefer to use SendMessageToCacadu or SendMessageToCacaSer instead,
        ''' if you're calling this method from Cacadu or CacaSer.
        ''' </remarks>
        Public Function SendMessageHelper(pipeObject As NamedPipesBase,
                                          recipient As String,
                                          originator As String,
                                          message As Object, Optional exception As Boolean = False) As Boolean

            Dim msgObj As New Ipc.NamedPipes.Message()

            msgObj.MessageId = Guid.NewGuid()
            msgObj.Recipient = recipient
            msgObj.Originator = originator
            msgObj.MessageType = message.GetType
            msgObj.MessageDateTime = DateTime.Now

            If exception Then msgObj.Description = "exception"

            If msgObj.MessageType Is GetType([String]) Then
                msgObj.Payload = System.Text.Encoding.UTF8.GetBytes(CStr(message))
            Else
                msgObj.Payload = message
            End If

            Dim msgByteArray As Byte() = Balora.Util.ObjectToByteArray(msgObj)

            Return pipeObject.SendMessage(msgByteArray)

        End Function

    End Module

End Namespace

