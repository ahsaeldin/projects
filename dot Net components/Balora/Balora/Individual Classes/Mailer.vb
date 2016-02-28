Option Strict On
Option Explicit On

Imports System.Net.Mail
Imports System.Windows.Forms
Imports System.ComponentModel

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'                                                   Mailer Class                                                   
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'commented next line as no need to avoid obfuscation of Mailer 
'<System.Reflection.ObfuscationAttribute(Feature:="renaming", ApplyToMembers:=True)>
Public Class Mailer
    Inherits BaloraBase
    Implements IDisposable

#Region "Properties"
    ''' <summary>
    ''' send to email
    ''' </summary>
    Private senderValue As String
    Public Property Sender() As String
        Get
            Return senderValue
        End Get
        Set(ByVal value As String)
            senderValue = value
        End Set
    End Property

    ''' <summary>
    ''' Collection of emails to be send to
    ''' </summary>
    Private receiversValue As New MailAddressCollection
    Public Property Receivers() As MailAddressCollection
        Get
            Return receiversValue
        End Get
        Set(ByVal value As MailAddressCollection)
            receiversValue = value
        End Set
    End Property

    ''' <summary>
    ''' Display name of sender
    ''' </summary>
    Private senderDisplayNameValue As String
    Public Property SenderDisplayName() As String
        Get
            Return senderDisplayNameValue
        End Get
        Set(ByVal value As String)
            senderDisplayNameValue = value
        End Set
    End Property

    ''' <summary>
    ''' Enter password directly as a string.
    ''' </summary>
    ''' <remarks></remarks>
    Private passwordValue As String
    Public Property Password() As String
        Get
            Return passwordValue
        End Get
        Set(ByVal value As String)
            passwordValue = value
        End Set
    End Property

    ''' <summary>
    ''' Enter Credential/Password as a secure string.
    ''' </summary>
    ''' <remarks></remarks>
    Private credentialValue As New System.Security.SecureString
    Public Property Credential() As System.Security.SecureString
        Get
            Return credentialValue
        End Get
        Set(ByVal value As System.Security.SecureString)
            credentialValue = value
        End Set
    End Property

    ''' <summary>
    ''' SMTP host name.
    ''' </summary>
    Private smtpHostValue As String
    Public Property SmtpHost() As String
        Get
            Return smtpHostValue
        End Get
        Set(ByVal value As String)
            smtpHostValue = value
        End Set
    End Property

    ''' <summary>
    ''' SMTP port number.
    ''' </summary>
    ''' <remarks></remarks>
    Private smtpPortValue As Integer
    Public Property SmtpPort() As Integer
        Get
            Return smtpPortValue
        End Get
        Set(ByVal value As Integer)
            smtpPortValue = value
        End Set
    End Property

    ''' <summary>
    ''' Enable/Disable SSL, (must provide ssl port)
    ''' </summary>
    Private isSslEnabledValue As Boolean
    Public Property IsSslEnabled() As Boolean
        Get
            Return isSslEnabledValue
        End Get
        Set(ByVal value As Boolean)
            isSslEnabledValue = value
        End Set
    End Property

    ''' <summary>
    ''' Email priority
    ''' </summary>
    ''' <remarks></remarks>
    Private priorityValue As MailPriority
    Public Property Priority() As MailPriority
        Get
            Return priorityValue
        End Get
        Set(ByVal value As MailPriority)
            priorityValue = value
        End Set
    End Property

    ''' <summary>
    ''' Email subject.
    ''' </summary>
    Private subjectValue As String
    Public Property Subject() As String
        Get
            Return subjectValue
        End Get
        Set(ByVal value As String)
            subjectValue = value
        End Set
    End Property

    ''' <summary>
    ''' Message body.
    ''' </summary>
    Private bodyValue As String
    Public Property Body() As String
        Get
            Return bodyValue
        End Get
        Set(ByVal value As String)
            bodyValue = value
        End Set
    End Property

    ''' <summary>
    ''' Enable sending message as an HTML document.
    ''' </summary>
    Private isBodyHtmlValue As Boolean
    Public Property IsBodyHtml() As Boolean
        Get
            Return isBodyHtmlValue
        End Get
        Set(ByVal value As Boolean)
            isBodyHtmlValue = value
        End Set
    End Property

    ''' <summary>
    ''' True, if the server requires user's credentials not the default credentials.
    ''' </summary>
    Private useDefaultCredentialsValue As Boolean
    Public Property UseDefaultCredentials() As Boolean
        Get
            Return useDefaultCredentialsValue
        End Get
        Set(ByVal value As Boolean)
            useDefaultCredentialsValue = value
        End Set
    End Property

    ''' <summary>
    ''' Is message has an attachments?
    ''' </summary>
    Private hasAttachmentsValue As Boolean
    Public Property HasAttachments() As Boolean
        Get
            Return hasAttachmentsValue
        End Get
        Set(ByVal value As Boolean)
            hasAttachmentsValue = value
        End Set
    End Property

    ''' <summary>
    ''' Pass attachments paths as an arraylist,
    ''' and it will check if files exists and
    ''' set HasAttachments to true.
    ''' If one or all paths isn't exists, it collect them
    ''' then display them in a message box.
    ''' </summary>
    ''' <remarks></remarks>
    Private attachmentsPathsValue As New ArrayList
    Public Property AttachmentsPaths() As ArrayList
        Get
            Return attachmentsPathsValue
        End Get
        Set(ByVal value As ArrayList)
            attachmentsPathsValue = value
            Dim tempPath As String = ""
            Dim invalidPaths As String = ""
            Try
                For Each attachmentPath As String In attachmentsPathsValue
                    tempPath = System.IO.Path.GetFullPath(attachmentPath)
                    If (System.IO.File.Exists(tempPath)) Then
                        HasAttachments = True
                    Else
                        invalidPaths = invalidPaths + tempPath + vbNewLine
                    End If
                Next
                If invalidPaths <> "" Then
                    invalidPaths = "The following attachment(s) file(s) not found:" +
                    vbNewLine + vbNewLine + invalidPaths
                    attachmentsPathsValue(0) = 0
                    MsgBox(invalidPaths, MsgBoxStyle.Critical, "Error")
                End If
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                MsgBox(ex.Message)
            End Try
        End Set
    End Property

    ''' <summary>
    ''' A value indicting if mailer is in sending state.
    ''' </summary>
    Private isSendingValue As Boolean
    Public Property IsSending() As Boolean
        Get
            Return isSendingValue
        End Get
        Set(ByVal value As Boolean)
            isSendingValue = value
        End Set
    End Property

    ''' <summary>
    ''' A value indicting if mailer is in timeout state.
    ''' </summary>
    Private isTimeOutValue As Boolean
    Public ReadOnly Property IsTimeOut() As Boolean
        Get
            Return isTimeOutValue
        End Get
    End Property

    ''' <summary>
    ''' A value indicating if the last call to send method is complete
    ''' and the message either sent or canceled or error occurred.
    ''' </summary>
    Private isCompletedValue As Boolean
    Public Property IsCompleted() As Boolean
        Get
            Return isCompletedValue
        End Get
        Set(ByVal value As Boolean)
            isCompletedValue = value
        End Set
    End Property

    ''' <summary>
    ''' 'If you specify the SendingTimeOut (milleSecond) property 
    ''' then I will take it, otherwise I will take SmtpClient.TimeOut.
    ''' </summary>
    Private sendingTimeOutValue As Integer
    Public Property SendingTimeOut() As Integer
        Get
            Return sendingTimeOutValue
        End Get
        Set(ByVal value As Integer)
            sendingTimeOutValue = value
        End Set
    End Property
#End Region

#Region "Methods"
    ''' <summary>
    ''' An Overridable method for sending the message after
    ''' filling the necessary properties.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overridable Sub Send()
        'zeroing values from previous calls.
        IsSending = False
        IsCompleted = False
        isTimeOutValue = False
        Dim msg As New MailMessage()
        msg.From = New MailAddress(Sender, SenderDisplayName)
        ' To(s) addresses
        For Each receiverAddr As MailAddress In Receivers
            msg.To.Add(receiverAddr)
        Next
        'Set to high priority
        msg.Priority = Priority
        msg.Subject = Subject
        'You can specify a plain text or HTML contents.
        msg.Body = Body
        msg.IsBodyHtml = IsBodyHtml
        'Attaching some data
        If HasAttachments Then
            'In the internal logic of "AttachmentsPaths" I set first item in it to 0,
            'if one or more of the paths isn't exists and arises a msgbox to declare that.
            If Not AttachmentsPaths(0) Is CStr(0) Then
                For Each singleAttachmentPath As String In AttachmentsPaths
                    msg.Attachments.Add(New Attachment(singleAttachmentPath))
                Next
            End If
        End If
        'Connecting to the server and configuring it.
        Dim smtpClient As New SmtpClient()
        smtpClient.Host = SmtpHost
        smtpClient.Port = SmtpPort
        smtpClient.EnableSsl = IsSslEnabled
        'The server requires user's credentials not the default credentials.
        smtpClient.UseDefaultCredentials = UseDefaultCredentials
        smtpClient.Credentials = New System.Net.NetworkCredential(Sender, Credential)
        smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network
        AddHandler smtpClient.SendCompleted, AddressOf smtpClient_SendCompleted
        'Setup the timer that will check if timeout occurs.
        Dim TimeOutCheckerTimer As New Timer
        TimeOutCheckerTimer.Enabled = False
        'If you specify the SendingTimeOut property then I will take it, otherwise I will take SmtpClient.TimeOut.
        If SendingTimeOut = Nothing Then
            TimeOutCheckerTimer.Interval = smtpClient.Timeout
        Else
            TimeOutCheckerTimer.Interval = SendingTimeOut
        End If

        AddHandler TimeOutCheckerTimer.Tick, AddressOf TimeOutChecker
        Try
            Dim userToken As Object = msg
            TimeOutCheckerTimer.Enabled = True
            IsSending = True
            'Use SendAsync to send the message asynchronously.
            smtpClient.SendAsync(msg, userToken)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Hodhod.WriteToErrorLog(ex.Message, ex.StackTrace, "Error sending mail:")
        End Try
    End Sub
#End Region

#Region "Events"
    Public Event SendCompleted(ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
    Public Event TimeOut()

    Private Sub smtpClient_SendCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
        IsSending = False
        IsCompleted = True
        RaiseEvent SendCompleted(e)
    End Sub

    ''' <summary>
    ''' Timer for checking timout and if then RaiseTimeOut()
    ''' </summary>
    ''' <param name="sender">We use it to get the original timer object to disable it. </param>
    Private Sub TimeOutChecker(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim timer As Timer = CType(sender, Timer)
        timer.Enabled = False
        If IsSending Then
            isTimeOutValue = True
            RaiseEvent TimeOut()
        End If
    End Sub
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class