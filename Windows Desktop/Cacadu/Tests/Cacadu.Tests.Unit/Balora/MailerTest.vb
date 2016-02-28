Imports System.Net.Mail
Imports Balora
Imports NUnit.Framework

Namespace BaloraUT
    <TestFixture()> _
    Public Class MailerTest

        Dim target As Mailer
        Dim manulReset As New Threading.ManualResetEvent(False)

        <TestFixtureSetUp()> _
        Public Sub SetupMethods()
            PrepareCacaduComponents()
            target = New Mailer()
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
        Public Sub SendTest()
            Console.WriteLine("Starting send test...")
            Dim actual As Boolean = False
            Dim expected As Boolean = True
            Dim smtpMailer As New Mailer
            With smtpMailer
                .Sender = "error_reporting_center@conderella.com"
                .SenderDisplayName = "IAM THE SENDER"
                With .Receivers
                    .Add(New MailAddress("cprinahmed@yahoo.com", "I AM THE RECEIVER"))
                    .Add(New MailAddress("mooparmghor@gmail.com", "I AM THE RECEIVER no 2"))
                End With
                With .Credential '===========================================================>
                    .AppendChar(CChar("J")) : .AppendChar(CChar("O")) : .AppendChar(CChar("l"))
                    .AppendChar(CChar("U")) : .AppendChar(CChar("m")) : .AppendChar(CChar("8"))
                    .AppendChar(CChar("m")) : .AppendChar(CChar("Y")) : .AppendChar(CChar("u"))
                    .AppendChar(CChar("R")) : .AppendChar(CChar("n")) : .AppendChar(CChar("O"))
                    .AppendChar(CChar("8")) : .AppendChar(CChar("9")) : .AppendChar(CChar("7"))
                    .AppendChar(CChar("5")) : .AppendChar(CChar("M")) : .AppendChar(CChar("n"))
                    .AppendChar(CChar("G")) : .AppendChar(CChar("n")) : .AppendChar(CChar("K"))
                    .AppendChar(CChar("b")) : .AppendChar(CChar("h"))
                End With
                .SmtpHost = "mail.conderella.com"
                .SmtpPort = 26
                .IsSslEnabled = False
                .Subject = "Unit Testing Mailer Class"
                .Body = "This is body"
                .IsBodyHtml = False
                .Priority = MailPriority.High
                'Dim paths As New ArrayList
                'paths.Add("TestingSendingAttachments1.txt")
                'paths.Add("TestingSendingAttachments2.txt")
                'paths.Add("TestingSendingAttachments3.txt")
                '.AttachmentsPaths = paths
                .Send()
            End With
            target = smtpMailer
            Console.WriteLine("Waiting for the message to be sent...")
            AddHandler target.SendCompleted, AddressOf Message_Completed
            manulReset.WaitOne()
            Console.WriteLine("Check if sending completed...")
            If target.IsCompleted Then actual = True : Console.WriteLine("Message Sent.")
            Assert.AreEqual(expected, actual)
        End Sub

        Private Sub Message_Completed(ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
            manulReset.Set()
        End Sub
    End Class
End Namespace
