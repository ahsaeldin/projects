'ErrorReportingUIConnector for ErrorReportingUI.
'Use it to connect between ErrorReportingUI form and its events.
Imports System.Net.Mail
Imports Balora.ErrorReporter

Namespace Connectors

    'How to reuse this connector:-
    'I did my best to make this connector reusable, however, 
    'we will just need to change the error lines in ConnectorSettingsAndResources.
    Friend Class ErrorReportingFormConnector

        Private errRep As ErrorReporter
        Public Sub New(ByVal errReporter As ErrorReporter)
            errRep = errReporter
        End Sub

#Region "NonReusableArea"
        ''' <summary>
        ''' This is the only place in which is not reusable,
        ''' however, it so simple just replace lines that will
        ''' generate errors with the corresponding in the new app.
        ''' </summary>
        ''' <remarks>
        ''' If you want to reuse the whole connector just fill 
        ''' this structure with the proper settings and strings.
        ''' </remarks>
        Structure ConnectorSettingsAndResources
            Public Shared PictureBoxAppIconImage As Image = My.Resources.Cacadu_Image_54x54
            '7 Nov, 2011 -- Disable main window icon for now.
            Public Shared ErrorReportFormIcon As Icon '= Icon.FromHandle(My.Resources.CacaIco.GetHicon())
            '7 Nov, 2011 -- Add to change Error Reporter window title name
            Public Shared WindowCaption As String = "Cacadu"
            Public Shared AttachmentOpenFileDialogTitle As String = My.Resources.ErrorReportOpenFileDialogTitle
            Public Shared ErrorReportUiHeadingText As String = My.Resources.ErrorReportUiHeading
            Public Shared TellUsMoreText As String = My.Resources.TellUsMore
            Public Shared AlternativeIPText As String = My.Resources.AlternativeIPText
            Public Shared PreparingReportText As String = My.Resources.PreparingReport
            Public Shared TransferringReportText As String = My.Resources.TransferringReport
            Public Shared SendingCanceledText As String = My.Resources.SendingCanceled
            Public Shared SendingMailFailText As String = My.Resources.SendingMailFail
            Public Shared CheckFirewallText As String = My.Resources.CheckFirewall
            Public Shared ErrorSendingErrorReportText As String = My.Resources.ErrorSendingErrorReport
            Public Shared TransferringCompletedText As String = My.Resources.TransferringCompleted
            Public Shared TransferringTimeOutText As String = My.Resources.TransferringTimeOut
            Public Shared CheckBoxAutomaticallySend As Boolean = My.Settings.CheckBoxAutoSendState
            Public Shared AddYourComments As String = My.Resources.AddYourCommentsHere
            Public Shared NoComments As String = My.Resources.NoCommentsFromUser
            ''' <summary>
            ''' Called by CheckBoxAutoSend_CheckChanged to save checkbox setting.
            ''' </summary>
            ''' <remarks></remarks>
            Public Shared Sub SaveFormSettings()
                With My.Settings
                    .CheckBoxAutoSendState = ErrorReportingUI.AutoSendCheckBox.Checked
                    ConnectorSettingsAndResources.CheckBoxAutomaticallySend = ErrorReportingUI.AutoSendCheckBox.Checked
                    .Save()
                End With
            End Sub
        End Structure
#End Region

        'Reusable methods that are abstracted from the whole picture 
        'so I can simplify reusing of this connector.
#Region "ReusableMethods"

        'A timer for closing ErrorReportingUI automatically after kam seconds keda.
        Private closeTimer As New System.Windows.Forms.Timer

        Structure ErrorReportingUITextColors
            Public Shared NormalStatusBarTextColor As Color = Color.Blue
            Public Shared FailStatusBarTextColor As Color = Color.Red
            Public Shared SelectingAllTextInCommentingBoxColor As Color = Color.Black
        End Structure

        ''' <summary>
        ''' Bind ErrorReportingUI events with corresponding methods here.
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub BindErrorReportingUIEvents()
            'Use this method to bind ErrorReporterForm and ErrorReporter events with its 
            'corresponding methods in this class to make a fully abstraction 
            'in ErrorReporterForm so I can use it @ any further project as a 
            'component of Balora.
            AddHandler errRep.TimeOut, AddressOf ErrorReporter_TimeOut
            AddHandler ErrorReportingUI.Load, AddressOf ErrorReportingUI_Load
            AddHandler errRep.SendCompleted, AddressOf ErrorReporter_SendCompleted
            AddHandler ErrorReportingUI.FormClosing, AddressOf ErrorReportingUI_FormClosing
            AddHandler ErrorReportingUI.CommentsTextBox.Click, AddressOf CommentsTextBox_Click
            AddHandler ErrorReportingUI.DoNotSendButton.Click, AddressOf DontSendButton_Click
            AddHandler ErrorReportingUI.AttachmentButton.Click, AddressOf ButtonAttachment_Click
            AddHandler ErrorReportingUI.CommentsTextBox.KeyDown, AddressOf CommentsTextBox_KeyDown
            AddHandler ErrorReportingUI.ErrorDescTextBox.KeyDown, AddressOf TextBoxErrorDesc_KeyDown
            AddHandler ErrorReportingUI.AttachmentPathTextBox.Click, AddressOf ButtonAttachment_Click
            AddHandler ErrorReportingUI.SendErrorReportButton.Click, AddressOf ButtonSendErrorReport_Click
            AddHandler ErrorReportingUI.AutoSendCheckBox.CheckedChanged, AddressOf CheckBoxAutoSend_CheckChanged
        End Sub

        Private Sub UnBindErrorReportingUIEvents()
            RemoveHandler errRep.TimeOut, AddressOf ErrorReporter_TimeOut
            RemoveHandler ErrorReportingUI.Load, AddressOf ErrorReportingUI_Load
            RemoveHandler errRep.SendCompleted, AddressOf ErrorReporter_SendCompleted
            RemoveHandler ErrorReportingUI.FormClosing, AddressOf ErrorReportingUI_FormClosing
            RemoveHandler ErrorReportingUI.CommentsTextBox.Click, AddressOf CommentsTextBox_Click
            RemoveHandler ErrorReportingUI.DoNotSendButton.Click, AddressOf DontSendButton_Click
            RemoveHandler ErrorReportingUI.AttachmentButton.Click, AddressOf ButtonAttachment_Click
            RemoveHandler ErrorReportingUI.CommentsTextBox.KeyDown, AddressOf CommentsTextBox_KeyDown
            RemoveHandler ErrorReportingUI.ErrorDescTextBox.KeyDown, AddressOf TextBoxErrorDesc_KeyDown
            RemoveHandler ErrorReportingUI.AttachmentPathTextBox.Click, AddressOf ButtonAttachment_Click
            RemoveHandler ErrorReportingUI.SendErrorReportButton.Click, AddressOf ButtonSendErrorReport_Click
            RemoveHandler ErrorReportingUI.AutoSendCheckBox.CheckedChanged, AddressOf CheckBoxAutoSend_CheckChanged
        End Sub

        ''' <summary>
        ''' We use this method to show the dialog up after rebinding it with its events here.
        ''' </summary>
        ''' <remarks>Use this method only to show ErrorReportingUI, don't call 
        ''' ErrorReportingUI.ShowDialog() directy.
        ''' </remarks>
        Public Sub ShowErrorReportingDialog(ByVal ex As Exception)
            errRep.ExceptionObject = ex
            ErrorReportingUI.ErrorDescTextBox.Text = ex.Message & vbNewLine & ex.StackTrace
            If ConnectorSettingsAndResources.CheckBoxAutomaticallySend Then
                SendSilently("Silently Sent")
            Else
                BindErrorReportingUIEvents()
                ErrorReportingUI.ShowDialog()
            End If
        End Sub

        ''' <summary>
        ''' Prepare the form, calling from ErrorReportingUI_Load
        ''' </summary>
        ''' <param name="formIcon">Pass the icon you want for the form</param>
        ''' <param name="errorReportUiHeading">Heading Text box @ tob</param>
        ''' <param name="tellUsMore">Tell us more plz box</param>
        ''' <param name="isAutomaticallySending">check box of automatically sending the report</param>
        ''' <param name="userIP">user IP</param>
        ''' <param name="userComments"></param>
        ''' <param name="exceptionObject"></param>
        ''' <remarks></remarks>
        Private Sub PreparingErrorReportingFormForLoading(ByVal formIcon As Icon,
                                                             ByVal errorReportUiHeading As String,
                                                             ByVal tellUsMore As String,
                                                             ByVal isAutomaticallySending As Boolean,
                                                             ByVal userIP As String,
                                                             ByVal userComments As String,
                                                             ByVal exceptionObject As Exception)
            With ErrorReportingUI
                .Text = ConnectorSettingsAndResources.WindowCaption
                'Set the top right icon of ErrorReportingUI form.
                .AppIconPictureBox.Image = ConnectorSettingsAndResources.PictureBoxAppIconImage
                'Set form icon.
                .Icon = ConnectorSettingsAndResources.ErrorReportFormIcon
                'Main title of the Form
                .LabelHeadingTextLabel.Text = errorReportUiHeading
                'Label above commenting text area.
                .PlzTellUsLabel.Text = tellUsMore
                'Check box state of automatically send
                .AutoSendCheckBox.Checked = isAutomaticallySending
                'Enable controls if disables from a previous call.
                Balora.Util.DisableButtonsAndTextBoxes(ErrorReportingUI, True)
                'Prepare SMTP sending strings
                PrepareSendingStrings()
                'مؤقتا لن نرسل الايبي نفسه حتي اجد طريقة اسرع
                'userIP = ErrorReporter.UserRealIp(My.Resources.UrlToCheckRealIP)
                errRep.UserRealIp = userIP
                'Just to make sure that window will be @ center.
                .StartPosition = FormStartPosition.CenterScreen

                With errRep
                    .UserComments = userComments
                    If .UserComments = ConnectorSettingsAndResources.AddYourComments Then
                        .UserComments = ConnectorSettingsAndResources.NoComments
                    End If
                    Balora.Hodhod.WriteToErrorLog(exceptionObject.Message, exceptionObject.StackTrace, "Error from" & """" & exceptionObject.Source & """")
                End With

            End With
        End Sub

        ''' <summary>
        ''' Calling from CommentsTextBox_Click to select all the texts in the box by clicking it.
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub SelectAllTextBoxTexts(textbox As TextBox)
            textbox.SelectAll()
            textbox.ForeColor = ErrorReportingUITextColors.SelectingAllTextInCommentingBoxColor
        End Sub

        ''' <summary>
        ''' Calling from "CommentsTextBox_KeyDown" by pressing Ctrl + A
        ''' </summary>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub SelectAllTextBoxTextsByHotKey(ByVal e As System.Windows.Forms.KeyEventArgs, textbox As TextBox)
            With ErrorReportingUI
                If e.KeyCode = Keys.A And e.Control Then
                    SelectAllTextBoxTexts(textbox)
                End If
            End With
        End Sub

        ''' <summary>
        ''' Show up Open File Dialog to choose an attachment, calling from ButtonAttachment_Click
        ''' </summary>
        Private Shared Sub ShowAttachmentOpenFileDialog()
            ErrorReportingUI.AttachmentPathTextBox.Text = Balora.Util.ShowOpenFileDialog(ConnectorSettingsAndResources.AttachmentOpenFileDialogTitle)
        End Sub

        ''' <summary>
        ''' This method creates a list of strings then pass this list to
        ''' message body. The list contains the html tags required to create
        ''' the report. In the future if you want to change or customize the
        ''' report, you can create the list as you want then pass it to body
        ''' @ the end of the method
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub PrepareErrorReport()
            Dim messageBody As New List(Of String)
            With errRep
                'Clear body if it filled from a previous call.
                .Body = ""
                '2 blank lines before comments
                .UserComments = "<br /><br />"
                If .UserComments = My.Resources.AddYourCommentsHere Then
                    'Nothing to send to user
                    .UserComments = My.Resources.NoCommentsFromUser
                Else
                    'Insert breaks between lines to make comments more readable.
                    .UserComments += Balora.Util.InsertBetweenTextBoxLines(ErrorReportingUI.CommentsTextBox, "<br /><br />")
                End If
                'Add heading with a green background
                messageBody.Add("<div style=" & Chr(34) & "background:#6C3" & Chr(34) & "><h1 align=" & Chr(34) & "center" & Chr(34) & ">" & Application.ProductName & " Error Report</h1></div>")
                'UserName/Machine Name and DateTime 
                messageBody.Add("<p>" & Environment.MachineName & "\" & Environment.UserName & "  " & DateAndTime.Now & "</p><br />")
                'UserComments
                messageBody.Add("<strong>User Comments: </strong>" & .UserComments & "<hr />")
                'ExceptionSource from ExceptionObject
                messageBody.Add("<strong>Exception Source: </strong>" & .ExceptionObject.Source & "<hr />")
                'ExceptionMessage from ExceptionObject 
                messageBody.Add("<strong>Exception Description: </strong>" & .FullExceptionType & " | " & "{" & Chr(34) & .ExceptionObject.Message & Chr(34) & "}" & "<hr />")
                'Is UnHandled Exception
                messageBody.Add("<strong>Is UnHandled Exception: </strong>" & .IsUnHandledException & "<hr />")
                'Falsing the variable for future calling that are "Handled Exception"
                .IsUnHandledException = False
                'Error File Name
                messageBody.Add("<strong>Error File Name: </strong>" & .ErrorFileName & "<hr />")
                'Error Method Nam
                messageBody.Add("<strong>Error Method Name: </strong>" & .ErrorMethodName & "<hr />")
                'Error Line Number
                messageBody.Add("<strong>Error Line Number: </strong>" & .ErrorLineNumber & "<hr />")
                'Error Column Number
                messageBody.Add("<strong>Error Column Number: </strong>" & .ErrorColumnNumber & "<hr />")
                'AppVersion
                messageBody.Add("<strong>" & Application.ProductName & " Version: </strong>" & .AppVersion & "<hr />")
                'BuildDate
                messageBody.Add("<strong>" & Application.ProductName & " Version Built on: </strong>" & .BuildDate & "<hr />")
                'FullWindowsVersion
                messageBody.Add("<strong>Windows Version: </strong>" & .FullWindowsVersion & "<hr />")
                'ExeName
                messageBody.Add("<strong>" & Application.ProductName & " Exe Name: </strong>" & .ExeName & "<hr />")
                'ExePath
                messageBody.Add("<strong>" & Application.ProductName & " Exe Path: </strong>" & .ExePath & "<hr />")
                'UserRealIp
                messageBody.Add("<strong>User IP: </strong>" & .UserRealIp & "<hr />")
                'AvailablePhysicalMemory / TotalPhysicalMemory 
                messageBody.Add("<strong>AvailablePhysicalMemory/TotalPhysicalMemory: </strong>" &
                            .AvailablePhysicalMemory & "/" & .TotalPhysicalMemory & "<hr />")
                'LoadedAssemplies
                messageBody.Add("<strong>Loaded Assemblies: </strong><br /><br />")
                messageBody.Add("<table width=" & Chr(34) & 800 & Chr(34) & "border=" & 1 & ">")
                Dim loadedAssempliesDic As New Dictionary(Of String, String)
                loadedAssempliesDic = .LoadedAssemblies
                For Each kvp As KeyValuePair(Of String, String) In loadedAssempliesDic
                    messageBody.Add("<tr><td>" & "<strong>" & kvp.Key & "</strong><br />" & kvp.Value & "<br />" & "</td></tr>")
                Next kvp
                messageBody.Add("</table><br />")
                messageBody.Add("<hr />")
                'StackTrace  
                messageBody.Add("<strong>Stack Trace: </strong><br /><br />")
                messageBody.Add("<table width=" & Chr(34) & 800 & Chr(34) & "border=" & 1 & ">")
                Dim stackTraceDic As New Stack(Of String)
                stackTraceDic = .StackTrace
                For Each stackEnry As String In stackTraceDic
                    messageBody.Add("<tr><td>" & stackEnry & "</td></tr>")
                Next stackEnry

                messageBody.Add("</table><br />")
                messageBody.Add("<hr />")
                'ExceptionObject 
                messageBody.Add("<strong>ExceptionObject: </strong>" & .ExceptionObject.ToString & "<hr />")
                For Each entry As String In messageBody
                    errRep.Body += entry
                    'My.Computer.FileSystem.WriteAllText("C:\Test.html", entry, True)
                Next
            End With
        End Sub

        ''' <summary>
        ''' Prepare the connection strings we will used to send the email.
        ''' </summary>
        ''' <remarks>
        ''' So simple, you have all the data you need to change
        ''' </remarks>
        Private Sub PrepareSendingStrings()
            Static isSendingStringsPrepared As Boolean
            'To avoid errors, we prepare once.
            If Not isSendingStringsPrepared Then
                isSendingStringsPrepared = True
            Else
                Exit Sub
            End If
            With errRep
                .Sender = "error_reporting_center@conderella.com"
                .SenderDisplayName = Application.ProductName & " Error Reporter"
                With .Receivers
                    .Add(New System.Net.Mail.MailAddress("error_reporting_center@conderella.com", "Conderella Error Reporting Center"))
                End With
                Try
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
                Catch ex As ObjectDisposedException
                    Balora.Hodhod.WriteToErrorLog(ex.Message, ex.StackTrace, "PrepareSendingStrings::Error adding credential {This secure string has already been disposed.}")
                Catch ex As InvalidOperationException
                    Balora.Hodhod.WriteToErrorLog(ex.Message, ex.StackTrace, "PrepareSendingStrings::Error adding credential {This secure string is read-only.}")
                Catch ex As ArgumentOutOfRangeException
                    Balora.Hodhod.WriteToErrorLog(ex.Message, ex.StackTrace, "PrepareSendingStrings::Error adding credential {Performing this operation would make the length of this secure string greater than 65536 characters.}")
                Catch ex As System.Security.Cryptography.CryptographicException
                    Balora.Hodhod.WriteToErrorLog(ex.Message, ex.StackTrace, "PrepareSendingStrings::Error adding credential {An error occurred while encrypting or decrypting the value of this secure string.}")
                End Try
                .SmtpHost = "mail.conderella.com"
                .SmtpPort = 26
                'Three minutes 
                .SendingTimeOut = 180000
                .IsSslEnabled = False
                .Subject = Application.ProductName & " error report from " & Environment.MachineName & "\" & Environment.UserName & "  {" & .ExceptionMessage & "}"
                .IsBodyHtml = True
                .Priority = System.Net.Mail.MailPriority.High
            End With
        End Sub

        ''' <summary>
        ''' Send error report by preparing the report first then,
        ''' calling ErrorReport Send method second.
        ''' </summary>
        ''' <remarks>Used in SendSilently and ClickButtonSendErrorReport.</remarks>
        Private Sub SendErrorReport()
            Dim paths As New ArrayList
            Try
                paths.Add(Balora.Settings.LogFilePath)
                If ErrorReportingUI.AttachmentPathTextBox.Text <> "" Then
                    paths.Add(ErrorReportingUI.AttachmentPathTextBox.Text)
                End If
            Catch ex As NotSupportedException
                Balora.Hodhod.WriteToErrorLog(ex.Message, ex.StackTrace, "SendErrorReport::Error adding an attachment path {The ArrayList is read-only. -or- The ArrayList has a fixed size.}")
            End Try
            errRep.AttachmentsPaths = paths
            PrepareErrorReport()
            errRep.Send()
        End Sub

        ''' <summary>
        ''' Send error report silently if user choose this
        ''' by checking a checkbox in the ErrorReportingUI
        ''' </summary>
        ''' <param name="comments">Comments added by me just tell that sending is silently</param>
        ''' <remarks>
        ''' You must call PrepareSendingStrings in this method
        ''' because you won't load the form so you won't call
        ''' PrepareSendingStrings in PreparingErrorReportingFormForLoading method.
        ''' </remarks>
        Public Sub SendSilently(ByVal comments As String)
            'To tell me that
            ErrorReportingUI.CommentsTextBox.Text = comments
            PrepareSendingStrings()
            SendErrorReport()
        End Sub

        ''' <summary>
        ''' Invoked when user click sending error 
        ''' report button (ButtonSendErrorReport_Click).
        ''' </summary>
        ''' <param name="statusBarTextBeforeSending">Status bar text just before sending</param>
        ''' <param name="statusBarTextAfterSending">Status bar text just after sending</param>
        ''' <param name="statusBarColor">Status bar text forecolor</param>
        ''' <remarks></remarks>
        Private Sub ClickButtonSendErrorReport(ByVal statusBarTextBeforeSending As String,
                                                        ByVal statusBarTextAfterSending As String,
                                                        ByVal statusBarColor As Color
                                                        )
            With ErrorReportingUI
                'Tell the user in status bar about preparing of report
                .LabelStartusbar.Text = statusBarTextBeforeSending
                'Display the marquee progress bar.
                .SendingStatusToolStripProgressBar.Visible = True
                .LabelStartusbar.Visible = True
                'Disable controls of ErrorReportingUI
                Balora.Util.DisableButtonsAndTextBoxes(ErrorReportingUI)
                'Set status bar text color "blue" is the default
                .LabelStartusbar.ForeColor = statusBarColor
                'Send report
                SendErrorReport()
                'Set text after sending start.
                ErrorReportingUI.LabelStartusbar.Text = statusBarTextAfterSending
            End With
        End Sub

        ''' <summary>
        ''' Disable the timer responsible for closing up the form.
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <remarks></remarks>
        Private Sub DisableCloseTimer(ByVal sender As Object)
            Using closeTimerAlias As System.Windows.Forms.Timer = CType(sender, System.Windows.Forms.Timer)
                closeTimerAlias.Enabled = False
            End Using
            closeTimer.Dispose()
        End Sub

        ''' <summary>
        ''' Handle sending complete event for the error report.
        ''' </summary>
        ''' <param name="AsyncCompletedEventArguments">passed by ErrorReporter_SendCompleted</param>
        ''' <param name="StatusBarSendingCanceledText">Status bar text if canceled.</param>
        ''' <param name="StatusBarSendingMailFailText">Status bar text if failed.</param>
        ''' <param name="StatusBarCheckFirewallText">Status bar text if an error occurs.</param>
        ''' <param name="StatusBarErrorSendingError">Status bar text if an error occurs. (written to error log)</param>
        ''' <param name="StatusBarTransferComplete">Status bar text if sent successfully</param>
        ''' <remarks></remarks>
        Private Sub SendComplete(ByVal AsyncCompletedEventArguments As System.ComponentModel.AsyncCompletedEventArgs,
                                                    ByVal StatusBarSendingCanceledText As String,
                                                    ByVal StatusBarSendingMailFailText As String,
                                                    ByVal StatusBarCheckFirewallText As String,
                                                    ByVal StatusBarErrorSendingError As String,
                                                    ByVal StatusBarTransferComplete As String
                                                    )
            'Get the Original MailMessage object.
            Using msg As MailMessage = CType(AsyncCompletedEventArguments.UserState, MailMessage)
                'Write out the subject.
                Dim subject As String = msg.Subject
                With ErrorReportingUI

                    'Hide the progress bar in status bar.
                    .SendingStatusToolStripProgressBar.Visible = False

                    With .LabelStartusbar
                        'If the process canceled
                        If AsyncCompletedEventArguments.Cancelled Then
                            .ForeColor = ErrorReportingUITextColors.FailStatusBarTextColor
                            .Text = StatusBarSendingCanceledText
                        End If
                    End With

                    'If there is no error
                    If AsyncCompletedEventArguments.Error Is Nothing Then
                        .LabelStartusbar.ForeColor = ErrorReportingUITextColors.NormalStatusBarTextColor
                        'Tell the user that transfer complete.
                        .LabelStartusbar.Text = StatusBarTransferComplete
                        'Setup a timer for closing the form automatically after a specified interval.
                        SetClosingTimer()
                    Else 'If there is an error then
                        'Set font color of status bar to red
                        .LabelStartusbar.ForeColor = ErrorReportingUITextColors.FailStatusBarTextColor
                        'Tell the user about the error.
                        .LabelStartusbar.Text = StatusBarSendingMailFailText & " " & StatusBarCheckFirewallText
                        'Enable controls, may the user try again.
                        Balora.Util.DisableButtonsAndTextBoxes(ErrorReportingUI, True)
                        Balora.Hodhod.WriteToErrorLog(AsyncCompletedEventArguments.Error.Message, "", StatusBarErrorSendingError)
                    End If

                End With
            End Using
        End Sub

        ''' <summary>
        ''' A method called by DontSendButton_Click and CloseTimer_Tick to close the form 
        ''' </summary>
        ''' <param name="dispose">If true then dispose the form not only closing it.</param>
        ''' <remarks></remarks>
        Private Sub CloseForm(ByVal dispose As Boolean)
            If dispose Then
                ErrorReportingUI.Close() ' ErrorReportingUI.Dispose
            Else
                ErrorReportingUI.Close()
            End If
        End Sub

        ''' <summary>
        ''' A method is being called by SendComplete method 
        ''' if message sent successfully to close up the form
        ''' automatically after a specified period.
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub SetClosingTimer()
            'Set closing timer.
            AddHandler closeTimer.Tick, AddressOf CloseTimer_Tick
            closeTimer.Interval = 3000
            closeTimer.Enabled = True
        End Sub

        ''' <summary>
        ''' A method is being called when a time out occurs.
        ''' </summary>
        ''' <param name="statusBarColor">Status bar text color if time out occurs</param>
        ''' <param name="statusBarTimeOutText">Status bar text if time out occurs</param>
        ''' <remarks>Called by ErrorReporter_TimeOut</remarks>
        Private Sub TimeOut(ByVal statusBarColor As Color,
                               ByVal statusBarTimeOutText As String)
            With ErrorReportingUI
                'Hide progress bar from status bar.
                .SendingStatusToolStripProgressBar.Visible = False
                'Set status bar font color to red.
                .LabelStartusbar.ForeColor = statusBarColor
                'Tell the user that he has a time out.
                .LabelStartusbar.Text = statusBarTimeOutText
                'Enable ErrorReportingUI controls, may the user try again.
                Balora.Util.DisableButtonsAndTextBoxes(ErrorReportingUI, True)
            End With
        End Sub
#End Region

        'I did my best to abstract these events, and make 
        'it simple to read, simple to understand as easy 
        'as you can imagine to reuse.
#Region "ErrorReportingFormEvents"
        ''' <summary>
        ''' 'Loading event of ErrorReportingUI
        ''' </summary>
        Private Sub ErrorReportingUI_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
            PreparingErrorReportingFormForLoading(ConnectorSettingsAndResources.ErrorReportFormIcon,
                                                  ConnectorSettingsAndResources.ErrorReportUiHeadingText,
                                                  ConnectorSettingsAndResources.TellUsMoreText,
                                                  ConnectorSettingsAndResources.CheckBoxAutomaticallySend,
                                                  ConnectorSettingsAndResources.AlternativeIPText,
                                                  ErrorReportingUI.CommentsTextBox.Text,
                                                  errRep.ExceptionObject)
        End Sub

        Private Sub ErrorReportingUI_FormClosing(sender As Object, e As FormClosingEventArgs)
            UnBindErrorReportingUIEvents()
        End Sub

        ''' <summary>
        '''     'Clicking of ButtonSendErrorReport button
        ''' </summary>
        ''' <remarks>Just pass the text that will be appear in status bar and the their colors.</remarks>
        Private Sub ButtonSendErrorReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            ClickButtonSendErrorReport(ConnectorSettingsAndResources.PreparingReportText,
                                       ConnectorSettingsAndResources.TransferringReportText,
                                       ErrorReportingUITextColors.NormalStatusBarTextColor)
        End Sub

        ''' <summary>
        ''' Handle the change of CheckBoxAutoSend in ErrorReporterUI
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub CheckBoxAutoSend_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
            ConnectorSettingsAndResources.SaveFormSettings()
        End Sub

        ''' <summary>
        ''' Clicking event handler of CommentsTextBox text box. 
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub CommentsTextBox_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            'Select the content of the text box @ loading the form.
            If ErrorReportingUI.CommentsTextBox.Text.Contains(ConnectorSettingsAndResources.AddYourComments) Then
                SelectAllTextBoxTexts(ErrorReportingUI.CommentsTextBox)
            End If
        End Sub

        ''' <summary>
        ''' Key down event handler of "send us your comments" text box.
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub CommentsTextBox_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            SelectAllTextBoxTextsByHotKey(e, ErrorReportingUI.CommentsTextBox)
        End Sub

        Private Sub TextBoxErrorDesc_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
            SelectAllTextBoxTextsByHotKey(e, ErrorReportingUI.ErrorDescTextBox)
        End Sub

        ''' <summary>
        '''  Clicking "Attachment button" event handler of ErrorReporterUI. 
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub ButtonAttachment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            ShowAttachmentOpenFileDialog()
        End Sub

        ''' <summary>
        ''' Clicking "Dont send button" event handler of ErrorReporterUI 
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub DontSendButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            CloseForm(False)
        End Sub

        ''' <summary>
        ''' Send complete event handler of ErrorReporter 
        ''' </summary>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub ErrorReporter_SendCompleted(ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
            SendComplete(e,
                     ConnectorSettingsAndResources.SendingCanceledText,
                     ConnectorSettingsAndResources.SendingMailFailText,
                     ConnectorSettingsAndResources.CheckFirewallText,
                     ConnectorSettingsAndResources.ErrorSendingErrorReportText,
                     ConnectorSettingsAndResources.TransferringCompletedText)
        End Sub

        ''' <summary>
        ''' 'TimeOut timer for sending process.
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub ErrorReporter_TimeOut()
            TimeOut(ErrorReportingUITextColors.FailStatusBarTextColor, ConnectorSettingsAndResources.TransferringTimeOutText)
        End Sub

        ''' <summary>
        '''  'A timer for closing ErrorReportingUI automatically after a specified interval in ErrorReporter_SendCompleted
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Private Sub CloseTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs)
            DisableCloseTimer(sender)
            CloseForm(True)
        End Sub
#End Region
    End Class
End Namespace