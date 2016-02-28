Imports Facebook
Imports System.Windows.Forms
Imports Actions.FacebookActions.LastFacebookNotification
Imports System.Runtime.Serialization

<Serializable()>
Public Class FbHelper
    Implements IDisposable

    Enum ErrorType
        NoError
        Cancelled
        Failed
        Generic
    End Enum

    Structure ResultMessage
        Dim result As ErrorType
        Dim message As String
    End Structure

#Region "Fields"
    'Declare it shared to avoid re-login and to allow logout using actons's options of logout.
    Private Shared _accessToken As String
    Private _userNameValue As String
    Private _lastNotificationURLValue As String
    Private _lastNotificationTitleValue As String
    Private _userProfilePictureValue As System.Drawing.Image
    Private _lastNotificationIconValue As System.Drawing.Image

    Private _appId As String = "235098143285148"
    Private _appSecret As String = "582c8474c6c4e707953e08880c363a2f"
    Private _extendedPermissions As String = "user_about_me,publish_stream,manage_notifications"
#End Region

#Region "Properties"
    Public Shared ReadOnly Property IsLogon As Boolean
        Get
            If _accessToken <> vbNullString Then Return True
            Return False
        End Get
    End Property

    Public ReadOnly Property UserProfilePicture() As System.Drawing.Image
        Get
            Return _userProfilePictureValue
        End Get
    End Property

    Public ReadOnly Property lastNotificationURL As String
        Get
            Return _lastNotificationURLValue
        End Get
    End Property

    Public ReadOnly Property LastNotificationTitle As String
        Get
            Return _lastNotificationTitleValue
        End Get
    End Property

    Public ReadOnly Property LastNotificationIcon As System.Drawing.Image
        Get
            Return _lastNotificationIconValue
        End Get
    End Property

    Public ReadOnly Property UserName As String
        Get
            Return _userNameValue
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub Login(Optional waitHandle As Threading.ManualResetEvent = Nothing)
        Dim resultStruct As New ResultMessage
        resultStruct.result = ErrorType.NoError
        resultStruct.message = ""

        Dim _manualRest As New Threading.ManualResetEvent(False)
        'We use a single thread because the WebBrowser in the login form
        'needs to be called from a single thread otherwise it raises up
        'this error from InitializeComponent method.
        '"ActiveX control e88121a3-9fa9-11cf-9d9f-00aa003a3aa3 cannot be instantiated because the current thread is not a single-threaded
        Dim singleApartmentThread As New Threading.Thread(Sub()
                                                              If _accessToken <> vbNullString Then
                                                                  If Not IsNothing(waitHandle) Then
                                                                      'Because I need to check first if handle is wait or not before calling set to avoid exception, I called WaitOne(0)
                                                                      'http://stackoverflow.com/questions/3313093/how-to-check-if-the-waithandle-was-set
                                                                      waitHandle.WaitOne(0)
                                                                  End If
                                                                  _manualRest.Set()
                                                                  _manualRest.Dispose()
                                                                  _manualRest = Nothing
                                                                  Exit Sub
                                                              End If

                                                              'open the Facebook Login Dialog and ask for user permissions.
                                                              Dim fbLoginDlg = New FacebookLoginForm(_appId, _extendedPermissions)

                                                              If fbLoginDlg.InvokeRequired Then
                                                                  fbLoginDlg.BeginInvoke(Sub() fbLoginDlg.ShowDialog())
                                                              Else
                                                                  fbLoginDlg.TopMost = True
                                                                  fbLoginDlg.TopLevel = True
                                                                  fbLoginDlg.ShowInTaskbar = False
                                                                  fbLoginDlg.ShowDialog()
                                                              End If

                                                              resultStruct = initLogin(fbLoginDlg.FacebookOAuthResult)

                                                              'Unfortunately, we must use it because fbLoginDlg 
                                                              'opened in another thread, so if you called it from
                                                              'the background worker, the thread won't wait till
                                                              'getting the access_token therefore nothing will
                                                              'work after that like getting last notification.
                                                              If Not IsNothing(waitHandle) Then
                                                                  'Because I need to check first if handle is wait or not before calling set to avoid exception, I called WaitOne(0)
                                                                  'http://stackoverflow.com/questions/3313093/how-to-check-if-the-waithandle-was-set
                                                                  waitHandle.WaitOne(0)
                                                              End If
                                                              _manualRest.Set()
                                                              _manualRest.Dispose()
                                                              _manualRest = Nothing
                                                          End Sub)
        singleApartmentThread.TrySetApartmentState(Threading.ApartmentState.STA)
        singleApartmentThread.Start()
        'Unfortunately, we must use it because fbLoginDlg 
        'opened in another thread, so if you called it from
        'the background worker, the thread won't wait till
        'getting the access_token therefore nothing will
        'work after that like getting last notification.
        _manualRest.WaitOne(50000) 'We set timeout to 50 seconds to give him an enough period to write his username and password.
        Select Case resultStruct.result
            Case ErrorType.Cancelled
                Throw New FacebookLogonCancelled("Facebook logon cancelled.")
            Case ErrorType.Failed
                Throw New FacebookLogonFailed(resultStruct.message)
            Case ErrorType.Generic
                Throw New FacebookGenericError(resultStruct.message)
        End Select
    End Sub

    Private Function initLogin(facebookOAuthResult As Facebook.FacebookOAuthResult) As ResultMessage
        If facebookOAuthResult Is Nothing Then
            ' the user closed the FacebookLoginDialog, so do nothing.
            'Balora.Alerter.SendMessageToCacadu("AlertOnError", "Facebook logon cancelled.")
            Dim errMessage As ResultMessage
            errMessage.result = ErrorType.Cancelled
            errMessage.message = "Facebook logon cancelled."
            Return errMessage 'cancelled
        End If

        ' Even though facebookOAuthResult is not null, it could had been an 
        ' OAuth 2.0 error, so make sure to check IsSuccess property always.
        If facebookOAuthResult.IsSuccess Then
            ' since our respone_type in FacebookLoginDialog was token,
            ' we got the access_token
            ' The user now has successfully granted permission to our app.
            _accessToken = facebookOAuthResult.AccessToken
            Return DoAfterLogonFetch()
            ' GraphApiParametersInPathExample()
        Else
            ' for some reason we failed to get the access token.
            ' most likely the user clicked don't allow.
            Dim errMessage As ResultMessage
            errMessage.message = facebookOAuthResult.ErrorDescription
            errMessage.result = ErrorType.Failed
            'Balora.Alerter.SendMessageToCacadu("AlertOnError", facebookOAuthResult.ErrorDescription)
            Return errMessage
        End If
    End Function

    Private Function DoAfterLogonFetch() As ResultMessage
        Return GetUserProfilePictureAndUserName()
    End Function

    Public Sub LogOut()
        Dim _manualRest As New Threading.ManualResetEvent(False)
        Dim fb As FacebookClient = New FacebookClient()

        Dim logoutUrl = fb.GetLogoutUrl(New With {Key .[next] = "https://www.facebook.com/connect/login_success.html", _
                                                  Key .access_token = _accessToken})

        Dim singleApartmentThread As New Threading.Thread(Sub()
                                                              Dim webBrowser = New WebBrowser()
                                                              AddHandler webBrowser.Navigated, Function(o, args)
                                                                                                   If args.Url.AbsoluteUri = "https://www.facebook.com/connect/login_success.html" Then
                                                                                                       Balora.Alerter.SendMessageToCacadu("FacebookLogOutSuccess", Me)
                                                                                                   End If
                                                                                                   _manualRest.Set()
                                                                                                   _manualRest.Dispose()
                                                                                                   _manualRest = Nothing
                                                                                                   Return True
                                                                                               End Function
                                                              webBrowser.Navigate(logoutUrl.AbsoluteUri)
                                                          End Sub)

        singleApartmentThread.TrySetApartmentState(Threading.ApartmentState.STA)
        singleApartmentThread.Start()
        _manualRest.WaitOne(10000)
        fb = Nothing
    End Sub
#End Region

#Region "Helper Methods"
    Public Function GetUserProfilePictureAndUserName() As ResultMessage
        Dim _waitHandle As New Threading.ManualResetEvent(False)
        Dim errMessage As New ResultMessage
        errMessage.result = ErrorType.NoError
        If _accessToken = vbNullString Then Return errMessage
        ' avoid using XAsync methods as it is marked obsolete.
        ' use XAsync only for .net 3.5/SL4/WP7
        ' use XTaskAsync methods instead.
        Dim fb = New FacebookClient(_accessToken)

        ' make sure to add the appropriate event handler
        ' before calling the async methods.
        ' GetCompleted     => GetAsync
        ' PostCompleted    => PostAsync
        ' DeleteCompleted  => DeleteAsync
        AddHandler fb.GetCompleted, Sub(o, e)
                                        ' incase you support cancellation, make sure to check
                                        ' e.Cancelled property first even before checking (e.Error != null).
                                        ' for this example, we can ignore as we don't allow this
                                        ' example to be cancelled.
                                        If e.Cancelled Then
                                        ElseIf e.[Error] IsNot Nothing Then
                                            ' error occurred
                                            'Balora.Alerter.SendMessageToCacadu("AlertOnError", MessageBox.Show(e.[Error].Message))
                                            errMessage.result = ErrorType.Generic
                                            errMessage.message = e.[Error].Message
                                        Else
                                            ' the request was completed successfully

                                            ' now we can either cast it to IDictionary<string, object> or IList<object>
                                            ' depending on the type.
                                            ' For this example, we know that it is IDictionary<string,object>.
                                            Dim result = DirectCast(e.GetResultData(), IDictionary(Of String, Object))

                                            _userNameValue = DirectCast(result("first_name"), String) & " " & DirectCast(result("last_name"), String)

                                            _userProfilePictureValue = FetchUserPicture(DirectCast(result("id"), String), "square")

                                            ' since this is an async callback, make sure to be on the right thread
                                            ' when working with the UI.
                                            'Me.BeginInvoke(New MethodInvoker(Function()
                                            '                                     lblFirstName.Text = "First Name: " & firstName
                                            '                                 End Function))
                                            Dim state As New Dictionary(Of String, Object)
                                            state.Add("photourl", Me.UserProfilePicture)
                                            state.Add("user_name", _userNameValue)
                                            Balora.Alerter.SendMessageToCacadu("FacebookLoginSuccess", state)
                                        End If
                                        _waitHandle.Set()
                                        _waitHandle.Dispose()
                                        _waitHandle = Nothing
                                    End Sub
        ' additional parameters can be passed and 
        ' must be assignable from IDictionary<string, object> or anonymous object
        Dim parameters = New Dictionary(Of String, Object)()
        parameters("fields") = "id,first_name,last_name"
        fb.GetTaskAsync("me", parameters)
        _waitHandle.WaitOne(10000)

        'Or
        'fb.GetAsync("me", new { fields = new[] { "first_name", "last_name" } });
        Return errMessage
    End Function

    Public Sub DisplayLastNotification(Optional ByVal state As LastFacebookNotificationOutputs = Nothing,
                                        Optional waitHandle As Threading.ManualResetEvent = Nothing)
        'Use state parameter to receive the notification message string by ref, so an external caller method can
        'get the result. Just like (LastFacebookNotificationOutputs)
        Dim errMessage = DisplayLastNotificationByFqlAsync(state, waitHandle)
        Select Case errMessage.result
            Case ErrorType.Cancelled
                Throw New FacebookLogonCancelled(errMessage.message)
            Case ErrorType.Generic
                Throw New FacebookGenericError(errMessage.message)
            Case ErrorType.Failed
                Throw New FacebookLogonFailed(errMessage.message)
        End Select
    End Sub

    Private Function DisplayLastNotificationByFqlAsync(Optional ByVal state As LastFacebookNotificationOutputs = Nothing,
                                                       Optional waitHandle As Threading.ManualResetEvent = Nothing) As ResultMessage
        Dim _waitHandle As New Threading.ManualResetEvent(False)
        Dim errMessage As New ResultMessage
        errMessage.result = ErrorType.NoError
        If _accessToken = vbNullString Then Return errMessage
        'Use state parameter to receive the notification message string by ref, so an external caller method can
        'get the result. Just like (LastFacebookNotificationOutputs)
        Dim fb = New FacebookClient(_accessToken)

        ' since FQL is internally a GET request,
        ' make sure to add the GET event handler.
        AddHandler fb.GetCompleted, Function(o, e)
                                        ' incase you support cancellation, make sure to check
                                        ' e.Cancelled property first even before checking (e.Error!=null).
                                        If e.Cancelled Then
                                            ' for this example, we can ignore as we don't allow this
                                            ' example to be cancelled.

                                            ' you can check e.Error for reasons behind the cancellation.
                                            Dim cancellationError = e.[Error]
                                            ' cancel occurred
                                            'Me.    (New MethodInvoker(Function()
                                            '                                     MessageBox.Show(e.[Error].Message)
                                            '                                 End Function))
                                            'Balora.Alerter.SendMessageToCacadu("AlertOnError", MessageBox.Show(e.[Error].Message))
                                            errMessage.result = ErrorType.Cancelled
                                            errMessage.message = e.[Error].Message
                                            _waitHandle.Set()
                                            _waitHandle.Dispose()
                                            _waitHandle = Nothing
                                            Return False
                                        ElseIf e.[Error] IsNot Nothing Then
                                            ' error occurred
                                            'Me.BeginInvoke(New MethodInvoker(Function()
                                            '                                     MessageBox.Show(e.[Error].Message)
                                            '                                 End Function))
                                            'Balora.Alerter.SendMessageToCacadu("AlertOnError", MessageBox.Show(e.[Error].Message))
                                            errMessage.result = ErrorType.Generic
                                            errMessage.message = e.[Error].Message
                                            _waitHandle.Set()
                                            _waitHandle.Dispose()
                                            _waitHandle = Nothing
                                            Return False
                                        Else
                                            ' the request was completed successfully

                                            ' now we can either cast it to IDictionary<string, object> or IList<object>
                                            ' depending on the type. or we could use dynamic.
                                            Dim result = DirectCast(e.GetResultData(), IDictionary(Of String, Object))
                                            Dim data = DirectCast(result("data"), IList(Of Object))
                                            If data.Count > 0 Then
                                                Dim lastNotificationObj = data(0)
                                                _lastNotificationURLValue = lastNotificationObj("href")
                                                _lastNotificationTitleValue = lastNotificationObj("title_text")
                                                _lastNotificationIconValue = FetchUserPicture(lastNotificationObj("sender_id"), "small")
                                                Balora.Alerter.SendMessageToCacadu("lastNotificationTitleHtml", Me)
                                                If Not IsNothing(waitHandle) Then waitHandle.Set()
                                                If Not IsNothing(state) AndAlso Not state.SaveToMember = vbNullString Then
                                                    state.ParentAction.SetOutput(state.SaveToMember, _lastNotificationTitleValue)
                                                End If
                                            End If
                                            _waitHandle.Set()
                                            _waitHandle.Dispose()
                                            _waitHandle = Nothing
                                            Return data

                                            'Since this is an async callback, make sure to be on the right thread
                                            ' when working with the UI.
                                            'Me.BeginInvoke(New MethodInvoker(Function()
                                            '                                     lblTotalFriends.Text = String.Format("You have {0} friend(s).", count)
                                            '                                 End Function))
                                            'Return True
                                        End If
                                    End Function

        'https://developers.facebook.com/docs/reference/fql/notification/
        Dim query As String = String.Format("SELECT title_text,href,icon_url,sender_id FROM notification WHERE recipient_id={0} AND is_unread = 1 AND is_hidden = 0", "me()")

        'SELECT notification_id, sender_id, title_html, body_html, href
        'FROM notification
        'WHERE recipient_id=me()
        'AND is_unread = 1
        'AND is_hidden = 0
        'AND created_time > <time>

        ' call the Query or QueryAsync method to execute a single fql query.
        fb.GetTaskAsync("fql", New With {Key .q = query})
        _waitHandle.WaitOne(10000)
        Return errMessage
    End Function

    Private Function FetchUserPicture(user_id As String, type As String) As System.Drawing.Image
        Dim img As System.Drawing.Image = Nothing
        If user_id <> vbNullString Then
            'Find more info about type here
            'https://developers.facebook.com/docs/reference/api/
            Dim profilePictureURL = String.Format("https://graph.facebook.com/{0}/picture?type={1}", user_id, type)
            img = Balora.Util.LoadImageFromUrl(profilePictureURL)
        End If
        Return img
    End Function
#End Region

#Region "Sample Methods For Future Development"
    Private Sub GraphApiParametersInPathExample()
        If _accessToken = vbNullString Then Exit Sub
        ' rather then creating a new object for parameter
        ' you can also embed simple parameters as part of the path.

        Try
            Dim fb = New FacebookClient(_accessToken)

            'Dim result As Object = fb.[Get]("me?fields=first_name,last_name")
            Dim result As Object = fb.[Get]("me/home?access_token=" & _accessToken)

            Dim firstName As Object = result.first_name

            'This is especially useful for paged data (result.paging.next and result.paging.previous)
            'and your path can also contain the full graph url (https://graph.facebook.com/"
            'var nextPath = "https://graph.facebook.com/me/likes?limit=3&access_token=xxxxxxxxxxx&offset=3";
            'dynamic nextResult = fb.Get(nextPath);
            Dim lastName As Object = result.last_name
        Catch ex As FacebookApiException
            MessageBox.Show(ex.Message)
        End Try
    End Sub
#End Region

#Region "Custom Exception Classes"
    <Serializable()> _
    Public Class FacebookLogonCancelled
        Inherits Exception

        Public Sub New()
            MyBase.New()
        End Sub
        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub
        Public Sub New(ByVal message As String, ByVal inner_exception As Exception)
            MyBase.New(message, inner_exception)
        End Sub
        Public Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
            MyBase.New(info, context)
        End Sub
    End Class

    <Serializable()> _
    Public Class FacebookLogonFailed
        Inherits Exception

        Public Sub New()
            MyBase.New()
        End Sub
        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub
        Public Sub New(ByVal message As String, ByVal inner_exception As Exception)
            MyBase.New(message, inner_exception)
        End Sub
        Public Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
            MyBase.New(info, context)
        End Sub
    End Class

    <Serializable()> _
    Public Class FacebookGenericError
        Inherits Exception

        Public Sub New()
            MyBase.New()
        End Sub
        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub
        Public Sub New(ByVal message As String, ByVal inner_exception As Exception)
            MyBase.New(message, inner_exception)
        End Sub
        Public Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)
            MyBase.New(info, context)
        End Sub
    End Class
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            _userProfilePictureValue.Dispose()
            _lastNotificationIconValue.Dispose()
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

