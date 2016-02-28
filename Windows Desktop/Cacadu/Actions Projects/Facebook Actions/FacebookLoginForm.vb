Imports Facebook
Imports System.Dynamic

Friend Class FacebookLoginForm

#Region "Constructors"

#Region "Fields"
    Private _fb As FacebookClient
    Private ReadOnly _loginUrl As Uri
    Private m_FacebookOAuthResult As FacebookOAuthResult
#End Region

    Public Sub New(appId As String, extendedPermissions As String)
        Me.New(New FacebookClient(), appId, extendedPermissions)
    End Sub

    Private Sub New(fb As FacebookClient, appId As String, extendedPermissions As String)
        If fb Is Nothing Then
            Throw New ArgumentNullException("fb")
        End If
        If String.IsNullOrWhiteSpace(appId) Then
            Throw New ArgumentNullException("appId")
        End If

        _fb = fb
        _loginUrl = GenerateLoginUrl(appId, extendedPermissions)

        InitializeComponent()
    End Sub
#End Region

#Region "Properties"
    Public Property FacebookOAuthResult() As FacebookOAuthResult
        Get
            Return m_FacebookOAuthResult
        End Get
        Private Set(value As FacebookOAuthResult)
            m_FacebookOAuthResult = value
        End Set
    End Property
#End Region

#Region "Event Handlers"
    Private Sub FacebookLoginForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ' make sure to use AbsoluteUri.
        fbWebBrowser.Navigate(_loginUrl.AbsoluteUri)
    End Sub

    Private Sub fbwebBrowser_Navigated(sender As System.Object, e As System.Windows.Forms.WebBrowserNavigatedEventArgs) Handles fbWebBrowser.Navigated
        ' whenever the browser navigates to a new url, try parsing the url.
        ' the url may be the result of OAuth 2.0 authentication.
        Dim oauthResult As FacebookOAuthResult = Nothing
        If _fb.TryParseOAuthCallbackUrl(e.Url, oauthResult) Then
            ' The url is the result of OAuth 2.0 authentication
            FacebookOAuthResult = oauthResult
            DialogResult = If(FacebookOAuthResult.IsSuccess, Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.No)
        Else
            ' The url is NOT the result of OAuth 2.0 authentication.
            FacebookOAuthResult = Nothing
        End If
    End Sub

    Private Sub FacebookLoginForm_Disposed(sender As Object, e As System.EventArgs) Handles Me.Disposed
        _fb = Nothing
        fbWebBrowser.Dispose()
    End Sub
#End Region

#Region "Methods"
    Private Function GenerateLoginUrl(appId As String, extendedPermissions As String) As Uri
        Dim parameters As Object = New ExpandoObject()
        parameters.client_id = appId
        parameters.redirect_uri = "https://www.facebook.com/connect/login_success.html"

        ' The requested response: an access token (token), an authorization code (code), or both (code token).
        parameters.response_type = "token"

        ' list of additional display modes can be found at http://developers.facebook.com/docs/reference/dialogs/#display
        parameters.display = "popup"

        ' add the 'scope' parameter only if we have extendedPermissions.
        If Not String.IsNullOrWhiteSpace(extendedPermissions) Then
            parameters.scope = extendedPermissions
        End If

        ' when the Form is loaded navigate to the login url.
        Return _fb.GetLoginUrl(parameters)
    End Function
#End Region

End Class