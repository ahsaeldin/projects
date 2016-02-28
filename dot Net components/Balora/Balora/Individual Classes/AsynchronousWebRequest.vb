Imports System
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.ComponentModel

'Class inspired by:
'http://www.informit.com/guides/content.aspx?g=dotnet&seqNum=582
'http://msdn.microsoft.com/query/dev10.query?appId=Dev10IDEF1&l=EN-US&k=k%28SYSTEM.NET.HTTPWEBREQUEST.BEGINGETRESPONSE%29;k%28TargetFrameworkMoniker-%22.NETFRAMEWORK%2cVERSION%3dV4.0%22%29;k%28DevLang-VB%29&rd=true

Friend Class AsynchronousWebRequest
    Inherits BaloraBase

    Private allDone As New ManualResetEvent(False)
    Private BUFFER_SIZE As Integer = 1024
    Private timeOutValue As Integer = 1000 * 60 * 1000

    Public Property TimeOut() As Integer
        Get
            Return timeOutValue
        End Get
        Set(ByVal value As Integer)
            timeOutValue = value
        End Set
    End Property

    'To check if request completed.
    Private isCompleteValue As Boolean = False
    Public Property IsComplete() As Boolean
        Get
            Return isCompleteValue
        End Get
        Set(ByVal value As Boolean)
            isCompleteValue = value
        End Set
    End Property

    'Returned HTML page content.
    Private contentValue As String
    Public Property Content() As String
        Get
            Return contentValue
        End Get
        Set(ByVal value As String)
            contentValue = value
        End Set
    End Property

    Private Class RequestState
        ' This class stores the State of the request.
        Private BUFFER_SIZE As Integer = 1024
        Public requestData As StringBuilder
        Public BufferRead() As Byte
        Public request As HttpWebRequest
        Public response As HttpWebResponse
        Public streamResponse As Stream

        Public Sub New()
            BufferRead = New Byte(BUFFER_SIZE) {}
            requestData = New StringBuilder("")
            request = Nothing
            streamResponse = Nothing
        End Sub 'New
    End Class 'RequestState

    Public Sub SendAsyncWebRequest(ByVal uri As String)
        Try
            ' Create a HttpWebrequest object to the desired URL. 
            Dim myHttpWebRequest As HttpWebRequest = CType(WebRequest.Create(uri), HttpWebRequest)

            ' Create an instance of the RequestState and assign the previous myHttpWebRequest
            ' object to its request field.  

            Dim myRequestState As New RequestState()
            myRequestState.request = myHttpWebRequest

            ' Start the asynchronous request.
            Dim result As IAsyncResult = CType(myHttpWebRequest.BeginGetResponse(New AsyncCallback(AddressOf ResponseCallback), myRequestState), IAsyncResult)

            ' this line implements the timeout, if there is a timeout, the callback fires and the request aborts.
            ThreadPool.RegisterWaitForSingleObject(result.AsyncWaitHandle, New WaitOrTimerCallback(AddressOf TimeoutCallback), myHttpWebRequest, TimeOut, True)

            ' The response came in the allowed time. The work processing will happen in the 
            ' callback function.
            allDone.WaitOne()

            ' Release the HttpWebResponse resource.
            myRequestState.response.Close()
        Catch e As Exception
            Alerter.REP("Exception during a web request", e, True)
        End Try
    End Sub

    Private Sub ResponseCallback(ByVal asynchronousResult As IAsyncResult)
        Try
            ' State of request is asynchronous.
            Dim myRequestState As RequestState = CType(asynchronousResult.AsyncState, RequestState)
            Dim myHttpWebRequest As HttpWebRequest = myRequestState.request
            myRequestState.response = CType(myHttpWebRequest.EndGetResponse(asynchronousResult), HttpWebResponse)

            ' Read the response into a Stream object.
            Dim responseStream As Stream = myRequestState.response.GetResponseStream()
            myRequestState.streamResponse = responseStream

            ' Begin the Reading of the contents of the HTML page and print it to the console.
            Dim asynchronousInputRead As IAsyncResult = responseStream.BeginRead(myRequestState.BufferRead, 0, 1024, New AsyncCallback(AddressOf ReadCallBack), myRequestState)
            Return
        Catch e As Exception
            Alerter.REP("Exception in response callback of a web request.", e, True)
        End Try
        allDone.Set()
    End Sub 'RespCallback

    Private Sub ReadCallBack(ByVal asyncResult As IAsyncResult)
        Try
            Dim myRequestState As RequestState = CType(asyncResult.AsyncState, RequestState)
            Dim responseStream As Stream = myRequestState.streamResponse
            Dim read As Integer = responseStream.EndRead(asyncResult)
            ' Read the HTML page and then print it to the console.
            If read > 0 Then
                myRequestState.requestData.Append(Encoding.ASCII.GetString(myRequestState.BufferRead, 0, read))
                Dim asynchronousResult As IAsyncResult = responseStream.BeginRead(myRequestState.BufferRead, 0, 1024, New AsyncCallback(AddressOf ReadCallBack), myRequestState)
                Return
            Else
                If myRequestState.requestData.Length > 1 Then
                    Content = myRequestState.requestData.ToString()
                End If
                responseStream.Close()
            End If

        Catch e As WebException
            Alerter.REP("Exception in read callback of a web request.", e, True)
        End Try
        allDone.Set()
    End Sub 'ReadCallBack 

    ' 2 minutes timeout
    ' Abort the request if the timer fires.
    Private Sub TimeoutCallback(ByVal state As Object, ByVal timedOut As Boolean)
        If timedOut Then
            Dim request As HttpWebRequest = CType(state, HttpWebRequest)

            If Not (request Is Nothing) Then
                request.Abort()
            End If
        End If
    End Sub 'TimeoutCallback
End Class
