Option Strict On
Option Explicit On

Imports Balora.PathsHelper

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'                                                   Balora Settings Class                                           
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Class Settings
    Inherits BaloraBase
    ' Settings class for the whole Balora.

    ' Let Balora user customize Error log file name.
    Private Shared logFileNameValue As String
    Public Shared Property LogFileName() As String
        Get
            Return logFileNameValue
        End Get
        Set(ByVal value As String)
            logFileNameValue = value
        End Set
    End Property

    Public Shared ReadOnly Property LogFilePath() As String
        Get
            Return GetCurrentExecutingDirectory() & Settings.LogFileName
        End Get
    End Property

    ' Enable/Disable error logging.
    Private Shared logErrorsValue As Boolean = True
    Public Shared Property LogErrors() As Boolean
        Get
            Return logErrorsValue
        End Get
        Set(ByVal value As Boolean)
            logErrorsValue = value
        End Set
    End Property

    ' Enable/Disable error logger windows.
    Private Shared showLogWindowValue As Boolean = True
    Public Shared Property ShowLogWindow() As Boolean
        Get
            Return showLogWindowValue
        End Get
        Set(ByVal value As Boolean)
            showLogWindowValue = value
        End Set
    End Property

    '<summary>
    'A hidden property used in Balora that will set by the calling
    'Application that will set it to a md5 value of a guid
    'The idea here
    'is to check the equity of a md5 value of a GUID specified here by me
    'And Cacadu will be the only who knows this vlaue and pass its hash to me.
    '</summary>
    '<remarks>
    'CAG means Calling Assembly Guid.
    'We make the property type of Object
    'rather than string to hide any hint
    ''about what CAG do
    '</remarks>
    Private Shared cAGValue As Object
    Public Shared Property CAG() As Object
        Get
            Return cAGValue
        End Get
        Set(ByVal value As Object)
            cAGValue = value
        End Set
    End Property

End Class

'Use to set Global Balora Events that will be processed 
'by caller such as exceptions sent by error reporter
'+Notes:- you can use AddNewError for passing exception
'+and messages from within Cacadore. And you can use 
'+Balora.Alerter.BoradCastExcpetion to display the error 
'+Reporter.
Public Class Alerter
    Inherits BaloraBase

#Region "Fields"
    Private Shared errorListValue As New List(Of ErrorEntry)
#End Region

#Region "Structures"
    'ErrorEntry
    'If you want to deliver a message from within Cacadore to outside,
    'like when you want to tell the user that he wants to add a duplicate 
    '"variable" to Globals class.
    '+Notes:- you can use AddNewError for passing exception
    '+and messages from within Cacadore. And you can use 
    '+Balora.Alerter.BoradCastExcpetion to display the error 
    '+Reporter.
    Public Structure ErrorEntry
        Dim exceptionObj As Exception
        'You can expose the message here rather than
        'hard coding a message box inside Cacadore, then
        'the caller may process the message as he wish...
        '(displaying a message box, status bar)
        Dim errorDescription As String
        Sub New(ByVal ex As Exception, description As String)
            exceptionObj = ex
            errorDescription = description
        End Sub
    End Structure
#End Region

#Region "Properties"
    'A public shared list of all errors.
    Public Shared ReadOnly Property ErrorList As System.Collections.Generic.List(Of ErrorEntry)
        Get
            Return errorListValue
        End Get
    End Property
#End Region

#Region "Methods"
    Shared Sub SendMessageToCacadu(message As String, Optional state As Object = Nothing)
        RaiseEvent AlertForm_MessageSent(message, state)
    End Sub

    Shared Sub ShowWaitWinodw(Optional description As String = "Loading")
        RaiseEvent WaitForm_Load(description)
    End Sub

    Shared Sub CloseWaitWinodw()
        RaiseEvent WaitForm_Close()
    End Sub

    ''' <summary>
    ''' RaiseUp error event for listeners with exception parameter and description 
    ''' and Boolean indicates if caller wants to display the error reporter or not.
    ''' then register the error in a list.
    ''' </summary>
    ''' Added justLog parameter for logging errors only without raising NewErrorOccur event.
    Shared Sub REP(d As String,
                   x As Object,
                   Optional bC As Boolean = False,
                   Optional jL As Boolean = False)
        raiseErrorUp(d, x, bC, jL)
    End Sub

    Private Shared Sub raiseErrorUp(errorDescription As String,
                                    ex As Object,
                                    Optional broadCast As Boolean = False,
                                    Optional justLog As Boolean = False)
        Dim _ex As Exception = CType(ex, Exception)
        AddNewError(errorDescription, _ex)
        STOWError(errorDescription)
        sendToErrorLogFile(errorDescription, _ex)
        onNewErrorOccur(errorDescription, _ex, broadCast, justLog)
    End Sub

    Private Shared Sub onNewErrorOccur(ByVal errorDescription As String, ByVal ex As Exception, ByVal broadCast As Boolean, ByVal justLog As Boolean)
        If Not justLog Then
            RaiseEvent NewErrorOccur(ex, errorDescription, broadCast)
        End If
    End Sub

    Private Shared Sub sendToErrorLogFile(ByVal errorDescription As String, ByVal ex As Exception)
        '"RaiseErrorUp::" just to mark the error item in error log window.
        If Not ex Is Nothing Then
            Hodhod.WriteToErrorLog(ex.Message, ex.StackTrace, "RaiseErrorUp::" & errorDescription)
        Else
            Hodhod.WriteToErrorLog("Not exception", "", "RaiseErrorUp::" & errorDescription)
        End If
    End Sub

    Private Shared Sub STOWError(ByVal errorDescription As String)
        If errorDescription <> "" Then Balora.Hodhod.STOW(errorDescription)
    End Sub

    Private Shared Sub AddNewError(ByVal errorDescription As String, ByVal ex As Exception)
        Dim newErrorEntry As New ErrorEntry(ex, errorDescription)
        errorListValue.Add(newErrorEntry)
    End Sub

    Public Shared Sub LoadAlertForm(caption As String,
                                      message As String,
                                      tag As String,
                                      showButtons As Boolean,
                                      image As Object,
                                      state As Object)
        RaiseEvent AlertForm_Load(caption, message, tag, showButtons, image, state)
    End Sub
#End Region

#Region "Events"
    Public Shared Event AlertForm_MessageSent(message As String, state As Object)

    Public Shared Event AlertForm_Load(caption As String,
                                       message As String,
                                       tag As String,
                                       showButtons As Boolean,
                                       image As Object,
                                       state As Object)
    Public Shared Event NewErrorOccur(ByVal e As System.Exception,
                                      description As String,
                                      broadCast As Boolean)
    Public Shared Event WaitForm_Load(description As String)
    Public Shared Event WaitForm_Close()
#End Region

End Class