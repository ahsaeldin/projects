Imports NLog
Imports System.IO
Imports System.Net
Imports System.Drawing
Imports System.Reflection
Imports Balora.PathsHelper
Imports System.Windows.Forms
Imports System.Collections.ObjectModel

'Every thing about logging is here.
Public Class Hodhod

#Region "Old School *Obsolete*"
    Public Shared Sub WriteToErrorLog(ByVal msg As String, ByVal stkTrace As String, ByVal title As String)
        Dim parametersQueue As New Generic.Queue(Of String)
        parametersQueue.Enqueue(msg)
        parametersQueue.Enqueue(stkTrace)
        parametersQueue.Enqueue(title)

        _writeToErrorLog(parametersQueue)

#If DEBUG Then
        'Massive important while debug mode. It will remove the hood of filth automatically.
        Debugger.Break()
#End If

    End Sub

    Private Shared Sub _writeToErrorLog(parametersQueue As Generic.Queue(Of String))
        Dim msg As String = parametersQueue.Dequeue
        Dim stkTrace As String = parametersQueue.Dequeue
        Dim title As String = parametersQueue.Dequeue
        'Disable logging'''''''''''''''''''''''''''
        If Settings.LogErrors = False Then Exit Sub
        '''''''''''''''''''''''''''''''''''''''''''
        Dim errorLogFileName As String = Settings.LogFileName
        'Static errorListBox As ListBox
        If Settings.ShowLogWindow Then
            If Debugger.IsAttached Then
                'AhSaElDin 20111229: Disable log window for now.
                'AhSaEldin 20111221: calling the form in a background thread
                'Dim logFormDlg As MethodToBeInjectedDelegate = New MethodToBeInjectedDelegate(Function() LoggingForm(errorListBox, Util.GetFirstAssemblyCalledName & " Error List"))
                'Dim logFormBackgroundWorker As New BackgroundWorkerHelper
                'logFormBackgroundWorker.WorkToDo = logFormDlg
                'logFormBackgroundWorker.Start()
            End If
        Else
            'errorListBox = Nothing
        End If
        Try
            ''check the file
            Dim fs As FileStream = New FileStream(GetCurrentExecutingDirectory() & errorLogFileName, FileMode.OpenOrCreate, FileAccess.ReadWrite)
            Dim s As StreamWriter = New StreamWriter(fs)
            s.Close()
            fs.Close()
            'log it
            Dim fs1 As FileStream = New FileStream(GetCurrentExecutingDirectory() & errorLogFileName, FileMode.Append, FileAccess.Write)
            Dim s1 As StreamWriter = New StreamWriter(fs1)
            s1.Write("Title: " & title & " ")
            Dim dateTimeString = DateTime.Now.ToString()
            s1.Write("(Date/Time: " & dateTimeString & ")" & vbCrLf)
            s1.Write("" & vbCrLf)
            s1.Write("Error Description: " & msg & vbCrLf)
            s1.Write("" & vbCrLf)
            s1.Write("StackTrace:" & vbNewLine & stkTrace & vbCrLf)
            s1.Write("=============================================================" & vbCrLf)
            s1.Close()
            fs1.Close()
#If DEBUG Then
            Process.Start(GetCurrentExecutingDirectory() & errorLogFileName)
#End If
            'errorListBox.Items.Add("[" & dateTimeString & "]" & title & ", " & msg)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
        End Try
    End Sub

    Public Shared Function LoggingForm(ByRef entiresListBox As ListBox, windowText As String) As Form
        'Static logWindow As Form
        'If logWindow Is Nothing Then
        '    logWindow = New Form
        '    With logWindow
        '        .Show()
        '        .TopMost = True
        '        .ClientSize = New System.Drawing.Size(356, 343)
        '        .Location = New System.Drawing.Point(5, 39)
        '        .Text = windowText
        '    End With
        '    entiresListBox = New ListBox
        '    With entiresListBox
        '        .Location = New System.Drawing.Point(12, 12)
        '        .Size = New System.Drawing.Size(332, 325)
        '        .Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
        '                              Or System.Windows.Forms.AnchorStyles.Left) _
        '                              Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        '    End With
        '    logWindow.Controls.Add(entiresListBox)
        'End If
        'Return logWindow
        Return Nothing
    End Function
#End Region

#Region "Using NLog"
    Dim logger As Logger = LogManager.GetLogger("CacaduLogger")

    'A stub method for my instrument logic.
    <Conditional("DEBUG")>
    Public Shared Sub STOW(message As String) 'STOW == Send To Output Window.
        'NLog help contents: http://nlog-project.org/wiki/Tutorial
        'How to properly log exceptions? http://nlog-project.org/wiki/How_to_properly_log_exceptions%3F
        'How to configure logging in a component? http://nlog-project.org/wiki/How_to_configure_logging_in_a_component%3F
        'Sentinel - Log Viewer http://sentinel.codeplex.com
        'Log2Console - A Generic Log Viewer (for Log4Net, NLog...) http://log2console.codeplex.com/

        'No need to use console, you have a logger now.
        'Console.WriteLine(message)
        Dim _log As Logger = LogManager.GetCurrentClassLogger
        _log.Debug(message)
    End Sub
#End Region
End Class
