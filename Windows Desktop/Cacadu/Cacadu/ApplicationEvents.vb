Imports Cacadu.IntelliProtectorService.IntelliProtector

Namespace My
    Delegate Sub errorRepoterDlg()
    ' The following events are available for MyApplication:

    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
    Partial Friend Class MyApplication
        Private Sub MyApplication_Startup(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) Handles Me.Startup
            CheckForExistingInstance()
            'Balora.Hodhod.STOW("Cacadu about to startup, showing splash screen...")
            If My.Settings.SS Then CacaduGloby.ShowSplashScreen()
            CheckForImportFile()
            'Init IntelliProtector
            'Balora.Hodhod.STOW("Init IntelliProtector...")
            If Init() Then
                WaitForInitialization()
                'If I prepare the Balora from MainForm connector, the path of LogFileName "CacaduErrorLog.txt" will not be set
                'so if any unhandled exception occur, i will get an error coz attachment isn't there.
                PrepareBaloraForCacadu()
                PrepareCacadoreForCacadu()
            End If
        End Sub

        Private Sub MyApplication_UnhandledException(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.UnhandledExceptionEventArgs) Handles Me.UnhandledException
            e.ExitApplication = False
            SendExceptionToConderella(e.Exception, True)
        End Sub

        Private Sub MyApplication_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown
            CacaduGloby.CloseSplashScreen()
            'Not shutting down Quartz will prevent Cacadu from closing completely.
            ShutdownQuartz()
            If Not IsNothing(SparkleAutoUpdater) Then SparkleAutoUpdater.Dispose()
            If Not IsNothing(ErrorReportingUI) AndAlso ErrorReportingUI.IsHandleCreated Then
                'Don't call "ErrorReportingUI.Dispose()" directly to avoid cross threading problems
                'if the form displayed from another thread. 
                Dim errorRepoterDelegateDlg As New errorRepoterDlg(AddressOf ErrorReportingUI.Dispose)
                ErrorReportingUI.Invoke(errorRepoterDelegateDlg)
            End If
            'AhSaElDin 20120102: disabled for now, as CacaSer is disabled.
            'If NamedPipesClient.IsConnected Then
            'NamedPipesClient.Disconnect()
            'NamedPipesClient.Dispose()
            'End If
        End Sub

        Public Sub CheckForExistingInstance()
#If Not Debug Then
            'Get number of processes of you program
            If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1 Then
                End
            End If
#End If
        End Sub
    End Class
End Namespace

