Imports System.Management

#Region "References"
'Windows Management Instrumentation
'   http://en.wikipedia.org/wiki/Windows_Management_Instrumentation
'Shut down system using C#
'   http://www.geekpedia.com/code36_Shut-down-system-using-Csharp.html
'How To Get Hardware Information (CPU ID, MainBoard Info, Hard Disk Serial, System Information , ...) "Best one"
'    http://www.codeproject.com/Articles/17973/How-To-Get-Hardware-Information-CPU-ID-MainBoard-I 
'How to shutdown the computer from C#
'   http://stackoverflow.com/questions/102567/how-to-shutdown-the-computer-from-c-sharp
'C# put pc to sleep or hibernate
'   http://stackoverflow.com/questions/2079813/c-sharp-put-pc-to-sleep-or-hibernate
'How to shutdown a locked PC (vb.net)?
'   http://stackoverflow.com/questions/3699703/how-to-shutdown-a-locked-pc-vb-net
'Shutdown/Restart VB.Net
'   http://www.codeproject.com/Articles/12815/Shutdown-Restart-VB-Net
'How to Shutdown, Restart, Log off your computer with VB.NET or C#
'   http://www.brangle.com/wordpress/2009/08/how-to-shutdown-restart-log-off-your-computer-with-vb-net-or-c
'Lock, Logoff, Reboot, Shutdown, Hibernate, Standby in .Net
'   http://www.c-sharpcorner.com/UploadFile/thiagu304/desktopfunctions02112007140806PM/desktopfunctions.aspx
'C# put pc to sleep or hibernate
'   http://stackoverflow.com/questions/2079813/c-sharp-put-pc-to-sleep-or-hibernate
'Application.SetSuspendState Method
'   http://msdn.microsoft.com/en-us/library/system.windows.forms.application.setsuspendstate.aspx
'Standby or Hibernate Windows Programmatically
'   http://www.blackwasp.co.uk/StandbyHibernate.aspx
'FIX: You receive a "Privilege not held" error message when you execute WMI operations by using the System.Management namespace in the .NET Framework 1.0 Service Pack 3
'   http://support.microsoft.com/kb/891864
#End Region

<Serializable()>
Public Class SysManage
    Inherits BaloraBase

    'Overload جامع if you want use it once.
    Public Function Shutdown(shutdownType As ShutdownTypeEnum) As Boolean
        Return _shutdown(shutdownType)
    End Function

    Public Function Shutdown() As Boolean
        Return _shutdown(ShutdownTypeEnum.Shutdown)
    End Function

    Public Function LogOff() As Boolean
        Return _shutdown(ShutdownTypeEnum.LogOff)
    End Function

    Public Function Restart() As Boolean
        Return _shutdown(ShutdownTypeEnum.Restart)
    End Function

    Public Function ForceLogOff() As Boolean
        Return _shutdown(ShutdownTypeEnum.ForcedLogOff)
    End Function

    Public Function ForceShutdown() As Boolean
        Return _shutdown(ShutdownTypeEnum.ForcedShutdown)
    End Function

    Public Function ForceReboot() As Boolean
        Return _shutdown(ShutdownTypeEnum.ForcedRestart)
    End Function

    Public Function Sleep() As Boolean
        Return _shutdown(ShutdownTypeEnum.Sleep)
    End Function

    Public Function Hibernate() As Boolean
        Return _shutdown(ShutdownTypeEnum.Hibernate)
    End Function

    Public Function ForceHibernate() As Boolean
        Return _shutdown(ShutdownTypeEnum.ForcedHibernate)
    End Function

    Public Function ForceSleep() As Boolean
        Return _shutdown(ShutdownTypeEnum.ForcedSleep)
    End Function

    Public Function CancelShutdown() As Boolean
        Return _cancelShutdown()
    End Function

    Private Function _shutdown(shutdownType As ShutdownTypeEnum) As Boolean
        If _handleSuspendState(shutdownType) Then Return True

        Return _shutdownUsingWMI(shutdownType)
    End Function

    Private Function _cancelShutdown() As Boolean
        Dim _process As Process = Nothing
        _process = Process.Start("shutdown.exe", "-a")
        Return Not IsNothing(_process)
    End Function

    Private Function _handleSuspendState(shutdownType As ShutdownTypeEnum) As Boolean
        Select Case shutdownType
            Case ShutdownTypeEnum.Sleep
                Return Win32API.SetSuspendState(False, False, True)
            Case ShutdownTypeEnum.Hibernate
                Return Win32API.SetSuspendState(True, False, True)
            Case ShutdownTypeEnum.ForcedSleep
                Return Win32API.SetSuspendState(False, True, True)
            Case ShutdownTypeEnum.ForcedHibernate
                Return Win32API.SetSuspendState(True, True, True)
            Case Else
                Return False
        End Select
    End Function

    Private Function _shutdownUsingShutdownCommand(ByVal shutdownType As ShutdownTypeEnum) As Boolean
        Dim _process As Process = Nothing
        Select Case shutdownType
            Case ShutdownTypeEnum.LogOff
                _process = Process.Start("shutdown.exe", "-l -t 0")
            Case ShutdownTypeEnum.Shutdown
                _process = Process.Start("shutdown.exe", "-s -t 0")
            Case ShutdownTypeEnum.Restart
                _process = Process.Start("shutdown.exe", "-r -t 0")
            Case ShutdownTypeEnum.ForcedLogOff
                _process = Process.Start("shutdown.exe", "-l -t 0")
            Case ShutdownTypeEnum.ForcedShutdown
                _process = Process.Start("shutdown.exe", "-s -f -t 0")
            Case ShutdownTypeEnum.ForcedRestart
                _process = Process.Start("shutdown.exe", "-s -f -t 0")
            Case Else
                Return False
        End Select
        Return Not IsNothing(_process)
    End Function

    Private Function _shutdownUsingExitWindowEx(ByVal shutdownType As ShutdownTypeEnum) As Boolean
        Dim _process As Process = Nothing
        Select Case shutdownType
            Case ShutdownTypeEnum.LogOff
                Return Win32API.ExitWindowsEx(ExitWindows.LogOff, ShutdownReason.MajorOther)
            Case ShutdownTypeEnum.Shutdown
                Return Win32API.ExitWindowsEx(ExitWindows.ShutDown, ShutdownReason.MajorOther)
            Case ShutdownTypeEnum.Restart
                Return Win32API.ExitWindowsEx(ExitWindows.Reboot, ShutdownReason.MajorOther)
            Case ShutdownTypeEnum.ForcedLogOff
                Return Win32API.ExitWindowsEx(ExitWindows.Force, ShutdownReason.MajorOther)
            Case ShutdownTypeEnum.ForcedShutdown
                Return Win32API.ExitWindowsEx(CType(4, ExitWindows), ShutdownReason.MajorOther)
            Case ShutdownTypeEnum.ForcedRestart
                Return Win32API.ExitWindowsEx(CType(6, ExitWindows), ShutdownReason.MajorOther)
            Case Else
                Return False
        End Select
    End Function

    Private Function _shutdownUsingWMI(ByVal shutdownType As ShutdownTypeEnum) As Boolean
        Dim _managementBaseObject = createManagementBaseObject()
        Dim _win32ManagementClass = createWin32_OperatingSystemManagementClass()
        Dim mboShutdownParams As ManagementBaseObject = setShutDownParameters(_win32ManagementClass, shutdownType)
        invokeShutdownMethod(_managementBaseObject, _win32ManagementClass, mboShutdownParams)
        'it will return false only it doesn't work.
        Return False
    End Function

    Private Function createManagementBaseObject() As ManagementBaseObject
        Dim _managementBaseObject As ManagementBaseObject = Nothing
        Return _managementBaseObject
    End Function

    Private Function createWin32_OperatingSystemManagementClass() As ManagementClass
        Dim _win32ManagementClass As New ManagementClass("Win32_OperatingSystem")
        Return _win32ManagementClass
    End Function

    Private Function setShutDownParameters(ByRef win32ManagementClass As ManagementClass, ByVal shutdownType As ShutdownTypeEnum) As ManagementBaseObject
        win32ManagementClass.[Get]()
        ' You can't shutdown without security privileges
        win32ManagementClass.Scope.Options.EnablePrivileges = True
        Dim mboShutdownParams As ManagementBaseObject = win32ManagementClass.GetMethodParameters("Win32Shutdown")
        mboShutdownParams("Flags") = shutdownType
        mboShutdownParams("Reserved") = "0"
        Return mboShutdownParams
    End Function

    Private Sub invokeShutdownMethod(ByVal _managementBaseObject As ManagementBaseObject,
                                            ByVal _win32ManagementClass As ManagementClass,
                                            ByVal mboShutdownParams As ManagementBaseObject)
        For Each _managementObject As ManagementObject In _win32ManagementClass.GetInstances()
            _managementBaseObject = _managementObject.InvokeMethod("Win32Shutdown", mboShutdownParams, Nothing)
        Next
    End Sub
End Class

