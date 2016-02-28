'Service Installer Class
Imports System.ServiceProcess
Imports System.Configuration.Install

#Region "References"
'Programmatically Install Window Service
'http://dotnetstep.blogspot.com/2009/06/programmatically-install-window-service.html
'ServiceProcessInstaller Class
'http://msdn.microsoft.com/en-us/library/system.serviceprocess.serviceprocessinstaller.aspx
'Writing Services Programmatically
'http://msdn.microsoft.com/en-us/library/aa984335%28v=vs.71%29.aspx
#End Region

Public Class Servnst
    Inherits BaloraBase

#Region "Fields"
    Private Shared _serviceControlValue As New ServiceController
#End Region

#Region "ServiceData"
    Interface IServiceData
        Property ServicePath As String
        Property UserName As String
        Property Password As String
        Property DisplayName As String
        Property Description As String
        Property ServiceName As String
        Property StartType As ServiceStartMode
        Property ServiceAccount As ServiceAccount
    End Interface

    ''' <summary>
    ''' All data need to be prepared before installing a service.
    ''' </summary>
    Class ServiceData
        Implements IServiceData
        Property ServicePath As String Implements IServiceData.ServicePath
        Property UserName As String Implements IServiceData.UserName
        Property Password As String Implements IServiceData.Password
        Property DisplayName As String Implements IServiceData.DisplayName
        Property Description As String Implements IServiceData.Description
        Property ServiceName As String Implements IServiceData.ServiceName
        Property StartType As ServiceStartMode Implements IServiceData.StartType
        Property ServiceAccount As ServiceAccount Implements IServiceData.ServiceAccount
    End Class
#End Region

#Region "Methods"
    ''' <summary>
    ''' Determines whether [is service installed] using [the specified service name].
    ''' </summary>
    ''' <param name="serviceName">Name of the service.</param><returns>
    '''   <c>true</c> if [is service installed] [the specified service name]; otherwise, <c>false</c>.
    ''' </returns>
    ''' <remarks>
    ''' For more information about ServiceController Class.
    ''' http://bit.ly/qzBKR9
    ''' </remarks>
    Public Shared Function IsServiceInstalled(serviceName As String) As Boolean
        Dim servStatus As ServiceControllerStatus = GetServiceStatus(serviceName)
        If servStatus <> Nothing Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Gets the service status.
    ''' </summary>
    ''' <param name="serviceName">Name of the service.</param><returns></returns>
    Public Shared Function GetServiceStatus(serviceName As String) As ServiceControllerStatus
        Dim controller As New ServiceController
        'The name of the computer that is running the service 
        'associated with this ServiceController instance. 
        'The default is the local computer (".").
        controller.MachineName = "."
        controller.ServiceName = serviceName
        Try
            Dim status As String = controller.Status.ToString()
            _serviceControlValue = controller
            Return controller.Status
        Catch ex As InvalidOperationException
            Return Nothing  'Service not found."
        End Try
    End Function

    ''' <summary>
    ''' Starts the service.
    ''' </summary>
    ''' <param name="serviceName">Name of the service.</param><returns></returns>
    Public Shared Function StartService(serviceName As String) As Boolean
        If IsServiceInstalled(serviceName) Then
            Try
                _serviceControlValue.Start()
                Return True
            Catch ex As Exception
                Alerter.REP("Error starting " & serviceName & " service", Nothing, False, True)
                Return False
            End Try
        Else
            Alerter.REP(serviceName & " service isn't installed", Nothing, False, True)
            Return False
        End If
    End Function

    ''' <summary>
    ''' Stops the service.
    ''' </summary>
    ''' <param name="serviceName">Name of the service.</param><returns></returns>
    Public Shared Function StopService(serviceName As String) As Boolean
        If IsServiceInstalled(serviceName) Then
            Try
                _serviceControlValue.Stop()
                Return True
            Catch ex As Exception
                Alerter.REP("Error stopping " & serviceName & " service", Nothing, False, True)
                Return False
            End Try
        Else
            Alerter.REP(serviceName & " service isn't installed", Nothing, False, True)
            Return False
        End If
    End Function

    ''' <summary>
    ''' Pauses the service.
    ''' </summary>
    ''' <param name="serviceName">Name of the service.</param><returns></returns>
    Public Shared Function PauseService(serviceName As String) As Boolean
        If IsServiceInstalled(serviceName) Then
            Try
                _serviceControlValue.Pause()
                Return True
            Catch ex As Exception
                Alerter.REP("Error pausing " & serviceName & " service", Nothing, False, True)
                Return False
            End Try
        Else
            Alerter.REP(serviceName & " service isn't installed", Nothing, False, True)
            Return False
        End If
    End Function

    ''' <summary>
    ''' Resumes the service.
    ''' </summary>
    ''' <param name="serviceName">Name of the service.</param><returns></returns>
    Public Shared Function ResumeService(serviceName As String) As Boolean
        If IsServiceInstalled(serviceName) Then
            Try
                _serviceControlValue.Continue()
                Return True
            Catch ex As Exception
                Alerter.REP("Error continuing " & serviceName & " service", Nothing, False, True)
                Return False
            End Try
        Else
            Alerter.REP(serviceName & " service isn't installed", Nothing, False, True)
            Return False
        End If
    End Function

    ''' <summary>
    ''' Installs the service.
    ''' </summary>
    ''' <param name="IServiceData">Instance of a class the implements IServiceData</param><returns></returns>
    Public Shared Function InstallService(IServiceData As IServiceData) As Boolean
        'Another way to install a service:
        'Dim cmdline As [String]() = {"asdfasdf"}
        'Dim Installer As New System.Configuration.Install.AssemblyInstaller(ExeFilename, cmdline)
        'Installer.UseNewContext = True
        'Installer.Install(Nothing)
        'Installer.Commit(Nothing)
        If IsServiceInstalled(IServiceData.ServiceName) Then
            Alerter.REP(IServiceData.ServiceName & " service already installed.", Nothing, False)
            Return False
        End If
        '==================================================================================================
        Dim ProcesServiceInstaller As New ServiceProcessInstaller()
        ProcesServiceInstaller.Account = IServiceData.ServiceAccount
        If IServiceData.UserName <> "" Then
            ProcesServiceInstaller.Username = IServiceData.UserName
            ProcesServiceInstaller.Password = IServiceData.Password
        End If
        '==================================================================================================
        Dim ServiceInstallerObj As New ServiceInstaller()
        Dim Context As InstallContext = New System.Configuration.Install.InstallContext()
        Dim path As [String] = [String].Format("/assemblypath={0}", IServiceData.ServicePath)
        Dim cmdline As [String]() = {path}
        Context = New System.Configuration.Install.InstallContext("", cmdline)
        '==================================================================================================
        With ServiceInstallerObj
            .Context = Context
            .DisplayName = IServiceData.DisplayName
            .Description = IServiceData.Description
            .ServiceName = IServiceData.ServiceName
            .StartType = IServiceData.StartType
            .Parent = ProcesServiceInstaller
        End With
        '==================================================================================================
        Dim state As New System.Collections.Specialized.ListDictionary()
        Try
            ServiceInstallerObj.Install(state)
            Return True
        Catch ex As Exception
            Alerter.REP("", ex, True)
            Return False
        End Try
        '==================================================================================================
    End Function

    ''' <summary>
    ''' Uninstalls the service.
    ''' </summary>
    ''' <param name="serviceName">Name of the service.</param><returns></returns>
    Public Shared Function UninstallService(serviceName As String) As Boolean
        'another way to uninstall a service:-
        'Dim cmdline As [String]() = {"asdfasdf"}
        'Dim Installer As New System.Configuration.Install.AssemblyInstaller(ExeFilename, cmdline)
        'Installer.UseNewContext = True
        'Installer.Uninstall(Nothing)
        If Not IsServiceInstalled(serviceName) Then
            Alerter.REP(serviceName & " service isn't installed.", Nothing, False)
            Return False
        End If
        Dim ServiceInstallerObj As New ServiceInstaller()
        'Dim Context As New InstallContext("<<log file path>>", Nothing)
        Dim Context As New InstallContext()
        ServiceInstallerObj.Context = Context
        ServiceInstallerObj.ServiceName = serviceName
        Try
            ServiceInstallerObj.Uninstall(Nothing)
            Return True
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("", ex, True)
            Return False
        End Try
    End Function
#End Region
End Class
