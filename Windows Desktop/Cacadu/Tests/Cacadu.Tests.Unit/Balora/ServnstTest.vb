Option Explicit Off
Option Strict Off

Imports Balora
Imports NUnit.Framework
Imports Balora.PathsHelper
Imports System.ServiceProcess


Namespace BaloraUT
    <TestFixture()> _
    Public Class ServnstTest

        Private _serivceNameToBeTested As String = "CacaSer"
        Private _serviceFileToBeTested As String = vbNullString

#If CONFIG = "Release" Then
        Dim exeDir As String = "D:\Rocknee\KV\Cacadu\Code\Trunk\Cacadu\bin\Release\"
#ElseIf CONFIG = "Debug" Then
        Dim exeDir As String = "D:\Rocknee\KV\Cacadu\Code\Trunk\Cacadu\bin\Debug\"
#End If

        'Use this property to force not to uninstall the service in TearDownMethods,
        'in case you want to debug the service in the IDE
        Private Property DoNotUninstall As Boolean = False

        <TestFixtureSetUp()> _
        Public Sub SetupMethods()

            Dim current As String = GetCurrentExecutingDirectory()
            If Not current.Contains("Rocknee") Then
                exeDir = current
            End If

            Balora.Hodhod.STOW("#######################################################")
            PrepareCacaduComponents()
            Dim currentServiceExeFilePath As String = exeDir & "CacaSer.exe"
            Dim newTempServiceExeFilePath As String = exeDir & "CacaSer" & DateTime.Now.Millisecond & ".exe"
            Try
                FileCopy(currentServiceExeFilePath, newTempServiceExeFilePath)
            Catch ex As Exception

                Assert.Inconclusive("Could not find Cacaser.exe file!")
            End Try
            _serviceFileToBeTested = newTempServiceExeFilePath
        End Sub

        <SetUp()> _
        Public Sub SetupTest()
            Balora.Hodhod.STOW("#######################################################")
            Console.WriteLine("Checking if the service is not installed...")
            If Not Servnst.IsServiceInstalled(_serivceNameToBeTested) Then
                Console.WriteLine("Service is not installed, installing service...")
                InstallServiceTest()
            End If
        End Sub

        <TestFixtureTearDown()> _
        Public Sub TearDownMethods()
            'Use this property 'DoNotUninstall' to force not to uninstall the service in TearDownMethods,
            'in case you want to debug the service in the IDE
            If Not DoNotUninstall Then
                Balora.Hodhod.STOW("#######################################################")
                Console.WriteLine("Test session ended, now uninstalling the service...")
                Servnst.UninstallService(_serivceNameToBeTested)
                'DeleteAllTempServicesFiles()
                Balora.Hodhod.STOW("#######################################################")
            End If
        End Sub

        <TearDown()> _
        Public Sub TearDownTest()

        End Sub

#Region "Debug CacaSer"

        ''' <summary>
        ''' Installs the service without uninstall.
        ''' </summary>
        ''' <remarks>
        ''' Use to debug OnStart service method (run this method then run the service and debugger will lanuch.)
        ''' </remarks>
        <Test()> _
        Public Sub DEBUG_InstallService()

            Dim actual As Boolean = False

            Dim isUninstalled As Boolean = True

            DoNotUninstall = True

            Dim expected As Boolean = True

            If Servnst.IsServiceInstalled(_serivceNameToBeTested) Then

                Console.WriteLine("Uninstalling service in order to installing it...")

                isUninstalled = Servnst.UninstallService(_serivceNameToBeTested)

            End If

            actual = _installService()

            Assert.AreEqual(True, actual)

        End Sub

        ''' <summary>
        ''' Starts the service without uninstall it.
        ''' in case you want to debug the service in the IDE.
        ''' </summary>
        ''' <remarks>
        ''' Why this method is useful?
        ''' In case you want to debug the service, just run this test after setting 
        ''' 'System.Diagnostics.Debugger.Launch()' statements in your desired lines 
        ''' you want to debug in the service.
        ''' </remarks>
        <Test()> _
        Public Sub DEBUG_StartService()

            Dim actual As Boolean = False

            RemoveOldServiceAndInstallNew(actual)

            actual = Servnst.StartService(_serivceNameToBeTested)

            DoNotUninstall = True

            Assert.AreEqual(True, actual)
        End Sub

        ''' <summary>
        ''' Uninstalls the service test.
        ''' </summary>
        <Test()> _
        Public Sub UninstallServiceTest()

            Balora.Hodhod.STOW("#######################################################")

            Balora.Hodhod.STOW("Starting 'UninstallServiceTest' test...")

            Dim actual As Boolean = False

            Dim expected As Boolean = True

            'No need to uninstall in TearDown because this is already uninstall.
            DoNotUninstall = True

            actual = Servnst.UninstallService(_serivceNameToBeTested)

            Balora.Hodhod.STOW(" 'UninstallService' Test Ended.")

            'wait until fully uninstalled.
            Threading.Thread.Sleep(3000)

            Assert.AreEqual(expected, actual)

        End Sub

        <Test()> _
        Public Sub DEBUG_DeleteAllTempServicesFiles()

            Dim files = From file In FileIO.FileSystem.GetFiles(exeDir, FileIO.SearchOption.SearchTopLevelOnly, "CacaSer???.exe")
                       Where Not file.Contains("CacaSer.exe")

            For Each file In files
                FileIO.FileSystem.DeleteFile(file)
            Next

        End Sub

#End Region

        ''' <summary>
        ''' Test Servnst.IsServiceInstalled by 
        ''' passing a name for a server that 
        ''' doesn't exists and expect False in return.
        ''' </summary>
        <Test()> _
        Public Sub ServiceNotFoundTest()
            Balora.Hodhod.STOW("#######################################################")
            Dim actual As Boolean = False
            Dim expected As Boolean = False
            Balora.Hodhod.STOW("Starting 'IsServiceInstalled' test...")
            actual = Servnst.IsServiceInstalled("NoServiceWithThisName")
            Balora.Hodhod.STOW("'ServiceNotFound' Test Ended.")
            Assert.AreEqual(expected, actual)
        End Sub

        ''' <summary>
        ''' Installs the service test.
        ''' </summary>
        <Test()> _
        Public Sub InstallServiceTest()
            Balora.Hodhod.STOW("#######################################################")
            Balora.Hodhod.STOW("Starting 'InstallServiceTest' test...")
            Dim actual As Boolean = False
            Dim isUninstalled As Boolean = True
            Dim expected As Boolean = True
            Console.WriteLine("Checking if the service is installed...")
            If Servnst.IsServiceInstalled(_serivceNameToBeTested) Then
                Console.WriteLine("Uninstalling service in order to installing it...")
                isUninstalled = Servnst.UninstallService(_serivceNameToBeTested)
            End If
            If isUninstalled Then
                Balora.Hodhod.STOW("Installing service...")
                actual = _installService()
            Else
                Assert.Inconclusive("CacaSer service can't be uninstalled.")
            End If
            Balora.Hodhod.STOW(" 'InstallService' Test Ended.")
            Assert.AreEqual(expected, actual)
        End Sub

        ''' <summary>
        ''' Just installs the service.
        ''' </summary><returns></returns>
        Private Function _installService() As Boolean

            Dim srvData As New Servnst.ServiceData
            With srvData
                .Description = "This service handles the execution of Cacadu tasks."
                .DisplayName = "Cacadu Service"
                .ServiceAccount = ServiceAccount.LocalSystem
                .ServiceName = _serivceNameToBeTested
                .ServicePath = _serviceFileToBeTested
                .StartType = ServiceStartMode.Automatic
            End With
            Return Servnst.InstallService(srvData)

        End Function

        ''' <summary>
        ''' Test Servnst.IsServiceInstalled.
        ''' </summary>
        <Test()> _
        Public Sub IsServiceInstalledTest()
            Balora.Hodhod.STOW("#######################################################")
            Balora.Hodhod.STOW("Starting 'IsServiceInstalledTest' test...")
            Dim actual As Boolean = False
            Dim expected As Boolean = True
            'Windows Search Service.
            actual = Servnst.IsServiceInstalled(_serivceNameToBeTested)
            Balora.Hodhod.STOW(" 'IsServiceInstalled' Test Ended.")
            Assert.AreEqual(expected, actual)
        End Sub

        ''' <summary>
        ''' Starts the service test.
        ''' </summary>
        <Test()> _
        Public Sub StartServiceTest()
            Balora.Hodhod.STOW("#######################################################")
            Balora.Hodhod.STOW("Starting 'StartServiceTest' test...")
            Dim actual As Boolean = False
            Dim expected As Boolean = True
            Dim state As ServiceControllerStatus = Servnst.GetServiceStatus(_serivceNameToBeTested)
            If state = ServiceControllerStatus.Stopped Then
                actual = Servnst.StartService(_serivceNameToBeTested)
            Else
                actual = Servnst.StopService(_serivceNameToBeTested)
                Threading.Thread.Sleep(1000)
                actual = Servnst.StartService(_serivceNameToBeTested)
            End If
            Balora.Hodhod.STOW(" 'StartService' Test Ended.")
            Assert.AreEqual(expected, actual)
        End Sub

        Public Sub RemoveOldServiceAndInstallNew(ByVal actual As Boolean)
            'Remove old service file
            If Servnst.IsServiceInstalled(_serivceNameToBeTested) Then
                actual = Servnst.UninstallService(_serivceNameToBeTested)
                If actual Then actual = _installService()
            Else
                actual = _installService()
            End If
        End Sub

        ''' <summary>
        ''' Stops the service test.
        ''' </summary>
        <Test()> _
        Public Sub StopServiceTest()
            Balora.Hodhod.STOW("#######################################################")
            Balora.Hodhod.STOW("Starting 'StopServiceTest' test...")
            Dim actual As Boolean = False
            Dim expected As Boolean = True
            Dim state As ServiceControllerStatus = Servnst.GetServiceStatus(_serivceNameToBeTested)
            If state = ServiceControllerStatus.Stopped Then
                Servnst.StartService(_serivceNameToBeTested)
                actual = Servnst.StopService(_serivceNameToBeTested)
            Else
                Threading.Thread.Sleep(1000)
                actual = Servnst.StopService(_serivceNameToBeTested)
                Threading.Thread.Sleep(1000)
                Servnst.StartService(_serivceNameToBeTested)
            End If
            Balora.Hodhod.STOW(" 'StopService' Test Ended.")
            Assert.AreEqual(expected, actual)
        End Sub

        ''' <summary>
        ''' Pauses the service test.
        ''' </summary>
        <Test()> _
        Public Sub PauseServiceTest()
            Balora.Hodhod.STOW("#######################################################")
            Balora.Hodhod.STOW("Starting 'PauseServiceTest' test...")
            Dim actual As Boolean = False
            Dim expected As Boolean = True
            Dim state As ServiceControllerStatus = Servnst.GetServiceStatus(_serivceNameToBeTested)
            If state = ServiceControllerStatus.Running Then
                actual = Servnst.PauseService(_serivceNameToBeTested)
            Else
                Servnst.StartService(_serivceNameToBeTested)
                Threading.Thread.Sleep(4000)
                actual = Servnst.PauseService(_serivceNameToBeTested)
            End If
            Balora.Hodhod.STOW(" 'PauseService' Test Ended.")
            Assert.AreEqual(expected, actual)
        End Sub

        ''' <summary>
        ''' Continues the service test.
        ''' </summary>
        <Test()> _
        Public Sub ContinueServiceTest()
            Balora.Hodhod.STOW("#######################################################")
            Balora.Hodhod.STOW("Starting 'ContinueServiceTest' test...")
            Dim actual As Boolean = False
            Dim expected As Boolean = True
            Dim state As ServiceControllerStatus = Servnst.GetServiceStatus(_serivceNameToBeTested)
            If Not state = ServiceControllerStatus.Paused Then
                Servnst.StartService(_serivceNameToBeTested)
                Threading.Thread.Sleep(5000)
                state = Servnst.GetServiceStatus(_serivceNameToBeTested)
                If state = ServiceControllerStatus.Running Then
                    Servnst.PauseService(_serivceNameToBeTested)
                Else
                    Balora.Hodhod.STOW("Error with start running the service, check your system.")
                End If
            End If
            actual = Servnst.ResumeService(_serivceNameToBeTested)
            Balora.Hodhod.STOW(" 'ContinueService' Test Ended.")
            Assert.AreEqual(expected, actual)
        End Sub

    End Class

End Namespace
