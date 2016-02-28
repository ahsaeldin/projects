Imports Balora
Imports NUnit.Framework

Namespace BaloraUT
    <TestFixture()> _
    Public Class SysManageTests

#Region "Fields"
        Private _sysManage As SysManage
#End Region

#Region "Setup Methods"
        <TestFixtureSetUp()> _
        Public Sub SetupMethods()
            Unit.PrepareCacaduComponents()
        End Sub

        <TestFixtureTearDown()> _
        Public Sub TearDownMethods()
        End Sub

        <SetUp()> _
        Public Sub SetupTest()
            _sysManage = New SysManage
        End Sub

        <TearDown()> _
        Public Sub TearDownTest()
        End Sub
#End Region

#Region "Tests"
        <Test()> _
        <Category("ShutDown Tests Be Careful")>
        <Ignore()>
        Public Sub TestLogOff()
            _sysManage.LogOff()
        End Sub

        <Test()> _
        <Category("ShutDown Tests Be Careful")>
              <Ignore()>
        Public Sub TestShutdown()
            _sysManage.Shutdown()
        End Sub

        <Test()> _
        <Category("ShutDown Tests Be Careful")>
              <Ignore()>
        Public Sub TestRestart()
            _sysManage.Restart()
        End Sub

        <Test()> _
        <Category("ShutDown Tests Be Careful")>
        <Ignore()>
        Public Sub TestForcedLogOff()
            _sysManage.ForceLogOff()
        End Sub

        <Test()> _
        <Category("ShutDown Tests Be Careful")>
        <Ignore()>
        Public Sub TestForcedShutdown()
            _sysManage.ForceShutdown()
        End Sub

        <Test()> _
        <Category("ShutDown Tests Be Careful")>
        <Ignore()>
        Public Sub TestForcedReboot()
            _sysManage.ForceReboot()
        End Sub

        <Test()> _
        <Category("ShutDown Tests Be Careful")>
        <Ignore()>
        Public Sub TestHibernate()
            _sysManage.Hibernate()
        End Sub

        <Test()> _
        <Category("ShutDown Tests Be Careful")>
        <Ignore()>
        Public Sub TestForceHibernate()
            _sysManage.ForceHibernate()
        End Sub

        <Test()> _
        <Category("ShutDown Tests Be Careful")>
        <Ignore()>
        Public Sub TestSleep()
            _sysManage.Sleep()
        End Sub

        <Test()> _
        <Category("ShutDown Tests Be Careful")>
        <Ignore()>
        Public Sub TestForceSleep()
            _sysManage.ForceSleep()
        End Sub
#End Region

    End Class
End Namespace