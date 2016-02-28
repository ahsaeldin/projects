Imports Cacadore
Imports NUnit.Framework
Imports System.Threading
Imports Cacadore.Actions.System

Namespace CacadoreUT
    <TestFixture()> _
    Public Class ShutdownPCTest

#Region "Fields"
        Private _manualGate As ManualResetEvent
        Private _shutdownPCAction As ShutdownPC
        Private Shutdown_Raised As Boolean = False
        Private Shutdown_Cancelled As Boolean = False
#End Region

#Region "Setup Methods"
        <TestFixtureSetUp()> _
        Public Sub SetupMethods()
            PrepareCacaduComponents()
        End Sub

        <TestFixtureTearDown()> _
        Public Sub TearDownMethods()
        End Sub

        <SetUp()> _
        Public Sub SetupTest()
            _manualGate = New ManualResetEvent(False)
            Shutdown_Raised = False
            Shutdown_Cancelled = False
        End Sub

        <TearDown()> _
        Public Sub TearDownTest()
        End Sub
#End Region

#Region "Tests"
        <Test()> _
        <Category("ShortRunning")>
        Public Sub TestCanceledProperty_SetToFalse_CheckIfNotExceuted()
            '@@@@@@@@@@@-3 shit
            Dim _inputs As New ShutdownPC.ShutdownPCInputs
            _inputs.ShutdownType = Balora.ShutdownTypeEnum.Sleep
            _inputs.AllowCancel = True
            _shutdownPCAction = New ShutdownPC(_inputs)
            _shutdownPCAction.SetActionId("SecoSecoYa7yawan")
            _shutdownPCAction.ParentTask = New Task("SecoSecoYaTwor")
            SetupEventsHandlers()
            _shutdownPCAction.CancelShutDown()
            Dim executionResult As ITaskState.ExecutionState = _shutdownPCAction.Execute()
            Assert.IsTrue(Shutdown_Cancelled)
            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Canceled)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub TestPeriodBeforeShutdown_SetTo60Seconds_CheckPeriodBeforeExecution()
            '@@@@@@@@@@@@-4 shit
            Dim _inputs As New ShutdownPC.ShutdownPCInputs
            _inputs.ShutdownType = Balora.ShutdownTypeEnum.Sleep
            _inputs.PeriodBeforeShutdown = 60
            _shutdownPCAction = New ShutdownPC(_inputs)
            SetupEventsHandlers()
            Dim reterviedInputs = CType(_shutdownPCAction.Inputs, ShutdownPC.ShutdownPCInputs)
            Assert.AreEqual(reterviedInputs.PeriodBeforeShutdown, 60)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub TestIsForcedAndAllowCancel_SetToTrue_CheckReturendEventArgs()
            '@@@@@@@@@@@-1 shit
            Dim _inputs As New ShutdownPC.ShutdownPCInputs
            _inputs.ShutdownType = Balora.ShutdownTypeEnum.Sleep
            _inputs.IsForced = True
            _inputs.AllowCancel = True
            _shutdownPCAction = New ShutdownPC(_inputs)
            _shutdownPCAction.SetActionId("SecoSecoYa7yawan")
            _shutdownPCAction.ParentTask = New Task("SecoSecoYaTwor")
            SetupEventsHandlers()
            Dim reterviedInputs = CType(_shutdownPCAction.Inputs, ShutdownPC.ShutdownPCInputs)
            _shutdownPCAction.CancelShutDown()
            _shutdownPCAction.Execute()
            Assert.AreEqual(reterviedInputs.ShutdownType, Balora.ShutdownTypeEnum.ForcedSleep)
            Assert.IsTrue(reterviedInputs.IsForced)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub TestConfirmBeforeExecute_SetToTrue_CheckTrueIfEventArgs()
            '@@@@@@@@@@@-2 shit
            Dim _inputs As New ShutdownPC.ShutdownPCInputs
            _inputs.ShutdownType = Balora.ShutdownTypeEnum.Sleep
            _inputs.ConfirmBeforeExecute = True
            _shutdownPCAction = New ShutdownPC(_inputs)
            SetupEventsHandlers()
            Dim reterviedInputs = CType(_shutdownPCAction.Inputs, ShutdownPC.ShutdownPCInputs)
            Assert.IsTrue(reterviedInputs.ConfirmBeforeExecute)
        End Sub

        <Test()> _
        <Category("ShutDown Tests Be Careful")>
        Public Sub TestShutdown()
            Dim _inputs As New ShutdownPC.ShutdownPCInputs
            _inputs.ShutdownType = Balora.ShutdownTypeEnum.Shutdown
            _shutdownPCAction = New ShutdownPC(_inputs)
            SetupEventsHandlers()
            AddHandler _shutdownPCAction.Before_Shutdown, AddressOf Before_Shutdown
            Dim executionResult As ITaskState.ExecutionState = _shutdownPCAction.Execute()
            Assert.IsTrue(Shutdown_Raised)
            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Done)
        End Sub

        <Test()> _
        <Category("ShutDown Tests Be Careful")>
        Public Sub TestLogOff()
            Dim _inputs As New ShutdownPC.ShutdownPCInputs
            _inputs.ShutdownType = Balora.ShutdownTypeEnum.LogOff
            _shutdownPCAction = New ShutdownPC(_inputs)
            SetupEventsHandlers()
            AddHandler _shutdownPCAction.Before_Shutdown, AddressOf Before_Shutdown
            Dim executionResult As ITaskState.ExecutionState = _shutdownPCAction.Execute()
            Assert.IsTrue(Shutdown_Raised)
            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Done)
        End Sub

        Private Sub SetupEventsHandlers()
            AddHandler _shutdownPCAction.Before_Shutdown, AddressOf Before_Shutdown
            AddHandler _shutdownPCAction.Shutdown_Canceled, AddressOf Shutdown_Canceled
        End Sub
#End Region

#Region "Event"
        Private Sub Before_Shutdown(shutdownType As Balora.ShutdownTypeEnum, shutdownEventArgs As ShutdownPC.ShutdownEventArgs)
            Shutdown_Raised = True
        End Sub

        Private Sub Shutdown_Canceled(shutdownType As Balora.ShutdownTypeEnum)
            Shutdown_Cancelled = True
        End Sub
#End Region
    End Class
End Namespace