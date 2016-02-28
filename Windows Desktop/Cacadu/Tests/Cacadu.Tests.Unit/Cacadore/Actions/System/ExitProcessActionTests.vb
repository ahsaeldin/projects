Imports Cacadore
Imports NUnit.Framework
Imports System.Threading
Imports Cacadore.Actions.System
Imports Cacadore.Actions.System.ExitProcess

Namespace CacadoreUT
    <TestFixture()> _
    Public Class ExitProcessActionTests

#Region "Fields"
        Private _manualGate As ManualResetEvent
        Private _isExceptionRaised As Boolean = False
#End Region

#Region "Setups"

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
        End Sub

        <TearDown()> _
        Public Sub TearDownTest()
        End Sub
#End Region

#Region "Tests"
        <Test()> _
        <Category("Test solo And Watch Results")>
        Public Sub CallExecute_EndProcess_CheckManuallyIfDone()
            Process.Start("calc") : Process.Start("calc")
            Threading.Thread.Sleep(1000)
            Dim actionInput As New ExitProcessInputs("calc")
            Dim exitProcessAction As New ExitProcess
            AddHandler exitProcessAction.Process_Exited, AddressOf ProcessExited
            actionInput.WaitForExitPeriod = -1
            exitProcessAction.Inputs = actionInput
            Dim executionResult As ITaskState.ExecutionState = exitProcessAction.Execute()
            _manualGate.WaitOne()
            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Done)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub CallExecute_TriggerExceptionToBeThrown_CheckTheReturnedException()
            'Prepare action
            Dim actionInput As New ExitProcessInputs("calc")
            Dim exitProcessAction As New ExitProcess
            exitProcessAction.Inputs = actionInput
            'Display them in a selected combo boxes.
            Dim possibleExceptionsDescriptionsList = exitProcessAction.ExceptionsManager.PossibleExceptionsDescriptions
            'The user choose this message.
            Dim selectedMessageFromSelecteiveCheckBoxes As String = "The process has already exited"
            'Ask ExceptionsManager to retrieve the corresponding ActionExceptionHandler to the selected message.
            Dim _actionExceptionHandler As ActionExceptionHandler = exitProcessAction.
                                                                    ExceptionsManager.
                                                                    GetSelectedExceptionHandler(selectedMessageFromSelecteiveCheckBoxes)
            'Now you can play with the handler and set the actions to take if an exception raised.
            _actionExceptionHandler.RunAnotherTaskOnError = True
            AddHandler exitProcessAction.ExceptionsManager.After_Exception_Raised, AddressOf After_Exception_Raised
            exitProcessAction.ParentTask = New Task("testingId")
            exitProcessAction.SetActionId("testingId")
            exitProcessAction.Execute()
            If Not _isExceptionRaised Then Assert.Inconclusive("Test fail, no exception raised.")
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub CallGetRunningProcessesNames_CheckReturnedIfNUll()
            Dim exitProcessAction As New ExitProcess
            Dim _procs = ExitProcess.GetRunningProcessesNames()
            Assert.NotNull(_procs)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub CallExecute_PassingNoProcessName_CheckNullAndFailed()
            Dim actionInput As New ExitProcessInputs("")
            Dim exitProcessAction As New ExitProcess
            AddHandler exitProcessAction.Process_Exited, AddressOf ProcessExited
            actionInput.WaitForExitPeriod = -1
            exitProcessAction.Inputs = actionInput
            exitProcessAction.ParentTask = New Task("testingId")
            exitProcessAction.SetActionId("testingId")
            Dim executionResult As ITaskState.ExecutionState = exitProcessAction.Execute()
            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Failed)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub CallDispose_CheckIfDisposed()
            Dim exitProcessAction As New ExitProcess
            exitProcessAction.Dispose()
            Assert.IsTrue(IsNothing(exitProcessAction.MyProcess))
        End Sub
#End Region

#Region "Events"
        Private Sub ProcessExited(sender As Object, e As EventArgs)
            _manualGate.Set()
        End Sub

        Private Sub After_Exception_Raised(exception As Exception, actionToTake As IExceptionHandler.RaisedExceptionArgs)
            _manualGate.Set()
            _isExceptionRaised = True
        End Sub
#End Region

    End Class
End Namespace