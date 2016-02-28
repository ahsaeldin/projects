Imports Cacadore
Imports NUnit.Framework
Imports System.Threading
Imports Cacadore.Actions.System
Imports System.Runtime.InteropServices
Imports Cacadore.Actions.System.StartApplication

Namespace CacadoreUT
    <TestFixture()> _
    Public Class StartApplicationActionTests

        Private Declare Auto Function IsIconic Lib "user32.dll" (ByVal hwnd As IntPtr) As Boolean

#Region "Fields"
        Private _manualGate As ManualResetEvent
#End Region

#Region "Properties"
        Property IsProcessExited As Boolean = False
#End Region

#Region "Setups and TearDowns"
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
            IsProcessExited = False

        End Sub

        <TearDown()> _
        Public Sub TearDownTest()
        End Sub

#End Region
#Region "Test"

#Region "Running Action Tests"
        <Test()> _
        <Category("ShortRunning")>
        Public Sub Execute_RunningCalculator_CheckIfDone()
            Dim actionInput As New StartApplicationInputs("calc") With {.WindowStyle = ProcessWindowStyle.Minimized}

            Dim startAppAction As New StartApplication

            actionInput.WaitForExitPeriod = 1000

            startAppAction.Inputs = actionInput

            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()

            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Done)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub Execute_OpenFileByOpenVerb_CheckIfDone()
            Dim startAppAction As New StartApplication
            'Create a temp file for testing
            Dim fileName As String = "AnyFileName.txt" : My.Computer.FileSystem.WriteAllText(fileName, "Any Text Here.", False)
            startAppAction.Inputs = New StartApplicationInputs(fileName) With {.WindowStyle = ProcessWindowStyle.Minimized}

            'Setting Open verb and send all verbs names to console.
            Dim supportedVerbsBuilder As System.Text.StringBuilder = SetVerb(startAppAction, "open")
            Balora.Hodhod.STOW(supportedVerbsBuilder.ToString())

            'Execute then delete the temp file.
            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()
            My.Computer.FileSystem.DeleteFile(fileName)

            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Done)
        End Sub

#End Region

#Region "Exception Tests"
        <Test()> _
        <Category("ShortRunning")>
        Public Sub Execute_TriggerInvalidOperationExceptionToBeThrown_CheckTheReturnedException()
            Dim startAppAction As New StartApplication

            startAppAction.SetActionId("SecoSecoYa7yawan")

            startAppAction.ParentTask = New Task("SecoSecoYaTwor")

            startAppAction.Inputs = New StartApplicationInputs("NoFileWithThisName")

            'Execute then delete the temp file.
            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()

            Dim _errors As StartApplicationErrors = CType(startAppAction.Errors, StartApplicationErrors)

            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Failed)
            Assert.IsTrue(_errors.ErrorList.Contains("The system cannot find the file specified"))
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub Execute_TriggerWin32ExceptionToBeThrown_CheckTheReturnedException()
            Dim startAppAction As New StartApplication

            'Create a temp file for testing
            Dim fileName As String = "AnyFileName.txt" : My.Computer.FileSystem.WriteAllText(fileName, "Any Text Here.", False)

            startAppAction.Inputs = New StartApplicationInputs(fileName) With {.WindowStyle = ProcessWindowStyle.Minimized}

            startAppAction.SetActionId("SecoSecoYa7yawan")

            startAppAction.ParentTask = New Task("SecoSecoYaTwor")

            startAppAction.MyProcess.StartInfo.Verb = "Invalid Verb Name"

            'Execute then delete the temp file.
            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()

            My.Computer.FileSystem.DeleteFile(fileName)

            Dim _errors As StartApplicationErrors = CType(startAppAction.Errors, StartApplicationErrors)

            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Failed)

            Assert.IsTrue(_errors.ErrorList.Contains("No application is associated with the specified file for this operation"))
        End Sub
#End Region

#Region "Events Tests"
        <Test()> _
        <Category("ShortRunning")>
        Public Sub Execute_RunningCMD_CatchingExitedEvent()

            Dim actionInput As New StartApplicationInputs("cmd")

            Dim startAppAction As New StartApplication

            AddHandler startAppAction.Process_Exited, AddressOf Application_Exited

            startAppAction.Inputs = actionInput

            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()

            If Not IsProcessExited Then
                _manualGate.Reset()
                _manualGate.WaitOne()
            End If

            Assert.IsTrue(IsProcessExited)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub Execute_RunningSort_CatchingOutputEvent()
            Dim actionInput As New StartApplicationInputs("ipconfig")

            'To wait before exit for errors.
            actionInput.WaitForExitPeriod = Timeout.Infinite

            Dim startAppAction As New StartApplication

            AddHandler startAppAction.Process_Exited, AddressOf Application_Exited

            startAppAction.Inputs = actionInput

            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()

            Dim _outputs As StartApplicationOutputs = CType(startAppAction.Outputs, StartApplicationOutputs)

            Dim StringToBeCheckedInOutput = "Windows IP Configuration"
            Dim isReturnedOutputContainsThisPharse = _outputs.Output.Contains(StringToBeCheckedInOutput)

            CheckStrinInOutput(_outputs, StringToBeCheckedInOutput, isReturnedOutputContainsThisPharse)

            Assert.IsTrue(isReturnedOutputContainsThisPharse)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub Execute_RunningSort_CatchingErrorEvent()
            Dim actionInput As New StartApplicationInputs("net.exe")

            'To wait before exit for errors.
            actionInput.WaitForExitPeriod = Timeout.Infinite

            actionInput.Arguments = "view asdfasdf"

            Dim startAppAction As New StartApplication

            AddHandler startAppAction.Process_Exited, AddressOf Application_Exited

            startAppAction.Inputs = actionInput

            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()

            Dim _errors As StartApplicationErrors = CType(startAppAction.Errors, StartApplicationErrors)

            Dim isReturnedOutputContainsThisPharse As Boolean

            Dim stringToBeCheckedInErrors As String = "The network path was not found"

            CheckStringInErrors(_errors,
                                isReturnedOutputContainsThisPharse,
                                stringToBeCheckedInErrors)

            Assert.IsTrue(isReturnedOutputContainsThisPharse)
        End Sub

#End Region

#Region "Process StartInfo Tests"
        <Test()> _
        <Category("ShortRunning")>
        Public Sub Execute_PingingCprinGold_CheckReturnedOutput()
            Dim actionInput As New StartApplicationInputs("PING") With {.WindowStyle = ProcessWindowStyle.Maximized,
                                                        .Arguments = "www.cpringold.com"}

            Dim startAppAction As New StartApplication

            'To wait before exit for errors.
            actionInput.WaitForExitPeriod = Timeout.Infinite

            AddHandler startAppAction.Process_Exited, AddressOf Application_Exited

            startAppAction.Inputs = actionInput

            'You can't get output if you don't disable shell and redirect output
            'however by making this you can't maximize or minimize the form.
            startAppAction.DisableShellExecute()
            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()

            Dim _outputs As StartApplicationOutputs = CType(startAppAction.Outputs, StartApplicationOutputs)

            Dim stringToBeCheckInOutput As String = "Pinging cpringold.com"

            Dim isReturnedOutputContainsThisPharse As Boolean

            CheckStrinInOutput(_outputs, stringToBeCheckInOutput, isReturnedOutputContainsThisPharse)

            Assert.IsTrue(isReturnedOutputContainsThisPharse)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub CreateNoWindow()
            Dim startAppAction As New StartApplication
            'Create a temp file for testing
            Dim fileName As String = "BlankFileName.txt" : My.Computer.FileSystem.WriteAllText(fileName, "Any Text Here.", False)
            startAppAction.Inputs = New StartApplicationInputs(fileName) With {.CreateNoWindow = True}
            'Setting Open verb and send all verbs names to console.
            Dim supportedVerbsBuilder As System.Text.StringBuilder = SetVerb(startAppAction, "open")
            Balora.Hodhod.STOW(supportedVerbsBuilder.ToString())
            'Execute then delete the temp file.
            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()
            My.Computer.FileSystem.DeleteFile(fileName)
            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Done)
        End Sub

        <Test()> _
        <Category("Test solo And Watch Results")>
        Public Sub Execute_ForceNotePadMaximize_CheckIfMaximized()
            Dim startAppAction As New StartApplication

            Dim _inputs As New StartApplicationInputs("notepad")

            _inputs.WindowStyle = ProcessWindowStyle.Maximized

            startAppAction.Inputs = _inputs

            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()

            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Done)
        End Sub

        <Test()> _
        <Category("Test solo And Watch Results")>
        Public Sub Execute_ForceNotePadMinmize_CheckIfMinimized()
            Dim startAppAction As New StartApplication

            Dim _inputs As New StartApplicationInputs("notepad")

            _inputs.WindowStyle = ProcessWindowStyle.Minimized

            startAppAction.Inputs = _inputs

            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()

            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Done)

        End Sub

        <Test()> _
        <Category("Test solo And Watch Results")>
        Public Sub Execute_ForceNotePadHidden_CheckIfHidden()
            Dim startAppAction As New StartApplication

            Dim _inputs As New StartApplicationInputs("calc")

            _inputs.WindowStyle = ProcessWindowStyle.Hidden

            _inputs.CreateNoWindow = True

            startAppAction.Inputs = _inputs

            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()

            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Done)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub WorkingDirectoryProperty()
            Dim startAppAction As New StartApplication

            Dim fileName As String = My.Computer.FileSystem.SpecialDirectories.Desktop & "\BlankFileName.txt"

            My.Computer.FileSystem.WriteAllText(fileName, "Any Text Here.", False)

            Dim _inputs As New StartApplicationInputs("BlankFileName.txt")

            'Magic is here, without this line StartApplication cannot find the file.
            _inputs.WorkingDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop

            startAppAction.Inputs = _inputs

            'Execute then delete the temp file.
            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()

            My.Computer.FileSystem.DeleteFile(fileName)

            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Done)
        End Sub
#End Region

#Region "Wait for exit/input idle tests"
        <Test()> _
        <Category("ShortRunning")>
        Public Sub WaitForExit()
            Dim actionInput As New StartApplicationInputs("calc") With {.WindowStyle = ProcessWindowStyle.Minimized}

            Dim startAppAction As New StartApplication

            startAppAction.Inputs = actionInput

            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()

            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Done)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub WaitPeriodForExit()
            Dim actionInput As New StartApplicationInputs("calc") With {.WindowStyle = ProcessWindowStyle.Minimized}

            Dim startAppAction As New StartApplication

            actionInput.WaitForExitPeriod = 5000

            startAppAction.Inputs = actionInput

            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()

            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Done)
        End Sub

#End Region

#Region "Checking Nulls"
        <Test()> _
        <Category("ShortRunning")>
        Public Sub Execute_PassingNoFile_CheckNullAndFailed()
            Dim startAppAction As New StartApplication

            startAppAction.SetActionId("SecoSecoYa7yawan")

            startAppAction.ParentTask = New Task("SecoSecoYaTwor")

            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()

            Dim _errors As StartApplicationErrors = CType(startAppAction.Errors, StartApplicationErrors)

            Dim isReturnedOutputContainsThisPharse As Boolean

            CheckStringInErrors(_errors, isReturnedOutputContainsThisPharse, "Cannot start process because a file name has not been provided")

            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Failed)

            Assert.IsTrue(isReturnedOutputContainsThisPharse)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub Execute_PassingBlankFileName_CheckNullAndFailed()
            Dim startAppAction As New StartApplication

            startAppAction.SetActionId("SecoSecoYa7yawan")

            startAppAction.ParentTask = New Task("SecoSecoYaTwor")

            startAppAction.Inputs = New StartApplicationInputs("")

            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()

            Dim _errors As StartApplicationErrors = CType(startAppAction.Errors, StartApplicationErrors)

            Dim isReturnedOutputContainsThisPharse As Boolean

            CheckStringInErrors(_errors, isReturnedOutputContainsThisPharse, "Cannot start process because a file name has not been provided")

            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Failed)

            Assert.IsTrue(isReturnedOutputContainsThisPharse)
        End Sub

#End Region

#Region "Disposing Tests"
        <Test()> _
        <Category("ShortRunning")>
        Public Sub TestDisposing()
            Dim startAppAction As New StartApplication

            'Create a temp file for testing
            Dim fileName As String = "AnyFileName.txt" : My.Computer.FileSystem.WriteAllText(fileName, "Any Text Here.", False)

            startAppAction.Inputs = New StartApplicationInputs(fileName) With {.WindowStyle = ProcessWindowStyle.Minimized}

            'Setting Open verb and send all verbs names to console.
            Dim supportedVerbsBuilder As System.Text.StringBuilder = SetVerb(startAppAction, "open")

            Balora.Hodhod.STOW(supportedVerbsBuilder.ToString())

            'Execute then delete the temp file.
            Dim executionResult As ITaskState.ExecutionState = startAppAction.Execute()

            My.Computer.FileSystem.DeleteFile(fileName)

            startAppAction.Dispose()

            Assert.AreEqual(executionResult, ITaskState.ExecutionState.Done)

            Assert.IsTrue(IsNothing(startAppAction.MyProcess))
        End Sub
#End Region
#End Region

#Region "Helper Methods"
        Private Function SetVerb(ByVal startAppAction As StartApplication, verbName As String) As System.Text.StringBuilder
            Dim supportedVerbsBuilder As System.Text.StringBuilder = New System.Text.StringBuilder(vbNullString)

            For Each _verb As String In startAppAction.MyProcess.StartInfo.Verbs

                If _verb.Contains(verbName) Then
                    startAppAction.MyProcess.StartInfo.Verb = _verb
                End If

                supportedVerbsBuilder.Append(_verb & "-")
            Next

            Return supportedVerbsBuilder
        End Function

        Private Shared Sub CheckStrinInOutput(ByVal _outputs As StartApplicationOutputs,
                                      ByVal StringToBeCheckedInOutput As String,
                                      ByRef isReturnedOutputContainsThisPharse As Boolean)
            For Each _outputString In _outputs.Output
                If _outputString.Contains(StringToBeCheckedInOutput) Then
                    isReturnedOutputContainsThisPharse = True
                End If
            Next
        End Sub

        Private Sub CheckStringInErrors(ByVal errors As StartApplicationErrors,
                                        ByRef isReturnedOutputContainsThisPharse As Boolean,
                                        ByVal stringToBeCheckedInErrors As String)
            For Each _error In errors.ErrorList
                If _error.Contains(stringToBeCheckedInErrors) Then
                    isReturnedOutputContainsThisPharse = True
                    Exit For
                End If
            Next
        End Sub

#Region "Helper Events Handlers"
        Private Sub Application_Exited(sender As Object, e As EventArgs)
            IsProcessExited = True

            _manualGate.Set()
        End Sub
#End Region
#End Region
    End Class
End Namespace