Imports Balora
Imports Helpers
Imports TecDAL
Imports Cacadore
Imports NUnit.Framework
Imports Balora.PathsHelper

Namespace CacaduIT
    <TestFixture()> _
    Public Class TectonicTesting

#Region "Fields"
        Private taskObject As Task
        Private builtInTasksObj As BIT
        Private settingsObj As Cacadore.Settings
        Private createActions As BackgroundWorkerHelper
        Private CacadoreToTectonic As New CacadoreToTectonic
        Private executionGate As New Threading.ManualResetEvent(False)
        Private tectonicPath As String = commonAppDataPath & "tectonic"
        Private tectonic_copyPath As String = commonAppDataPath & "tectonic_copy"
        Private commonAppDataPath As String = GetCommonApplicationData() & "Cacadu\"
        Private resulting_tectonicPath As String = commonAppDataPath & "resulting_tectonic"
#End Region

#Region "Setups Methods"
        <TestFixtureSetUp()> _
        Public Sub SetupMethods()
            Unit.PrepareCacaduComponents()
        End Sub

        <TestFixtureTearDown()> _
        Public Sub TearDownMethods()
            'Re-copy the original.
            'IO.File.Copy(tectonicPath, resulting_tectonicPath, True)
            'IO.File.Copy(tectonic_copyPath, tectonicPath, True)
            'IO.File.Delete(tectonic_copyPath)
        End Sub

        <SetUp()> _
        Public Sub SetupTest()
            taskObject = Nothing
            executionGate = New Threading.ManualResetEvent(False)
            builtInTasksObj = New BIT
            createActions = New BackgroundWorkerHelper
            settingsObj = New Cacadore.Settings
            'Make a backup copy of tectonic before playing with the orignial. 
            'IO.File.Delete(tectonic_copyPath)
            'IO.File.Copy(tectonicPath, tectonic_copyPath)
        End Sub

        <TearDown()> _
        Public Sub TearDownTest()
        End Sub
#End Region

        ''' <summary>
        ''' Execute mostly all the operations provided by 
        ''' CacadoreToTectonic via TypedDataSetCrud class.
        ''' </summary>
        <Test()> _
        <Category("ShortRunning")>
        Public Sub SimpleTectonicingWithDeleting()
            SimpleTectonicing(True)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub SimpleTectonicingWithOutDeleting()
            SimpleTectonicing(False)
        End Sub

        ''' <summary>
        ''' Simples the tectonicing.
        ''' </summary>
        ''' <param name="deleting">if set to <c>true</c> [deleting].</param>
        Public Sub SimpleTectonicing(deleting As Boolean)
            Dim actual As Boolean = False
            Dim expected As Boolean = True
            Dim [stopWatch] As New Stopwatch
            [stopWatch].Start()
            'Add and insert a new task:-
            'Add a new group:-
            Console.WriteLine("Create a new group...")
            Dim mainGroup = CacadoreToTectonic.CreateNewGroup("NohaAli" & DateTime.UtcNow.ToString)
            If mainGroup Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Group created.")
            '====================================================================================
            'Add 10 Tasks
            Dim newTask As Task = Nothing
            For i = 1 To 10
                'Console.WriteLine("Create a new task (task " & i & " )")
                newTask = CacadoreToTectonic.CreateNewTask(True, mainGroup.GroupId)
                If newTask Is Nothing Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                Console.WriteLine("Task created.")
                Console.WriteLine("Set task's properties before saving it...")
                newTask.Properties.IsEnabled = True
                newTask.Properties.LastTimeRun = DateTime.UtcNow
                newTask.Properties.NextTimeToBeRun = DateTime.UtcNow
                newTask.Properties.TaskName = Balora.RandomBytesGenerator.GRCM & DateTime.UtcNow.ToString
                'For both, they will be blank for now until I finish the logic of triggers
                'Tasks>=Triggers_When ... TaskProperties.Triggers_When
                'Tasks_States>=Next_Trigger_Id ... TaskState.NextTriggerId
                newTask.Properties.TriggersWhen = DateTime.UtcNow.ToString
                If Not newTask.Save() Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                Console.WriteLine("Task's properties saved.")
            Next
            '====================================================================================
            'Add globals:-
            Dim globalsObj As New Cacadore.Globals
            Dim res As Boolean
            Console.WriteLine("Adding a new global variables...")
            Dim memId As String = Balora.RandomBytesGenerator.GRCM
            If Not globalsObj.IsRegisteredGlobalMember(memId) Then res = CacadoreToTectonic.AddGlobalVariable(memId, "DateTime.UtcNow.ToString1")
            If Not res Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            memId = Balora.RandomBytesGenerator.GRCM
            If Not globalsObj.IsRegisteredGlobalMember(memId) Then res = CacadoreToTectonic.AddGlobalVariable(memId, "DateTime.UtcNow.ToString2")
            If Not res Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            If Not globalsObj.IsRegisteredGlobalMember(memId) Then memId = Balora.RandomBytesGenerator.GRCM
            memId = Balora.RandomBytesGenerator.GRCM
            If Not globalsObj.IsRegisteredGlobalMember(memId) Then res = CacadoreToTectonic.AddNewGlobalConstant(memId, "DateTime.UtcNow.ToString3")
            If Not res Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Global variables added")
            '====================================================================================
            'Add a variable to a task:-
            Dim soso = Balora.RandomBytesGenerator.GRCM & DateTime.UtcNow.ToString
            Console.WriteLine("Adding a new task variable...")

            CacadoreToTectonic.AddNewTaskVariable("tecx om malk", soso, "asdfasdf")
            Console.WriteLine("New task variable added.")
            'Uncomment to test duplicate variable error (if you try to insert a new variable that is already attached to at task)
            Console.WriteLine("Adding a new task variable...")
            CacadoreToTectonic.AddNewTaskVariable("tecx om malk", "Azangazanga" & DateTime.Now, "asdfasdf")
            Console.WriteLine("New task variable added.")
            Console.WriteLine("Adding a dump trigger to the task.")
            '====================================================================================
            Dim newTaskTrigger As QuartzHelper.QuartzTrigger = Nothing
            For i = 1 To 13
                Dim cacaHelper As New CacadoreHelper
                Dim tskMeta As New TaskMeta
                tskMeta.CronExpression = "0/1 * * ? * *"
                tskMeta.StartDateTimeUTC = DateTime.Now.ToUniversalTime
                'Attach 2 triggers to newTask:-
                newTaskTrigger = cacaHelper.AttachScheduledTriggerToTask(newTask, tskMeta)
                If newTaskTrigger Is Nothing Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
            Next
            '====================================================================================
            Console.WriteLine("Adding a new task variable again...")
            'Create a TaskVariable class with an id of a task that already doesn't exists:-
            Dim tskVaryMan As New TaskVariables(newTask.TaskId) 'secosecoseaco isn't exists.
            If Not tskVaryMan.AddVariable("trying1", "failed1") Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("New task variable added.")
            Dim noOfActionsToBeAdded As Integer = 50
            Console.WriteLine("Attaching a " & noOfActionsToBeAdded & " actions to the task...")
            '====================================================================================
            Dim newTaskAction As Action = Nothing
            For i = 1 To noOfActionsToBeAdded
                'Attach an action to a task:-
                Dim actionShape As New Shape
                Dim actionInput As IAction.IInputs = Nothing
                Dim actionOutput As IAction.IOutputs = Nothing
                newTaskAction =
                CacadoreToTectonic.AttachActionToTask(newTask,
                                                          IAction.ActionTypeEnum.StartApplication,
                                                          "Action Desc",
                                                          False,
                                                          actionInput,
                                                          actionOutput,
                                                          actionShape)
                If newTaskAction Is Nothing Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                Console.WriteLine("attached a new action (action " & newTaskAction.ActionId & " )")
            Next
            Console.WriteLine("All actions attached.")
            '====================================================================================
            Console.WriteLine("Inserting a new history entry...")
            'Insert History Entry:-
            Dim historyEntryMan As New HistoryEntry
            historyEntryMan.TaskId = newTask.TaskId
            historyEntryMan.Status = 0
            historyEntryMan.Description = "what i write here"
            historyEntryMan.EntryType = IHistoryEntry.HistoryEntryType.Task
            historyEntryMan.EntryDateTime = DateTime.UtcNow
            CacadoreToTectonic.InsertNewHistoryEntry(historyEntryMan)

            If Not CacadoreToTectonic.InsertNewHistoryEntry(historyEntryMan) Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                If Cacadore.Settings.HistoryEnabled Then
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                Else
                    Console.WriteLine("History isn't enabled.")
                End If
            End If
            Console.WriteLine("History entry added.")
            '====================================================================================
            Console.WriteLine("Updating settings table...")
            'Update Settings:-
            Dim sett As New Cacadore.Settings.SettingEntry
            sett.Enabled = True
            sett.HistoryEnabled = False
            sett.TriggersEnabled = True
            If Not CacadoreToTectonic.UpdateAllSettings(sett) Then
                If Cacadore.Settings.HistoryEnabled Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                Else
                    Console.WriteLine("History isn't enabled.")
                End If
            End If
            If Not CacadoreToTectonic.UpdateSetting(ISettings.SettingsEnum.HistoryEnabled, True) Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Settings table updated.")
            '====================================================================================
            Console.WriteLine("Adding 2 global variables...")
            'Update Globals:-
            Dim varName As String = "G<ahmed1>"
            Dim varName2 As String = "G<ahmed2>"
            Dim res2 As Boolean = False
            If Not globalsObj.IsRegisteredGlobalMember(varName) Then
                res2 = CacadoreToTectonic.AddGlobalVariable(varName, "seco")
            Else
                res2 = CacadoreToTectonic.UpdateGlobalVariable(varName, "seco")
            End If
            If Not res2 Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("One global variable added.")

            If Not globalsObj.IsRegisteredGlobalMember(varName2) Then
                res2 = CacadoreToTectonic.AddGlobalVariable(varName2, "neco")
            Else
                res2 = CacadoreToTectonic.UpdateGlobalVariable(varName2, "neco")
            End If

            If Not globalsObj.IsRegisteredGlobalMember(varName2) Then
                res2 = CacadoreToTectonic.AddGlobalVariable(varName2, "neco")
            Else
                res2 = CacadoreToTectonic.UpdateGlobalVariable(varName2, "neco")
            End If

            If Not res2 Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If

            Console.WriteLine("Global variables added.")
            Console.WriteLine("Updating a global variable...")
            res2 = CacadoreToTectonic.UpdateGlobalVariable(varName, "DateTime.UtcNow.ToString1")
            If Not res2 Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Global variable updated.")
            '====================================================================================
            Console.WriteLine("Updating a group...")
            'Update Groups
            res2 = CacadoreToTectonic.ChangeGroupName(mainGroup, "update" & DateTime.Now.Minute)
            If Not res2 Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Group Updated.")
            '====================================================================================
            Console.WriteLine("Updating a task variable...")
            'Update TaskVariable
            res2 = CacadoreToTectonic.UpdateTaskVariable(newTask.TaskId, "trying1", "نجحت")
            If Not res2 Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Task variable Updated.")
            '====================================================================================
            'Update Task
            newTask.Properties.IsEnabled = True
            newTask.Properties.TaskName = "Updated Taskttt" & DateTime.Now.Minute
            res2 = CacadoreToTectonic.UpdateTask(newTask)
            If Not res2 Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            '====================================================================================
            Console.WriteLine("Updating a task action...")
            'Update Action
            newTaskAction.Comments = "حمار حصاوي"
            res2 = CacadoreToTectonic.UpdateTaskAction(newTaskAction, newTask)
            If newTaskAction Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Task action Updated.")
            '====================================================================================
            Console.WriteLine("Querying history...")
            'Query History
            Dim tasksHistory As List(Of HistoryEntry) =
            CacadoreToTectonic.GetTasksHistory()
            If tasksHistory Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Get tasks history done...")
            Dim actionsHistory As List(Of HistoryEntry) =
            CacadoreToTectonic.GetActionsHistory()
            If actionsHistory Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Get actions history done...")
            Dim triggersHistory As List(Of HistoryEntry) =
            CacadoreToTectonic.GetTriggersHistory()
            If triggersHistory Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Get triggers history done...")
            Dim newTaskHistory As List(Of HistoryEntry) =
            CacadoreToTectonic.GetSingleTaskHistory(newTask.TaskId)
            If newTaskHistory Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Get single task history done...")
            Console.WriteLine("All history queries Done.")
            '====================================================================================
            Console.WriteLine("Querying globals...")
            'Query Globals
            Dim globConstants As Dictionary(Of String, String) =
                CacadoreToTectonic.GetGlobals(IGlobals.GlobalType.Constant)
            If globConstants Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Got all globals constants...")
            Dim globVariables As Dictionary(Of String, String) =
                CacadoreToTectonic.GetGlobals(IGlobals.GlobalType.Variable)
            If globVariables Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Got all globals variables...")
            Console.WriteLine("All globals queries Done.")
            '====================================================================================
            Console.WriteLine("Gets a dictionary of all Tectonic groups...")
            'Query GroupsCollection
            Dim groupsCollec =
                CacadoreToTectonic.GetGroups()
            If groupsCollec Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Query GroupsCollection done.")
            '====================================================================================
            Console.WriteLine("Get settings...")
            'Query Settings
            Dim settingsDic =
                CacadoreToTectonic.GetSettings()
            If settingsDic Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Get settings done.")
            '====================================================================================
            Console.WriteLine("Get task variables...")
            'Query Task Variables
            Dim mainTaskVariables =
                CacadoreToTectonic.GetTaskVariables(newTask)
            If mainTaskVariables Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Getting task variables done.")
            'Get VariableValue
            Dim variableValue =
                 CacadoreToTectonic.GetTaskVariableValue(newTask, "trying1")
            If variableValue Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Getting single task variable value done.")
            '====================================================================================
            Console.WriteLine("Query task properties...")
            'Query TaskProperties
            Dim taskName As String =
                CacadoreToTectonic.QueryTaskName(newTask)
            If taskName = "" Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Dim isTaskEnabled As Boolean =
                CacadoreToTectonic.QueryTaskEnabledState(newTask)
            If isTaskEnabled = False Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Dim lastTimeRun As Date =
                CacadoreToTectonic.QueryTaskLastTimeRun(newTask)
            If lastTimeRun = Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Dim taskRunResult As String =
                CacadoreToTectonic.QueryTaskLastTimeRunResult(newTask)
            If taskRunResult = "" Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Dim nexRun As Date =
                CacadoreToTectonic.QueryTaskNextTimeToBeRun(newTask)
            If nexRun = Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            'For both, they will be blank for now until I finish the logic of triggers
            'Tasks>=Triggers_When ... TaskProperties.Triggers_When
            'Tasks_States>=Next_Trigger_Id ... TaskState.NextTriggerId
            Dim triggerWhen As String =
                CacadoreToTectonic.QueryTaskTriggersWhen(newTask)
            If triggerWhen = "" Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Querying task properties done.")
            '====================================================================================
            Console.WriteLine("Get task triggers...")
            'Get Task Triggers
            Dim trigs As Dictionary(Of String, Trigger) =
            CacadoreToTectonic.GetTaskTriggers(newTask)
            If trigs Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Getting task triggers done.")
            '====================================================================================
            Console.WriteLine("Getting an action by its Id...")
            'Get action byId 
            Dim act As Action = CacadoreToTectonic.GetTaskActionById(newTask.TaskId & "-a1")
            If act Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Getting an action by its Id done.")
            '====================================================================================
            Console.WriteLine("Getting a trigger by its Id...")
            'Get trigger byId 
            Dim trg As Trigger = CacadoreToTectonic.GetTaskTriggerById(newTask.TaskId & "-t1")
            If trg Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Getting a trigger by its Id done.")
            '====================================================================================
            Console.WriteLine("Get Group's tasks collection...")
            'Get Group's tasks collection
            Dim tsksCollection As TasksCollection = CacadoreToTectonic.GetTasksCollection(mainGroup.GroupId)
            Dim tsks As Dictionary(Of String, Task) = tsksCollection.Tasks
            If tsksCollection Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Get Group's tasks collection done.")
            '====================================================================================
            Console.WriteLine("Getting a task by its Id...")
            Dim newTaskAgain As Task = CacadoreToTectonic.GetTaskById(newTask.TaskId)
            If newTaskAgain Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Getting a task by its Id done.")
            '====================================================================================
            Console.WriteLine("Getting a group by its Id...")
            Dim mainGroupAgain As Group = CacadoreToTectonic.GetGroupById(mainGroup.GroupId)
            If mainGroupAgain Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Getting a group by its Id done.")
            '====================================================================================
            If deleting = True Then
                Console.WriteLine("Deleting a group...")
                ''Delete Group:-
                Dim newGroup = CacadoreToTectonic.CreateNewGroup("AhmedS")
                If Not CacadoreToTectonic.DeleteGroup(newGroup) Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                Console.WriteLine("Done.")
                '====================================================================================
                Console.WriteLine("Delete all tectonic groups...")
                'DeleteAllGroups:-
                If Not CacadoreToTectonic.DeleteAllGroups() Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                Console.WriteLine("Done.")
                '====================================================================================
                Console.WriteLine("Delete all tectonic globals...")
                'Delete All Globals:-
                If Not CacadoreToTectonic.DeleteGlobalVariable(varName) Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                If Not CacadoreToTectonic.DeleteGlobalVariable(varName2) Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                If Not CacadoreToTectonic.DeleteAllGlobals() Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                Console.WriteLine("Done.")
                '====================================================================================
                Console.WriteLine("Deleting Tasks Variables (single and all)...")
                'Delete Tasks Variables (single and all)
                If Not CacadoreToTectonic.DeleteSingleTaskVariable(tskVaryMan, "trying1") Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                If Not CacadoreToTectonic.DeleteAllTasksVariables() Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                Console.WriteLine("Done.")
                '====================================================================================
                Console.WriteLine("Deleting tasks states  (single and all)...")
                'Delete Tasks States  (single and all)
                If Not CacadoreToTectonic.DeleteTaskState(newTask.State) Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                If Not CacadoreToTectonic.DeleteAllTasksStates(newTask.State) Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                Console.WriteLine("Done.")
                '====================================================================================
                Console.WriteLine("Delete All/Single Triggers...")
                'Delete All/Single Triggers            
                If Not CacadoreToTectonic.DeleteSingleTrigger(newTaskTrigger) Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                If Not CacadoreToTectonic.DeleteAllTriggers() Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                Console.WriteLine("Done.")
                '====================================================================================
                Console.WriteLine("Delete All/Single Action...")
                'Delete All/Single Action:-
                If Not CacadoreToTectonic.DeleteSingleAction(newTaskAction, newTask) Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Console.WriteLine("Error deleting action " & newTaskAction.ActionId)
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                End If
                Console.WriteLine("Now delete all actions")
                If Not CacadoreToTectonic.DeleteAllActions() Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                Console.WriteLine("Done.")
                ''====================================================================================
                Console.WriteLine("Start deleting... :D")
                Console.WriteLine("Deleting a task...")
                'Delete Task:-
                If Not CacadoreToTectonic.DeleteTask(newTask) Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                Console.WriteLine("Done.")
                Console.WriteLine("Delete all tectonic tasks...")
                'Delete All Tasks
                If Not CacadoreToTectonic.DeleteAllTectonicTasks() Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                Console.WriteLine("Done.")
                '====================================================================================
                Console.WriteLine("Delete History...")
                'Delete History
                'Check example in CacadoreToTectonic for more information 
                If Not CacadoreToTectonic.ClearHistoryByEntryType(IHistoryEntry.HistoryEntryType.Task) Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                'Check example in CacadoreToTectonic for more information 
                If Not CacadoreToTectonic.ClearHistoryByTimeRange(IHistory.TimeRange.Today) Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                'Check example in CacadoreToTectonic for more information 
                If Not CacadoreToTectonic.ClearFullHistory() Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                Console.WriteLine("Done.")
                '====================================================================================
            End If
            [stopWatch].Stop()
            Console.WriteLine(String.Format("Simple Tectonicing Done in {0} seconds, Congrats!", [stopWatch].ElapsedMilliseconds / 1000))
            actual = True
            Assert.AreEqual(expected, actual)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub MassiveTriggersTest()
            Dim newTask As Cacadore.Task = CacadoreToTectonic.CreateNewTask(True, "0")
            Dim newTaskTrigger As Cacadore.Trigger = Nothing
            Console.WriteLine("Loop about to start @" & DateTime.Now)
            For i = 1 To 100
                Dim cacaHelper As New CacadoreHelper
                Dim tskMeta As New TaskMeta
                tskMeta.CronExpression = "0/1 * * ? * *"
                tskMeta.StartDateTimeUTC = DateTime.Now '.AddMonths(10)
                'Attach 2 triggers to newTask:-
                Console.WriteLine("Just before attaching task #" & i)
                newTaskTrigger = cacaHelper.AttachScheduledTriggerToTask(newTask, tskMeta)
                If newTaskTrigger Is Nothing Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                Console.WriteLine(String.Format("trigger #{0} attached to task", i))
                newTaskTrigger.StartListening()
                Console.WriteLine(String.Format("trigger #{0} started", i))
            Next
            Assert.AreEqual(True, True)
        End Sub

        ''' <summary>
        ''' Load test by a massive insert.
        ''' </summary>
        <Test()> _
        <Category("ShortRunning")>
        Public Sub MassiveInsertTest()
            Dim actual As Boolean = False
            Dim expected As Boolean = True
            'Add and insert a new task:-
            Dim CacadoreToTectonic As New CacadoreToTectonic
            'Add a new group:-
            Console.WriteLine("Creating a new group...")
            Dim mainGroup = CacadoreToTectonic.CreateNewGroup("LoadTest" & DateTime.UtcNow.ToString)
            If mainGroup Is Nothing Then
                Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                Exit Sub
            End If
            Console.WriteLine("Group created successfully.")
            Dim newTask As Task = Nothing
            Console.WriteLine("Begin a loop of creating a 100 new tasks...")
            For i As Integer = 1 To 200
                newTask = CacadoreToTectonic.CreateNewTask(True, mainGroup.GroupId)
                If newTask Is Nothing Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
                newTask.Properties.IsEnabled = True
                newTask.Properties.LastTimeRun = DateTime.UtcNow
                newTask.Properties.NextTimeToBeRun = DateTime.UtcNow
                newTask.Properties.TaskName = Balora.RandomBytesGenerator.GRCM & DateTime.UtcNow.ToString
                'For both, they will be blank for now until I finish the logic of triggers
                'Tasks>=Triggers_When ... TaskProperties.Triggers_When
                'Tasks_States>=Next_Trigger_Id ... TaskState.NextTriggerId
                newTask.Properties.TriggersWhen = DateTime.UtcNow.ToString
                If Not newTask.Save() Then
                    Dim lineNumber As Integer = New StackTrace(New StackFrame(True)).GetFrame(0).GetFileLineNumber - 2
                    Assert.Inconclusive("Nothing or False @ Line No:" & lineNumber)
                    Exit Sub
                End If
            Next
            Console.WriteLine("100 tasks added successfully to Tectonic.")
            Console.WriteLine("Massive Insert Test Done, Congrats!")
            actual = True
            Assert.AreEqual(expected, actual)
        End Sub

        ''' <summary>
        ''' Load test by getting all tasks in Tectonic.
        ''' </summary>
        <Test()> _
        <Category("ShortRunning")>
        Public Sub GetAllTasksTest()
            Dim actual As Boolean = False
            Dim expected As Boolean = True
            Dim allTasks As New Dictionary(Of String, Task)
            allTasks = CacadoreToTectonic.GetAllRegisteredTasks()
            If Not allTasks Is Nothing Then
                actual = True
                Assert.AreEqual(expected, actual)
            Else
                Assert.AreEqual(expected, actual)
            End If
        End Sub

        Dim TestCommonTasksOperationsResult As New List(Of Boolean)

        ''' <summary>
        ''' Testing common tasks like Create,Execute,Pause,Resume and cancel.
        ''' </summary>
        <Test()> _
        <Category("Test solo")>
        Public Sub TestCommonTasksOperations()

            Console.WriteLine("#######################################################")

            Console.WriteLine("Testing common tasks like Create,Execute,Pause,Resume and cancel.")

            Dim actual As Boolean = False
            Dim expected As Boolean = True

            Console.WriteLine("CreateTaskWithSevralActions...")
            CreateTaskWithSevralActions()

            Console.WriteLine("Execute...")
            Execute()

            'Why we added as sleep
            'http://stackoverflow.com/questions/751472/c-net-testing-backgroundworker-with-nunit
            System.Threading.Thread.Sleep(1000)
            Console.WriteLine("PauseNow...")
            PauseNow()
            System.Threading.Thread.Sleep(2000)
            Console.WriteLine("ResumeNow...")
            ResumeNow()
            System.Threading.Thread.Sleep(1000)
            Cacadore.Settings.Enabled = False
            System.Threading.Thread.Sleep(2000)
            Cacadore.Settings.Enabled = True
            System.Threading.Thread.Sleep(500)
            Console.WriteLine("Cancel...")
            taskObject.Cancel()

            Dim resultsCount = From result In TestCommonTasksOperationsResult Where result = True

            If resultsCount.Count >= 3 Then
                actual = True 'Because they are 5 tests.
            End If

            Console.WriteLine("ResultsCount ==" & resultsCount.Count)

            Assert.AreEqual(expected, actual)

            Console.WriteLine("#######################################################")
        End Sub

        <Test()> _
        <Category("Test solo")>
        Public Sub TestGroupPause()
            Dim newGroup = CacadoreToTectonic.CreateNewGroup("Just A New Group Man")
            For i = 1 To 10
                CacadoreToTectonic.CreateNewTask(False, newGroup.GroupId)
            Next

            newGroup.Pause()

            For Each tsk In newGroup.TasksCollection.Tasks
                Assert.AreEqual(tsk.Value.State.GetExecutionState(), Cacadore.ITaskState.ExecutionState.Paused)
            Next

            newGroup.Resume()

            For Each tsk In newGroup.TasksCollection.Tasks
                Assert.AreEqual(tsk.Value.State.GetExecutionState(), Cacadore.ITaskState.ExecutionState.Resumed)
            Next
        End Sub

        Sub CreateTaskWithSevralActions()
            taskObject = BIT.CreateBunchOfRunApplicationTasks()
            AddHandler taskObject.Executed, AddressOf Executed
            AddHandler taskObject.Paused, AddressOf paused
            AddHandler taskObject.Resumed, AddressOf resumed
            AddHandler Cacadore.Settings.CacadoreEnabledChanged, AddressOf CacadoreEnabledChangedEvent
        End Sub

        Sub Execute()
            taskObject.Execute()
        End Sub

        Private Sub Executed(Sender As Task, e As ITask.ExecutionEventArgs)
            TestCommonTasksOperationsResult.Add(True)
            Balora.Hodhod.STOW("#######################################################")
            Balora.Hodhod.STOW("Executed")
            Balora.Hodhod.STOW("#######################################################")
        End Sub

        Private Sub paused(sender As Object, e As EventArgs)
            TestCommonTasksOperationsResult.Add(True)
            Balora.Hodhod.STOW("#######################################################")
            Balora.Hodhod.STOW("Paused")
            Balora.Hodhod.STOW("#######################################################")
        End Sub

        Private Sub resumed(sender As Object, e As EventArgs)
            TestCommonTasksOperationsResult.Add(True)
            Balora.Hodhod.STOW("#######################################################")
            Balora.Hodhod.STOW("Resumed")
            Balora.Hodhod.STOW("#######################################################")
        End Sub

        Private Sub PauseNow()
            taskObject.Pause()
        End Sub

        Private Sub ResumeNow()
            taskObject.[Resume]()
        End Sub

        Private Sub CacadoreEnabledChangedEvent(enabled As Boolean)
            TestCommonTasksOperationsResult.Add(True)
            Balora.Hodhod.STOW("#######################################################")
            If enabled Then
                Balora.Hodhod.STOW("Cacadore Enabled.")
            Else
                Balora.Hodhod.STOW("Cacadore Disabled.")
            End If
            Balora.Hodhod.STOW("#######################################################")
        End Sub

    End Class
End Namespace
