'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   TypedDataSetCrud Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
Imports Balora
Imports Cacadore
Imports Balora.Datler
Imports Balora.PathsHelper

Public Class TypedDataSetCrud
    Implements ICrud
    Implements IDisposable
    Implements IBroadcaster

    '++You can add various queries methods here,
    '++Like, GetAllVariables or GetAllTasks
    '++It will be funny for just getting the whole thing
    '++from within...
#Region "Fields"
    Delegate Sub MethodToBeInjectedDlg()
    Private _tectonicListeners As New ArrayList
    Private _timeTakenToFillAllTablesValue As TimeSpan
    Public Shared TectonicDataSetObj As New TectonicDataSet
    Private Shared _connection As System.Data.Common.DbConnection 'New Devart.Data.SQLite.SQLiteConnection
    'Private Shared callingDictionary As New Dictionary(Of String, Integer)

#Region "Adapters Instances"
    Private Shared _tasksTableAdapter As New TectonicDataSetTableAdapters.tasksTableAdapter
    Private Shared _groupsTableAdapter As New TectonicDataSetTableAdapters.groupsTableAdapter
    Private Shared _globalsTableAdapter As New TectonicDataSetTableAdapters.globalsTableAdapter
    Private Shared _globallookupTableAdapterAdapter As New TectonicDataSetTableAdapters.global_lookupTableAdapter
    Private Shared _tasksStatusTableAdapter As New TectonicDataSetTableAdapters.tasks_statesTableAdapter
    Private Shared _tasksVariablesTableAdapter As New TectonicDataSetTableAdapters.task_variablesTableAdapter
    Private Shared _triggersTableAdapter As New TectonicDataSetTableAdapters.triggersTableAdapter
    Private Shared _actionsTableAdapter As New TectonicDataSetTableAdapters.actionsTableAdapter
    Private Shared _historyTableAdapter As New TectonicDataSetTableAdapters.historyTableAdapter
    Private Shared _settingsTableAdapter As New TectonicDataSetTableAdapters.settingsTableAdapter
#End Region
#End Region
#Region "Properties"

    Private Shared _tectonicPathValue As String
    Public Shared ReadOnly Property TectonicPath As String
        Get
            Return _tectonicPathValue
        End Get
    End Property

    Private Shared _tectonicBackupPathValue As String
    Public Shared ReadOnly Property TectonicBackupPath As String
        Get
            Return _tectonicBackupPathValue
        End Get
    End Property

    Private Shared _tectonicTempPath As String
    Public Shared ReadOnly Property TectonicTempPath As String
        Get
            Return _tectonicTempPath
        End Get
    End Property

    Public Shared ReadOnly Property Connection() As System.Data.Common.DbConnection
        Get
            Return _connection
        End Get
    End Property

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Initializes a new instance of the <see cref="TypedDataSetCrud" /> class.
    ''' </summary>
    ''' <remarks>
    ''' This constructor fills all Tectonic tables using 
    ''' adapters in TectonicDataSetTableAdapters namespace.
    ''' </remarks>
    Public Sub New()
        'Dim bckGrndWrkHelper As New Balora.BackgroundWorkerHelper
        'Dim workToDoDelegate As MethodToBeInjectedDelegate = New MethodToBeInjectedDelegate(AddressOf myConstructor)
        'bckGrndWrkHelper.WorkToDo = workToDoDelegate
        'bckGrndWrkHelper.Start()
        myConstructor()
    End Sub

    Private Sub setTectonicPaths()
        Dim commonAppDataPath As String = GetCommonApplicationData() & "Cacadu\"
#If DEBUG Then
        _tectonicPathValue = commonAppDataPath & "working_tectonic"
        _tectonicBackupPathValue = commonAppDataPath & "working_tectonicB"
        _tectonicTempPath = commonAppDataPath & "working_tectonicT"
#Else
        _tectonicPathValue = commonAppDataPath & "tectonic"
        _tectonicBackupPathValue = commonAppDataPath & "tectonicB"
        _tectonicTempPath = commonAppDataPath & "tectonicT"
#End If
        commonAppDataPath = ""
    End Sub

    Public Shared Sub CloseConnection()
        If Not IsNothing(_connection) Then
            _connection.Close()

            _tasksVariablesTableAdapter.Connection.Close()
            _tasksStatusTableAdapter.Connection.Close()
            _tasksTableAdapter.Connection.Close()
            _groupsTableAdapter.Connection.Close()
            _globalsTableAdapter.Connection.Close()
            _triggersTableAdapter.Connection.Close()
            _actionsTableAdapter.Connection.Close()
            _historyTableAdapter.Connection.Close()
            _settingsTableAdapter.Connection.Close()
            _globallookupTableAdapterAdapter.Connection.Close()


            _connection = Nothing
            _tasksVariablesTableAdapter = Nothing
            _tasksStatusTableAdapter = Nothing
            _tasksTableAdapter = Nothing
            _groupsTableAdapter = Nothing
            _globalsTableAdapter = Nothing
            _triggersTableAdapter = Nothing
            _actionsTableAdapter = Nothing
            _historyTableAdapter = Nothing
            _settingsTableAdapter = Nothing
            _globallookupTableAdapterAdapter = Nothing

            TectonicDataSetObj.Dispose()
            TectonicDataSetObj = Nothing
        End If
    End Sub

    Private Sub setConnectionStrings()
        'Will be used in malformation method (_connection is a back field for Me.Connection property).
        _connection = _settingsTableAdapter.Connection
        'Setting connection string manually.
        Dim _connectionString As String = "Data Source=" & _tectonicPathValue & ";AutoVacuum=Full;"
        _tasksVariablesTableAdapter.Connection.ConnectionString = _connectionString
        _tasksStatusTableAdapter.Connection.ConnectionString = _connectionString
        _tasksTableAdapter.Connection.ConnectionString = _connectionString
        _groupsTableAdapter.Connection.ConnectionString = _connectionString
        _globalsTableAdapter.Connection.ConnectionString = _connectionString
        _triggersTableAdapter.Connection.ConnectionString = _connectionString
        _actionsTableAdapter.Connection.ConnectionString = _connectionString
        _historyTableAdapter.Connection.ConnectionString = _connectionString
        _settingsTableAdapter.Connection.ConnectionString = _connectionString
        _globallookupTableAdapterAdapter.Connection.ConnectionString = _connectionString
    End Sub

    'AhSaElDin 20111227: Change it to a function to get a return value if success or not.
    Private Function doMalformationCheck() As Boolean
        Dim malformationCheckingResult As String = DevartSQLiteDatabase.CheckIfMalformed(TecDAL.TypedDataSetCrud.Connection)
        If malformationCheckingResult <> "ok" And malformationCheckingResult <> "" Then
            If FileIO.FileSystem.FileExists(TectonicBackupPath) Then

                'mald stand for malformed database
                MsgBox(My.Resources.mald,
                     MsgBoxStyle.Critical,
                     "Cacadu")

                Try
                    'Restore Tectonic...
                    FileCopy(TectonicBackupPath, TectonicPath)
                Catch ex As Exception
                    If Debugger.IsAttached Then Debugger.Break()
                    Alerter.REP("", ex, True)
                    Return False
                End Try

            Else
                'AhSaElDin 20111227: what if backup isn't there.
                'dd stand for database damaged.
                MsgBox(My.Resources.dd,
                     MsgBoxStyle.Critical,
                     "Cacadu")
                Return False
            End If
        End If
        Return True
    End Function

    Private Sub fillAllTables()
        _tasksVariablesTableAdapter.Fill(TectonicDataSetObj.task_variables)
        _tasksStatusTableAdapter.Fill(TectonicDataSetObj.tasks_states)
        _tasksTableAdapter.Fill(TectonicDataSetObj.tasks)
        _groupsTableAdapter.Fill(TectonicDataSetObj.groups)
        _globalsTableAdapter.Fill(TectonicDataSetObj.globals)
        _triggersTableAdapter.Fill(TectonicDataSetObj.triggers)
        _actionsTableAdapter.Fill(TectonicDataSetObj.actions)
        _historyTableAdapter.Fill(TectonicDataSetObj.history)
        _settingsTableAdapter.Fill(TectonicDataSetObj.settings)
        _globallookupTableAdapterAdapter.Fill(TectonicDataSetObj.global_lookup)
    End Sub

    ''' <summary>
    ''' Private method to handle Constructor job,
    ''' you can use for either hiding the implementation 
    ''' from reflectors or init tables in a separate thread.
    ''' </summary>
    Private Sub myConstructor()
        setTectonicPaths()
        setConnectionStrings()
        'Return true if not malformed database
        If doMalformationCheck() Then
            fillAllTables()
        Else
            Debug.Print("write my.resource.dd message in statas bar")
        End If
    End Sub
#End Region
#Region "Crud Methods"

#Region "Main Crud Methods"

    ''' <summary>
    ''' Query tectonic with the command text
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' Here I'm using Datler itself rather than using TectonicDataSet because of 2 reasons
    ''' 1.I don't know how can I query multiple tables using TectonicDataSet, however, I can
    ''' do it simply using Datler.
    ''' 2.If I used datler here, the code will be very simple and I won't write select case
    ''' statement to handle all tables.
    ''' </remarks>
    Private Function Query(ByVal cacadoreObj As ICrudable, ByVal ParamArray args() As String) As Object Implements ICrud.Query
        Dim res As Object = "nothing"
        'Dim arg As String = vbNullString
        'If args.Length > 0 Then arg = args(0)
        'Console.WriteLine(String.Format("A query of {0} has started... with arg(0)= {1}", cacadoreObj.GetType.Name, arg))
        Select Case cacadoreObj.GetType.Name
            Case "Commands"
                If args.Contains(CStr(Commands.CommandsEnum.GetTaskProperty)) Then
                    res = GetTaskProperty(args)
                ElseIf args.Contains(CStr(Commands.CommandsEnum.GetActionById)) Then
                    res = GetActionById(args)
                ElseIf args.Contains(CStr(Commands.CommandsEnum.GetTriggerById)) Then
                    res = GetTriggerById(args)
                ElseIf args.Contains(CStr(Commands.CommandsEnum.GetTaskById)) Then
                    res = GetTaskbyId(args)
                ElseIf args.Contains(CStr(Commands.CommandsEnum.GetGroupById)) Then
                    res = GetGroupById(args)
                ElseIf args.Contains(CStr(Commands.CommandsEnum.GetAllTasks)) Then
                    res = GetAllTasks(args)
                ElseIf args.Contains(CStr(Commands.CommandsEnum.GetAllTasksStates)) Then
                    res = GetAllTasksStates(args)
                ElseIf args.Contains(CStr(Commands.CommandsEnum.GetUnfullfilledTasks)) Then
                    res = GetUnfullfilledTasks(args)
                ElseIf args.Contains(CStr(Commands.CommandsEnum.GetAllTriggers)) Then
                    res = GetAllTriggers(args)
                ElseIf args.Contains(CStr(Commands.CommandsEnum.GetTimeOutTriggers)) Then
                    res = GetTimeOutTriggers(args)
                ElseIf args.Contains(CStr(Commands.CommandsEnum.GetTaskExecutionState)) Then
                    res = GetTaskExecutionState(args)
                ElseIf args.Contains(CStr(Commands.CommandsEnum.IsRegisterdTask)) Then
                    res = IsRegisteredTask(args(1))
                ElseIf args.Contains(CStr(Commands.CommandsEnum.IsRegisterdTaskVariable)) Then
                    res = IsVariableAttachedToThisTask(args(2), args(1))
                ElseIf args.Contains(CStr(Commands.CommandsEnum.GetTaskTriggersCount)) Then
                    res = GetTaskTriggersCount(args(1))
                ElseIf args.Contains(CStr(Commands.CommandsEnum.GetTaskActionsCount)) Then
                    res = GetTaskActionsCount(args(1))
                ElseIf args.Contains(CStr(Commands.CommandsEnum.IsRegisterdAction)) Then
                    res = IsRegisteredAction(args(1))
                End If
            Case "History"
                res = GetHistory(args)
            Case "Globals"
                res = GetGlobals(args)
            Case "Settings"
                res = GetSettings(args)
            Case "TaskVariables"
                If args(0).Contains("GetVariableValue") Then
                    res = GetVariable(args)
                Else
                    res = GetTaskVariables(args)
                End If
            Case "GroupsCollection"
                res = GetAllGroups(args)
            Case "TriggersCollection"
                res = GetTaskTriggers(args)
            Case "ActionsCollection"
                res = GetTaskActions(args(0))
            Case "TasksCollection"
                res = GetGroupTasks(args)
        End Select
        args = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Inserts the specified cacadore obj into tectonic.
    ''' </summary>
    ''' <param name="cacadoreObj">The cacadore obj received by Cacadore.</param>
    ''' <param name="args">The args that you may need in inserting</param><returns></returns>
    Private Function Insert(ByVal cacadoreObj As ICrudable, ByVal ParamArray args() As String) As Boolean Implements ICrud.Insert
        Dim res As Boolean
        'Dim arg As String = vbNullString
        'If args.Length > 0 Then arg = args(0)
        'Console.WriteLine(String.Format("An Insert of {0} has started... with arg(0)= {1}", cacadoreObj.GetType.Name, arg))
        Dim baseType As String = cacadoreObj.GetType.BaseType().Name
        'AhSaElDin 20111213: Added next line to support class like ScheduleTrigger.
        Dim baseTypeOfBaseType = cacadoreObj.GetType.BaseType().BaseType.Name
        If baseType = "Trigger" Or baseTypeOfBaseType = "Trigger" Then
            res = UpdateTrigger(CType(cacadoreObj, Trigger), True)
        ElseIf baseType = "Action" Or baseTypeOfBaseType = "Action" Then
            res = UpdateAction(CType(cacadoreObj, Action), True)
        End If
        Select Case cacadoreObj.GetType.Name
            Case "Task"
                res = UpdateTask(CType(cacadoreObj, Task))
            Case "Group"
                res = UpdateGroup(CType(cacadoreObj, Group))
            Case "GlobalEntry"
                res = UpdateGlobals(CType(cacadoreObj, Cacadore.GlobalEntry), True)
            Case "Variable"
                res = UpdateTaskVariable(CType(cacadoreObj, Variable), True)
            Case "HistoryEntry"
                res = InsertNewHistoryEntry(cacadoreObj)
            Case "Settings" 'No need to insert in this table, coz it has only one row So update it only
        End Select
        args = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Updates the specified cacadore obj in Tectonic.
    ''' </summary>
    ''' <param name="cacadoreObj">The cacadore obj.</param>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function Update(ByVal cacadoreObj As ICrudable, ByVal ParamArray args() As String) As Boolean Implements ICrud.Update
        Dim res As Boolean = False
        'Dim arg As String = vbNullString
        'If args.Length > 0 Then arg = args(0)
        'Console.WriteLine(String.Format("An Update of {0} has started... with arg(0)= {1}", cacadoreObj.GetType.Name, arg))
        Dim baseType As String = cacadoreObj.GetType.BaseType().Name
        'AhSaElDin 20111213: Added next line to support class like ScheduleTrigger.
        Dim baseTypeOfBaseType = cacadoreObj.GetType.BaseType().BaseType.Name
        If baseType = "Trigger" Or baseTypeOfBaseType = "Trigger" Then 'No need until now for updating triggers table
            res = UpdateTrigger(CType(cacadoreObj, Trigger))
        ElseIf baseType = "Action" Or baseTypeOfBaseType = "Action" Then
            res = UpdateAction(CType(cacadoreObj, Action))
        End If
        Select Case cacadoreObj.GetType.Name
            Case "Commands"
                If args.Contains(CStr(Commands.CommandsEnum.MoveUpActionExecutionPrecedence)) Then
                    moveUpActionExecutionPrecedence(args(1), CInt(args(2)))
                ElseIf args.Contains(CStr(Commands.CommandsEnum.MoveDownActionExecutionPrecedence)) Then
                    moveDownActionExecutionPrecedence(args(1), CInt(args(2)))
                End If
            Case "Task"
                res = UpdateTask(CType(cacadoreObj, Task))
            Case "Settings"
                res = UpdateSettings(CType(cacadoreObj, Cacadore.Settings), args)
            Case "GlobalEntry"
                res = UpdateGlobals(CType(cacadoreObj, Cacadore.GlobalEntry))
            Case "Group"
                res = UpdateGroup(CType(cacadoreObj, Group))
            Case "Variable"
                res = UpdateTaskVariable(CType(cacadoreObj, Variable), )
            Case "Trigger" 'No need until now for updating triggers table
        End Select
        args = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Deletes the specified cacadore obj from Tectonic.
    ''' </summary>
    ''' <param name="cacadoreObj">The cacadore obj.</param>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function Delete(ByVal cacadoreObj As ICrudable, ByVal ParamArray args() As String) As Boolean Implements ICrud.Delete
        Dim res As Boolean = False
        'Dim arg As String = vbNullString
        'If args.Length > 0 Then arg = args(0)
        'Console.WriteLine(String.Format("A Delete of {0} has started... with arg(0)= {1}", cacadoreObj.GetType.Name, arg))
        Dim baseType As String = cacadoreObj.GetType.BaseType().Name
        'AhSaElDin 20111213: Added next line to support class like ScheduleTrigger.
        Dim baseTypeOfBaseType = cacadoreObj.GetType.BaseType().BaseType.Name
        If baseType = "Trigger" Or baseTypeOfBaseType = "Trigger" Then 'No need until now for updating triggers table
            res = DeleteTrigger(CType(cacadoreObj, Trigger))
        ElseIf baseType = "Action" Or baseTypeOfBaseType = "Action" Then
            res = DeleteAction(CType(cacadoreObj, Action))
        End If
        Select Case cacadoreObj.GetType.Name
            Case "Commands"
                If args.Contains(CStr(Commands.CommandsEnum.DeleteAllTasks)) Then
                    res = DeleteAllTasks()
                ElseIf args.Contains(CStr(Commands.CommandsEnum.DeleteAllGroups)) Then
                    res = DeleteAllGroups()
                ElseIf args.Contains(CStr(Commands.CommandsEnum.DeleteAllGlobals)) Then
                    res = DeleteAllGlobalsVariables()
                ElseIf args.Contains(CStr(Commands.CommandsEnum.DeleteAllVariables)) Then
                    res = DeleteAllVariables()
                ElseIf args.Contains(CStr(Commands.CommandsEnum.DeleteAllTasksStates)) Then
                    res = DeleteAllTasksStates()
                ElseIf args.Contains(CStr(Commands.CommandsEnum.DeleteAllActions)) Then
                    res = DeleteAllActions()
                ElseIf args.Contains(CStr(Commands.CommandsEnum.DeleteAllTriggers)) Then
                    res = DeleteAllTriggers()
                End If
            Case "Task"
                res = DeleteTask(CType(cacadoreObj, Task))
            Case "Group"
                res = DeleteGroup(CType(cacadoreObj, Group))
            Case "Globals"
                res = DeleteGlobalVariable(args(0))
            Case "TaskVariables"
                res = DeleteVariable(CType(cacadoreObj, TaskVariables), args)
            Case "TaskState"
                res = DeleteTaskState(args(0))
            Case "Trigger"
                res = DeleteTrigger(CType(cacadoreObj, Trigger))
            Case "Action"
                res = DeleteAction(CType(cacadoreObj, Action))
            Case "History"
                If args.Contains(Cacadore.History.ClearType.ByEntryType.ToString) Then
                    Dim historyEntryObj = CType(cacadoreObj, History)
                    res = DeleteHistoryByEntryType(CType(args(1), IHistoryEntry.HistoryEntryType))
                    'args(1) = entry type you want to delete as a parameter in ParamArray args()
                ElseIf args.Contains(Cacadore.History.ClearType.ByTimeRange.ToString) Then
                    res = DeleteHistoryByTimeRange(args)
                ElseIf args.Contains(Cacadore.History.ClearType.All.ToString) Then
                    res = DeleteFullHistory()
                End If
        End Select
        args = Nothing
        Return res
    End Function
#End Region

#Region "Query"

    ''' <summary>
    ''' Gets the history of the specified entry type in args().
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function GetHistory(ByVal ParamArray args() As String) As List(Of HistoryEntry)
        Dim taskId As String = ""
        Dim entryType As IHistoryEntry.HistoryEntryType = Nothing
        Dim tasksHistoryRows As System.Data.EnumerableRowCollection(Of TectonicDataSet.historyRow) = Nothing
        If args(0).ToString = "0" Then 'Task
            entryType = IHistoryEntry.HistoryEntryType.Task
        ElseIf args(0).ToString = "1" Then 'Action
            entryType = IHistoryEntry.HistoryEntryType.Action
        ElseIf args(0).ToString = "2" Then 'Trigger
            entryType = IHistoryEntry.HistoryEntryType.Trigger
        Else
            taskId = args(0)
        End If
        If taskId = "" Then
            tasksHistoryRows = From row As TectonicDataSet.historyRow In TectonicDataSetObj.history
            Where row.entry_type = entryType
        Else
            tasksHistoryRows = From row As TectonicDataSet.historyRow In TectonicDataSetObj.history
            Where row.task_id = taskId
        End If
        Dim historyEntiresList As New List(Of HistoryEntry)
        Dim histEntry As HistoryEntry = Nothing
        For Each row In tasksHistoryRows
            histEntry = New HistoryEntry
            histEntry.EntryDateTime = CDate(row.date_entry_date) 'DateTime.FromBinary(CLng(row.date_entry_date))
            histEntry.EntryType = CType(row.entry_type, IHistoryEntry.HistoryEntryType)
            histEntry.Description = row.message
            histEntry.Status = CType(row.status, ITaskState.ExecutionState)
            histEntry.TaskId = row.task_id
            historyEntiresList.Add(histEntry)
        Next
        args = Nothing
        entryType = Nothing
        tasksHistoryRows = Nothing
        Return historyEntiresList
    End Function

    ''' <summary>
    ''' Gets the settings.
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function GetSettings(ByVal ParamArray args() As String) As Boolean
        Dim settingsRow As System.Data.EnumerableRowCollection(Of TectonicDataSet.settingsRow) = Nothing
        settingsRow = From row As TectonicDataSet.settingsRow In TectonicDataSetObj.settings
        For Each row In settingsRow
            If CDbl(args(0)) = ISettings.SettingsEnum.CacadoreEnabled Then 'CacadoreEnabled
                settingsRow = Nothing
                Return CBool(row.is_cacadore_enabled)
            ElseIf CDbl(args(0)) = ISettings.SettingsEnum.HistoryEnabled Then 'HistoryEnabled
                settingsRow = Nothing
                Return CBool(row.is_history_enabled)
            ElseIf CDbl(args(0)) = ISettings.SettingsEnum.TriggersEnabled Then 'TriggersEnabled
                settingsRow = Nothing
                Return CBool(row.is_triggers_enabled)
            End If
        Next
        args = Nothing
        settingsRow = Nothing
        Return Nothing 'Nothing executed, nothing returned.
    End Function

    ''' <summary>
    ''' Gets the globals where args(0) == row.type column.
    ''' </summary>
    ''' <param name="args">The args. args contains 0 or 1, 0 means constant, 1 means variable</param><returns></returns>
    Private Function GetGlobals(ByVal ParamArray args() As String) As Dictionary(Of String, String)
        Dim globalsRows As System.Data.EnumerableRowCollection(Of TectonicDataSet.globalsRow) = Nothing
        globalsRows = From row As TectonicDataSet.globalsRow In TectonicDataSetObj.globals
                      Where row.type = CInt(args(0))
        Dim globalsEntriesDic As New Dictionary(Of String, String)
        For Each row In globalsRows
            globalsEntriesDic.Add(row.name, row.value)
        Next
        args = Nothing
        globalsRows = Nothing
        Return globalsEntriesDic
    End Function

    ''' <summary>
    ''' Gets all registered groups
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function GetAllGroups(ByVal ParamArray args() As String) As Dictionary(Of String, Group)
        Dim groupsRows As System.Data.EnumerableRowCollection(Of TectonicDataSet.groupsRow) = Nothing
        groupsRows = From row As TectonicDataSet.groupsRow In TectonicDataSetObj.groups
        Dim groupsDic As New Dictionary(Of String, Group)
        Dim tempGroup As Group = Nothing
        For Each row In groupsRows
            tempGroup = GroupRowToGroupObject(row)
            groupsDic.Add(tempGroup.GroupId, tempGroup)
        Next
        args = Nothing
        tempGroup.Dispose()
        groupsRows = Nothing
        Return groupsDic
    End Function

    ''' <summary>
    ''' Gets the variable value using its name and its taskId.
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function GetVariable(ByVal ParamArray args() As String) As String
        Dim variablesRows As System.Data.EnumerableRowCollection(Of TectonicDataSet.task_variablesRow) = Nothing
        variablesRows = From row As TectonicDataSet.task_variablesRow In TectonicDataSetObj.task_variables
                        Where row.task_id = args(1) And row.variable_name = args(2)
        For Each row In variablesRows
            variablesRows = Nothing
            Return row.variable_value
        Next
        args = Nothing
        variablesRows = Nothing
        Return "Nothing Found"
    End Function

    ''' <summary>
    ''' Gets the task variables (name, value), using taskId.
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function GetTaskVariables(ByVal ParamArray args() As String) As Dictionary(Of String, String)
        Dim variablesRows As System.Data.EnumerableRowCollection(Of TectonicDataSet.task_variablesRow) = Nothing
        variablesRows = From row As TectonicDataSet.task_variablesRow In TectonicDataSetObj.task_variables
                        Where row.task_id = args(0)
        Dim variablesDic As New Dictionary(Of String, String)
        For Each row In variablesRows
            variablesDic.Add(row.variable_name, row.variable_value)
        Next
        args = Nothing
        variablesRows = Nothing
        Return variablesDic
    End Function

    ''' <summary>
    ''' Gets the task property value using task id.
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function GetTaskProperty(ByVal ParamArray args() As String) As Object
        Dim tasksRows As System.Data.EnumerableRowCollection(Of TectonicDataSet.tasksRow) = Nothing
        tasksRows = From row As TectonicDataSet.tasksRow In TectonicDataSetObj.tasks
                        Where row.id = args(1)
        For Each row In tasksRows
            If args(2).Contains(TaskProperties.PropertiesNames.TaskName.ToString) Then
                args = Nothing
                tasksRows = Nothing
                Return row.name
            ElseIf args(2).Contains(TaskProperties.PropertiesNames.IsTaskEnabled.ToString) Then
                args = Nothing
                tasksRows = Nothing
                Return row.is_enables
            ElseIf args(2).Contains(TaskProperties.PropertiesNames.LastTimeRunResult.ToString) Then
                args = Nothing
                tasksRows = Nothing
                Return row.last_time_run_result
            ElseIf args(2).Contains(TaskProperties.PropertiesNames.LastTimeRun.ToString) Then
                args = Nothing
                tasksRows = Nothing
                Return row.date_last_time_run 'DateTime.FromBinary(CLng(row.date_last_time_run))
            ElseIf args(2).Contains(TaskProperties.PropertiesNames.NextTimeToBeRun.ToString) Then
                args = Nothing
                tasksRows = Nothing
                Return row.date_nexttime_run 'DateTime.FromBinary(CLng(row.date_nexttime_run))
            ElseIf args(2).Contains(TaskProperties.PropertiesNames.TriggerWhen.ToString) Then
                args = Nothing
                tasksRows = Nothing
                Return row.triggers_when
            End If
        Next
        args = Nothing
        tasksRows = Nothing
        Return "Property not found"
    End Function

    ''' <summary>
    ''' Gets all task properties using task id.
    ''' </summary>
    ''' <param name="taskId">The task id.</param><returns></returns>
    ''' <remarks >
    ''' Because this method is only called from GetGroupTasks(), you don't need to have
    ''' a ParamArray in the parameters like others. But you can change as needed in the
    ''' future.
    ''' </remarks>
    Private Function GetTaskProperties(taskId As String) As TaskProperties
        Dim tasksRows As System.Data.EnumerableRowCollection(Of TectonicDataSet.tasksRow) = Nothing
        tasksRows = From row As TectonicDataSet.tasksRow In TectonicDataSetObj.tasks
                        Where row.id = taskId
        Dim tskProperties As New TaskProperties(taskId)
        For Each row In tasksRows
            tskProperties.TaskName = row.name
            tskProperties.TriggersWhen = row.triggers_when
            tskProperties.IsEnabled = CBool(row.is_enables)
            Try
                If row.date_last_time_run <> " " Then tskProperties.LastTimeRun = CDate(row.date_last_time_run) 'DateTime.FromBinary(CLng(row.date_last_time_run))
                If row.date_nexttime_run <> " " Then tskProperties.NextTimeToBeRun = CDate(row.date_nexttime_run) 'DateTime.FromBinary(CLng(row.date_nexttime_run))
            Catch ex As Exception
                'Balora.Hodhod.STOW(ex.Message)
            End Try
            tskProperties.SetLastTimeRunResult(row.last_time_run_result)
            tasksRows = Nothing
            Return tskProperties
        Next
        tskProperties.Dispose()
        tskProperties = Nothing
        tasksRows = Nothing
        Return Nothing
    End Function

    Private Function GetTaskExecutionState(args As String()) As ITaskState.ExecutionState
        Dim tskState = GetTaskState(args(1))
        Dim state = tskState.TaskExecutionState
        tskState.Dispose()
        args = Nothing
        tskState = Nothing
        Return state
    End Function

    ''' <summary>
    ''' Gets the state of the task.
    ''' </summary>
    ''' <param name="taskId">The task id.</param><returns></returns>
    ''' <remarks >
    ''' Because this method is only called from GetGroupTasks(), you don't need to have
    ''' a ParamArray in the parameters like others. But you can change as needed in the
    ''' future.
    ''' </remarks>
    Private Function GetTaskState(taskId As String) As TaskState
        Dim tasksStateRows As System.Data.EnumerableRowCollection(Of TectonicDataSet.tasks_statesRow) = Nothing
        tasksStateRows = From row As TectonicDataSet.tasks_statesRow In TectonicDataSetObj.tasks_states
                        Where row.task_id = taskId
        Dim tskState As New TaskState(taskId)
        For Each row In tasksStateRows
            'AhSaElDin 20111231: property is writable for now.
            'If Not row.Isnext_trigger_idNull Then Util.SetPrivateField(tskState, "nextTriggerIdValue", row.next_trigger_id)
            If Not row.Isnext_trigger_idNull Then tskState.NextTriggerId = row.next_trigger_id
            tskState.TaskExecutionState = CType(row.task_status, ITaskState.ExecutionState)
            'AhSaElDin 20111223: support for CurrentTriggerId & PreviousTriggerId.
            If Not row.Iscurrent_trigger_idNull Then tskState.CurrentTriggerId = row.current_trigger_id
            If Not row.Isprevious_trigger_idNull Then tskState.PreviousTriggerId = row.previous_trigger_id
            Try
                'AhSaElDin 20111231: property is writable for now.
                'Util.SetPrivateField(tskState, "currentExecutingActionIdValue", row.current_executed_action)
                tskState.CurrentExecutingActionId = row.current_executed_action
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                'justLog enabled
                Alerter.REP("", ex, False, True)
            End Try
            tasksStateRows = Nothing
            Return tskState
        Next
        tskState.Dispose()
        tskState = Nothing
        tasksStateRows = Nothing
        Return Nothing
    End Function

    ''' <summary>
    ''' Gets the task triggers using task Id.
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function GetTaskTriggers(ByVal ParamArray args() As String) As Dictionary(Of String, Trigger)
        Dim triggersRows As System.Data.EnumerableRowCollection(Of TectonicDataSet.triggersRow) = Nothing
        triggersRows = From row As TectonicDataSet.triggersRow In TectonicDataSetObj.triggers
                       Where row.task_id = args(0)
        Dim triggersDic As New Dictionary(Of String, Trigger)
        Dim tempTrigger As Trigger
        For Each row In triggersRows
            'AhSaElDin 20111229: Invoke CreateTriggerByType of BIT 
            'rather(Commands.CreateTriggerByType) Because the later 
            'cannot return QuartzTrigger Type
            'tempTrigger = CType(Commands.CreateTriggerByType(CType(row.type, ITrigger.TriggerTypeEnum)), Trigger)
            tempTrigger = CType(Balora.Util.InvokeMethodWithinAssembly("helpers.dll", "Helpers.QuartzHelper.QuartzUtils", "CreateTriggerByType", {CType(row.type, ITrigger.TriggerTypeEnum)}), Trigger)
            tempTrigger.SetParentTaskId(row.task_id)
            tempTrigger.SetTriggerId(row.trigger_id)

            tempTrigger.MisFireCommand = CType(row.misfireCommand, ITrigger.MisFireCommandsEnum)
            tempTrigger.Description = row.description

            'AhSaElDin 20111222: Supporting the new properties of trigger class (InnerTrigger & Task).
            If Not row.Isinner_triggerNull() Then tempTrigger.InnerTrigger = Util.ByteArrayToObject(row.inner_trigger)
            If Not row.IstaskNull() Then tempTrigger.Task = CType(Util.ByteArrayToObject(row.task), Task)
            'AhSaElDin 20120115: supporting a new property of trigger class (calendar class)
            If Not row.IscalendarNull() Then tempTrigger.Calendar = CType(Util.ByteArrayToObject(row.calendar), Object)

            triggersDic.Add(tempTrigger.TriggerId, tempTrigger)
        Next
        args = Nothing
        triggersRows = Nothing
        Return triggersDic
    End Function

    ''' <summary>
    ''' Gets the task actions using task Id.
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function GetTaskActions(taskId As String) As Dictionary(Of String, Action)
        Dim actionsRows As System.Data.EnumerableRowCollection(Of TectonicDataSet.actionsRow) = GetTaskActionsRow(taskId)
        Dim actionsDic As New Dictionary(Of String, Action)
        For Each row In actionsRows
            actionsDic.Add(row.action_id, ActionRowToActionObject(row))
        Next
        actionsRows = Nothing
        Return actionsDic
    End Function

    Private Function GetTaskActionsRow(taskId As String) As System.Data.EnumerableRowCollection(Of TectonicDataSet.actionsRow)
        Dim actionsRows As System.Data.EnumerableRowCollection(Of TectonicDataSet.actionsRow) = Nothing
        actionsRows = From row As TectonicDataSet.actionsRow In TectonicDataSetObj.actions
                      Where row.task_id = taskId
                      Order By row.execution_precedence Ascending
        Return actionsRows
    End Function

    ''' <summary>
    ''' Gets all actions in Tectonic.
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function GetAllActions(ByVal ParamArray args() As String) As Dictionary(Of String, Action)
        Dim actionsRows As System.Data.EnumerableRowCollection(Of TectonicDataSet.actionsRow) = Nothing
        actionsRows = From row As TectonicDataSet.actionsRow In TectonicDataSetObj.actions
        Dim actionsDic As New Dictionary(Of String, Action)
        For Each row In actionsRows
            actionsDic.Add(row.action_id, ActionRowToActionObject(row))
        Next
        args = Nothing
        actionsRows = Nothing
        Return actionsDic
    End Function

    ''' <summary>
    ''' Gets all triggers in Tectonic.
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function GetAllTriggers(ByVal ParamArray args() As String) As Dictionary(Of String, Trigger)
        Dim triggersRows As System.Data.EnumerableRowCollection(Of TectonicDataSet.triggersRow) = Nothing
        triggersRows = From row As TectonicDataSet.triggersRow In TectonicDataSetObj.triggers
        Dim triggersDic As New Dictionary(Of String, Trigger)
        Dim tempTrigger As Trigger
        For Each row In triggersRows
            'AhSaElDin 20111229: Invoke CreateTriggerByType of BIT 
            'rather(Commands.CreateTriggerByType) Because the later 
            'cannot return QuartzTrigger Type
            'tempTrigger = CType(Commands.CreateTriggerByType(CType(row.type, ITrigger.TriggerTypeEnum)), Trigger)
            tempTrigger = CType(Balora.Util.InvokeMethodWithinAssembly("helpers.dll",
                                                                       "Helpers.QuartzHelper.QuartzUtils",
                                                                       "CreateTriggerByType",
                                                                       {CType(row.type, ITrigger.TriggerTypeEnum)}), 
                                                                       Trigger)
            tempTrigger.SetParentTaskId(row.task_id)
            tempTrigger.SetTriggerId(row.trigger_id)
            tempTrigger.Description = row.description
            tempTrigger.MisFireCommand = CType(row.misfireCommand, ITrigger.MisFireCommandsEnum)
            'AhSaElDin 20111222: tectonizing the new properties of trigger class (InnerTrigger & Task).
            If Not row.Isinner_triggerNull() Then tempTrigger.InnerTrigger = Util.ByteArrayToObject(row.inner_trigger)
            If Not row.IstaskNull() Then tempTrigger.Task = CType(Util.ByteArrayToObject(row.task), Task)
            'AhSaElDin 20120115: tectonizing Calendar property.
            If Not row.IscalendarNull Then tempTrigger.Calendar = CType(Util.ByteArrayToObject(row.calendar), Object)
            triggersDic.Add(tempTrigger.TriggerId, tempTrigger)
        Next
        args = Nothing
        triggersRows = Nothing
        Return triggersDic
    End Function

    ''' <summary>
    ''' Gets all tasks that have the same group Id.
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function GetGroupTasks(ByVal ParamArray args() As String) As Object
        Dim tasksRows As System.Data.EnumerableRowCollection(Of TectonicDataSet.tasksRow) = Nothing
        tasksRows = From row As TectonicDataSet.tasksRow In TectonicDataSetObj.tasks
                    Where row.group_id = args(0)
        Dim tasksDic As New Dictionary(Of String, Task)
        Dim tempTask As Task
        For Each row In tasksRows
            tempTask = TaskRowToTaskObject(row)
            tasksDic.Add(tempTask.TaskId, tempTask)
        Next
        args = Nothing
        tasksRows = Nothing
        Return tasksDic
    End Function

    ''' <summary>
    ''' Gets an action using its id.
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function GetActionById(ByVal ParamArray args() As String) As Action
        Dim actRow As TectonicDataSet.actionsRow = TectonicDataSetObj.actions.FindByaction_id(args(1))
        Dim act = ActionRowToActionObject(actRow)
        actRow = Nothing
        args = Nothing
        Return act
    End Function

    ''' <summary>
    ''' Gets a trigger using its id.
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function GetTriggerById(ByVal ParamArray args() As String) As ITrigger
        Dim triggerRow As TectonicDataSet.triggersRow = TectonicDataSetObj.triggers.FindBytrigger_id(args(1))
        Dim trg = TriggerRowToTriggerObject(triggerRow)
        triggerRow = Nothing
        args = Nothing
        Return trg
    End Function

    ''' <summary>
    ''' Gets a task using its id.
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function GetTaskbyId(ByVal ParamArray args() As String) As Task
        Dim taskRow As TectonicDataSet.tasksRow = TectonicDataSetObj.tasks.FindByid(args(1))
        Dim tsk = TaskRowToTaskObject(taskRow)
        taskRow = Nothing
        args = Nothing
        Return tsk
    End Function

    ''' <summary>
    ''' Gets a group using its id.
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function GetGroupById(ByVal ParamArray args() As String) As Group
        Dim groupRow As TectonicDataSet.groupsRow = TectonicDataSetObj.groups.FindByid(args(1))
        Dim grp = GroupRowToGroupObject(groupRow)
        args = Nothing
        groupRow = Nothing
        Return grp
    End Function

    ''' <summary>
    ''' Gets all tasks registered in Tectonic.
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function GetAllTasks(ByVal ParamArray args() As String) As Dictionary(Of String, Task)
        Dim tasksRows As System.Data.EnumerableRowCollection(Of TectonicDataSet.tasksRow) = Nothing
        tasksRows = From row As TectonicDataSet.tasksRow In TectonicDataSetObj.tasks
        Dim dict As New Dictionary(Of String, Task)
        For Each row In tasksRows
            dict.Add(row.id, TaskRowToTaskObject(row))
        Next
        tasksRows = Nothing
        args = Nothing
        Return dict
    End Function

    ''' <summary>
    ''' Gets all tasks states organized in a dictionary Of ("taskId", TaskState)
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function GetAllTasksStates(ByVal ParamArray args() As String) As Dictionary(Of String, TaskState)
        Dim tasksStatesRows As System.Data.EnumerableRowCollection(Of TectonicDataSet.tasks_statesRow) = Nothing
        tasksStatesRows = From row As TectonicDataSet.tasks_statesRow In TectonicDataSetObj.tasks_states
        Dim dict As New Dictionary(Of String, TaskState)
        For Each row In tasksStatesRows
            dict.Add(row.task_id, GetTaskState(row.task_id))
        Next
        tasksStatesRows = Nothing
        args = Nothing
        Return dict
    End Function

    Private Function GetUnfullfilledTasks(ByVal ParamArray args() As String) As Dictionary(Of String, Task)
        Dim tasksDict As New Dictionary(Of String, Task)
        Dim statesDict As Dictionary(Of String, TaskState) = GetAllTasksStates()
        For Each state In statesDict
            If state.Value.TaskExecutionState = ITaskState.ExecutionState.Unfulfilled Then
                tasksDict.Add(state.Key, Commands.GetTaskById(state.Key))
            End If
        Next
        args = Nothing
        statesDict = Nothing
        Return tasksDict
    End Function

    Private Function GetTimeOutTriggers(ByVal ParamArray args() As String) As Dictionary(Of String, ITrigger)
        Dim result As Integer
        Dim trgDict As New Dictionary(Of String, ITrigger)
        Dim allTrgs = GetAllTriggers()
        For Each trg In allTrgs
            If trg.Value.EndTimeUTC.HasValue Then result = DateTime.Compare(trg.Value.EndTimeUTC.Value, DateTime.UtcNow)
            If result < 0 Then
                trgDict.Add(trg.Value.TriggerId, trg.Value)
            End If
        Next
        args = Nothing
        allTrgs = Nothing
        Return trgDict
    End Function

    Private Function GetTaskTriggersCount(taskId As String) As Integer
        Dim queryResult = From tskTrgs As TectonicDataSet.triggersRow
                          In TectonicDataSetObj.triggers
                          Where tskTrgs.task_id = taskId
                          Select tskTrgs
        Dim count = queryResult.Count
        queryResult = Nothing
        Return count
    End Function

    Private Function GetTaskActionsCount(taskId As String) As Integer
        Dim queryResult = From tskActions As TectonicDataSet.actionsRow
                          In TectonicDataSetObj.actions
                          Where tskActions.task_id = taskId
                          Select tskActions
        Dim count = queryResult.Count
        queryResult = Nothing
        Return count
    End Function

#Region "Row Converters"
    ''' <summary>
    ''' Converts actions row to Action object.
    ''' </summary>
    ''' <param name="actRow">The act row.</param><returns></returns>
    Private Function ActionRowToActionObject(actRow As TectonicDataSet.actionsRow) As Action
        If Not actRow Is Nothing Then
            Dim act As Action = Bridge.Util.CreateActionByType(CType(actRow.action_type, IAction.ActionTypeEnum))
            act.Shape = CType(Util.BlobToObject(actRow.shape), Shape)
            act.Result = CType(actRow.result, ITaskState.ExecutionState)
            act.Enabled = CBool(actRow.is_inactive)
            act.Comments = actRow.description
            act.SetActionId(actRow.action_id)
            act.SetParentTaskId(actRow.task_id)
            If Not IsNothing(actRow.ExceptionsManager) Then act.ExceptionsManager = CType(Util.BlobToObject(actRow.ExceptionsManager), ActionExceptionsManager)
            If Not IsNothing(actRow.input) Then act.Inputs = CType(Util.BlobToObject(actRow.input), IAction.IInputs)
            If Not IsNothing(actRow.output) Then act.Outputs = CType(Util.BlobToObject(actRow.output), IAction.IOutputs)
            actRow = Nothing
            Return act
        Else
            actRow = Nothing
            Alerter.REP(My.Resources.ActNotFound, Nothing, False)
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Converts Triggers row to Trigger object.
    ''' </summary>
    ''' <param name="triggerRow">The trigger row.</param><returns></returns>
    Private Function TriggerRowToTriggerObject(triggerRow As TectonicDataSet.triggersRow) As ITrigger
        If Not triggerRow Is Nothing Then
            'AhSaElDin 20111229: Invoke CreateTriggerByType of BIT 
            'rather(Commands.CreateTriggerByType) Because the later 
            'cannot return QuartzTrigger Type
            'Dim trg As ITrigger = Commands.CreateTriggerByType(CType(triggerRow.type, ITrigger.TriggerTypeEnum))
            Dim returnObj As Object = Balora.Util.InvokeMethodWithinAssembly("helpers.dll", "Helpers.QuartzHelper.QuartzUtils", "CreateTriggerByType", {CType(triggerRow.type, ITrigger.TriggerTypeEnum)})
            Dim trg As ITrigger = CType(returnObj, ITrigger)
            trg.SetParentTaskId(triggerRow.task_id)
            trg.SetTriggerId(triggerRow.trigger_id)
            trg.Description = triggerRow.description
            trg.MisFireCommand = CType(triggerRow.misfireCommand, ITrigger.MisFireCommandsEnum)
            'AhSaElDin 20111222: Supporting the new properties of trigger class (InnerTrigger & Task).
            If Not triggerRow.Isinner_triggerNull() Then trg.InnerTrigger = Util.ByteArrayToObject(triggerRow.inner_trigger)
            If Not triggerRow.IstaskNull() Then trg.Task = CType(Util.ByteArrayToObject(triggerRow.task), Task)
            If Not triggerRow.IscalendarNull Then trg.Calendar = CType(Util.ByteArrayToObject(triggerRow.calendar), Object)
            triggerRow = Nothing
            Return trg
        Else
            triggerRow = Nothing
            Alerter.REP(My.Resources.TrgNotFound, Nothing, False)
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Converts tasks row to task object.
    ''' </summary>
    ''' <param name="taskRow">The task row.</param><returns></returns>
    Private Function TaskRowToTaskObject(taskRow As TectonicDataSet.tasksRow) As Task
        If Not taskRow Is Nothing Then
            Dim tsk As Task
            tsk = New Task(taskRow.id, False)
            tsk.ParentGroupId = taskRow.group_id
            tsk.WaitBetweenActions = CBool(taskRow.is_wait_between_actions)
            tsk.BreakOnFailed = CBool(taskRow.break_on_fail)
            tsk.Properties = GetTaskProperties(taskRow.id)
            tsk.SetTaskState(GetTaskState(taskRow.id))
            'No need for assign those:-
            'tempTask.TaskVariables.Variables
            'tempTask.MyActions
            'tempTask.MyTriggers
            'tempTask.History
            taskRow = Nothing
            Return tsk
        Else
            taskRow = Nothing
            Alerter.REP(My.Resources.TskNotFound, Nothing, False)
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Convert Group row to Group object.
    ''' </summary>
    ''' <param name="groupRow">The group row.</param><returns></returns>
    Private Function GroupRowToGroupObject(groupRow As TectonicDataSet.groupsRow) As Group
        If Not groupRow Is Nothing Then
            Dim tempGroup As Group
            tempGroup = New Group(groupRow.id)
            tempGroup.Name = groupRow.name
            'No need for tempGroup.TasksCollection 
            groupRow = Nothing
            Return tempGroup
        Else
            groupRow = Nothing
            Alerter.REP(My.Resources.GrpNotFound, Nothing, False)
            Return Nothing
        End If
    End Function

#End Region

#Region "SQL"

    ''' <summary>
    ''' Executes a text SQL command.
    ''' </summary>
    ''' <param name="commandText">The command text.</param>
    ''' <param name="queryParameters">The query parameters.</param><returns></returns>
    Private Function ExecuteTextSQLCommand(ByVal commandText As String, Optional queryParameters As Dictionary(Of String, Object) = Nothing) As DataTable
        'Ex: Calling using a sql command directly:-''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Dim sqlCommandText = "SELECT * FROM tasks"
        'Dim resultDataTable As DataTable = TectonicCruding.Query(sqlCommandText)
        ''How to Retrieve Datatable
        'For Each row As DataRow In resultDataTable.Rows
        '    Dim values(row.ItemArray.Count - 1) As Object
        '    values = row.ItemArray
        '    For Each obj In values
        '        Balora.Util.SendToOutputWindow(obj.ToString)
        '    Next
        'Next
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'You have 2 options now to execute the query:-
        '1.You can simply call any adapter in TectonicDataSetTableAdapters to fill the desired table,
        'however you will need to pass the table name in order to detect the corresponding tableadapter
        '+and you will need to specify a damn select case statement to handle all tables ya man.

        'Dim queryResultTable As Object
        'Dim selectCommand As New SQLite.SQLiteCommand()
        'selectCommand.CommandText = commandText
        'tasksTableAdapter.Adapter.SelectCommand = selectCommand
        'queryResultTable = New TectonicDataSet.tasksDataTable
        'tasksTableAdapter.Fill(queryResultTable)
        'Return queryResultTable


        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'or
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        'Simply execute query using Datler directly and this will enable you to query more than one table and 
        'with few lines:-
        '' '' '' ''Dim dbConn As New SQLite.SQLiteConnection
        '' '' '' ''Dim CacaduDatabase As New Balora.Datler.SQLiteDatabase
        '' '' '' ''Dim ConnectionFactory As New Balora.Datler.ConnectionFactory()
        '' '' '' ''Try
        '' '' '' ''    ConnectionFactory.DatabaseProviderName = Datler.Dbprovider.System_Data_SQLite
        '' '' '' ''    dbConn = CType(ConnectionFactory.CreateConnection(), SQLite.SQLiteConnection)
        '' '' '' ''    dbConn.ConnectionString = _tasksTableAdapter.Connection.ConnectionString
        '' '' '' ''Catch ex As Exception
        '' '' '' ''    Balora.Alerter.RaiseErrorUp("", ex, True)
        '' '' '' ''    dbConn = Nothing
        '' '' '' ''    Return Nothing
        '' '' '' ''End Try
        '' '' '' ''CacaduDatabase.DatabaseConnection = dbConn
        ' '' '' '' ''Execute query and retrieve Datatable
        '' '' '' ''Dim queryResult As DataTable = CacaduDatabase.Query(commandText, queryParameters)
        '' '' '' ''CacaduDatabase.DatabaseConnection.Close()
        '' '' '' ''Return queryResult
        Return Nothing
    End Function
#End Region
#End Region

#Region "Update & Insert"
    Private Function InsertNewTrigger(ByVal cacadoreObj As ICrudable) As Boolean
        Dim res As Boolean = True
        Dim tempTrigger As Trigger = CType(cacadoreObj, Trigger)
        If IsRegisteredTrigger(tempTrigger.TriggerId) Then
            Balora.Alerter.REP(My.Resources.TrgIsAlrdyAttached, Nothing, False)
            res = True
        End If
        If res Then
            Dim newTriggerRow As TectonicDataSet.triggersRow
            newTriggerRow = TectonicDataSetObj.triggers.NewtriggersRow
            With tempTrigger
                newTriggerRow.task_id = .ParentTaskId
                newTriggerRow.trigger_id = .TriggerId
                newTriggerRow.type = .TriggerType
                newTriggerRow.description = .Description
                newTriggerRow.misfireCommand = .MisFireCommand
                'AhSalEdin 20111222: Making InnerTrigger & Task properties persistable.
                If Not IsNothing(.InnerTrigger) Then newTriggerRow.inner_trigger = Util.ObjectToByteArray(.InnerTrigger)
                If Not IsNothing(.Task) Then newTriggerRow.task = Util.ObjectToByteArray(.Task)
                'AhSaElDin 20120115: tectonizing Calendar property.
                If Not IsNothing(.Calendar) Then newTriggerRow.calendar = Util.ObjectToByteArray(.Calendar)
            End With
            Try
                TectonicDataSetObj.triggers.Rows.Add(newTriggerRow)
                _triggersTableAdapter.Update(TectonicDataSetObj.triggers)
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Balora.Alerter.REP("", ex, True)
                res = False
            End Try
            newTriggerRow = Nothing
        End If
        Return res
    End Function

    ''' <summary>
    ''' Inserts a new history entry.
    ''' </summary>
    ''' <param name="cacadoreObj">The cacadore obj.</param><returns></returns>
    Private Function InsertNewHistoryEntry(ByVal cacadoreObj As ICrudable) As Boolean
        Dim res As Boolean = True
        Dim tempHistoryEntry As HistoryEntry = CType(cacadoreObj, HistoryEntry)
        Dim newHistoryEntryRow As TectonicDataSet.historyRow
        newHistoryEntryRow = TectonicDataSetObj.history.NewhistoryRow
        With tempHistoryEntry
            newHistoryEntryRow.task_id = .TaskId
            'Save DateTime as a long value to avoid miss-parse of it
            newHistoryEntryRow.date_entry_date = CStr(.EntryDateTime) 'CStr(.EntryDateTime.ToBinary)
            newHistoryEntryRow.entry_type = .EntryType
            If Not .Description = vbNullString Then newHistoryEntryRow.message = .Description
            newHistoryEntryRow.status = CStr(.Status)
            newHistoryEntryRow.no = GetNewHistoryTableIncrementValue()
        End With
        Try
            TectonicDataSetObj.history.Rows.Add(newHistoryEntryRow)
            _historyTableAdapter.Update(TectonicDataSetObj.history)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Balora.Alerter.REP("", ex, True)
            res = False
        End Try
        newHistoryEntryRow = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Updates the tasks table using Task instance.
    ''' </summary>
    ''' <param name="task">The task.</param><returns></returns>
    ''' <remarks>'This function works for both update/insert</remarks>
    Private Function UpdateTask([task] As Task) As Boolean
        Dim res As Boolean = True
        Dim insert As Boolean = False
        Dim isRegisterd As Boolean = IsRegisteredTask([task].TaskId)
        If Not IsRegisteredTask([task].TaskId) Then
            insert = True
        Else
            insert = False
        End If

        If insert AndAlso isRegisterd Then
            Balora.Alerter.REP(My.Resources.TskIsAlrdyDefined, Nothing, False)
        End If
        Dim taskRow As TectonicDataSet.tasksRow
        If insert Then
            taskRow = TectonicDataSetObj.tasks.NewtasksRow
        Else
            taskRow = TectonicDataSetObj.tasks.FindByid([task].TaskId)
            If taskRow Is Nothing Then
                Alerter.REP(My.Resources.NotTask, Nothing, False)
                Return False
            End If
        End If
        With [task]
            If insert Then
                taskRow.id = [task].TaskId
            End If
            taskRow.group_id = .ParentGroupId
            'for unknown reason it equals -1 if set directly
            If .WaitBetweenActions Then
                taskRow.is_wait_between_actions = 1
            Else
                taskRow.is_wait_between_actions = 0
            End If

            If .BreakOnFailed Then
                taskRow.break_on_fail = 1
            Else
                taskRow.break_on_fail = 0
            End If

            If .Properties.LastTimeRun.HasValue Then
                taskRow.date_last_time_run = CStr(.Properties.LastTimeRun.Value)
            Else
                taskRow.date_last_time_run = " "
            End If

            taskRow.last_time_run_result = CStr(.State.TaskExecutionState)

            If .Properties.NextTimeToBeRun.HasValue Then
                'Dim nextTimeToBeRunFormats = .Properties.NextTimeToBeRun.Value.GetDateTimeFormats("F"c)
                'If nextTimeToBeRunFormats.Count >= 2 Then
                '    taskRow.date_nexttime_run = nextTimeToBeRunFormats(1)
                'Else
                '    taskRow.date_nexttime_run = nextTimeToBeRunFormats(0)
                'End If
                taskRow.date_nexttime_run = CStr(.Properties.NextTimeToBeRun.Value)
            Else
                taskRow.date_nexttime_run = " "
            End If
            taskRow.name = .Properties.TaskName
            'For both, they will be blank for now until I finish the logic of triggers
            'Tasks>=Triggers_When ... TaskProperties.Triggers_When
            'Tasks_States>=Next_Trigger_Id ... TaskState.NextTriggerId
            If .Properties.TriggersWhen = "" Then
                taskRow.triggers_when = " "
            Else
                taskRow.triggers_when = .Properties.TriggersWhen
            End If
            'for unknown reason it equals -1 if set directly
            If .Properties.IsEnabled Then
                taskRow.is_enables = 1
            Else
                taskRow.is_enables = 0
            End If
        End With
        If insert Then
            TectonicDataSetObj.tasks.Rows.Add(taskRow)
        End If
        Try
            'Insert Task
            _tasksTableAdapter.Update(TectonicDataSetObj.tasks)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Balora.Alerter.REP(My.Resources.NwTskAddedErr, ex, True)
            res = False
        End Try
        If res Then
            'msg box solve here man
            If insert Then
                UpdateTaskState([task].State, True)
            Else
                UpdateTaskState([task].State)
            End If
        End If
        taskRow = Nothing

        Return res
    End Function

    ''' <summary>
    ''' Updates/Insert the tasks states table.
    ''' </summary>
    ''' <param name="taskState">State of the task.</param>
    ''' <param name="insert">if set to <c>true</c> [insert].</param><returns></returns>
    ''' <remarks>
    ''' This functions needs "insert" to be true in order to do insert stuff
    ''' in the future i may make it like UpdateTask, so it won't need insert
    ''' boolean parameter to be set and it will detect in the row in insert
    ''' or update. plz do it. search for instance of "Optional insert As Boolean = False"
    ''' in this document.
    ''' </remarks>
    Private Function UpdateTaskState(ByVal taskState As TaskState, Optional insert As Boolean = False) As Boolean
        Dim res As Boolean = True
        If insert Then
            If Not IsRegisteredTask(taskState.taskId) Then
                Balora.Alerter.REP(My.Resources.NotTask, Nothing, False)
                res = False
            End If
        End If
        Dim taskStateRow As TectonicDataSet.tasks_statesRow = Nothing
        If res Then
            If insert Then
                taskStateRow = TectonicDataSetObj.tasks_states.Newtasks_statesRow
            Else
                taskStateRow = TectonicDataSetObj.tasks_states.FindBytask_id(taskState.taskId)
                If taskStateRow Is Nothing Then
                    Alerter.REP(My.Resources.NotTask, Nothing, False)
                    res = False
                End If
            End If
        End If

        If res Then
            With taskState
                If insert Then
                    taskStateRow.task_id = .taskId
                End If
                If .CurrentExecutingActionId = Nothing Then
                    taskStateRow.current_executed_action = "n/a"
                Else
                    taskStateRow.current_executed_action = .CurrentExecutingActionId
                End If
                'For both, they will be blank for now until I finish the logic of triggers
                'Tasks>=Triggers_When ... TaskProperties.Triggers_When
                'Tasks_States>=Next_Trigger_Id ... TaskState.NextTriggerId
                taskStateRow.next_trigger_id = .NextTriggerId
                taskStateRow.task_status = .TaskExecutionState
                'AhSaElDin 20111223: support for CurrentTriggerId & PreviousTriggerId.
                taskStateRow.current_trigger_id = .CurrentTriggerId
                taskStateRow.previous_trigger_id = .PreviousTriggerId
            End With
            Try
                If insert Then
                    TectonicDataSetObj.tasks_states.Rows.Add(taskStateRow)
                End If
                _tasksStatusTableAdapter.Update(TectonicDataSetObj.tasks_states)
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Balora.Alerter.REP(My.Resources.AdingTskStErr, ex, False)
                res = False
            End Try
        End If
        taskStateRow = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Updates the settings.
    ''' </summary>
    Private Function UpdateSettings(ByVal settingObject As Cacadore.Settings, args As String()) As Boolean
        Dim res As Boolean = True
        Dim setRow As TectonicDataSet.settingsRow
        setRow = CType(TectonicDataSetObj.settings.Rows(0), TectonicDataSet.settingsRow)
        Select Case args(0)
            Case CStr(Cacadore.ISettings.SettingsEnum.CacadoreEnabled)
                If CBool(args(1)) Then
                    setRow.is_cacadore_enabled = 1
                Else
                    setRow.is_cacadore_enabled = 0
                End If
            Case CStr(Cacadore.ISettings.SettingsEnum.HistoryEnabled)
                If CBool(args(1)) Then
                    setRow.is_history_enabled = 1
                Else
                    setRow.is_history_enabled = 0
                End If
            Case CStr(Cacadore.ISettings.SettingsEnum.TriggersEnabled)
                If CBool(args(1)) Then
                    setRow.is_triggers_enabled = 1
                Else
                    setRow.is_triggers_enabled = 0
                End If
        End Select
        Try
            _settingsTableAdapter.Update(TectonicDataSetObj.settings)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Balora.Alerter.REP(My.Resources.ErrUpSettings, ex, False)
            res = False
        End Try
        setRow = Nothing
        args = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Updates/Insert the globals table
    ''' </summary>
    ''' <param name="globalEntry">The global entry.</param>
    ''' <param name="insert">if set to <c>true</c> [insert].</param><returns></returns>
    ''' <remarks>
    '''  This functions needs "insert" to be true in order to do insert stuff
    '''  in the future i may make it like UpdateTask, so it won't need insert
    '''  boolean parameter to be set and it will detect in the row in insert
    '''  or update. plz do it. search for instance of "Optional insert As Boolean = False"
    '''  in this document.
    ''' </remarks>
    Private Function UpdateGlobals(ByVal globalEntry As Cacadore.GlobalEntry, Optional insert As Boolean = False) As Boolean
        Dim res As Boolean = True
        Dim globalsRow As TectonicDataSet.globalsRow
        Dim isRegisterd = IsRegisteredVariable(globalEntry.Name, True)
        If insert Then
            If isRegisterd Then
                Balora.Alerter.REP("'" & globalEntry.Name & "' " & My.Resources.GrpIsAlrdyDefined, Nothing, True)
                res = False
            End If
            globalsRow = TectonicDataSetObj.globals.NewglobalsRow
        Else
            globalsRow = TectonicDataSetObj.globals.FindByname(globalEntry.Name)
            If globalsRow Is Nothing Then
                Alerter.REP(My.Resources.GlobNotDefined, Nothing, False)
                res = False
            End If
        End If
        If res Then
            If insert Then
                globalsRow.name = globalEntry.Name
            End If
            globalsRow.value = globalEntry.Value
            globalsRow.type = globalEntry.type
            If insert Then
                TectonicDataSetObj.globals.Rows.Add(globalsRow)
            End If
            Try
                _globalsTableAdapter.Update(TectonicDataSetObj.globals)
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Balora.Alerter.REP("", ex, True)
                res = False
            End Try
        End If
        globalsRow = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Updates the groups table.
    ''' </summary>
    ''' <param name="group">The group.</param><returns></returns>
    Private Function UpdateGroup(ByVal [group] As Group) As Boolean
        Dim res As Boolean = True
        Dim insert As Boolean = False
        Dim isRegistered As Boolean = IsRegisteredGroup([group].GroupId)
        If Not isRegistered Then
            insert = True
        End If
        If insert Then
            If isRegistered Then
                Balora.Alerter.REP(My.Resources.GrpIsAlrdyDefined, Nothing, False)
                Return False
                res = False
            End If
        Else
            If Not isRegistered Then
                Balora.Alerter.REP(My.Resources.GrpIsNotDefined, Nothing, False)
                res = False
            End If
        End If
        If res Then
            Dim newGroupsRow As TectonicDataSet.groupsRow
            If insert Then
                newGroupsRow = TectonicDataSetObj.groups.NewgroupsRow
            Else
                newGroupsRow = TectonicDataSetObj.groups.FindByid([group].GroupId)
            End If
            With [group]
                If insert Then
                    newGroupsRow.id = .GroupId
                End If
                newGroupsRow.name = .Name
            End With
            Try
                If insert Then
                    TectonicDataSetObj.groups.Rows.Add(newGroupsRow)
                End If
                _groupsTableAdapter.Update(TectonicDataSetObj.groups)
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Balora.Alerter.REP("", ex, True)
                res = False
            End Try
            newGroupsRow = Nothing
        End If
        Return res
    End Function

    ''' <summary>
    ''' Updates the task variable.
    ''' </summary>
    ''' <param name="tskVariable">The TSK variable.</param>
    ''' <param name="insert">if set to <c>true</c> [insert].</param><returns></returns>
    ''' <remarks>
    ''' This functions needs "insert" to be true in order to do insert stuff
    ''' in the future i may make it like UpdateTask, so it won't need insert
    ''' boolean parameter to be set and it will detect in the row in insert
    ''' or update. plz do it. search for instance of "Optional insert As Boolean = False"
    ''' in this document.
    ''' </remarks>
    Private Function UpdateTaskVariable(ByVal tskVariable As Cacadore.Variable, Optional insert As Boolean = False) As Boolean
        Dim res As Boolean = True
        TectonicDataSetObj.task_variables.Columns(2).AllowDBNull = True
        If insert Then
            Dim taskVariableRow As TectonicDataSet.task_variablesRow
            'If a variable with this name @ kvp.Key isn't already registered or 
            Dim isRegisterd = IsRegisteredVariable(tskVariable.Name)
            'If a variable with this name @ kvp.Key is registered for another task. AND
            Dim isAttached = IsVariableAttachedToThisTask(tskVariable.Name, tskVariable.TaskId)
            'The task is already registered to avoid adding a variable for a task that doesn't exists
            Dim taskIsRegistered = IsRegisteredTask(tskVariable.TaskId)
            If Not isRegisterd Or Not isAttached And taskIsRegistered Then
                taskVariableRow = TectonicDataSetObj.task_variables.Newtask_variablesRow
                taskVariableRow.variable_name = tskVariable.Name
                If tskVariable.Value = vbNullString Then
                    tskVariable.Value = " "
                End If
                taskVariableRow.variable_value = tskVariable.Value
                taskVariableRow.task_id = tskVariable.TaskId
                If insert Then
                    taskVariableRow.no = GetNewTaskVariablesTableIncrementValue()
                    TectonicDataSetObj.task_variables.Rows.Add(taskVariableRow)
                End If
            ElseIf isAttached Then
                If insert Then
                    Balora.Alerter.REP(String.Format("'{0}' {1}", tskVariable.Name, My.Resources.VrIsAddedToTsk), Nothing, True)
                    res = False
                End If
            End If
            If res Then
                Try
                    _tasksVariablesTableAdapter.Update(TectonicDataSetObj.task_variables)
                Catch ex As Exception
                    If Debugger.IsAttached Then Debugger.Break()
                    Balora.Alerter.REP("", ex, True)
                    res = False
                End Try
            End If
            taskVariableRow = Nothing
        Else
            Dim result = From row In TectonicDataSetObj.task_variables
                         Where row.task_id = tskVariable.TaskId And row.variable_name = tskVariable.Name
            For Each row As TectonicDataSet.task_variablesRow In result
                row.variable_name = tskVariable.Name
                row.variable_value = tskVariable.Value
            Next
            Try
                _tasksVariablesTableAdapter.Update(TectonicDataSetObj.task_variables)
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Balora.Alerter.REP("", ex, True)
                res = False
            End Try
            result = Nothing
        End If
        Return res
    End Function

    ''' <summary>
    ''' Updates the actions table
    ''' </summary>
    ''' <param name="action">The action.</param>
    ''' <param name="insert">if set to <c>true</c> [insert].</param><returns></returns>
    ''' <remarks>
    ''' This functions needs "insert" to be true in order to do insert stuff
    ''' in the future i may make it like UpdateTask, so it won't need insert
    ''' boolean parameter to be set and it will detect in the row in insert
    ''' or update. plz do it. search for instance of "Optional insert As Boolean = False"
    ''' in this document.
    ''' </remarks>
    Private Function UpdateAction([action] As Action, Optional insert As Boolean = False) As Boolean
        Dim res As Boolean = True
            If insert Then
                If IsRegisteredAction([action].ActionId) Then
                    Balora.Alerter.REP(My.Resources.ActIsAlrdyAttached, Nothing, False)
                    res = False
                End If
            Else
                If Not IsRegisteredAction([action].ActionId) Then
                    Balora.Alerter.REP(My.Resources.ActIsNotSet, Nothing, False)
                    res = False
                End If
            End If
            If res Then
                Dim actionRow As TectonicDataSet.actionsRow
                If insert Then
                    actionRow = TectonicDataSetObj.actions.NewactionsRow
                Else
                    actionRow = TectonicDataSetObj.actions.FindByaction_id([action].ActionId)
                End If
                With [action]
                    If insert Then
                        actionRow.task_id = .TaskId
                        actionRow.action_id = .ActionId
                        Dim actionsCount As Integer = GetTaskActionsCount([action].TaskId)
                        actionRow.execution_precedence = actionsCount + 1
                    End If
                    actionRow.action_type = .ActionType
                    If .Enabled Then
                        actionRow.is_inactive = 1
                    Else
                        actionRow.is_inactive = 0
                    End If
                    actionRow.result = .Result
                    If Not IsNothing(.ExceptionsManager) Then actionRow.ExceptionsManager = Util.ObjectToBlob(.ExceptionsManager)
                    If Not IsNothing(.Inputs) Then actionRow.input = Util.ObjectToBlob(.Inputs)
                    If Not IsNothing(.Outputs) Then actionRow.output = Util.ObjectToBlob(.Outputs)
                    If Not IsNothing(.Shape) Then actionRow.shape = Util.ObjectToBlob(.Shape)
                    actionRow.description = .Comments
                End With
                Try
                    If insert Then
                        TectonicDataSetObj.actions.Rows.Add(actionRow)
                    End If
                    _actionsTableAdapter.Update(TectonicDataSetObj.actions)
                Catch ex As Exception
                    If Debugger.IsAttached Then Debugger.Break()
                    If ex.Message.Contains("The database disk image is malformed") Then
                    'Balora.Hodhod.STOW("The database disk image is malformed")
                    End If
                    Balora.Alerter.REP("", ex, True)
                    res = False
                End Try
                actionRow = Nothing
            End If
            Return res
    End Function

    ''' <summary>
    ''' Inserts/updates a new trigger.
    ''' </summary>
    ''' <param name="trg">The TRG.</param>
    ''' <param name="insert">if set to <c>true</c> [insert].</param><returns></returns>
    Private Function UpdateTrigger(ByVal trg As ITrigger, Optional insert As Boolean = False) As Boolean
        Dim res As Boolean = True
        Dim tempTrigger As Trigger = CType(trg, Trigger)
        If insert Then
            If IsRegisteredTrigger(tempTrigger.TriggerId) Then
                Balora.Alerter.REP(My.Resources.TrgIsAlrdyAttached, Nothing, False)
                res = False
            End If
        Else
            If Not IsRegisteredTrigger(tempTrigger.TriggerId) Then
                Balora.Alerter.REP(My.Resources.TrgNotFound, Nothing, False)
                res = False
            End If
        End If
        If res Then
            Dim newTriggerRow As TectonicDataSet.triggersRow
            If insert Then
                newTriggerRow = TectonicDataSetObj.triggers.NewtriggersRow
            Else
                newTriggerRow = TectonicDataSetObj.triggers.FindBytrigger_id(trg.TriggerId)
            End If
            With tempTrigger
                If insert Then
                    newTriggerRow.task_id = .ParentTaskId
                    newTriggerRow.trigger_id = .TriggerId
                End If
                newTriggerRow.type = .TriggerType
                newTriggerRow.description = .Description
                newTriggerRow.misfireCommand = .MisFireCommand
                If IsNothing(.InnerTrigger) OrElse IsNothing(.Task) Then
                    Balora.Alerter.REP("Task is nothing, Can't update trigger", Nothing, False, True)
                End If
                'AhSalEdin 20111222: Making InnerTrigger & Task properties persistable.
                If Not IsNothing(.InnerTrigger) Then newTriggerRow.inner_trigger = Util.ObjectToByteArray(.InnerTrigger)
                If Not IsNothing(.Task) Then newTriggerRow.task = Util.ObjectToByteArray(.Task)
                'AhSaElDin 20120115: tectonizing Calendar property.
                If Not IsNothing(.Calendar) Then newTriggerRow.calendar = Util.ObjectToByteArray(.Calendar)
            End With
            Try
                If insert Then
                    TectonicDataSetObj.triggers.Rows.Add(newTriggerRow)
                End If
                _triggersTableAdapter.Update(TectonicDataSetObj.triggers)
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Balora.Alerter.REP("", ex, True)
                res = False
            End Try
            newTriggerRow = Nothing
        End If
        Return res
    End Function
#End Region

#Region "Delete"
    ''' <summary>
    ''' Sends a deleted task to history.
    ''' </summary>
    ''' <param name="tskName">Name of the TSK.</param>
    ''' <param name="tskId">The TSK id.</param><returns></returns>
    Private Function SendDeletedTaskToHistory(tskName As String, tskId As String) As Boolean
        Dim historyEntryMan As New HistoryEntry
        historyEntryMan.TaskId = tskId
        historyEntryMan.Status = ITaskState.ExecutionState.Deleted
        historyEntryMan.Description = "Task '" & tskName & "' deleted."
        historyEntryMan.EntryType = IHistoryEntry.HistoryEntryType.Task
        historyEntryMan.EntryDateTime = DateTime.UtcNow
        Dim res = InsertNewHistoryEntry(historyEntryMan)
        historyEntryMan.Dispose()
        historyEntryMan = Nothing
        Return res
    End Function

#Region "Delete Single Row"
    ''' <summary>
    ''' Deletes a task from tasks table.
    ''' </summary>
    ''' <param name="task">The task.</param><returns></returns>
    Private Function DeleteTask(ByVal [task] As Task) As Boolean
        Dim res As Boolean = True
        Using [task]
            Dim tasksRowToBeDeleted As TectonicDataSet.tasksRow
            tasksRowToBeDeleted = TectonicDataSetObj.tasks.FindByid([task].TaskId)
            Try
                'If you intend to delete a task, you will delete all referencing tables.
                'Delete all actions associated with this task from actions table
                'Delete task's variables in task variables
                'Delete from Task States
                'Delete all triggers that has the taskId
                DeleteTaskFromAllTables([task].TaskId)
                'Delete the task row itself
                tasksRowToBeDeleted.Delete()
                _tasksTableAdapter.Update(TectonicDataSetObj.tasks)
                'Write in history about deleted task.
                SendDeletedTaskToHistory([task].Properties.TaskName, [task].TaskId)
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Balora.Alerter.REP("", ex, True)
                res = False
            End Try
            tasksRowToBeDeleted = Nothing
        End Using
        Return res
    End Function

    ''' <summary>
    ''' Deletes a group from Groups table.
    ''' </summary>
    ''' <param name="group">The group.</param><returns></returns>
    Private Function DeleteGroup(ByVal [group] As Group) As Boolean
        Dim res As Boolean = True
        Using [group]
            Dim groupsRowToBeDeleted As TectonicDataSet.groupsRow
            groupsRowToBeDeleted = TectonicDataSetObj.groups.FindByid([group].GroupId)
            Try
                groupsRowToBeDeleted.Delete()
                _groupsTableAdapter.Update(TectonicDataSetObj.groups)
                MoveTasksToRootGroup([group].GroupId)
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Balora.Alerter.REP("", ex, True)
                res = False
            End Try
            groupsRowToBeDeleted = Nothing
        End Using
        Return res
    End Function

    ''' <summary>
    ''' After you delete a task from, deletes the task from all tables.
    ''' </summary>
    ''' <param name="taskid">The taskid.</param>
    Private Sub DeleteTaskFromAllTables(taskid As String)
        DeleteAllAssociatedActions(taskid)
        DeleteAllTaskVariables(taskid)
        DeleteTaskState(taskid)
        DeleteAllTaskTriggers(taskid)
    End Sub

    ''' <summary>
    ''' Deletes the state of the task from tasks states.
    ''' </summary>
    ''' <param name="taskid">The taskid.</param><returns></returns>
    Private Function DeleteTaskState(taskid As String) As Boolean
        Dim res As Boolean = True
        Dim result = From taskStatesRow In TectonicDataSetObj.tasks_states Where taskStatesRow.task_id = taskid
        For Each taskState As TectonicDataSet.tasks_statesRow In result
            taskState.Delete()
        Next
        Try
            _tasksStatusTableAdapter.Update(TectonicDataSetObj.tasks_states)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("", ex, True)
            res = False
        End Try
        result = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Delete a single variable row from task variables table using task id.
    ''' </summary>
    ''' <param name="taskVariablesObj">The task variables obj.</param>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function DeleteVariable(taskVariablesObj As TaskVariables, ByVal ParamArray args() As String) As Boolean
        Dim res As Boolean = False
        Using taskVariablesObj
            Dim taskVariables = From tasksVariablesRow In TectonicDataSetObj.task_variables
                                Where tasksVariablesRow.task_id = taskVariablesObj.TaskId And tasksVariablesRow.variable_name = args(0)
            For Each row As TectonicDataSet.task_variablesRow In taskVariables
                row.Delete()
            Next
            Try
                _tasksVariablesTableAdapter.Update(TectonicDataSetObj.task_variables)
            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Alerter.REP("", ex, True)
                res = False
            End Try
            taskVariables = Nothing
            Return res
        End Using
    End Function

    ''' <summary>
    ''' Deletes a global variable using its name.
    ''' </summary>
    ''' <param name="globalVariableName">Name of the global variable.</param><returns></returns>
    Private Function DeleteGlobalVariable(globalVariableName As String) As Boolean
        Dim res As Boolean = True
        Dim globalsRowToBeDeleted As TectonicDataSet.globalsRow
        globalsRowToBeDeleted = TectonicDataSetObj.globals.FindByname(globalVariableName)
        Try
            If Not globalsRowToBeDeleted Is Nothing Then
                globalsRowToBeDeleted.Delete()
                _globalsTableAdapter.Update(TectonicDataSetObj.globals)
            Else
                Balora.Alerter.REP(My.Resources.GlobNotDefined, Nothing, False)
                res = False
            End If
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Balora.Alerter.REP("", ex, True)
            res = False
        End Try
        globalsRowToBeDeleted = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Deletes a trigger using its id.
    ''' </summary>
    ''' <param name="trigger">The trigger.</param><returns></returns>
    Private Function DeleteTrigger(trg As ITrigger) As Boolean
        Dim res As Boolean = True
        Dim result = From row In TectonicDataSetObj.triggers
                         Where row.trigger_id = trg.TriggerId
        Try
            For Each row As TectonicDataSet.triggersRow In result
                row.Delete()
            Next
            res = CBool(_triggersTableAdapter.Update(TectonicDataSetObj.triggers))
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("", ex, True)
            res = False
        End Try
        result = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Deletes an action using its id.
    ''' </summary>
    ''' <param name="action">The action.</param><returns></returns>
    Private Function DeleteAction([action] As Action) As Boolean
        Dim res As Boolean = True
        Dim result = From row In TectonicDataSetObj.actions
             Where row.action_id = [action].ActionId
        Try
            For Each row As TectonicDataSet.actionsRow In result
                row.Delete()
            Next
            res = CBool(_actionsTableAdapter.Update(TectonicDataSetObj.actions))
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("", ex, True)
            res = False
        End Try
        result = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Deletes the type of the history by entry type.
    ''' </summary>
    ''' <param name="entryType">Type of the entry.</param><returns></returns>
    Private Function DeleteHistoryByEntryType(entryType As IHistoryEntry.HistoryEntryType) As Boolean
        Dim res As Boolean = True
        Try
            For Each row In TectonicDataSetObj.history
                If row.entry_type = entryType Then
                    row.Delete()
                End If
            Next
            _historyTableAdapter.Update(TectonicDataSetObj.history)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Balora.Alerter.REP("", ex, True)
            res = False
        End Try
        Return res
    End Function

    'Back to CacadoreToTectonic.ClearHistoryByTimeRange to get more info about the logic behind args here
    ''' <summary>
    ''' Deletes the history by time range.
    ''' </summary>
    ''' <param name="args">The args.</param><returns></returns>
    Private Function DeleteHistoryByTimeRange(ByVal ParamArray args() As String) As Boolean
        Dim res As Boolean = True
        Dim timeRangeType As IHistory.TimeRange = CType(args(1), IHistory.TimeRange)
        Dim customDate As Date
        Select Case timeRangeType
            Case IHistory.TimeRange.Everything
            Case IHistory.TimeRange.CustomDate
                customDate = CDate(args(2))
            Case IHistory.TimeRange.LastFourHours
            Case IHistory.TimeRange.LastHour
            Case IHistory.TimeRange.LastTwoHours
            Case IHistory.TimeRange.Today
        End Select
        Try
            For Each row In TectonicDataSetObj.history
                row.Delete()
            Next
            _historyTableAdapter.Update(TectonicDataSetObj.history)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Balora.Alerter.REP("", ex, True)
            res = False
        End Try
        timeRangeType = Nothing
        args = Nothing
        Return res
    End Function
#End Region

#Region "Delete All Rows"

    ''' <summary>
    ''' Deletes the full history.
    ''' </summary><returns></returns>
    Private Function DeleteFullHistory() As Boolean
        Try
            For Each row In TectonicDataSetObj.history
                row.Delete()
            Next
            _historyTableAdapter.Update(TectonicDataSetObj.history)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Balora.Alerter.REP("", ex, True)
            Return False
        End Try
        Return True
    End Function

    ''' <summary>
    ''' Deletes all associated actions.
    ''' </summary>
    ''' <param name="taskid">The taskid.</param><returns></returns>
    Private Function DeleteAllAssociatedActions(taskid As String) As Boolean
        Dim res As Boolean = True
        Dim taskActions = From actionRow In TectonicDataSetObj.actions Where actionRow.task_id = taskid
        For Each [action] As TectonicDataSet.actionsRow In taskActions
            [action].Delete()
            [action] = Nothing
        Next
        Try
            _actionsTableAdapter.Update(TectonicDataSetObj.actions)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("", ex, True)
            res = False
        End Try
        taskActions = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Deletes all task variables. It doesn't delete all tasks variables table, only the task's variable.
    ''' </summary>
    ''' <param name="taskid">The taskid.</param><returns></returns>
    Private Function DeleteAllTaskVariables(taskid As String) As Boolean
        Dim res As Boolean = True
        Dim taskVariables = From tasksVariablesRow In TectonicDataSetObj.task_variables Where tasksVariablesRow.task_id = taskid
        For Each taskVariable As TectonicDataSet.task_variablesRow In taskVariables
            taskVariable.Delete()
            taskVariable = Nothing
        Next
        Try
            _tasksVariablesTableAdapter.Update(TectonicDataSetObj.task_variables)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("", ex, True)
            res = False
        End Try
        taskVariables = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Deletes all task triggers.
    ''' </summary>
    ''' <param name="taskid">The taskid.</param><returns></returns>
    Private Function DeleteAllTaskTriggers(taskid As String) As Boolean
        Dim res As Boolean = True
        Dim result = From triggerRow In TectonicDataSetObj.triggers Where triggerRow.task_id = taskid
        For Each triggerRow As TectonicDataSet.triggersRow In result
            triggerRow.Delete()
            triggerRow = Nothing
        Next
        Try
            _triggersTableAdapter.Update(TectonicDataSetObj.triggers)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("", ex, True)
            res = False
        End Try
        result = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Deletes all tasks in Tectonic.
    ''' </summary><returns></returns>
    Private Function DeleteAllTasks() As Boolean
        Try
            DeleteAllActions()
            DeleteAllTriggers()
            DeleteAllVariables()
            For Each taskRow In TectonicDataSetObj.tasks
                taskRow.Delete()
                taskRow = Nothing
            Next
            _tasksTableAdapter.Update(TectonicDataSetObj.tasks)
            Return True
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Balora.Alerter.REP("", ex, True)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Deletes all tasks states.
    ''' </summary><returns></returns>
    Private Function DeleteAllTasksStates() As Boolean
        Try
            For Each stateRow In TectonicDataSetObj.tasks_states
                stateRow.Delete()
                stateRow = Nothing
            Next
            _tasksStatusTableAdapter.Update(TectonicDataSetObj.tasks_states)
            Return True
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Balora.Alerter.REP("", ex, True)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Deletes all actions.
    ''' </summary><returns></returns>
    Private Function DeleteAllActions() As Boolean
        'This delete All actions rows.
        For Each actionRow In TectonicDataSetObj.actions
            actionRow.Delete()
            actionRow = Nothing
        Next
        Try
            _actionsTableAdapter.Update(TectonicDataSetObj.actions)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("", ex, True)
            Return False
        End Try
        Return True
    End Function

    ''' <summary>
    ''' Deletes all variables table.
    ''' </summary><returns></returns>
    Private Function DeleteAllVariables() As Boolean
        For Each row In TectonicDataSetObj.task_variables
            row.Delete()
            row = Nothing
        Next
        Try
            _tasksVariablesTableAdapter.Update(TectonicDataSetObj.task_variables)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("", ex, True)
            Return False
        End Try
        Return True
    End Function

    ''' <summary>
    ''' Deletes all triggers.
    ''' </summary><returns></returns>
    Private Function DeleteAllTriggers() As Boolean
        For Each row In TectonicDataSetObj.triggers
            row.Delete()
            row = Nothing
        Next
        Try
            _triggersTableAdapter.Update(TectonicDataSetObj.triggers)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("", ex, True)
            Return False
        End Try
        Return True
    End Function

    ''' <summary>
    ''' Deletes all groups.
    ''' </summary><returns></returns>
    Private Function DeleteAllGroups() As Boolean
        For Each row In TectonicDataSetObj.groups
            row.Delete()
            row = Nothing
        Next
        Try
            _groupsTableAdapter.Update(TectonicDataSetObj.groups)
            MoveAllTasksToRootGroup()
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Balora.Alerter.REP("", ex, True)
            Return False
        End Try
        Return True
    End Function

    ''' <summary>
    ''' Deletes all globals variables.
    ''' </summary><returns></returns>
    Private Function DeleteAllGlobalsVariables() As Boolean
        Try
            For Each row In TectonicDataSetObj.globals
                row.Delete()
                row = Nothing
            Next
            _globalsTableAdapter.Update(TectonicDataSetObj.globals)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Balora.Alerter.REP("", ex, True)
            Return False
        End Try
        Return True
    End Function
#End Region
#End Region
    'Util, Reusable functions against Tectonic.
#Region "Reusable Utils"
    ''' <summary>
    ''' Gets a new group id.
    ''' </summary><returns></returns>
    Public Shared Function gngi() As String
        Return getNewGroupId()
    End Function

    ''' <summary>
    ''' Gets a new task id.
    ''' </summary><returns></returns>
    Public Shared Function gnti() As String
        Return getNewTaskId()
    End Function

    ''' <summary>
    ''' Gets  groups table rows count.
    ''' </summary><returns></returns>
    Private Shared Function GetGroupsCount() As Integer
        Dim allGroups = From groupRow As TectonicDataSet.groupsRow In TectonicDataSetObj.groups Select groupRow.id
        Dim count = allGroups.Count
        allGroups = Nothing
        Return count
    End Function

    ''' <summary>
    ''' Determines whether [is registered group] [the specified group id].
    ''' </summary>
    ''' <param name="groupId">The group id.</param><returns>
    '''   <c>true</c> if [is registered group] [the specified group id]; otherwise, <c>false</c>.
    ''' </returns>
    Private Shared Function IsRegisteredGroup(groupId As String) As Boolean
        Dim duplicatedName = From groupRow As TectonicDataSet.groupsRow In TectonicDataSetObj.groups Select groupRow.id
                             Where id = groupId
        Dim res = CType(duplicatedName.Count, Boolean)
        duplicatedName = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Determines whether [is registered task] [the specified task id].
    ''' </summary>
    ''' <param name="taskId">The task id.</param><returns>
    '''   <c>true</c> if [is registered task] [the specified task id]; otherwise, <c>false</c>.
    ''' </returns>
    Private Shared Function IsRegisteredTask(taskId As String) As Boolean
        'Dim tasksTable As TectonicDataSet.tasksDataTable = _tasksTableAdapter.GetData
        Dim duplicatedName = From taskRow As TectonicDataSet.tasksRow
                             In TectonicDataSetObj.tasks
                             Select taskRow.id
                             Where id = taskId
        Dim res = CType(duplicatedName.Count, Boolean)
        Return res
    End Function

    ''' <summary>
    ''' Determines whether [is registered variable] [the specified variable name].
    ''' </summary>
    ''' <param name="variableName">Name of the variable.</param>
    ''' <param name="global">if set to <c>true</c> [global].</param><returns>
    '''   <c>true</c> if [is registered variable] [the specified variable name]; otherwise, <c>false</c>.
    ''' </returns>
    Private Shared Function IsRegisteredVariable(variableName As String,
                                                 Optional [global] As Boolean = False) As Boolean
        If Not [global] Then
            Dim duplicatedName = From taskVariableRow As TectonicDataSet.task_variablesRow
                                 In TectonicDataSetObj.task_variables
                                 Select taskVariableRow.variable_name, taskVariableRow.task_id
                                 Where variable_name = variableName
            Dim duplicatedNameCount = duplicatedName.Count
            Dim res = CType(duplicatedNameCount, Boolean)
            duplicatedName = Nothing
            Return res
        Else
            'Dim globalsTable As TectonicDataSet.globalsDataTable = _globalsTableAdapter.GetData
            Dim duplicatedName = From globalsRow As TectonicDataSet.globalsRow
                                 In TectonicDataSetObj.globals
                                 Select globalsRow.name
                                 Where name = variableName
            Dim duplicatedNameCount = duplicatedName.Count
            Dim res = CType(duplicatedNameCount, Boolean)
            duplicatedName = Nothing
            Return res
        End If
    End Function

    ''' <summary>
    ''' Determines whether [is variable attached to this task] [the specified variable name].
    ''' </summary>
    ''' <param name="variableName">Name of the variable.</param>
    ''' <param name="taskId">The task id.</param><returns>
    '''   <c>true</c> if [is variable attached to this task] [the specified variable name]; otherwise, <c>false</c>.
    ''' </returns>
    Private Function IsVariableAttachedToThisTask(variableName As String, taskId As String) As Boolean
        Dim isAttached = From taskVariableRow As TectonicDataSet.task_variablesRow
                         In TectonicDataSetObj.task_variables
                         Select taskVariableRow.task_id, taskVariableRow.variable_name
                         Where variable_name = variableName And task_id = taskId
        If isAttached.Count > 0 Then
            isAttached = Nothing
            Return True
        Else
            isAttached = Nothing
            Return False
        End If
    End Function

    ''' <summary>
    ''' Gets the history table rows count.
    ''' </summary><returns></returns>
    Private Shared Function GetHistoryTableRowsCount() As Integer
        Dim rowsCount = From historyRow In TectonicDataSetObj.history
        Dim res = CType(rowsCount.Count, Integer)
        rowsCount = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Gets the new history table increment value.
    ''' </summary><returns></returns>
    Private Shared Function GetNewHistoryTableIncrementValue() As Integer
        'First we will get the rows count
        'Increment it
        'Check if it already exists in the table
        'If it exists then start recursion 
        Dim rowsCount = From historyRow In TectonicDataSetObj.history
        Dim res = getIncrementRowNumber(rowsCount.Count, "History")
        rowsCount = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Gets the new variable table increment value.
    ''' </summary><returns></returns>
    Private Shared Function GetNewTaskVariablesTableIncrementValue() As Integer
        Dim rowsCount = From row In TectonicDataSetObj.task_variables
        Dim res = getIncrementRowNumber(rowsCount.Count, "TaskVariables")
        rowsCount = Nothing
        Return res
    End Function

    Private Shared Function getIncrementRowNumber(rowsCount As Integer, tableName As String, Optional newValue As Integer = -1) As Integer
        'First we will get the rows count
        'Increment it
        'Check if it already exists in the table
        'If it exists then start recursion 
        Dim newIncrementValue As Integer
        If newValue > 0 Then
            newIncrementValue = newValue
        ElseIf newValue = -1 Then
            newIncrementValue = rowsCount + 1
        End If
        If isAddedIncrementNumber(newIncrementValue, tableName) Then
            Return getIncrementRowNumber(rowsCount, tableName, newIncrementValue + 1)
        End If
        Return newIncrementValue
    End Function

    Private Shared Function isAddedIncrementNumber(noToBeChecked As Integer, tableName As String) As Boolean
        Dim checkIfExists As EnumerableRowCollection(Of Long) = Nothing
        If tableName = "TaskVariables" Then
            checkIfExists = From variableRow In TectonicDataSetObj.task_variables
                            Select variableRow.no
                            Where no = noToBeChecked
        ElseIf tableName = "History" Then
            checkIfExists = From historyRow In TectonicDataSetObj.history Select historyRow.no
                            Where no = noToBeChecked
        End If
        If checkIfExists.Count > 0 Then
            checkIfExists = Nothing
            Return True
        Else
            checkIfExists = Nothing
            Return False
        End If
    End Function

    ''' <summary>
    ''' Determines whether [is registered trigger] [the specified trigger id].
    ''' </summary>
    ''' <param name="triggerId">The trigger id.</param><returns>
    '''   <c>true</c> if [is registered trigger] [the specified trigger id]; otherwise, <c>false</c>.
    ''' </returns>
    Private Shared Function IsRegisteredTrigger(triggerId As String) As Boolean
        Dim duplicatedName = From triggerRow
                             In TectonicDataSetObj.triggers
                             Select triggerRow.trigger_id
                             Where trigger_id = triggerId
        Dim res = CType(duplicatedName.Count, Boolean)
        duplicatedName = Nothing
        Return res
    End Function

    ''' <summary>
    ''' Determines whether [is registered action] [the specified action id].
    ''' </summary>
    ''' <param name="actionId">The action id.</param><returns>
    '''   <c>true</c> if [is registered action] [the specified action id]; otherwise, <c>false</c>.
    ''' </returns>
    Private Shared Function IsRegisteredAction(actionId As String) As Boolean
        Dim duplicatedName = From actionRow In TectonicDataSetObj.actions
                             Select actionRow.action_id
                             Where action_id = actionId
        Dim res As Boolean = CType(duplicatedName.Count, Boolean)
        duplicatedName = Nothing
        Return res
    End Function

    'Complex logic to handle "Ungrouped" group.
    'فيها أخطاء وعايزة تتظبط
    Private Shared Function CheckUngroupedGroup() As String
        Dim res As String = vbNullString
        Dim id As String = Commands.GetGroupIdByName("Ungrouped")
        Dim isThereUngroupedGroup As Boolean = IsRegisteredGroup("0")
        'لو مافيش بالمعرف دة هنعملها
        If Not isThereUngroupedGroup AndAlso id = "" Then
            Using newUngroupedGroup As New Group("0")
                newUngroupedGroup.Name = "Ungrouped"
                newUngroupedGroup.Save()
                res = "0"
            End Using
        Else
            'لو فيه جروب واخدة المعرف دة هنشيك علي اسمها الأول
            Dim groupName As String = Commands.GetGroupName("0")
            'لو مش بالاسم دة...
            If groupName <> "Ungrouped" Then
                'هاشوف الاسم متاخد والا لأ، ولو مش متاخد هعمل واحدة جديدة بأسمه
                If id = "" Then
                    Dim newId As String = gngi()
                    Using newUngroupedGroup As New Group(newId)
                        newUngroupedGroup.Name = "Ungrouped"
                        newUngroupedGroup.Save()
                        res = newId
                    End Using
                Else
                    'ولو متاخد هاخده مباشرة
                    res = id
                End If
            Else
                'لو فيه جروب بنفس المعرف ونفس الاسم هاخدها عالطول
                res = "0"
            End If
        End If
        Return res
    End Function

    ''' <summary>
    ''' Moves the tasks of a specified group to the root group.
    ''' </summary>
    ''' <param name="groupId">The group id.</param><returns></returns>
    Private Shared Function MoveTasksToRootGroup(groupId As String) As Boolean
        'Dim ungroupedGroupId As String = CheckUngroupedGroup()
        For Each row In TectonicDataSetObj.tasks
            If groupId = row.group_id Then
                row.group_id = CStr(0) 'ungroupedGroupId
            End If
        Next
        Try
            _tasksTableAdapter.Update(TectonicDataSetObj.tasks)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Balora.Alerter.REP("", ex, True)
        End Try
        Return True
    End Function

    ''' <summary>
    ''' Moves all tasks to root group.
    ''' </summary><returns></returns>
    Private Shared Function MoveAllTasksToRootGroup() As Boolean
        'Dim ungroupedGroupId As String = CheckUngroupedGroup()
        For Each row In TectonicDataSetObj.tasks
            row.group_id = CStr(0) 'ungroupedGroupId
        Next
        Try
            _tasksTableAdapter.Update(TectonicDataSetObj.tasks)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Balora.Alerter.REP("", ex, True)
        End Try
        Return True
    End Function

    Private Shared Function getNewGroupId() As String
        Dim newId As String = generateNewId()
        If Not IsRegisteredGroup(newId) Then
            Return newId
        Else
            Return getNewGroupId() 'This recursion didn't tested yet.
        End If
    End Function

    Private Shared Function getNewTaskId() As String
        Dim newId As String = Balora.RandomBytesGenerator.GRCM()
        If Not IsRegisteredTask(newId) Then
            Return newId
        Else
            Return getNewTaskId() ':$ first time in my life to use recursion
        End If
    End Function

    Private Shared Function generateNewId() As String
        Return Balora.RandomBytesGenerator.GRCM()
    End Function

    Private Function moveUpActionExecutionPrecedence(actionId As String, newPrecedence As Integer) As Boolean
        'Kinda a piece of art, handle drag or move up.
        Dim taskId As String = Commands.ExtractTaskId(actionId)
        Dim tskActions = GetTaskActionsRow(taskId)
        Dim sourceIndex As Integer = 0

        Dim queryResult = From acty In tskActions Where acty.action_id = actionId
        If queryResult.Count > 0 Then
            Dim acty = queryResult.First
            sourceIndex = CInt(acty.execution_precedence)
            acty.execution_precedence = newPrecedence
            _actionsTableAdapter.Update(acty)
            acty = Nothing
        End If

        For Each acty In tskActions
            If acty.action_id <> actionId AndAlso acty.execution_precedence <= sourceIndex AndAlso acty.execution_precedence >= newPrecedence Then
                acty.execution_precedence = acty.execution_precedence + 1
                _actionsTableAdapter.Update(acty)
                acty = Nothing
            End If
        Next
        tskActions = Nothing
        Return True
    End Function

    Private Function moveDownActionExecutionPrecedence(actionId As String, newPrecedence As Integer) As Boolean
        'Kinda a piece of art, handle drag or move down.
        Dim taskId As String = Commands.ExtractTaskId(actionId)
        Dim tskActions = GetTaskActionsRow(taskId)
        Dim sourceIndex As Integer = 0

        Dim queryResult = From acty In tskActions Where acty.action_id = actionId
        If queryResult.Count > 0 Then

            Dim acty = queryResult.First
            sourceIndex = CInt(acty.execution_precedence)
            Dim oldExecutionPrecedence As Integer = CInt(acty.execution_precedence)
            acty.execution_precedence = newPrecedence + 1
            _actionsTableAdapter.Update(acty)

            If acty.execution_precedence - oldExecutionPrecedence = 1 Then
                Dim result = From acty1 In tskActions Where acty.execution_precedence = newPrecedence + 1 And acty1.action_id <> actionId
                If result.Count > 1 Then result.First.execution_precedence = result.First.execution_precedence - 1
                _actionsTableAdapter.Update(result.First)
                tskActions = Nothing
                Return True
            End If
        End If

        For Each acty In tskActions
            If acty.action_id <> actionId AndAlso acty.execution_precedence <= newPrecedence + 1 AndAlso Not acty.execution_precedence <= 1 Then
                acty.execution_precedence = acty.execution_precedence - 1
                _actionsTableAdapter.Update(acty)
                acty = Nothing
            End If
        Next
        tskActions = Nothing
        Return False
    End Function
#End Region
#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                TectonicDataSetObj.Dispose()
                _tasksTableAdapter.Dispose()
                _groupsTableAdapter.Dispose()
                _globalsTableAdapter.Dispose()
                _tasksStatusTableAdapter.Dispose()
                _tasksVariablesTableAdapter.Dispose()
                _triggersTableAdapter.Dispose()
                _actionsTableAdapter.Dispose()
                _historyTableAdapter.Dispose()
                _settingsTableAdapter.Dispose()
                _timeTakenToFillAllTablesValue = Nothing
                TectonicDataSetObj = Nothing
                _groupsTableAdapter = Nothing
                _globalsTableAdapter = Nothing
                _tasksStatusTableAdapter = Nothing
                _tasksVariablesTableAdapter = Nothing
                _triggersTableAdapter = Nothing
                _actionsTableAdapter = Nothing
                _historyTableAdapter = Nothing
                _settingsTableAdapter = Nothing
            End If
            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

#Region "IBroadcaster"
    Public Sub AddListener(listener As Balora.IListener) Implements Balora.IBroadcaster.AddListener
        _tectonicListeners.Add(listener)
    End Sub

    Private Sub NotifyListeners(message As Object) Implements Balora.IBroadcaster.NotifyListeners
        For Each listener As IListener In _tectonicListeners
            listener.Notify(message)
        Next
    End Sub

    Public Sub RemoveListener(listener As Balora.IListener) Implements Balora.IBroadcaster.RemoveListener
        _tectonicListeners.Remove(listener)
    End Sub
#End Region
#End Region

End Class