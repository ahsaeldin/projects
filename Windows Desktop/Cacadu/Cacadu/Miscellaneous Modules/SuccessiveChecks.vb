Imports TecDAL
Imports Balora
Imports Cacadore
Imports System.IO
Imports Balora.Datler
Imports System.Threading

Module SuccessiveChecks

    'لو هترجع تشغلهم تاني اتأكد الأول
    'من استهلاكهم للرام علي المدي الطويل
    'Use to switch on/off all checks.
    Friend _switch As Boolean = False
    Private _startNow As Boolean = False

    Friend Sub PrepareSuccessiveChecks()

        Balora.xTimer.Switch = _switch

        If Not _switch Then Exit Sub

        _dueTimeForStartUnfulfilledTasksChecker = 1000 * 60 * 20
        _unfulfilledTasksCheckerTimerInterval = 1000 * 60 * 20

        _dueTimeForStartMalformationChecker = 1000 * 60 * 10
        _malformationCheckerTimerInterval = 1000 * 60 * 10

        _pendingHistoryEntiresTimerInterval = 1000 * 60 * 25
        _dueTimeForpendingHistoryEntiresTasksChecker = 1000 * 60 * 25

        _tectonicMalformationCheckerXTimer = New Balora.xTimer(_dueTimeForStartMalformationChecker,
                                                               _malformationCheckerTimerInterval,
                                                               _startNow,
                                                               Sub() tectonicMalformetionCheckerTimer(New Object))

        _unfulfilledTasksCheckerXTimer = New Balora.xTimer(_dueTimeForStartUnfulfilledTasksChecker,
                                                           _unfulfilledTasksCheckerTimerInterval,
                                                           _startNow,
                                                           Sub() checkInterruptedTasks(New Object))

        _pendingHistoryTimerXTimer = New Balora.xTimer(_dueTimeForpendingHistoryEntiresTasksChecker,
                                                       _pendingHistoryEntiresTimerInterval,
                                                       Sub() PushPendingHistoryEntires(New Object))
    End Sub


#Region "Malformation Checking Area"
    Private _malformationCheckingCount As Int64 = 0
    Private _malformationCheckerTimerInterval As Integer
    Private _dueTimeForStartMalformationChecker As Integer
    Private _tectonicMalformationCheckerXTimer As Balora.xTimer

    Private Sub tectonicMalformetionCheckerTimer(stateInfo As Object)

        If _switch = False Then
            Balora.Hodhod.STOW("Successive checks are disabled, skip check for Tectonic malformetion.")
            Exit Sub
        End If

        _malformationCheckingCount = _malformationCheckingCount + 1
        Balora.Hodhod.STOW(String.Format("Start Checking #{0} if Tectonic is malformed", _malformationCheckingCount))

        If Cacadore.Watcher.IsBusy Then

            Balora.Hodhod.STOW("Tectonic is busy write now, skip checking and wait for the next tick.")

            Exit Sub
        End If

        Balora.Hodhod.STOW("Disable Tectonic malformation checker to prevent successive calls...")

        Try

            FileCopy(TypedDataSetCrud.TectonicPath, TypedDataSetCrud.TectonicTempPath)

        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("Error creating a temporary backup of database.", ex, True)
            'Enable the timer again.
            _tectonicMalformationCheckerXTimer.Resume()
            Exit Sub
        End Try

        'Compare files...
        Dim isEqual As Boolean = Balora.Util.CompareFiles(
            TypedDataSetCrud.TectonicTempPath, TypedDataSetCrud.TectonicBackupPath)

        'Tectonic changed...
        If Not isEqual Then

            Balora.Hodhod.STOW("Database files are not equal, checking if Tectonic is malformed...")

            Dim malformationCheckingResult As String

            SyncLock Cacadore.Settings.CrudObject
                malformationCheckingResult = DevartSQLiteDatabase.CheckIfMalformed(TecDAL.TypedDataSetCrud.Connection)
            End SyncLock

            'if not malformed then update TectonicB by Tectonic
            If malformationCheckingResult = "ok" Then

                Balora.Hodhod.STOW("Tectonic is not malformed, now make a new backup of Tectonic...")

                Try
                    'Update TectonicB by Tectonic
                    FileCopy(TypedDataSetCrud.TectonicPath, TypedDataSetCrud.TectonicBackupPath)
                Catch ex As Exception
                    If Debugger.IsAttached Then Debugger.Break()
                    Alerter.REP("Error creating a database backup.", ex, True)
                    'Enable the timer again.
                    _tectonicMalformationCheckerXTimer.Resume()
                    Exit Sub
                End Try

            ElseIf malformationCheckingResult = "" Then

                Balora.Hodhod.STOW("Cannot check if Tectonic is malformed because Tectonic may be locked so we can't check malformation, nothing to do and wait for next tick...")

            Else 'If malformed then replace Tectonic by TectonicB... 

                Balora.Hodhod.STOW("Tectonic is malformed, displaying error message to inform user about Tectonic malformation...")

                'mald stands for malformed database
                MsgBox(My.Resources.mald,
                       MsgBoxStyle.Critical,
                       "Cacadu")

                Balora.Hodhod.STOW("Now restoring Tectonic from TectonicB...")

                Try
                    If FileIO.FileSystem.FileExists(TypedDataSetCrud.TectonicBackupPath) Then
                        'Restore Tectonic...
                        FileCopy(TypedDataSetCrud.TectonicBackupPath, TypedDataSetCrud.TectonicPath)
                    Else

                        'AhSaElDin 20111227: what if backup isn't there.
                        'dd stand for database damaged.
                        MsgBox(My.Resources.dd,
                             MsgBoxStyle.Critical,
                             "Cacadu")

                    End If
                Catch ex As Exception
                    If Debugger.IsAttached Then Debugger.Break()
                    MsgBox("Failed to restore database, please reinstall Cacadu.", MsgBoxStyle.Critical, "Cacadu")
                    Alerter.REP("Error restoring database backup.", ex, True)
                End Try

            End If

        End If

        Try
            'Delete Temp Tectonic File
            File.Delete(TypedDataSetCrud.TectonicTempPath)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("Error deleting the temporary backup of database.", ex, True)
        End Try
    End Sub
#End Region

#Region "Unfulfilled Tasks Checking"
    Dim _unfulfilledTasksCheckingCount As Int64 = 0
    Private _unfulfilledTasksCheckerTimerInterval As Integer
    Private _dueTimeForStartUnfulfilledTasksChecker As Integer
    Private _unfulfilledTasksCheckerXTimer As New Balora.xTimer

    ''' <summary>
    ''' Checks the interrupted tasks and delete all timeout triggers.
    ''' </summary>
    ''' <param name="stateInfo">The state info.</param>
    ''' <remarks>
    ''' This function doesn't collect UnFullFilled tasks,
    ''' it just mark them as UnFullFilled in Tectonic.
    ''' If you want to get them, use Watcher.UnFullFilledTasks
    ''' </remarks> 
    Public Sub checkInterruptedTasks(stateInfo As Object)

        If _switch = False Then
            Balora.Hodhod.STOW("Successive checks are disabled, skip check for unfulfilled tasks.")
            Exit Sub
        End If

        _unfulfilledTasksCheckingCount = _unfulfilledTasksCheckingCount + 1
        Balora.Hodhod.STOW("Start Checking #" & _unfulfilledTasksCheckingCount & " if there are unfulfilled tasks..")

        If Cacadore.Watcher.IsBusy Then

            Balora.Hodhod.STOW("Tectonic is busy write now, skip checking and wait for the next tick.")

            Exit Sub

        End If

        Balora.Hodhod.STOW("Get all tasks states from task_states table...")

        'Get all tasks states from task_states table
        Dim tasksStatesDic As Dictionary(Of String, TaskState) = CacadoreToTectonic.GetAllRegisteredTasksStates()

        'Query for all tasks that has value other than "N/A" in CurrentExecutingActionId field.
        Dim runningTasksStates = From tsk In tasksStatesDic
                                 Where tsk.Value.CurrentExecutingActionId <> "n/a" And tsk.Value.TaskExecutionState <> ITaskState.ExecutionState.Unfulfilled

        Dim tasksStatesCount As Integer = runningTasksStates.Count()

        If tasksStatesCount > 0 Then
            SyncLock Cacadore.Settings.CrudObject

                'For each task that has a state of running, we will check if it is watched by Watcher...")
                'If not watched by Watcher, then it is unfulfilled task.")
                SyncLock Cacadore.Watcher.WatchedTasks
                    For Each kvp In runningTasksStates
                        If Not Watcher.IsTaskWatched(kvp.Key) Then

                            Balora.Hodhod.STOW("A task with id " & kvp.Key & " has a state of running in Tectonic, however, it is not watched by Watcher.")
                            Balora.Hodhod.STOW("Task " & kvp.Key & " is unfulfilled task.")
                            Balora.Hodhod.STOW("Getting task " & kvp.Key & " using its id...")

                            Dim tsk As Task = Commands.GetTaskById(kvp.Key)

                            Balora.Hodhod.STOW("Setting task's nextTriggerIdValue &  currentExecutingActionIdValue fields.")

                            'AhSaElDin 20111231: property is writable for now.
                            'Balora.Util.SetPrivateField(tsk.State, "nextTriggerIdValue", kvp.Value.NextTriggerId)
                            tsk.State.NextTriggerId = kvp.Value.NextTriggerId

                            'AhSaElDin 20111231: property is writable for now.
                            'Balora.Util.SetPrivateField(tsk.State, "currentExecutingActionIdValue", kvp.Value.CurrentExecutingActionId)
                            tsk.State.CurrentExecutingActionId = kvp.Value.CurrentExecutingActionId

                            Balora.Hodhod.STOW("Change task " & kvp.Key & " state to Unfulfilled...")

                            tsk.State.ChangeExecutionState(ITaskState.ExecutionState.Unfulfilled)

                            If Not IsNothing(tsk) Then tsk.Dispose()
                            If Not IsNothing(tsk) Then tsk = Nothing
                        End If
                    Next
                End SyncLock
                Cacadore.Commands.DeleteTimeOutTriggers()
            End SyncLock
        End If
    End Sub

#End Region

#Region "Pushing History Entries"
    Private _pendingHistoryEntiresTimerInterval As Integer
    Private _dueTimeForpendingHistoryEntiresTasksChecker As Integer
    Private _pendingHistoryTimerXTimer As  Balora.xTimer

    Public Sub PushPendingHistoryEntires(stateInfo As Object)
        'The logic here needs to be more accurate.
        If _switch = False Then
            Balora.Hodhod.STOW("Successive checks are disabled, skip pushing pending history entires.")
            Exit Sub
        End If
        If Not Cacadore.Settings.HistoryEnabled Then Exit Sub
#If DEBUG Then
        My.Computer.FileSystem.WriteAllText("c:\koko.txt", "PushPendingHistoryEntires:: started", True)
#End If

        SyncLock pendingHistoryEntries
            If pendingHistoryEntries.Count > 0 AndAlso Not Watcher.IsExecuting AndAlso Cacadore.Settings.HistoryEnabled Then
#If DEBUG Then
                My.Computer.FileSystem.WriteAllText(
                     "c:\koko.txt", String.Format("PushPendingHistoryEntires:: pushing {0} item now" & vbNewLine, pendingHistoryEntries.Count), True)
#End If
                For i = 1 To pendingHistoryEntries.Count - 1
                    Dim entry = pendingHistoryEntries.Dequeue()
                    Cacadore.History.AddNewEntry(entry)
                Next
            End If
        End SyncLock
#If DEBUG Then
        My.Computer.FileSystem.WriteAllText("c:\koko.txt", "PushPendingHistoryEntires:: ended", True)
#End If
    End Sub
#End Region
End Module
