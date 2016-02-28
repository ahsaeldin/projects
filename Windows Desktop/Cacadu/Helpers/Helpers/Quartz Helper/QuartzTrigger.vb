'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   QuartzTrigger Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
'Just here we disable it in order to dynamically use Cacadore.Settings.Scheduler in StartListening and similar methods.
Option Strict Off

Imports Quartz
Imports Cacadore
Imports Quartz.Impl
Imports System.ComponentModel

Namespace QuartzHelper
    ''' <summary>
    ''' QuartzTrigger is a helper class inherits Cacadore.Triggers.ScheduleTrigger class
    ''' just to abstract the scheduling thing from the bone of Cacadore; so I can use 
    ''' Quartz now or change it in the future.
    ''' </summary>
    ''' <remarks>
    ''' Just to abstract my own implementation of Now(), StartListening() 
    ''' and StopListening() into this class, so Cacadore won't be coupled 
    ''' with Quartz.
    ''' </remarks>
    <Serializable()> Public Class QuartzTrigger
        Inherits Cacadore.Triggers.ScheduleTrigger
        Implements ITriggerListener
        Implements IDisposable

#Region "Fields"
        Private _triggerIdValue As String
        Private _calendarCValue As Object
        Private startListenResult As Object
        Private _innerTriggerValue As Quartz.ITrigger
#End Region

#Region "Properties"
        Public Overrides Property InnerTrigger As Object
            Set(ByVal value As Object)
                _innerTriggerValue = value
            End Set
            Get
                Return _innerTriggerValue
            End Get
        End Property

        Public Overrides Property Calendar As Object
            Get
                Return _calendarCValue
            End Get
            Set(value As Object)
                'If IsNothing(value) Then
                '    Dim sch As IScheduler = Cacadore.Settings.Scheduler
                '    Dim result As Boolean = sch.DeleteCalendar(Me.TriggerId & "Calendar")
                '    InnerTrigger.CalendarName = ""
                'End If
                _calendarCValue = value
            End Set
        End Property

        Public Overrides ReadOnly Property StartTimeUTC As Nullable(Of DateTime)
            Get
                Dim temp = InnerTrigger.StartTimeUtc
                If IsNothing(temp) Then
                    Return New Nullable(Of DateTime)
                Else
                    Return temp.Datetime
                End If
            End Get
        End Property

        Public Overrides ReadOnly Property EndTimeUTC As Nullable(Of DateTime)
            Get
                Dim temp = InnerTrigger.EndTimeUtc
                If IsNothing(temp) Then
                    Return New Nullable(Of DateTime)
                Else
                    Return temp.Datetime
                End If
            End Get
        End Property

        Public Overrides ReadOnly Property PreviousFireTimeUTC As Nullable(Of DateTime)
            Get
                Dim nextTimeList As IList = Nothing

                If StartTimeUTC.Value < Date.UtcNow Then
                    nextTimeList = TriggerUtils.ComputeFireTimesBetween(InnerTrigger,
                                                                        Calendar,
                                                                        StartTimeUTC,
                                                                        Date.UtcNow)
                End If

                Dim temp As Date?
                If Not IsNothing(nextTimeList) AndAlso nextTimeList.Count > 0 Then temp = nextTimeList.Item(nextTimeList.Count - 1).Datetime
                If IsNothing(temp) Then
                    Return New Nullable(Of DateTime)
                Else
                    Return temp
                End If
            End Get
        
        End Property

        Public Overrides ReadOnly Property NextFireTimeUTC As Nullable(Of DateTime)
            Get
                Dim temp As Nullable(Of DateTime) = getNextFireTime()
                If IsNothing(temp) Then
                    Return New Nullable(Of DateTime)
                Else
                    Return temp
                End If
            End Get
        End Property

        'AhSaElDin 20111228: Make it writable property.
        Private _descriptionValue As String = Cacadore.ITrigger.TriggersDescriptions.ScheduleTrigger
        Public Overrides Property Description As String
            Get
                Return _descriptionValue
            End Get
            Set(value As String)
                _descriptionValue = value
            End Set
        End Property

        Public Overrides ReadOnly Property TriggerType As Cacadore.ITrigger.TriggerTypeEnum
            Get
                Dim trgTypeName As String = InnerTrigger.GetType().Name
                If trgTypeName = "CronTriggerImpl" Then
                    Return Cacadore.ITrigger.TriggerTypeEnum.CronTrigger
                ElseIf trgTypeName = "SimpleTriggerImpl" Then
                    Return Cacadore.ITrigger.TriggerTypeEnum.SimpleTrigger
                Else
                    Return Cacadore.ITrigger.TriggerTypeEnum.ScheduleTrigger
                End If
            End Get
        End Property

        Public Overrides ReadOnly Property MayFireAgain As Boolean
            Get
                If TriggerState = Quartz.TriggerState.None Then
                    Return False
                Else
                    Return True
                End If
            End Get
        End Property

        Public ReadOnly Property TriggerState As Quartz.TriggerState
            Get
                Dim sch As IScheduler = Cacadore.Settings.Scheduler
                Dim trgState As Quartz.TriggerState = Quartz.TriggerState.None
                If Not sch.IsShutdown Then
                    Dim tempTrg As Quartz.ITrigger = Me.InnerTrigger
                    Dim trgKey As New TriggerKey(tempTrg.Key.Name, Task.ParentGroupId)
                    If Not IsNothing(Task) Then trgState = sch.GetTriggerState(trgKey)
                Else
                    Balora.Hodhod.STOW("Cannot get trigger state as scheduler is shutdown.")
                End If
                Return trgState
            End Get
        End Property

        Public Property TriggerListnerName As String
#End Region

#Region "Methods"
        Public Overrides Function StartListening() As Object
            Dim startingDate As Date? = _startListening()
            If Not IsNothing(startingDate) Then Balora.Alerter.SendMessageToCacadu("ClearStatusBarState")
            Return startingDate
        End Function

        ''' <summary>
        ''' Stops the listening.
        ''' </summary>
        ''' <remarks>
        ''' StopListening is just like task.UnAttachTrigger beside deleting the job from quartz.
        ''' </remarks>
        Public Overrides Function StopListening() As Boolean
            Dim sch As IScheduler = Cacadore.Settings.Scheduler
            Dim innTrg As Quartz.ITrigger = CType(InnerTrigger, Quartz.ITrigger)
            Dim groupName = innTrg.Key.Group

            If Not IsNothing(innTrg.JobKey) Then
                Dim jobName = innTrg.JobKey.Name
                Dim jobKey As New JobKey(jobName, groupName)
                Dim isDeleted As Boolean = sch.DeleteJob(jobKey)
            End If

            Dim triggerName = innTrg.Key.Name
            Dim trgKey As New TriggerKey(triggerName, groupName)
            Dim isUnscheduled As Boolean = sch.UnscheduleJob(trgKey)

            Dim newDescription As String
            newDescription = "Will no longer run " & Description.ToLower
            newDescription = newDescription.Replace("run run", "run")
            If Not IsNothing(Me.Task) AndAlso Not IsNothing(Me.Task.State) Then
                Watcher.AddNewEntry(Me.Task.TaskId,
                                    Me.Task.Properties.TaskName,
                                    IHistoryEntry.HistoryEntryType.Trigger,
                                    newDescription,
                                    DateTime.Now,
                                    Me.Task.State.GetExecutionState())
            End If
            Return MyBase.Delete()
        End Function
#End Region

#Region "Helper Methods"
        'Just schedule not start the trigger. Used in _startListening or if you want
        'to schedule a trigger but start it later.
        Private Function schedule(ByRef firstDate As Date?) As IScheduler
            Dim sch As IScheduler = checkSchedulerObject()

            Dim trg = QuartzUtils.GetTrigger(TriggerId)

            If Not IsNothing(trg) Then Return sch

            Dim tempTrg As Quartz.ITrigger = Me.InnerTrigger

            Dim jobdetails As IJobDetail = setJobDetails()

            Try
                firstDate = sch.ScheduleJob(jobdetails, InnerTrigger).DateTime
                Return sch
            Catch ex As Exception
                Dim justLog As Boolean = True
                If Debugger.IsAttached Then
                    Debugger.Break()
                    justLog = False
                End If
                Balora.Alerter.REP(ex.Message,
                                   ex,
                                   False,
                                   justLog)
                Return Nothing
            End Try
        End Function

        Private Function setJobDetails() As IJobDetail
            Dim jobdetails As IJobDetail = Nothing
            Try

                'AhSaElDin 20120922: JobName & TriggerName both will be equal to TriggerId.
                'Becasue TriggerId is variable, not constant as Taskid, so we can bind the Same
                'Task to a lot of triggers.
                Dim tempTrg As Quartz.ITrigger = Me.InnerTrigger
                jobdetails = JobBuilder.Create(Of QuartzJob)().WithIdentity(Me.TriggerId, tempTrg.Key.Group).Build()
                jobdetails.JobDataMap.Add(Me.Task.TaskId, Me.Task)
                jobdetails.JobDataMap.Add(Me.TriggerId, Me)

            Catch ex As Exception
                If Debugger.IsAttached Then Debugger.Break()
                Debug.Print("")
            End Try
            Return jobdetails
        End Function

        Private Function _startListening() As Date?
            Dim _date As Date
            Dim sch As IScheduler = schedule(_date)
            If IsNothing(sch) Then Return Nothing
            setTriggerListener(sch)
            sch.Start()
            Cacadore.Settings.Scheduler = sch
            Watcher.AddScheduledTrigger(Me)
            Watcher.AddNewEntry(Me.Task.TaskId,
                                Me.Task.Properties.TaskName,
                                IHistoryEntry.HistoryEntryType.Trigger,
                                Description,
                                DateTime.Now,
                                Cacadore.ITaskState.ExecutionState.Scheduled)
            quartzTrigger_Scheduled(Me.Task, New EventArgs)
            Return _date
        End Function

        Private Sub setTriggerListener(ByVal sch As IScheduler)
            TriggerListnerName = DateTime.Now.ToBinary
            'To enable ITriggerListener events below.
            ' Me.InnerTrigger.AddTriggerListener(TriggerListnerName)
            sch.ListenerManager.AddTriggerListener(Me)
        End Sub

        Private Function checkSchedulerObject() As IScheduler
            Dim sch As IScheduler = Nothing
            If Not IsNothing(Cacadore.Settings.Scheduler) Then
                sch = Cacadore.Settings.Scheduler
            Else
                Dim sf As ISchedulerFactory = New StdSchedulerFactory
                Cacadore.Settings.Scheduler = sf.GetScheduler()
                sch = Cacadore.Settings.Scheduler
            End If
            Return sch
        End Function

        Private Function getNextFireTime() As Nullable(Of DateTime)
            Dim nxtFireTimeUTC As Date?
            If IsNothing(InnerTrigger) Then Return Nothing
            Select Case TriggerType
                Case Cacadore.ITrigger.TriggerTypeEnum.SimpleTrigger
                    Dim trg As ISimpleTrigger = InnerTrigger
                    If trg.RepeatInterval.TotalSeconds = 0 And trg.RepeatCount = 0 Then Return Nothing
                    If PreviousFireTimeUTC.HasValue Then nxtFireTimeUTC = PreviousFireTimeUTC.Value.Add(trg.RepeatInterval)
                Case Cacadore.ITrigger.TriggerTypeEnum.CronTrigger
                    Dim trg As ICronTrigger = InnerTrigger
                    Dim croXGen As New CronExpression(trg.CronExpressionString)
                    nxtFireTimeUTC = croXGen.GetNextValidTimeAfter(DateTime.UtcNow).Value.DateTime
            End Select
            Return nxtFireTimeUTC
        End Function

        Private Sub setCurrentTriggerId()
            If Me.Task.State.GetExecutionState() = ITaskState.ExecutionState.Running Then
                Dim tempTrg As Quartz.ITrigger = Me.InnerTrigger
                Me.Task.State.CurrentTriggerId = tempTrg.Key.Name
            Else
                Me.Task.State.CurrentTriggerId = ""
            End If
        End Sub

        Private Sub updateNextTriggerId(ByVal nextTriggersUTC As SortedList(Of Date, String))
            If nextTriggersUTC.Count > 0 Then
                Dim trgId As String = nextTriggersUTC.Values.First
                If QuartzHelper.QuartzUtils.WillTriggerFireAgain(_innerTriggerValue.Key) Then
                    Me.Task.State.NextTriggerId = nextTriggersUTC.Values.First 'First item is the nearest trigger id.
                End If
            End If
        End Sub

        Private Sub setPreviousTriggerId(ByVal previousTriggerUTC As SortedList(Of Date, String))
            If previousTriggerUTC.Count > 0 Then Me.Task.State.PreviousTriggerId = previousTriggerUTC.Values.Last 'Last item is the latest trigger id.
        End Sub

        Private Sub updateTaskState(Optional save As Boolean = False)
            setCurrentTriggerId()
            Watcher.Watch(Task)
        End Sub

        Friend Function prepareNotifationMessage(message As String) As ArrayList
            Dim messageList As New ArrayList
            Dim tempTrg As Quartz.ITrigger = Me.InnerTrigger
            messageList.Add(tempTrg.Key.Name)
            messageList.Add(message)
            Return messageList
        End Function

        Public Overrides Sub PauseTrigger()
            Dim sch As IScheduler = Cacadore.Settings.Scheduler

            Dim tempTrg As Quartz.ITrigger = Me.InnerTrigger

            Dim jbKey As New JobKey(Me.Task.TaskId, tempTrg.Key.Group)
            Dim trgKey As New TriggerKey(tempTrg.Key.Name, tempTrg.Key.Group)

            sch.PauseJob(jbKey)
            sch.PauseTrigger(trgKey)
        End Sub

        Public Overrides Sub ResumeTrigger()
            Dim sch As IScheduler = Cacadore.Settings.Scheduler

            Dim tempTrg As Quartz.ITrigger = Me.InnerTrigger

            Dim jbKey As New JobKey(Me.Task.TaskId, tempTrg.Key.Group)
            Dim trgKey As New TriggerKey(tempTrg.Key.Name, tempTrg.Key.Group)

            sch.ResumeJob(jbKey)
            sch.ResumeTrigger(trgKey)
        End Sub

        Private Sub quartzTrigger_Scheduled(Sender As Cacadore.Task, e As System.EventArgs)
            updateTaskState()
            Me.Task.UpdateProperties(PreviousFireTimeUTC, Me.Task.State.TaskPreviousExecutionState.ToString, NextFireTimeUTC) 'Dangerous Tectonic Trip
            Me.Task.Notify(prepareNotifationMessage("Scheduled"))
            Me.Task.Save() 'Dangerous Tectonic Trip
        End Sub
#End Region

#Region "Implemented Interfaces"
#Region "ITriggerListener Impelmention"
        Private ReadOnly Property name As String Implements Quartz.ITriggerListener.Name
            Get
                Return TriggerListnerName
            End Get
        End Property


        Public Sub TriggerComplete(trigger As Quartz.ITrigger, context As Quartz.IJobExecutionContext, triggerInstructionCode As Quartz.SchedulerInstruction) Implements Quartz.ITriggerListener.TriggerComplete
            IsExecuting = False
        End Sub

        Public Sub TriggerFired(trigger As Quartz.ITrigger, context As Quartz.IJobExecutionContext) Implements Quartz.ITriggerListener.TriggerFired
            IsExecuting = True
            'What will happen if I could not delete this? Notify وهل ليها لزمة وهي اصلا موجودة في
            Watcher.AddNewEntry(Me.Task.TaskId,
                                Me.Task.Properties.TaskName,
                                IHistoryEntry.HistoryEntryType.Trigger,
                                "Task Fired.",
                                DateTime.Now,
                                Me.Task.State.GetExecutionState()) 'Dangerous Tectonic Trip

            If NextFireTimeUTC.HasValue AndAlso PreviousFireTimeUTC.HasValue Then
                Dim preFireTime As Date = PreviousFireTimeUTC.Value.ToLocalTime
                Dim nextFireTime As Date = NextFireTimeUTC.Value.ToLocalTime
                Me.Task.UpdateProperties(preFireTime, ITaskState.ExecutionState.Done, nextFireTime)
                Me.Task.Save()
            End If
        End Sub

        Public Sub TriggerMisfired(trigger As Quartz.ITrigger) Implements Quartz.ITriggerListener.TriggerMisfired
            RaiseEvent MisFired(Task, Me)
            Select Case MisFireCommand
                Case Cacadore.ITrigger.MisFireCommandsEnum.RunNow
                    Watcher.AddNewEntry(Me.Task.TaskId,
                                        Me.Task.Properties.TaskName,
                                        IHistoryEntry.HistoryEntryType.Trigger,
                                        "Misfired and will be executed now.",
                                        DateTime.Now,
                                        Me.Task.State.GetExecutionState())
                Case Cacadore.ITrigger.MisFireCommandsEnum.AskMe
                    Watcher.AddNewEntry(Me.Task.TaskId,
                                        Me.Task.Properties.TaskName,
                                        IHistoryEntry.HistoryEntryType.Trigger,
                                        "Misfired. Asking to be executed now...",
                                        DateTime.Now,
                                        Me.Task.State.GetExecutionState())
            End Select
        End Sub

        Public Function VetoJobExecution(trigger As Quartz.ITrigger, context As Quartz.IJobExecutionContext) As Boolean Implements Quartz.ITriggerListener.VetoJobExecution
            'Returns true if job execution should be vetoed, false otherwise.
            Dim sch As IScheduler = Settings.Scheduler
            Dim trgState = sch.GetTriggerState(trigger.Key)
            If trgState = TriggerState.Paused Then
                Console.WriteLine("Trigger is in a PAUSED state, vetoing job!")
                Return True
            Else
                Return False
            End If
            Return False
        End Function
#End Region
#End Region

#Region "Events"
        Public Event Fired(Sender As Cacadore.Task, e As System.EventArgs)
        Public Event MisFired(Sender As Cacadore.Task, trigger As QuartzTrigger)
#End Region

#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.

                _innerTriggerValue = Nothing
                _calendarCValue = Nothing
            End If
            Me.disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        Protected Overrides Sub Finalize()
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(False)
            MyBase.Finalize()
        End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region
    End Class
End Namespace

#Region "References"
'DateTimePicker.CustomFormat Property 
'	http://msdn.microsoft.com/en-us/library/system.windows.forms.datetimepicker.customformat.aspx
'	http://stackoverflow.com/questions/93472/datetimepicker-pick-both-date-and-time

'Lesson 6: Cron Trigger:
'	http://quartznet.sourceforge.net/tutorial/lesson_6.html

'Cron Trigger Class:
'	http://quartznet.sourceforge.net/apidoc

'Cron Trigger Tutorial:
'   http://www.quartz-scheduler.org/documentation/quartz-1.x/tutorials/crontrigger

'Cron Expression helper:
'   http://www.abunchofutils.com/utils/developer/cron-expression-helper

'Cron, crontab generator:
'	http://bit.ly/rWuPHj

'More intelligent calendaring options:
'   http://www.winautomation.com/forum/feature-requests

'Quartz.Impl.Calendar Namespace
'   http://quartznet.sourceforge.net/apidoc/topic2144.html

'*.Examine the 15 examples of quartz.net
'   Example 2 all about simple triggers
'   Example 3 all about cron triggers
'   Example 4 jobdatamap
'   Example 5 misfire example
'   Example 6 excpetions
'   Example 7 InterruptableJob
'   Example 8 calendars
'   Example 9 joblisteners
'   Example 10,11 load & plugins exampls
'   Example 12 server
'   Example 13 clustering
'   Example 14 priorty
'   Example 15 xml example
#End Region
