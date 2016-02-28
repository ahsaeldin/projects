'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   QuartzUtils Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
Imports Balora
Imports Cacadore
Imports Quartz
Imports Balora.CroXGen

Namespace QuartzHelper
    ''' <summary>
    ''' Provide methods to create triggers using Quartz
    ''' </summary>
    Public Class QuartzUtils

#Region "Triggers Creation"
        Public Shared Function CreateCronTrigger(triggerName As String,
                                                 groupName As String,
                                                 jobName As String,
                                                 startDateTimeUtc As DateTime?,
                                                 endDateTimeUtc As DateTime?,
                                                 cronExpression As String,
                                                 misFire As Cacadore.ITrigger.MisFireCommandsEnum,
                                                 Optional calenderName As String = vbNullString) As ICronTrigger
            Dim trgBuilder As TriggerBuilder = TriggerBuilder.Create()

            trgBuilder.WithIdentity(triggerName, groupName)

            trgBuilder.ForJob(jobName, groupName)

            If calenderName <> vbNullString Then trgBuilder.ModifiedByCalendar(calenderName)

            If startDateTimeUtc.HasValue Then trgBuilder.StartAt(startDateTimeUtc)

            If endDateTimeUtc.HasValue Then trgBuilder.EndAt(endDateTimeUtc)

            Dim cronTrg As ICronTrigger = trgBuilder.WithCronSchedule(cronExpression).Build

            Select Case misFire
                Case Cacadore.ITrigger.MisFireCommandsEnum.RunNow
                    cronTrg = trgBuilder.WithCronSchedule(cronExpression, Function(x) x.WithMisfireHandlingInstructionFireAndProceed).Build
                Case Cacadore.ITrigger.MisFireCommandsEnum.Skip, Cacadore.ITrigger.MisFireCommandsEnum.AskMe
                    cronTrg = trgBuilder.WithCronSchedule(cronExpression, Function(x) x.WithMisfireHandlingInstructionDoNothing).Build
            End Select

            Return cronTrg
        End Function

        Public Shared Function CreateSimpleTrigger(triggerName As String,
                                                   groupName As String,
                                                   jobName As String,
                                                   startTimeUTC As Date,
                                                   endTimeUTC As Date?,
                                                   repeatCount As Integer,
                                                   repeatInterval As TimeSpan,
                                                   repeatUnit As TimeUnit,
                                                   misFire As Cacadore.ITrigger.MisFireCommandsEnum) As ISimpleTrigger
            Dim trgBuilder As TriggerBuilder = TriggerBuilder.Create()

            trgBuilder.WithIdentity(triggerName, groupName)

            trgBuilder.StartAt(startTimeUTC)

            trgBuilder.ForJob(jobName, groupName)

            If endTimeUTC.HasValue Then trgBuilder.EndAt(endTimeUTC.Value)

            Dim simTrg As ISimpleTrigger = Nothing

            Select Case repeatUnit

                Case TimeUnit.Seconds

                    Select Case misFire
                        Case Cacadore.ITrigger.MisFireCommandsEnum.RunNow
                            simTrg = trgBuilder.WithSimpleSchedule(Function(x) x.WithIntervalInSeconds(repeatInterval.TotalSeconds).WithMisfireHandlingInstructionFireNow.WithRepeatCount(repeatCount)).Build()
                        Case Cacadore.ITrigger.MisFireCommandsEnum.Skip, Cacadore.ITrigger.MisFireCommandsEnum.AskMe
                            simTrg = trgBuilder.WithSimpleSchedule(Function(x) x.WithIntervalInSeconds(repeatInterval.TotalSeconds).WithMisfireHandlingInstructionIgnoreMisfires.WithRepeatCount(repeatCount)).Build()
                    End Select

                Case TimeUnit.Minutes
                    Select Case misFire
                        Case Cacadore.ITrigger.MisFireCommandsEnum.RunNow
                            simTrg = trgBuilder.WithSimpleSchedule(Function(x) x.WithIntervalInMinutes(repeatInterval.TotalMinutes).WithMisfireHandlingInstructionFireNow.WithRepeatCount(repeatCount)).Build()
                        Case Cacadore.ITrigger.MisFireCommandsEnum.Skip, Cacadore.ITrigger.MisFireCommandsEnum.AskMe
                            simTrg = trgBuilder.WithSimpleSchedule(Function(x) x.WithIntervalInMinutes(repeatInterval.TotalMinutes).WithMisfireHandlingInstructionIgnoreMisfires.WithRepeatCount(repeatCount)).Build()
                    End Select

                Case TimeUnit.Hours
                    Select Case misFire
                        Case Cacadore.ITrigger.MisFireCommandsEnum.RunNow
                            simTrg = trgBuilder.WithSimpleSchedule(Function(x) x.WithIntervalInHours(repeatInterval.TotalHours).WithMisfireHandlingInstructionFireNow.WithRepeatCount(repeatCount)).Build()
                        Case Cacadore.ITrigger.MisFireCommandsEnum.Skip, Cacadore.ITrigger.MisFireCommandsEnum.AskMe
                            simTrg = trgBuilder.WithSimpleSchedule(Function(x) x.WithIntervalInHours(repeatInterval.TotalHours).WithMisfireHandlingInstructionIgnoreMisfires.WithRepeatCount(repeatCount)).Build()
                    End Select

            End Select

            Return simTrg
        End Function

        Public Shared Function CreateRunOnceTrigger(triggerName As String,
                                           groupName As String,
                                           jobName As String,
                                           startTimeUTC As Date,
                                           endTimeUTC As Date?,
                                           misFire As Cacadore.ITrigger.MisFireCommandsEnum) As ISimpleTrigger
            Dim trgBuilder As TriggerBuilder = TriggerBuilder.Create()
            trgBuilder.WithIdentity(triggerName, groupName)
            trgBuilder.StartAt(startTimeUTC)

            If endTimeUTC.HasValue AndAlso endTimeUTC.Value > startTimeUTC Then trgBuilder.EndAt(endTimeUTC.Value)

            Dim simTrg As ISimpleTrigger = Nothing

            Select Case misFire
                Case Cacadore.ITrigger.MisFireCommandsEnum.RunNow, Cacadore.ITrigger.MisFireCommandsEnum.AskMe
                    simTrg = trgBuilder.WithSimpleSchedule(Function(x) x.WithMisfireHandlingInstructionFireNow).Build()
                Case Cacadore.ITrigger.MisFireCommandsEnum.Skip
                    simTrg = trgBuilder.WithSimpleSchedule(Function(x) x.WithMisfireHandlingInstructionIgnoreMisfires).Build()
            End Select

            Return simTrg
        End Function

        ''' <summary>
        ''' Just a copy of Cacadore.Commands.CreateTriggerByType that 
        ''' creates new type of the trigger by specifying ITrigger.TriggerTypeEnum.
        ''' We copy this function here to extend it by adding CronTrigger support,
        ''' and for not to make Cacadore depends on BIT if we support CronTrigger 
        ''' directly in the main method.
        ''' </summary>
        ''' <param name="triggerType">Type of the trigger.</param><returns></returns>
        Public Shared Function CreateTriggerByType(triggerType As Cacadore.ITrigger.TriggerTypeEnum) As Cacadore.ITrigger
            Select Case triggerType
                Case Cacadore.ITrigger.TriggerTypeEnum.DumpTrigger
                    Return New Triggers.DumpTrigger
                Case Cacadore.ITrigger.TriggerTypeEnum.ScheduleTrigger
                    Return New Triggers.ScheduleTrigger
                Case Cacadore.ITrigger.TriggerTypeEnum.CronTrigger
                    Return New QuartzTrigger
                Case Cacadore.ITrigger.TriggerTypeEnum.SimpleTrigger
                    Return New QuartzTrigger
            End Select
            Return Nothing
        End Function
#End Region

#Region "Helpers"
#Region "General Helpers"
        Public Shared Sub ParseSimpleTrigger(trg As QuartzHelper.QuartzTrigger, ByRef parsedExpr As Balora.CroXGen.SimpleCron)
            Dim trgTypeName As String = trg.InnerTrigger.GetType().Name
            If trgTypeName = "CronTriggerImpl" Then Exit Sub
            Dim tempTrg As Quartz.ISimpleTrigger = trg.InnerTrigger
            If trg.Description.Contains("hour") Then
                parsedExpr.n = CInt(tempTrg.RepeatInterval.TotalHours)
                parsedExpr.unit = Balora.TimeUnit.Hours
            ElseIf trg.Description.Contains("minute") Then
                parsedExpr.n = CInt(tempTrg.RepeatInterval.TotalMinutes)
                parsedExpr.unit = Balora.TimeUnit.Minutes
            ElseIf trg.Description.Contains("second") Then
                parsedExpr.n = CInt(tempTrg.RepeatInterval.TotalSeconds)
                parsedExpr.unit = Balora.TimeUnit.Seconds
            End If
        End Sub

        Public Shared Function WillTriggerFireAgain(trgKey As TriggerKey) As Boolean
            Dim sch As IScheduler = Cacadore.Settings.Scheduler
            Dim trgState = sch.GetTriggerState(trgKey)
            If trgState = TriggerState.Normal Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Shared Function GetToday() As WeekDays
            Return DateTime.Now.DayOfWeek + 1
        End Function

        Public Shared Function GetTomorrow() As WeekDays
            Return DateTime.Now.DayOfWeek + 2
        End Function

        Public Shared Function GetAfterTomorrow() As WeekDays
            Return DateTime.Now.DayOfWeek + 3
        End Function
#End Region

#Region "Pause & Resume"
        Public Shared Sub PauseTaskTriggers(taskId As String)
            Dim tsk = Commands.GetTaskById(taskId)
        End Sub

        Public Shared Sub PauseTaskTriggers(tsk As Cacadore.Task)

        End Sub

        Public Shared Sub PauseTrigger(triggerId As String)
            If Not IsTriggerPaused(triggerId) Then
                Dim sch As IScheduler = Cacadore.Settings.Scheduler
                Dim groupName As String = GetTriggerGroupName(triggerId)
                Dim trgKey As New TriggerKey(triggerId, groupName)
                sch.PauseTrigger(trgKey)
            End If
        End Sub

        Public Shared Sub ResumeTrigger(triggerId As String)
            If IsTriggerPaused(triggerId) Then
                Dim sch As IScheduler = Cacadore.Settings.Scheduler
                Dim groupName As String = GetTriggerGroupName(triggerId)
                Dim trgKey As New TriggerKey(triggerId, groupName)
                sch.ResumeTrigger(trgKey)
                Balora.Alerter.SendMessageToCacadu("ClearStatusBarState")
            End If
        End Sub

        Public Shared Sub ResumeAllTriggers()
            Dim sch As IScheduler = Cacadore.Settings.Scheduler
            sch.ResumeAll()
            Balora.Alerter.SendMessageToCacadu("ClearStatusBarState")
        End Sub

        Public Shared Sub PauseAllTriggers()
            Dim sch As IScheduler = Cacadore.Settings.Scheduler
            sch.PauseAll()
            Alerter.SendMessageToCacadu("AllTriggersPaused")
        End Sub

        Public Shared Sub PauseScheduler()
            Dim isch As IScheduler = Cacadore.Settings.Scheduler
            isch.Standby()
        End Sub

        Public Shared Sub ResumeScheduler()
            Dim isch As IScheduler = Cacadore.Settings.Scheduler
            isch.Start()
        End Sub
#End Region

#Region "Trigger State"
        Public Shared Function IsTriggerStarted(quTrg As QuartzTrigger)
            Return Cacadore.Watcher.ScheduledTriggers.ContainsKey(quTrg.TriggerId)
        End Function

        Public Shared Function GetTrigger(triggerId As String) As Quartz.ITrigger
            Dim sch As IScheduler = Cacadore.Settings.Scheduler
            Dim trgGroupName As String = GetTriggerGroupName(triggerId)
            Dim trgKey As New TriggerKey(triggerId, trgGroupName)
            Dim innerTrg As Quartz.ITrigger = sch.GetTrigger(trgKey)
            Return innerTrg
        End Function

        Public Shared Function GetTriggerGroupName(triggerId As String) As String
            Dim trgTask = Cacadore.Commands.GetTriggerTask(triggerId)
            If Not IsNothing(trgTask) Then Return trgTask.ParentGroupId
            Return vbNullString
        End Function

        Public Shared Function GetTriggerState(triggerId As String) As TriggerState
            Dim trgState As TriggerState = Nothing
            Dim sch As IScheduler = Cacadore.Settings.Scheduler
            Dim trgGroupName As String = GetTriggerGroupName(triggerId)
            Dim trgKey As New TriggerKey(triggerId, trgGroupName)
            If trgGroupName <> vbNullString Then trgState = sch.GetTriggerState(trgKey)
            Return trgState
        End Function

        Public Shared Function IsTriggerScheduled(triggerId As String) As Boolean
            Dim trgState As TriggerState = GetTriggerState(triggerId)
            If trgState = TriggerState.Normal Then Return True
            Return False
        End Function

        Public Shared Function IsTriggerUnScheduled(triggerId As String) As Boolean
            Dim trgState As TriggerState = GetTriggerState(triggerId)
            If trgState = TriggerState.None Then Return True
            Return False
        End Function

        Public Shared Function IsTriggerPaused(triggerId As String) As Boolean
            Dim trgState As TriggerState = GetTriggerState(triggerId)
            If trgState = TriggerState.Paused Then Return True
            Return False
        End Function

        Public Shared Function IsThereTriggerPaused() As Boolean
            Dim sch As IScheduler = Cacadore.Settings.Scheduler
            Dim triggersSet = sch.GetPausedTriggerGroups()
            If triggersSet.Count > 0 Then
                Return True
            End If
            Return False
        End Function
#End Region
#End Region
    End Class
End Namespace

