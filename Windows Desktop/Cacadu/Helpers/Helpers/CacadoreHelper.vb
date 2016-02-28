'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   CacadoreHelper Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/

Imports Quartz
Imports Cacadore
Imports Helpers.QuartzHelper

Public Class CacadoreHelper

#Region "Methods"
    Public Function CreateTestableTask(tskData As TaskMeta) As Task
        Dim tsk As Task = BIT.CreateBunchOfRunApplicationTasks()
        tsk.BreakOnFailed = tskData.BreakOnFailed
        tsk.WaitBetweenActions = tskData.WaitBetweenActions
        tsk.Properties.IsEnabled = tskData.IsEnabled
        tsk.Properties.TaskName = tskData.Name
        tskData.Dispose()
        tskData = Nothing
        Return tsk
    End Function

    Private Function createTriggerObject(quTrg As QuartzHelper.QuartzTrigger,
                                         tskMeta As TaskMeta,
                                         parentGroupId As String) As Quartz.ITrigger
        Dim repeatCount As Integer = tskMeta.RepeatCount
        Dim cronTrg As Quartz.ITrigger = Nothing
        If tskMeta.TriggerType = Cacadore.ITrigger.TriggerTypeEnum.CronTrigger Then
            cronTrg = QuartzUtils.CreateCronTrigger(quTrg.TriggerId,
                                                    parentGroupId,
                                                    quTrg.TriggerId,
                                                    tskMeta.StartDateTimeUTC,
                                                    tskMeta.EndDateTimeUTC,
                                                    tskMeta.CronExpression,
                                                    tskMeta.MisFire)
        Else
            If tskMeta.RunOnce Then
                cronTrg = QuartzUtils.CreateRunOnceTrigger(quTrg.TriggerId,
                                                           parentGroupId,
                                                           quTrg.TriggerId,
                                                           tskMeta.StartDateTimeUTC,
                                                           tskMeta.EndDateTimeUTC,
                                                           tskMeta.MisFire)
                Return cronTrg
            End If
            Dim parsedExpr As Balora.CroXGen.SimpleCron
            Balora.CroXGen.IsSimpleCronExperssion(tskMeta.CronExpression, parsedExpr)
            Select Case parsedExpr.unit
                Case Balora.TimeUnit.Seconds
                    cronTrg = QuartzUtils.CreateSimpleTrigger(quTrg.TriggerId,
                                                              parentGroupId,
                                                              quTrg.TriggerId,
                                                              tskMeta.StartDateTimeUTC,
                                                              tskMeta.EndDateTimeUTC,
                                                              repeatCount,
                                                              TimeSpan.FromSeconds(parsedExpr.n),
                                                              parsedExpr.unit,
                                                              tskMeta.MisFire)
                Case Balora.TimeUnit.Minutes
                    cronTrg = QuartzUtils.CreateSimpleTrigger(quTrg.TriggerId,
                                                              parentGroupId,
                                                              quTrg.TriggerId,
                                                              tskMeta.StartDateTimeUTC,
                                                              tskMeta.EndDateTimeUTC,
                                                              repeatCount,
                                                              TimeSpan.FromMinutes(parsedExpr.n),
                                                              parsedExpr.unit,
                                                              tskMeta.MisFire)
                Case Balora.TimeUnit.Hours
                    cronTrg = QuartzUtils.CreateSimpleTrigger(quTrg.TriggerId,
                                                              parentGroupId,
                                                              quTrg.TriggerId,
                                                              tskMeta.StartDateTimeUTC,
                                                              tskMeta.EndDateTimeUTC,
                                                              repeatCount,
                                                              TimeSpan.FromHours(parsedExpr.n),
                                                              parsedExpr.unit,
                                                              tskMeta.MisFire)
            End Select
        End If
        Return cronTrg
    End Function

    Private Function createNewInnerTrigger(trgType As Cacadore.ITrigger.TriggerTypeEnum) As Quartz.ITrigger

        If trgType = Cacadore.ITrigger.TriggerTypeEnum.CronTrigger Then
            'Added next to line just never save a blank values for Task and Trigger in Triggers table,
            'as AttachTrigger commit directly to Tectonic before collecting them.
            'If next 2 lines make problems in the future delete them, however you will get this error...
            '"RaiseErrorUp::Task is nothing, Can't update trigger"
            '0 0/3 * 1/1 * ? * == any dumpy cron experssion.
            Return Quartz.TriggerBuilder.Create().WithCronSchedule("0 0/3 * 1/1 * ? *").Build()
        Else
            Return Quartz.TriggerBuilder.Create().WithSimpleSchedule().Build()
        End If
    End Function

    Private Sub setCalender(ByRef quTrg As QuartzTrigger, tskMeta As TaskMeta)
        If tskMeta.CalendarCronExpression <> "" Then
            Dim sch As Quartz.IScheduler = Cacadore.Settings.Scheduler
            quTrg.Description = String.Format("{0} Except ({1})", quTrg.Description, tskMeta.CalendarCronExpressionDescription)
            Dim croCal As New Quartz.Impl.Calendar.CronCalendar(tskMeta.CalendarCronExpression)
            quTrg.Calendar = croCal
            Dim calName As String = quTrg.TriggerId & "Calendar"
            sch.AddCalendar(calName, quTrg.Calendar, True, True)
            quTrg.InnerTrigger.CalendarName = calName
        End If
    End Sub

    Public Function AttachScheduledTriggerToTask(tsk As Task, ByRef tskMeta As TaskMeta) As QuartzHelper.QuartzTrigger
        If IsNothing(tsk) Then Return Nothing

        Dim quTrg As New QuartzHelper.QuartzTrigger

        'To avoid error while cruding @ AttachTrigger
        quTrg.InnerTrigger = createNewInnerTrigger(tskMeta.TriggerType)

        quTrg.Task = tsk

        tsk.AttachTrigger(quTrg)

        quTrg.Description = tskMeta.CronExpressionDescription

        setCalender(quTrg, tskMeta)

        Dim cronTrg As Quartz.ITrigger = createTriggerObject(quTrg, tskMeta, tsk.ParentGroupId)

        If IsNothing(cronTrg) Then Return Nothing

        tskMeta.CronExpressionDescription = Balora.CroXGen.FormatExpressionDescrption(tskMeta.CronExpressionDescription,
                                                                                      tskMeta.StartDateTimeUTC,
                                                                                      tskMeta.EndDateTimeUTC,
                                                                                      tskMeta.RunOnce)
        tsk.Properties.TriggersWhen = tskMeta.CronExpressionDescription

        With quTrg
            .MisFireCommand = tskMeta.MisFire
            .InnerTrigger = cronTrg
            tsk.MyTriggers.Triggers(quTrg.TriggerId).InnerTrigger = cronTrg
            .Update(tsk)
        End With

        tskMeta.LastAttachedTriggerId = quTrg.TriggerId

        Cacadore.Watcher.Watch(tsk)
        Return quTrg
    End Function
#End Region
End Class

