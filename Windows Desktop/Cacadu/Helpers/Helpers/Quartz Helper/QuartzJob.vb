Imports Cacadore

'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   QuartzJob Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
Namespace QuartzHelper
    ''' <summary>
    ''' QuartzJob is a helper class, you can use it as a quartz job passed to 
    ''' job details when we assign MyTask property to the desired task you want 
    ''' to execute.
    ''' </summary>
    Friend Class QuartzJob
        Implements Quartz.IJob
      
        Public Sub Execute(context As Quartz.IJobExecutionContext) Implements Quartz.IJob.Execute
            Dim tempArr = context.JobDetail.Key.Name.Split("-")

            Dim taskId = tempArr(0)
            Dim trgId As String = context.JobDetail.Key.Name
            Dim myJobDataMap As Quartz.JobDataMap = context.JobDetail.JobDataMap
            Dim quartzTrigger As QuartzTrigger = CType(myJobDataMap(trgId), QuartzTrigger)
            Dim taskToBeExecuted As Cacadore.Task = CType(myJobDataMap(taskId), Cacadore.Task)

            If context.ScheduledFireTimeUtc.Value.Date.CompareTo(context.FireTimeUtc.Value.Date) < 0 Then
                'This is a misfire
                Select Case quartzTrigger.MisFireCommand
                    Case ITrigger.MisFireCommandsEnum.AskMe
                        Balora.Alerter.SendMessageToCacadu("AskUserForMisFire", taskToBeExecuted)
                    Case ITrigger.MisFireCommandsEnum.RunNow
                        'This won't be handled because it called automatically.
                    Case ITrigger.MisFireCommandsEnum.Skip
                        Watcher.AddNewEntry(taskToBeExecuted.TaskId,
                                            taskToBeExecuted.Properties.TaskName,
                                            IHistoryEntry.HistoryEntryType.Trigger,
                                            "Misfired, and will be skipped.",
                                            DateTime.Now,
                                            taskToBeExecuted.State.GetExecutionState())
                        Exit Sub
                End Select
            Else
                taskToBeExecuted.Execute()
            End If
            myJobDataMap(taskId) = Nothing
        End Sub
    End Class
End Namespace


