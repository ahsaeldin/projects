Option Strict Off

Imports Cacadore
Imports Balora
Imports Helpers
Imports NUnit.Framework
Imports Balora.PathsHelper
Imports Helpers.QuartzHelper

Namespace CacaduIT
    <TestFixture()> _
    <Serializable()> Public Class TriggersTest
#Region "Fields"
        'Must be declared as shared, because for unknown reason, nunit/resharper 
        'internally creates a new instance from this class which clears all fields.
        <NonSerialized()> Private Shared isAnExcludingHolidaysTest As Boolean = False
        <NonSerialized()> Private Shared isDeSerializingCalenderObjectTest As Boolean = False
        <NonSerialized()> Private excludeMinuate As Integer = 0
        <NonSerialized()> Private startDateEdit As New DevExpress.XtraEditors.DateEdit()
        <NonSerialized()> Private endDateEdit As New DevExpress.XtraEditors.DateEdit()
        '<NonSerialized()> Private substracedThreeSecondsForProcessing As Integer = -3
        <NonSerialized()> Private isFiredBefore As Boolean
        <NonSerialized()> Private oldActualFireDateTime As New DateTime
        <NonSerialized()> Private actualFireDateTime As New DateTime
        <NonSerialized()> Private firstExpectedFireTime As New DateTime
        <NonSerialized()> Private cancelTask As Boolean = False
        <NonSerialized()> Private pauseTask As Boolean = False
        'Private _useService As Boolean = True
        <NonSerialized()> Private _tskMeta As New TaskMeta
        <NonSerialized()> Private WithEvents tsk As Task
        'Private WithEvents _namedPipesClient As Balora.Ipc.NamedPipes.PipeClient
        <NonSerialized()> Private _executionGate As New Threading.ManualResetEvent(False)
        <NonSerialized()> Private _commonAppDataPath As String = GetCommonApplicationData() & "Cacadu\"
        <NonSerialized()> Private _tectonicPath As String = _commonAppDataPath & "tectonic"
        <NonSerialized()> Private _tectonic_copyPath As String = _commonAppDataPath & "tectonic_copy"
        <NonSerialized()> Private _resulting_tectonicPath As String = _commonAppDataPath & "resulting_tectonic"
        <NonSerialized()> Private _isTestRunTaskAtSelectedTimes As Integer = 0
        <NonSerialized()> Private CacadoreHelperObj As CacadoreHelper
        <NonSerialized()> Private CacadoreToTectonic As CacadoreToTectonic
        <NonSerialized()> Private ErrorReportingUI As Balora.ErrorReporter.ErrorReportingForm
#End Region

#Region "Setups & TearDowns"
        <TestFixtureSetUp()> _
        Public Sub SetupMethods()
            Unit.PrepareCacaduComponents()
            'Make a backup copy of tectonic before playing with the orignial. 
            'IO.File.Delete(_tectonic_copyPath)
            'IO.File.Copy(_tectonicPath, _tectonic_copyPath)

            '_namedPipesClient = New Balora.Ipc.NamedPipes.PipeClient(My.Resources.NPN, IO.Pipes.PipeOptions.Asynchronous)

            'If _useService Then
            '    Dim servnstTestObj As New Balora.BaloraUnitTests.ServnstTest
            '    If Servnst.IsServiceInstalled(My.Resources.CSN) Then servnstTestObj.UninstallServiceTest()
            '    servnstTestObj.SetupMethods()
            '    servnstTestObj.DEBUG_StartService()
            'End If
            CacadoreHelperObj = New CacadoreHelper
            CacadoreToTectonic = New CacadoreToTectonic
            ErrorReportingUI = New Balora.ErrorReporter.ErrorReportingForm
        End Sub

        <TestFixtureTearDown()> _
        Public Sub TearDownMethods()
            SendMessageToConsole("Tearing down all tests...")
            'If _useService Then
            '    Dim servnstTestObj As New Balora.BaloraUnitTests.ServnstTest
            '    If Servnst.IsServiceInstalled(My.Resources.CSN) Then servnstTestObj.UninstallServiceTest()
            'End If

            'Re-copy the original.
            'IO.File.Copy(_tectonicPath, _resulting_tectonicPath, True)
            'IO.File.Copy(_tectonic_copyPath, _tectonicPath, True)
            'IO.File.Delete(_tectonic_copyPath)
        End Sub

        <SetUp()> _
        Public Sub SetupTest()
            isAnExcludingHolidaysTest = False
            isDeSerializingCalenderObjectTest = False
            Try
                If Not TestContext.CurrentContext.Test.Name.Contains("TestTriggerMisFire") Then
                    _tskMeta = New TaskMeta
                    tsk = CreateNewTask(_tskMeta) 'Parameter is out and ByRef.
                End If
            Catch ex As Exception
                'Added try-catch to avoid raisded exception when using Reshaprter
                'Reshaper has a problem with TestContext class
                _tskMeta = New TaskMeta
                tsk = CreateNewTask(_tskMeta) 'Parameter is out and ByRef.
            End Try
            cancelTask = False
            pauseTask = False
            firstExpectedFireTime = New DateTime
            actualFireDateTime = New DateTime
            isFiredBefore = False
            oldActualFireDateTime = New DateTime
            startDateEdit = New DevExpress.XtraEditors.DateEdit()
            endDateEdit = New DevExpress.XtraEditors.DateEdit()
            _isTestRunTaskAtSelectedTimes = 0
        End Sub

        <TearDown()> _
        Public Sub TearDownTest()
        End Sub
#End Region

#Region "Tests"
#Region "Simple Scheduling"
        ''' <summary>
        ''' Repeat task every 5 seconds.
        ''' </summary>
        <Test()> _
        <Category("ShortRunning")>
        Public Sub TestEvery5Seconds()
            Dim everySeconds As Integer = 5
            SendMessageToConsole("Starting TestEvery5Seconds test...")
            SendMessageToConsole("This test set a trigger to run your task every " & everySeconds & " seconds.")
            Assert.AreEqual(True, SimpleScheduleTest(everySeconds, TimeUnit.Seconds))
        End Sub

        ''' <summary>
        ''' Repeat task every 1 minutes.
        ''' </summary>
        <Test()> _
        <Category("ShortRunning")>
        Public Sub TestEvery1Minuates()
            SendMessageToConsole("Starting TestEvery1Minuates test...")
            Assert.AreEqual(True, SimpleScheduleTest(1, TimeUnit.Minutes))
        End Sub

        ''' <summary>
        ''' Repeat task 1 hour.
        ''' </summary>
        <Test()> _
        <Category("LongRunning")>
        Public Sub TestEvery1Hour()
            SendMessageToConsole("Starting TestEvery2Hours...")
            Assert.AreEqual(True, SimpleScheduleTest(1, TimeUnit.Hours))
        End Sub

        ''' <summary>
        ''' Tests the run task once @ prespecified start time.
        ''' </summary>
        ''' <remarks>You only need start time.</remarks> 
        <Test()> _
        <Category("ShortRunning")>
        Public Sub TestRunTaskOnce()
            Dim actual As Boolean = False
            Dim expected As Boolean = True
            Dim startDateEdit As New DevExpress.XtraEditors.DateEdit()
            Dim startDateTime As New DateTime
            startDateEdit.Properties.Mask.EditMask = "G"
            startDateEdit.DateTime = DateTime.Now.AddSeconds(60)
            startDateTime = startDateEdit.DateTime
            firstExpectedFireTime = Quartz.TriggerUtils.GetEvenMinuteDate(startDateTime)
            Dim startDateTimeUTC = startDateTime.ToUniversalTime
            Dim cronExpression As Dictionary(Of String, String) = Balora.CroXGen.GenerateRunOnceExpression(startDateTime)
            Assert.AreEqual(True, WaitForTrigger(startDateEdit.DateTime, startDateTimeUTC, Nothing, cronExpression.First.Key, cronExpression.First.Value))
        End Sub
#End Region

#Region "Common Schedule Tests"
        ''' <summary>
        ''' Test canceling a task
        ''' </summary>
        <Test()> _
        <Category("Test solo")>
        Public Sub TestTaskCancel()
            SendMessageToConsole("Starting TestTaskCancel test...")
            cancelTask = True
            Assert.AreEqual(True, SimpleScheduleTest(1, TimeUnit.Seconds))
        End Sub

        ''' <summary>
        ''' Test canceling a task
        ''' </summary>
        <Test()> _
        <Category("Test solo")>
        Public Sub TestPauseResumeTask()
            SendMessageToConsole("Starting TestPauseResumeTask test...")
            pauseTask = True
            Assert.AreEqual(True, SimpleScheduleTest(1, TimeUnit.Minutes))
        End Sub

        <Test()> _
        <Category("Test solo And Watch Results")>
        Public Sub TestTriggerMisFire()

            'Get all registered triggers.
            Dim allTriggers As Dictionary(Of String, Trigger) = CacadoreToTectonic.GetAllTriggers()
            'Get all triggers that has an inner trigger.
            Dim kvp = From trg In allTriggers.Values Where Not IsNothing(trg.InnerTrigger)
            'If there is no trigger with inner trigger, run any test to create one.
            If kvp.Count = 0 Then
                Assert.Inconclusive("Run any test to create an inner trigger first.")
            End If
            'Get first one of them.
            Dim qutrg As QuartzHelper.QuartzTrigger = CType(kvp.First, QuartzTrigger)
            AddHandler qutrg.MisFired, AddressOf trg_MisFired
            'Start selected past trigger and wait for misFire.
            qutrg.StartListening()
            _executionGate.Reset()
            _executionGate.WaitOne()
        End Sub

        <Test()> _
        <Category("Test solo And Watch Results")>
        Public Sub TestWaitBetweenActions()
            'This test is just exactly like
            tsk.WaitBetweenActions = True
            TestEvery1Minuates()
        End Sub
#End Region

#Region "Advanced Scheduling"
        ''' <summary>
        ''' Tests the run task at selected times. like 07:30:00, 07:33:33, 07:55:44.
        ''' </summary>
        <Test()> _
        <Category("LongRunning")>
        Public Sub TestRunTaskAtCertinTime()
            Dim startDateTime As DateTime = DateTime.Now.AddSeconds(15)
            firstExpectedFireTime = Quartz.TriggerUtils.GetEvenMinuteDate(startDateTime)
            Dim startDateTimeUTC = startDateTime.ToUniversalTime

            Dim date1 As DateTime = startDateTime.AddMinutes(1)
            Dim date2 As DateTime = startDateTime.AddHours(1).AddMinutes(5).AddSeconds(3)
            Dim date3 As DateTime = startDateTime.AddMinutes(7).AddSeconds(2)
            firstExpectedFireTime = date1 'It is a specific so no need to call evenXXX methods.

            SendMessageToConsole("date1 = " & date1)
            SendMessageToConsole("date2 = " & date2)
            SendMessageToConsole("date3 = " & date3)

            SendMessageToConsole("FirstExpectedFireTime = " & firstExpectedFireTime)
            SendMessageToConsole("SecondExpectedFireTime = " & Quartz.TriggerUtils.GetEvenSecondDate(date2))
            SendMessageToConsole("ThirdExpectedFireTime = " & Quartz.TriggerUtils.GetEvenSecondDate(date3))

            Dim sortedDateTimes As New SortedSet(Of DateTime)
            sortedDateTimes.Add(date2)
            sortedDateTimes.Add(date3)
            sortedDateTimes.Add(date1)

            Dim cronExpression As Dictionary(Of String, String) = CroXGen.GenerateComplexExpressions(sortedDateTimes,
                                                                                                     Nothing,
                                                                                                     Nothing,
                                                                                                     Nothing,
                                                                                                     Nothing)
            For Each expr In cronExpression
                CreateTrigger(startDateTime, startDateTimeUTC, Nothing, expr.Key, expr.Value)
            Next

            _executionGate.Reset() : _executionGate.WaitOne()
        End Sub

        <Test()> _
        <Category("LongRunning")>
        Public Sub TestRunTaskAtCertinDaysOfMonth()
            Dim startDateTime As DateTime = DateTime.Now.AddSeconds(15)
            Dim startDateTimeUTC = startDateTime.ToUniversalTime

            Dim dayOfMonth As Integer = Now.Day
            Dim monthName As Months = CType(Now.Month, Months)

            Dim timeToFire As DateTime = DateTime.Now.AddMinutes(5)

            actualFireDateTime = Quartz.TriggerUtils.GetEvenMinuteDate(timeToFire)

            SendMessageToConsole("FirstExpectedFireTime = " & firstExpectedFireTime)

            Dim sortedDateTimes As New SortedSet(Of DateTime)
            Dim sortedDaysOfMonth As New SortedSet(Of Integer)
            Dim months As New SortedSet(Of Months)
            Dim years As New SortedSet(Of Integer)

            sortedDateTimes.Add(timeToFire)
            sortedDaysOfMonth.Add(dayOfMonth)
            months.Add(monthName)
            years.Add(2012)
            Dim cronExpression As Dictionary(Of String, String) = CroXGen.GenerateComplexExpressions(sortedDateTimes,
                                                                                           sortedDaysOfMonth,
                                                                                           months,
                                                                                           Nothing,
                                                                                           years)
            CreateTrigger(startDateTime, startDateTimeUTC, Nothing, cronExpression.First.Key, cronExpression.First.Value)
            _executionGate.Reset() : _executionGate.WaitOne()
        End Sub

        ''' <summary>
        ''' Tests the run task at tomorrow and after Tomorrow @ 9:55:10 am, 10:30:09am.
        ''' </summary>
        <Test()> _
        <Category("VeryLongRunning")>
        Public Sub TestRunTaskAtCertinDays()
            '9:55:10 am, 10:30:09am @ today & after today
            Dim tomorrow As WeekDays = QuartzHelper.QuartzUtils.GetTomorrow
            Dim afterTomorrow As WeekDays = QuartzHelper.QuartzUtils.GetAfterTomorrow

            Dim startDateTime As DateTime = DateTime.Now
            Dim endDateTime As DateTime = DateTime.Now.AddDays(3)
            Dim startDateTimeUTC = startDateTime.ToUniversalTime
            Dim endDateTimeUTC As DateTime = endDateTime.ToUniversalTime

            Dim date1 As DateTime = #9:55:10 AM#
            Dim date2 As DateTime = #10:30:09 AM#
            firstExpectedFireTime = New DateTime(Now.Year, Now.Month, Now.AddDays(1).Day, date1.Hour, date1.Minute, date1.Second)

            SendMessageToConsole("date1 = " & date1 & " @ " & tomorrow.ToString & " and " & afterTomorrow.ToString)
            SendMessageToConsole("date2 = " & date2 & " @ " & tomorrow.ToString & " and " & afterTomorrow.ToString)

            SendMessageToConsole("FirstExpectedFireTime = " & firstExpectedFireTime)
            SendMessageToConsole("SecondExpectedExpectedFireTime = " &
                                 New DateTime(Now.Year, Now.Month, Now.AddDays(1).Day, date2.Hour, date2.Minute, date2.Second))
            SendMessageToConsole("ThirdExpectedExpectedFireTime = " &
                                 New DateTime(Now.Year, Now.Month, Now.AddDays(2).Day, date1.Hour, date1.Minute, date1.Second))
            SendMessageToConsole("FourthExpectedFireTime = " &
                                 New DateTime(Now.Year, Now.Month, Now.AddDays(2).Day, date2.Hour, date2.Minute, date2.Second))

            Dim sortedDateTimes As New SortedSet(Of DateTime)
            sortedDateTimes.Add(date1)
            sortedDateTimes.Add(date2)

            Dim sortedDays As New SortedList(Of WeekDays, String)
            sortedDays.Add(tomorrow, tomorrow.ToString)
            sortedDays.Add(afterTomorrow, afterTomorrow.ToString)

            Dim cronExpression As Dictionary(Of String, String) = Balora.CroXGen.GenerateComplexExpressions(sortedDateTimes,
                                                                                                  Nothing,
                                                                                                  Nothing,
                                                                                                  sortedDays,
                                                                                                  Nothing)
            For Each expr In cronExpression
                CreateTrigger(startDateTime, startDateTimeUTC, endDateTimeUTC, expr.Key, expr.Value)
            Next

            _executionGate.Reset() : _executionGate.WaitOne()
        End Sub

#End Region

#Region "CronExpression Generator (CronXGen) Tests"
        ''' <summary>
        ''' Tests generating a complex cron expressions with day of week field.
        ''' </summary>
        <Test()> _
        <Category("ShortRunning")>
        Public Sub TestGenerateComplexCronExpressions()
            Dim times As New SortedSet(Of DateTime)
            Dim daysOfMonth As New SortedSet(Of Integer)
            Dim months As New SortedSet(Of Months)
            Dim daysOfWeek As New SortedList(Of WeekDays, String)
            Dim years As New SortedSet(Of Integer)
            times.Add(#10:11:12 AM#)
            times.Add(#11:12:13 AM#)

            'daysOfMonth.Add(1)
            'daysOfMonth.Add(25)
            months.Add(PublicEnums.Months.JUN)
            months.Add(PublicEnums.Months.OCT)
            daysOfWeek.Add(WeekDays.MON, WeekDays.MON.ToString)
            daysOfWeek.Add(WeekDays.WED, WeekDays.WED.ToString)
            years.Add(2012)
            years.Add(2015)
            Dim cronExpression As Dictionary(Of String, String) = CroXGen.GenerateComplexExpressions(times,
                                                 daysOfMonth,
                                                 months,
                                                 daysOfWeek,
                                                 years)
            For Each expr In cronExpression
                CreateTrigger(Now, Now.ToUniversalTime, Nothing, expr.Key, expr.Value)
            Next

            Assert.AreEqual(cronExpression.Keys.Contains("12 11 10 ? JUN,OCT MON,WED 2012,2015"), True)
            Assert.AreEqual(cronExpression.Keys.Contains("13 12 11 ? JUN,OCT MON,WED 2012,2015"), True)
        End Sub

        ''' <summary>
        ''' Tests generating a complex cron expressions without day of week field.
        ''' </summary>
        <Test()> _
        <Category("ShortRunning")>
        Public Sub TestGenerateComplexCronExpressionsWithOutWeekDays()
            Dim times As New SortedSet(Of DateTime)
            Dim daysOfMonth As New SortedSet(Of Integer)
            Dim months As New SortedSet(Of Months)
            Dim daysOfWeek As New SortedList(Of WeekDays, String)
            Dim years As New SortedSet(Of Integer)
            times.Add(#10:11:12 AM#)
            times.Add(#11:12:13 AM#)

            daysOfMonth.Add(1)
            daysOfMonth.Add(25)

            months.Add(PublicEnums.Months.JUN)
            months.Add(PublicEnums.Months.OCT)

            'daysOfWeek.Add(WeekDays.MON)
            'daysOfWeek.Add(WeekDays.WED)

            years.Add(2012)
            years.Add(2015)

            Dim cronExpression As Dictionary(Of String, String) = CroXGen.GenerateComplexExpressions(times,
                                                                                                     daysOfMonth,
                                                                                                     months,
                                                                                                     daysOfWeek,
                                                                                                     years)
            For Each expr In cronExpression
                CreateTrigger(Now, Now.ToUniversalTime, Nothing, expr.Key, expr.Value)
            Next

            Assert.AreEqual(cronExpression.Keys.Contains("12 11 10 1,25 JUN,OCT ? 2012,2015"), True)
            Assert.AreEqual(cronExpression.Keys.Contains("13 12 11 1,25 JUN,OCT ? 2012,2015"), True)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub TestGenerateComplexCronExprWithOrdinalDayOfWeekNumber()
            Dim times As New SortedSet(Of DateTime)
            Dim daysOfMonth As New SortedSet(Of Integer)
            Dim months As New SortedSet(Of Months)
            Dim daysOfWeek As New SortedList(Of WeekDays, String)
            Dim years As New SortedSet(Of Integer)

            times.Add(#10:11:12 AM#)
            times.Add(#11:12:13 AM#)

            months.Add(PublicEnums.Months.JUN)
            months.Add(PublicEnums.Months.OCT)

            daysOfWeek.Add(WeekDays.Ordinal, "3#1")

            years.Add(2012)

            Dim cronExpression As Dictionary(Of String, String) = CroXGen.GenerateComplexExpressions(times,
                                                                                           daysOfMonth,
                                                                                           months,
                                                                                           daysOfWeek,
                                                                                           years)

            For Each expr In cronExpression
                CreateTrigger(Now, Now.ToUniversalTime, Nothing, expr.Key, expr.Value)
            Next

            Assert.AreEqual(cronExpression.Keys.Contains("12 11 10 ? JUN,OCT 3#1 2012"), True)
        End Sub
#End Region

#Region "Calendar Testing"
        <Test()> _
        <Category("Test solo")>
        Public Sub TestExcludeHolidays()
            isAnExcludingHolidaysTest = True

            Dim minuatlyCronExpr As Dictionary(Of String, String) = Balora.CroXGen.GenerateEveryNTimeunitExpression(1, TimeUnit.Minutes)
            Dim _excluedCronExpression As Dictionary(Of String, String) = Balora.CroXGen.GenerateEveryNTimeunitExpression(2, TimeUnit.Minutes)

            firstExpectedFireTime = Quartz.TriggerUtils.GetEvenMinuteDate(DateTime.Now)
            Dim excludeDateTime = firstExpectedFireTime.AddMinutes(2).ToUniversalTime
            excludeMinuate = firstExpectedFireTime.Minute + 2

            SendMessageToConsole("FirstExpectedFireTime = " & firstExpectedFireTime)
            SendMessageToConsole("Excluded Holiday Minute = " & excludeMinuate)

            Dim quTrg = CreateTrigger(Now,
                                      Quartz.TriggerUtils.GetEvenMinuteDate(DateTime.UtcNow),
                                      DateTime.UtcNow.AddMinutes(4),
                                      minuatlyCronExpr.First.Key,
                                      minuatlyCronExpr.First.Value,
                                      _excluedCronExpression.First)

            Dim croCal As Quartz.Impl.Calendar.CronCalendar = CType(quTrg.Calendar, 
                                                                    Quartz.Impl.Calendar.CronCalendar)

            Dim result = croCal.IsTimeIncluded(excludeDateTime)

            Assert.IsTrue(result, "Time was included when it was not supposed to be")
        End Sub

        ''' <summary>
        ''' Testing deSerializing Calendar from Triggers table then checks it Calendar still works
        ''' </summary>
        <Test()> _
        <Category("Test solo")>
        Public Sub TestDeSerializingCalenderObject()
            isDeSerializingCalenderObjectTest = True

            'Get all registered triggers.
            Dim allTriggers As Dictionary(Of String, Trigger) = CacadoreToTectonic.GetAllTriggers()

            'Get all triggers that has an inner trigger.
            Dim kvp = From trg In allTriggers.Values Where Not IsNothing(trg.Calendar)

            'Get first one of them.
            Dim qutrg As QuartzHelper.QuartzTrigger = CType(kvp.First, QuartzTrigger)

            If IsNothing(qutrg) Then
                Assert.Inconclusive("Cannot find a calendar in Triggers table.")
            End If

            firstExpectedFireTime = Quartz.TriggerUtils.GetEvenMinuteDate(DateTime.Now)
            Dim excludeDateTime As DateTime = firstExpectedFireTime.AddMinutes(2).ToUniversalTime()

            excludeMinuate = excludeDateTime.Minute

            SendMessageToConsole("FirstExpectedFireTime = " & firstExpectedFireTime)
            SendMessageToConsole("Excluded Holiday Minute = " & excludeMinuate)

            'Get Calendar
            Dim croCal As Quartz.Impl.Calendar.CronCalendar = CType(qutrg.Calendar, 
                                                                    Quartz.Impl.Calendar.CronCalendar)

            Dim result = croCal.IsTimeIncluded(excludeDateTime)

            Assert.IsTrue(result, "Time was included when it was not supposed to be")
        End Sub
#End Region
#End Region

#Region "Helper Methods"

        Private Function CreateTrigger(ByVal startDateTime As Date,
                               ByVal startDateTimeUTC As Date?,
                               ByVal endDateTimeUTC As Date?,
                               ByVal cronExpression As String,
                               ByVal cronExpressionDescription As String,
                               Optional holiday As KeyValuePair(Of String, String) = Nothing) As QuartzHelper.QuartzTrigger
            SendMessageToConsole("A new cron expression created: " & cronExpression)

            _tskMeta.StartDateTimeUTC = startDateTimeUTC
            _tskMeta.EndDateTimeUTC = endDateTimeUTC
            _tskMeta.CronExpression = cronExpression
            _tskMeta.CalendarCronExpression = holiday.Key
            _tskMeta.CalendarCronExpressionDescription = holiday.Value

            If _tskMeta.StartDateTimeUTC.HasValue Then
                cronExpressionDescription = cronExpressionDescription & " starting from " & startDateTimeUTC.Value.ToLocalTime
            End If
            If _tskMeta.EndDateTimeUTC.HasValue Then
                cronExpressionDescription = cronExpressionDescription & " and ending at " & endDateTimeUTC.Value.ToLocalTime
            End If

            _tskMeta.CronExpressionDescription = cronExpressionDescription

            SendMessageToConsole("Creating and attaching a new trigger to the task...")
            Dim qutrg As QuartzHelper.QuartzTrigger = CacadoreHelperObj.AttachScheduledTriggerToTask(tsk, _tskMeta)
            SendMessageToConsole("Trigger: " & _tskMeta.LastAttachedTriggerId & " created and scheduled.")

            qutrg.StartListening()

            SendMessageToConsole("And now trigger is waiting to be triggered @start time=" & startDateTime)
            SendMessageToConsole("And @firstExpectedFireTime=" & firstExpectedFireTime)
            If endDateTimeUTC.HasValue Then
                SendMessageToConsole("And @endDateTime = " & endDateTimeUTC.Value.ToLocalTime)
            Else
                SendMessageToConsole("And @endDateTime = " & "n/a")
            End If
            If qutrg.NextFireTimeUTC.HasValue Then
                SendMessageToConsole("qutrg.NextFireTimeUTC = " & qutrg.NextFireTimeUTC.Value.ToLocalTime.ToLongTimeString)
            End If

            Return qutrg
        End Function

        Private Function WaitForTrigger(ByVal startDateTime As Date,
                                       ByVal startDateTimeUTC As Date?,
                                       ByVal endDateTimeUTC As Date?,
                                       ByVal cronExpression As String,
                                       ByVal cronExpressionDescription As String) As Boolean
            CreateTrigger(startDateTime, startDateTimeUTC, endDateTimeUTC, cronExpression, cronExpressionDescription)
            SendMessageToConsole("Waiting for the task to start its execution...")
            _executionGate.Reset() : _executionGate.WaitOne()
            SendMessageToConsole("Test Done, Congrts :D")
            Return True
        End Function

        ''' <summary>
        ''' Schedule a job for running starting @startDateTime 
        ''' variable and ending @endDateTime, with repetition 
        ''' defined according to n and unit parameter. 
        ''' </summary>
        ''' <param name="n">The n.</param>
        ''' <param name="unit">The unit.</param>
        Private Function SimpleScheduleTest(n As Integer, unit As TimeUnit) As Boolean
            Dim actual As Boolean = False
            Dim expected As Boolean = True

            startDateEdit.Properties.Mask.EditMask = "G"
            endDateEdit.Properties.Mask.EditMask = "G"

            Select Case unit
                Case TimeUnit.Seconds
                    startDateEdit.DateTime = DateTime.Now.AddSeconds(15)
                    endDateEdit.DateTime = DateTime.Now.AddSeconds(30)
                    firstExpectedFireTime = Quartz.TriggerUtils.GetEvenSecondDate(startDateEdit.DateTime)
                Case TimeUnit.Minutes
                    startDateEdit.DateTime = DateTime.Now.AddMinutes(0.1)
                    endDateEdit.DateTime = DateTime.Now.AddMinutes(5.2)
                    firstExpectedFireTime = Quartz.TriggerUtils.GetEvenMinuteDate(startDateEdit.DateTime)
                Case TimeUnit.Hours
                    startDateEdit.DateTime = DateTime.Now.AddMinutes(0.4)
                    endDateEdit.DateTime = DateTime.Now.AddHours(3.1)
                    firstExpectedFireTime = Quartz.TriggerUtils.GetEvenHourDate(startDateEdit.DateTime)
            End Select

            Dim startDateTime = startDateEdit.DateTime
            Dim endDateTime = endDateEdit.DateTime
            Dim startDateTimeUTC = startDateEdit.DateTime.ToUniversalTime
            Dim endDateTimeUTC = endDateEdit.DateTime.ToUniversalTime

            Dim cronExpression As Dictionary(Of String, String) = Balora.CroXGen.GenerateEveryNTimeunitExpression(n, unit)

            Return WaitForTrigger(startDateTime, startDateTimeUTC, endDateTimeUTC, cronExpression.First.Key, cronExpression.First.Value)
        End Function

        Public Function CreateNewTask(ByRef tskMeta As TaskMeta) As Task
            'Parameter is out and byref.
            With tskMeta
                .BreakOnFailed = True
                .WaitBetweenActions = False
                .IsEnabled = True
                .Name = "Unit Tested Task"
                .ParentGroupId = CStr(0)
            End With

            Dim createdTask As Task = CacadoreHelperObj.CreateTestableTask(tskMeta)

            tskMeta.TaskId = createdTask.TaskId

            Return createdTask
        End Function

        Public Sub SendMessageToConsole(msg As String)
            Console.WriteLine(DateTime.Now.ToLongTimeString & ":" & msg)
        End Sub
#End Region

#Region "Task Events"

        Private Sub tsk_Fired(sender As Cacadore.Task, e As Cacadore.ITask.ExecutionEventArgs) Handles tsk.Fired
            actualFireDateTime = DateTime.Now

            If Not isFiredBefore Then

                If actualFireDateTime.ToString = firstExpectedFireTime.ToString Then
                    SendMessageToConsole(
                          "actualFireDateTime (" & actualFireDateTime & ") = firstExpectedFireTime(" & firstExpectedFireTime & ")")
                Else
                    SendMessageToConsole(
                        "actualFireDateTime (" & actualFireDateTime & ") & is not equal to firstExpectedFireTime(" & firstExpectedFireTime & ")")
                End If
            Else
                Dim interval As TimeSpan = actualFireDateTime - oldActualFireDateTime
                SendMessageToConsole("New actualFireDateTime = " & actualFireDateTime)
                SendMessageToConsole("Total seconds between ticking: " & interval.TotalSeconds)
                SendMessageToConsole("Total minutes between ticking: " & interval.TotalMinutes)
                SendMessageToConsole("Total hours between ticking: " & interval.TotalHours)
            End If

            If Not isFiredBefore Then
                isFiredBefore = True
            End If

            oldActualFireDateTime = actualFireDateTime

            Dim qutrg As QuartzHelper.QuartzTrigger = CType(e.trigger, QuartzHelper.QuartzTrigger)

            If qutrg.NextFireTimeUTC.HasValue Then
                SendMessageToConsole("qutrg.NextFireTimeUTC = " & qutrg.NextFireTimeUTC.Value.ToLocalTime.ToLongTimeString)
            End If
        End Sub

        Private Sub tsk_Before_Executing(sender As Cacadore.Task, e As System.EventArgs) Handles tsk.Before_Executing
            SendMessageToConsole("Just before executing task...")
            If isAnExcludingHolidaysTest Or isDeSerializingCalenderObjectTest Then
                If DateTime.Now.Minute = excludeMinuate Then
                    Assert.Inconclusive("Calendar Test Failed")
                End If
            End If
        End Sub

        Private Sub tsk_Canceled(sender As Cacadore.Task, e As System.EventArgs) Handles tsk.Canceled
            SendMessageToConsole("Task Canceled")
        End Sub

        Private Sub tsk_Executed(sender As Cacadore.Task, e As Cacadore.ITask.ExecutionEventArgs) Handles tsk.Executed
            SendMessageToConsole("Task executed.")
            If Not sender.MayFiredAgain Then
                If Not IsNothing(_executionGate) Then _executionGate.Set()
                Dim qutrg As QuartzHelper.QuartzTrigger = CType(e.trigger, QuartzHelper.QuartzTrigger)
                If Not IsNothing(qutrg) Then
                    If qutrg.NextFireTimeUTC.HasValue Then
                        SendMessageToConsole("qutrg.NextFireTimeUTC = " & qutrg.NextFireTimeUTC.Value.ToLocalTime.ToLongTimeString)
                    End If
                End If
            Else
                _isTestRunTaskAtSelectedTimes = _isTestRunTaskAtSelectedTimes + 1
                If _isTestRunTaskAtSelectedTimes = 2 Then
                    _executionGate.Set()
                End If
            End If
        End Sub

        Private Sub tsk_Executing(sender As Cacadore.Task, e As System.ComponentModel.ProgressChangedEventArgs) Handles tsk.Executing
            SendMessageToConsole("Task is executing...")
            Dim currentState = sender.State.GetExecutionState()
            Dim ssssss = sender.State.TaskExecutionState
            If cancelTask And currentState = ITaskState.ExecutionState.Running Then
                sender.Cancel()
                _executionGate.Reset() : _executionGate.WaitOne()
            ElseIf pauseTask Then
                sender.Pause()
                pauseTask = False
            End If
        End Sub

        Private Sub tsk_Paused(sender As Cacadore.Task, e As System.EventArgs) Handles tsk.Paused
            SendMessageToConsole("Task Paused.")
            sender.[Resume]()
            _executionGate.Set()
        End Sub

        Private Sub tsk_Resumed(sender As Cacadore.Task, e As System.EventArgs) Handles tsk.Resumed
            SendMessageToConsole("Task Resumed")
        End Sub

        Private Sub tsk_Scheduled(sender As Cacadore.Task, e As Cacadore.ITask.ScheduleEventArgs) Handles tsk.Scheduled
            SendMessageToConsole("Task Scheduled")
        End Sub

        Sub trg_MisFired(Sender As Cacadore.Task, trigger As QuartzTrigger)
            SendMessageToConsole("MisFired event raised")
            _executionGate.Set()
        End Sub
#End Region

    End Class
End Namespace
