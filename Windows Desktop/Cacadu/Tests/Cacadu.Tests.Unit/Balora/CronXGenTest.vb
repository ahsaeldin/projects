Imports Balora
Imports NUnit.Framework

Namespace BaloraUT
    <TestFixture()> _
    Public Class CronXGenTest

        <TestFixtureSetUp()> _
        Public Sub SetupMethods()
            PrepareCacaduComponents()
        End Sub

        <TestFixtureTearDown()> _
        Public Sub TearDownMethods()
        End Sub

        <SetUp()> _
        Public Sub SetupTest()
        End Sub

        <TearDown()> _
        Public Sub TearDownTest()
        End Sub

#Region "Methods"

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

            Assert.AreEqual(cronExpression.Keys.Contains("12 11 10 ? JUN,OCT 3#1 2012"), True)

        End Sub

#End Region


    End Class
End Namespace