Public Class CroXGen
    Inherits BaloraBase

    Structure SimpleCron
        Dim n As Integer
        Dim unit As TimeUnit
    End Structure

#Region "Generate Full Expressions Methods"
    ''' <summary>
    ''' Generates the every N time unit cron expression.
    ''' </summary>
    ''' <remarks>
    ''' Generates cron expression for repeat every n seconds/minutes/hours.
    ''' </remarks>
    Public Shared Function GenerateEveryNTimeunitExpression(n As Integer,
                                                            unit As TimeUnit,
                                                            Optional startDate As Date? = Nothing) As Dictionary(Of String, String)
        Dim expr As String
        Dim returnedExpr As New Dictionary(Of String, String)
        Select Case unit
            Case TimeUnit.Seconds
                If Not startDate.HasValue Then
                    expr = "0/" & n & " * * ? * *"
                Else
                    expr = startDate.Value.Second.ToString & "/" & n & " * * ? * *"
                End If
            Case TimeUnit.Minutes
                If Not startDate.HasValue Then
                    expr = "0 0/" & n & " * 1/1 * ? *"
                Else
                    expr = "0 " & startDate.Value.Minute.ToString & "/" & n & " * 1/1 * ? *"
                End If
            Case TimeUnit.Hours
                If Not startDate.HasValue Then
                    expr = "0 0 0/" & n & " 1/1 * ? *"
                Else
                    expr = "0 0 " & startDate.Value.Hour.ToString & "/" & n & " 1/1 * ? *"
                End If

            Case Else
                expr = ""
        End Select
        returnedExpr.Add(expr, "Every " & n & " " & unit.ToString)
        Return returnedExpr
    End Function

    ''' <summary>
    ''' Generates a run once expression.
    ''' </summary>
    Public Shared Function GenerateRunOnceExpression(utcDateTime As DateTime) As Dictionary(Of String, String)
        Dim second As String = utcDateTime.Second.ToString
        Dim minute As String = utcDateTime.Minute.ToString
        Dim hour As String = utcDateTime.Hour.ToString
        Dim dayOfMonth As String = utcDateTime.Day.ToString
        Dim month As String = utcDateTime.Month.ToString
        Dim dayOfWeek As String = "?"
        Dim year As String = utcDateTime.Year.ToString

        Dim expr As String = second & " " &
                             minute & " " &
                             hour & " " &
                             dayOfMonth & " " &
                             month & " " &
                             dayOfWeek & " " & year

        'We commented the "year", because it raises "Based on configured 
        'schedule, the given trigger will never fire" error for unknown reason. 
        Dim returnedExpr As New Dictionary(Of String, String)
        returnedExpr.Add(expr, "Run once at " & utcDateTime.ToLocalTime)
        Return returnedExpr
    End Function

    ''' <summary>
    ''' Generates the complex cron expressions.
    ''' </summary>
    ''' <param name="times">The times.</param>
    ''' <param name="dayOfMonth">The day of month.</param>
    ''' <param name="months">The months.</param>
    ''' <param name="dayOfWeek">The day of week. (You can add a unique week day and if ordinal, you can add only one ordinal, logical huh?!)</param>
    ''' <param name="years">The years.</param><returns></returns>
    Public Shared Function GenerateComplexExpressions(times As SortedSet(Of DateTime),
                                                      dayOfMonth As SortedSet(Of Integer),
                                                      months As SortedSet(Of Months),
                                                      dayOfWeek As SortedList(Of WeekDays, String),
                                                      years As SortedSet(Of Integer)) As Dictionary(Of String, String)
        Dim cronExprsList As New Dictionary(Of String, String)
        Dim dayOfMonthField As String = BuildDaysOfMonthField(dayOfMonth)
        Dim monthsField As String = BuildMonthsField(months)
        Dim dayOfWeekField As String = BuildDayOfWeekField(dayOfWeek)
        Dim yearsField As String = BuildYearsField(years)
        AdjustDayOfWeekAndDayOfMonthFields(dayOfMonthField, dayOfWeekField)
        BuildComplexCronExpressionsList(times, cronExprsList, dayOfMonthField, monthsField, dayOfWeekField, yearsField)
        Return cronExprsList
    End Function

#End Region

#Region "Build Fields Methods"
    Private Shared Sub BuildComplexCronExpressionsList(ByVal times As SortedSet(Of DateTime),
                                                       ByVal cronExprsList As Dictionary(Of String, String),
                                                       ByVal dayOfMonthField As String,
                                                       ByVal monthsField As String,
                                                       ByVal dayOfWeekField As String,
                                                       ByVal yearsField As String)
        'لو فاضي يبقي كل الأوقات
        If times.Count = 0 Then
            Dim _timeField As String = BuildTimeField(Nothing)
            Dim expr As String = _timeField & dayOfMonthField & monthsField & dayOfWeekField & yearsField
            Dim cronExpressionDescription As String = BuildCronExpressionDescription("*",
                                                                                     dayOfMonthField,
                                                                                     monthsField,
                                                                                     dayOfWeekField,
                                                                                     yearsField)
            cronExprsList.Add(expr, cronExpressionDescription)
        Else
            For Each _time In times
                Dim _timeField As String = BuildTimeField(_time)
                Dim expr As String = _timeField & dayOfMonthField & monthsField & dayOfWeekField & yearsField
                Dim cronExpressionDescription As String = BuildCronExpressionDescription(_time.ToLongTimeString,
                                                                                         dayOfMonthField,
                                                                                         monthsField,
                                                                                         dayOfWeekField,
                                                                                         yearsField)
                cronExprsList.Add(expr, cronExpressionDescription)
            Next
        End If
    End Sub

#Region "Fields Methods"
    ''' <summary>
    ''' Generates the time field only of cron expression with the trailing space at end.
    ''' </summary>
    ''' <param name="dTime">The d time.</param><returns></returns>
    Private Shared Function BuildTimeField(dTime As DateTime?) As String
        Dim timeField As String = vbNullString
        'لو فاضي يبقي كل الأوقات
        If Not dTime.HasValue Then
            timeField = "*" & " " &
                        "*" & " " &
                        "*" & " "
        Else
            'دة السيناريو العادي.
            timeField = dTime.Value.Second & " " &
            dTime.Value.Minute & " " &
            dTime.Value.Hour & " "
        End If
        Return timeField
    End Function

    ''' <summary>
    ''' Builds the days of month field with a trailing space at end.
    ''' </summary>
    Private Shared Function BuildDaysOfMonthField(dayOfMonth As SortedSet(Of Integer)) As String
        Dim daysOfMonthField As String = vbNullString
        If IsNothing(dayOfMonth) OrElse dayOfMonth.Count = 0 Then
            Return "* "
        End If
        For Each _day In dayOfMonth
            If Not dayOfMonth.Last = _day Then
                daysOfMonthField = daysOfMonthField & _day & ","
            Else
                daysOfMonthField = daysOfMonthField & _day & " "
            End If
        Next
        Return daysOfMonthField
    End Function

    ''' <summary>
    ''' Builds the months field with a trailing space at end.
    ''' </summary>
    ''' <param name="months">The months.</param><returns></returns>
    Private Shared Function BuildMonthsField(months As SortedSet(Of Months)) As String
        Dim monthsField As String = vbNullString
        If IsNothing(months) OrElse months.Count = 0 Then
            Return "* "
        End If
        For Each _month In months
            If Not months.Last = _month Then
                monthsField = monthsField & _month.ToString & ","
            Else
                monthsField = monthsField & _month.ToString & " "
            End If
        Next
        Return monthsField
    End Function

    ''' <summary>  
    ''' Builds the day of week field with a trailing space at end.
    ''' </summary>
    ''' <param name="days">The days.</param><returns></returns>
    Private Shared Function BuildDayOfWeekField(days As SortedList(Of WeekDays, String)) As String
        Dim dayOfWeekField As String = vbNullString
        If IsNothing(days) OrElse days.Count = 0 Then
            Return "? "
        End If
        For Each _day In days
            Dim dayValue As String
            If _day.Key = WeekDays.Ordinal Then
                dayValue = _day.Value
            Else
                dayValue = CStr(_day.Key.ToString)
            End If
            If Not days.Last.Value = _day.Value Then
                dayOfWeekField = dayOfWeekField & dayValue & ","
            Else
                dayOfWeekField = dayOfWeekField & dayValue & " "
            End If
        Next
        Return dayOfWeekField
    End Function

    ''' <summary>
    ''' Builds the years field with no trailing space @ end.
    ''' </summary>
    ''' <param name="years">The years.</param><returns></returns>
    Private Shared Function BuildYearsField(years As SortedSet(Of Integer)) As String
        Dim yearsField As String = vbNullString
        If IsNothing(years) OrElse years.Count = 0 Then
            Return "*"
        End If
        For Each _year In years
            If Not years.Last = _year Then
                yearsField = yearsField & _year & ","
            Else
                yearsField = yearsField & _year
            End If
        Next
        Return yearsField
    End Function
#End Region

#Region "Build Fields Descriptions Methods"
    Private Shared Function BuildCronExpressionDescription([time] As String,
                                                           dayOfMonth As String,
                                                           month As String,
                                                           dayOfWeek As String,
                                                           years As String) As String
        Dim exprDesc As String = [time]
        BuildDayOfWeekDescription(dayOfWeek, exprDesc)
        BuildDayOfMonthDescription(dayOfMonth, exprDesc)
        BuildMonthDescription(month, exprDesc)
        BuildYearsDescription(years, exprDesc)
        ReplaceDoubleSpaceWithASpace(exprDesc)
        Return exprDesc
    End Function

#Region "Fields Description Methods"
    Private Shared Sub BuildDayOfWeekDescription(ByVal dayOfWeek As String,
                                                 ByRef exprDesc As String)
        If exprDesc = "*" Then
            exprDesc = "Every second, "
        End If

        If dayOfWeek.Contains("*") Then
            exprDesc = exprDesc & " at all week days "
        ElseIf dayOfWeek.Contains("?") Then

        ElseIf dayOfWeek.Contains("#") Then
            exprDesc = exprDesc & " at the " & BuildOrdinalDayDescription(dayOfWeek) & " every month, "
        Else
            Dim dayOfWeekList As List(Of String) = dayOfWeek.Split(CChar(",")).ToList()
            Dim completeDayOfWeek As String = vbNullString
            Dim _completeDayOfWeekField As String = Nothing
            For Each _day In dayOfWeekList
                Dim _DayName As String = CompleteDayName(_day.Trim)
                _completeDayOfWeekField = _completeDayOfWeekField & _DayName & ","
            Next
            'remove the ending ","
            _completeDayOfWeekField = Strings.Left(_completeDayOfWeekField, Len(_completeDayOfWeekField) - 1)
            exprDesc = exprDesc & " every " & _completeDayOfWeekField
        End If
    End Sub

    Private Shared Sub BuildDayOfMonthDescription(ByVal dayOfMonth As String,
                                                  ByRef exprDesc As String)
        If dayOfMonth.Contains("*") Then
            exprDesc = exprDesc & " every day"
        ElseIf dayOfMonth.Contains("?") Then

        Else
            exprDesc = exprDesc & " at day(s) " & dayOfMonth
        End If
    End Sub

    Private Shared Sub BuildMonthDescription(ByVal month As String,
                                             ByRef exprDesc As String)
        If month <> "" AndAlso Not month.Contains("*") Then
            exprDesc = exprDesc & " in the month(s) " & month
        End If
    End Sub

    Private Shared Sub BuildYearsDescription(ByVal years As String,
                                             ByRef exprDesc As String)
        If years <> "" AndAlso Not years.Contains("*") Then
            exprDesc = exprDesc & " during the year(s) " & years
        End If
    End Sub

    Private Shared Sub ReplaceDoubleSpaceWithASpace(ByRef exprDesc As String)
        exprDesc = exprDesc.Replace("  ", " ")
    End Sub

    Private Shared Function BuildOrdinalDayDescription(dayOfWeek As String) As String
        dayOfWeek = dayOfWeek.Trim
        Dim dayName As String = CType(Left(dayOfWeek, 1), WeekDays).ToString
        dayName = CompleteDayName(dayName.Trim)
        Dim ordinal As String = CType(Right(dayOfWeek, 1), OrdinalNumbers).ToString
        Return ordinal & " " & dayName
    End Function
#End Region
#End Region
#End Region

#Region "Helper Methods"
    Public Shared Function FormatExpressionDescrption(description As String,
                                                      startDate As Date?,
                                                      endDate As Date?,
                                                      isRunOnce As Boolean) As String
        description = description.Replace("Seconds", "second(s)")
        description = description.Replace("Minutes", "minute(s)")
        description = description.Replace("Hours", "hour(s)")

        If startDate.HasValue Then
            If Not isRunOnce Then
             description = String.Format("{0} starting from {1}", description, startDate.Value.ToLocalTime())
            End If
        End If

        If endDate.HasValue And Not isRunOnce Then
            description = String.Format("{0} and ending at {1}", description, endDate.Value.ToLocalTime)
        End If

        Return description
    End Function

    Private Shared Function CompleteDayName(shortDayName As String) As String
        Select Case shortDayName
            Case "Sun", "SUN"
                Return "Sunday"
            Case "Mon", "MON"
                Return "Monday"
            Case "Tue", "TUE"
                Return "Tuesday"
            Case "Wed", "WED"
                Return "Wednesday"
            Case "Thu", "THU"
                Return "Thursday"
            Case "Fri", "FRI"
                Return "Friday"
            Case "Sat", "SAT"
                Return "Saturday"
            Case Else
                Return ""
        End Select
    End Function

#Region "Checks Methods"
    ''' <summary>
    ''' Adjusts the day of week and day of month fields.
    ''' </summary>
    ''' <param name="dayOfMonthField">The day of month field.</param>
    ''' <param name="dayOfWeekField">The day of week field.</param>
    ''' <remarks>
    ''' Because the limitation of the current version of Quartz, 
    ''' we must adjust these 2 fields because there is no support 
    ''' for specifying both a day-of-week and a day-of-month value 
    ''' is not complete (you must currently use the '?' character 
    ''' in one of these fields).
    ''' </remarks> 
    Private Shared Sub AdjustDayOfWeekAndDayOfMonthFields(ByRef dayOfMonthField As String,
                                                          ByRef dayOfWeekField As String)

        Dim isDayOfMonthHasSpecificValue As Boolean = If(dayOfMonthField <> "? " AndAlso
                                                         dayOfMonthField <> "* " AndAlso
                                                         dayOfMonthField <> vbNullString, True, False)

        Dim isDayOfMonthHasNoSpecificValue As Boolean = If(dayOfMonthField = "? ", True, False)

        Dim isDayOfMonthHasAllValue As Boolean = If(dayOfMonthField = "* ", True, False)

        'Checks against dayOfMonthField & dayOfWeekField...
        Dim isDayOfWeekHasSpecifiedValue As Boolean = If(dayOfWeekField <> "? " AndAlso
                                                         dayOfWeekField <> "* " AndAlso
                                                         dayOfWeekField <> vbNullString, True, False)

        Dim isDayOfWeekHasNoSpecificValue As Boolean = If(dayOfWeekField = "? ", True, False)

        Dim isDayOfWeekHasAllValue As Boolean = If(dayOfWeekField = "* ", True, False)

        If isDayOfMonthHasSpecificValue AndAlso isDayOfWeekHasSpecifiedValue Then 'Like: 1 MON become ?,MON
            dayOfMonthField = "? "
        ElseIf isDayOfMonthHasNoSpecificValue AndAlso isDayOfWeekHasNoSpecificValue Then 'Like: ? ? become * ?
            dayOfMonthField = "* "
        ElseIf isDayOfMonthHasAllValue AndAlso isDayOfWeekHasAllValue Then 'Like: * MON become ? MON
            dayOfMonthField = "? "
        ElseIf isDayOfMonthHasAllValue AndAlso isDayOfWeekHasNoSpecificValue Then 'Like: 1 * become 1 ?
            dayOfWeekField = "? "
        ElseIf isDayOfMonthHasAllValue AndAlso isDayOfWeekHasSpecifiedValue Then 'Like: * MON,FRI become ? MON,FRI
            dayOfMonthField = "? "
        End If

    End Sub
#End Region
#End Region

#Region "Parsing Methods"
    Public Shared Function ParseSimpleCronExperssion(expr As String, ByRef exprStruct As SimpleCron) As Boolean
        Return IsSimpleCronExperssion(expr, exprStruct)
    End Function

    Public Shared Function IsSimpleCronExperssion(expr As String, ByRef exprStruct As SimpleCron) As Boolean
        Dim exprArr = expr.Split(CChar(" "))
        If exprArr(0).Contains("/") Then
            If exprArr(1) = "*" Then
                If exprArr(2) = "*" Then
                    If exprArr(3) = "?" Then
                        If exprArr(4) = "*" Then
                            If exprArr(5) = "*" Then
                                exprStruct.n = extractNTimes(exprArr(0))
                                exprStruct.unit = TimeUnit.Seconds
                                Return True
                            End If
                        End If
                    End If
                End If
            End If
        End If

        If exprArr(0) = "0" Then
            If exprArr(1).Contains("/") Then
                If exprArr(2) = "*" Then
                    If exprArr(3) = "1/1" Then
                        If exprArr(4) = "*" Then
                            If exprArr(5) = "?" Then
                                If exprArr(6) = "*" Then
                                    exprStruct.n = extractNTimes(exprArr(1))
                                    exprStruct.unit = TimeUnit.Minutes
                                    Return True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        If exprArr(0) = "0" Then
            If exprArr(1) = "0" Then
                If exprArr(2).Contains("/") Then
                    If exprArr(3) = "1/1" Then
                        If exprArr(4) = "*" Then
                            If exprArr(5) = "?" Then
                                If exprArr(6) = "*" Then
                                    exprStruct.n = extractNTimes(exprArr(2))
                                    exprStruct.unit = TimeUnit.Hours
                                    Return True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return False
    End Function

    Private Shared Function extractNTimes(str As String) As Integer
        Dim whereSlash As Integer = str.IndexOf("/")
        Dim result = str.Split(CChar("/"))
        Return CInt(result.Last())
    End Function
#End Region
End Class
