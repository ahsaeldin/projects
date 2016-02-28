Option Strict Off

Imports Balora
Imports Helpers
Imports Cacadu.UI
Imports Cacadu.UI.ScheduleTriggerForm
Imports DevExpress.XtraEditors.Controls
Imports Quartz.Impl.Triggers

Namespace Connectors
    Friend Class ScheduleTriggerFormConnector

#Region "Structures & Enums"
        Structure ParsedComplexCronExpression
            Public Seconds As Quartz.Collection.TreeSet(Of Integer )
            Public Minutes As Quartz.Collection.TreeSet(Of Integer)
            Public Hours As Quartz.Collection.TreeSet(Of Integer)

            Public Days As Quartz.Collection.TreeSet(Of Integer)
            Public Months As Quartz.Collection.TreeSet(Of Integer)
            Public DayOfWeek As Quartz.Collection.TreeSet(Of Integer)
            Public Years As Quartz.Collection.TreeSet(Of Integer)

            Public Ordinal As Integer
            Public HasOrdinal As Integer
            Public OrdinalDayName As Integer
        End Structure
#End Region

#Region "Properties"
        Shared Property ScheduleForm As UI.ScheduleTriggerForm
#End Region

#Region "Form Related"
        Shared Sub LoadScheduleTriggerForm()
            SetIconToCacaduIcon(ScheduleForm)
            SetBoundControls()
            OnExcludeHolidaysCheckBox_CheckedChanged()
            SetDefaultsCombosValues()
        End Sub
#End Region


#Region "BackstageView"
        Shared Function GetSelectedTab(ByVal selectedTabPageName As String) As DevExpress.XtraBars.Ribbon.BackstageViewClientControl
            With ScheduleForm
                If selectedTabPageName = "SimpleScheduleBackstageViewTabItem" Then
                    Return .SimpleScheduleBackstageViewClientControl
                ElseIf selectedTabPageName = "AdvancedScheduleBackstageViewTabItem" Then
                    Return .AdvancedScheduleBackstageViewClientControl
                ElseIf selectedTabPageName = "HolidaysBackstageViewTabItem" Then
                    Return .HolidaysBackstageViewClientControl
                ElseIf selectedTabPageName = "MissFireBackstageViewTabItem" Then
                    Return .MissFireBackstageViewClientControl
                End If
            End With
            Return Nothing
        End Function

        Shared Sub SetSelectedTab(tab As DevExpress.XtraBars.Ribbon.BackstageViewTabItem)
            ScheduleTriggerForm.ScheduleTriggerBackstage.SelectedTab = tab
        End Sub

        Shared Sub CopySimpleControls(selecteBackstageViewClientControl As DevExpress.XtraBars.Ribbon.BackstageViewClientControl)
            With ScheduleForm
                selecteBackstageViewClientControl.Controls.Add(.StartDateTimeLabelControl)
                selecteBackstageViewClientControl.Controls.Add(.StartDateTimeDateEdit)
                selecteBackstageViewClientControl.Controls.Add(.ToLabelControl)
                selecteBackstageViewClientControl.Controls.Add(.EndDateTimeLabelControl)
                selecteBackstageViewClientControl.Controls.Add(.EndDateTimeDateEdit)
            End With
        End Sub

        Shared Sub MoveCommonControlsBetweenTabs(ByVal selectedTabPageName As String)
            With ScheduleForm
                Dim selecteBackstageViewClientControl As DevExpress.XtraBars.Ribbon.BackstageViewClientControl = GetSelectedTab(selectedTabPageName)
                If Not IsNothing(selecteBackstageViewClientControl) Then
                    If Not selectedTabPageName = "MissFireBackstageViewTabItem" And Not selectedTabPageName = "HolidaysBackstageViewTabItem" Then
                        CopySimpleControls(selecteBackstageViewClientControl)
                    End If
                End If
            End With
        End Sub

        Shared Sub ProcessSimpleTab(startDate As DateTime,
                                    endDate As DateTime?)
            Dim timeCount As Integer
            Dim timeUnit As TimeUnit
            Dim isRunOnce As Boolean
            'Default value if not set
            Dim repeatCount As Integer = SimpleTriggerImpl.RepeatIndefinitely

            With ScheduleForm
                startDate = CDate(.StartDateTimeDateEdit.EditValue)
                Dim endDateEditValue = .EndDateTimeDateEdit.EditValue
                If Not IsNothing(endDateEditValue) Then endDate = .EndDateTimeDateEdit.DateTime

                If .RunTaskEveryCheckEdit.Checked Then
                    timeCount = CInt(.TimesCountSpinEdit.EditValue)
                    timeUnit = CType(.TimeUnitComboBoxEdit.SelectedIndex, PublicEnums.TimeUnit)
                ElseIf .RunTaskOnceCheckEdit.Checked Then
                    isRunOnce = True
                End If

                If .StopAfterCheckEdit.Checked Then
                    repeatCount = .TimesSpinEdit.Value
                End If

                Dim quTrg = CreateSimpleScheduleTrigger(startDate, endDate, timeCount, timeUnit, isRunOnce, repeatCount)

                'One More Check before proceed.
                Dim isStarted As Boolean = QuartzHelper.QuartzUtils.IsTriggerScheduled(quTrg.TriggerId)
                If Not isStarted Then quTrg.StartListening()
            End With
        End Sub

        Shared Sub ProcessAdvancedTab(startDate As DateTime,
                              endDate As DateTime?)
            Dim quTrgList = CreateAdvancedTrigger(startDate, endDate)
            For Each trg In quTrgList
                'One More Check before proceed.
                Dim isStarted As Boolean = QuartzHelper.QuartzUtils.IsTriggerScheduled(trg.TriggerId)
                If Not isStarted Then trg.StartListening()
            Next
        End Sub

        Shared Sub ProcessSelectedTab()
            With ScheduleForm
                Dim startDate As DateTime : Dim endDate As DateTime?
                startDate = CDate(.StartDateTimeDateEdit.EditValue)
                Dim endDateEditValue = .EndDateTimeDateEdit.EditValue
                If Not IsNothing(endDateEditValue) Then endDate = .EndDateTimeDateEdit.DateTime

                If Not IsEditMode() Then
                    If .ScheduleTriggerBackstage.SelectedTab Is .SimpleScheduleBackstageViewTabItem Then
                        ProcessSimpleTab(startDate, endDate)
                    ElseIf .ScheduleTriggerBackstage.SelectedTab Is .AdvancedScheduleBackstageViewTabItem Then
                        ProcessAdvancedTab(startDate, endDate)
                    ElseIf IsAdvancedTabSet() Then
                        'Case where clicing OK button from Misfire or Holidays tab
                        ProcessAdvancedTab(startDate, endDate)
                    Else
                        'Case where clicing OK button from Misfire or Holidays tab
                        ProcessSimpleTab(startDate, endDate)
                    End If
                Else
                    If IsAdvancedTabSet() Then
                        ProcessAdvancedTab(startDate, endDate)
                    Else
                        ProcessSimpleTab(startDate, endDate)
                    End If
                End If

            End With
        End Sub
#End Region

#Region "Buttons"
        Shared Sub OnCancel()
            ScheduleTriggerForm.Dispose()
            'Never dispose it because user may open it 1000 times and every time it recreated.
            'So either you will lose Form properties or you will re init them.
            'ScheduleTriggerForm.Dispose()
        End Sub

        Shared Sub OnOK()
            CacaduGloby.ShowWaitForm("Wait...")
            SetStatusBarText("Wait...")

            ProcessSelectedTab()

            SetStatusBarText("Triggers settings updated.")
            CacaduGloby.CloseWaitForm()
            ScheduleTriggerForm.Dispose()
            TriggersFormConnector.FillTriggersList(TriggersFormConnector.TaskObj)
            MainForm.TaskPropertiesPropertyGrid.Refresh()
            'Never dispose it because user may open it 1000 times and every time it recreated.
            'So either you will lose Form properties or you will re init them.
            'ScheduleTriggerForm.Dispose()
        End Sub
#End Region

#Region "Events Handlers"
        Shared Sub OnExcludeHolidaysCheckBox_CheckedChanged()
            ScheduleTriggerForm.HolidaysBackstageViewTabItem.Enabled = ScheduleTriggerForm.ExcludeHolidaysCheckBox.Checked
        End Sub
#End Region

#Region "General Helpers"
#Region "Init & Filling Controls"
        Shared Sub SetBoundControls()
            With ScheduleForm
                'Simple
                .TimeUnitComboBoxEdit.DataBindings.Add("Enabled", .RunTaskEveryCheckEdit, "Checked")
                .TimesCountSpinEdit.DataBindings.Add("Enabled", .RunTaskEveryCheckEdit, "Checked")
                .StopAfterCheckEdit.DataBindings.Add("Enabled", .RunTaskEveryCheckEdit, "Checked")
                .TimesSpinEdit.DataBindings.Add("Enabled", .StopAfterCheckEdit, "Checked")

                'Advanced
                .AddTimeSimpleButton.DataBindings.Add("Enabled", .RepeatTimeCheckEdit, "Checked")
                .RepeatTimeCheckedComboBoxEdit.DataBindings.Add("Enabled", .RepeatTimeCheckEdit, "Checked")
                .RepeatDayOfMonthCheckedComboBoxEdit.DataBindings.Add("Enabled", .RepeatDayOfMonthCheckEdit, "Checked")
                .RepeatMonthCheckedComboBoxEdit.DataBindings.Add("Enabled", .RepeatMonthCheckEdit, "Checked")
                .RepeatDayOfWeekCheckedComboBoxEdit.DataBindings.Add("Enabled", .RepeatDayOfWeekCheckEdit, "Checked")
                .DaysNameComboBoxEdit.DataBindings.Add("Enabled", .OrdinalDayOfWeekCheckEdit, "Checked")
                .OrdinalComboBoxEdit.DataBindings.Add("Enabled", .OrdinalDayOfWeekCheckEdit, "Checked")
                .RepeatYearCheckedComboBoxEdit.DataBindings.Add("Enabled", .RepeatYearsCheckEdit, "Checked")

                'Calendar
                .HolidayAddTimeSimpleButton.DataBindings.Add("Enabled", .HolidayTimeCheckEdit, "Checked")
                .HolidayTimeCheckedComboBoxEdit.DataBindings.Add("Enabled", .HolidayTimeCheckEdit, "Checked")
                .HolidayDayOfMonthCheckedComboBoxEdit.DataBindings.Add("Enabled", .HolidayDayOfMonthCheckEdit, "Checked")
                .HolidayMonthCheckedComboBoxEdit.DataBindings.Add("Enabled", .HolidayMonthCheckEdit, "Checked")
                .HolidayDayOfWeekCheckedComboBoxEdit.DataBindings.Add("Enabled", .HolidayDayOfWeekCheckEdit, "Checked")
                .HolidayDaysNameComboBoxEdit.DataBindings.Add("Enabled", .HolidayOrdinalDayOfWeekCheckEdit, "Checked")
                .HolidayOrdinalComboBoxEdit.DataBindings.Add("Enabled", .HolidayOrdinalDayOfWeekCheckEdit, "Checked")
                .HolidayYearCheckedComboBoxEdit.DataBindings.Add("Enabled", .HolidayYearsCheckEdit, "Checked")
            End With
        End Sub

        Shared Sub SetDefaultsCombosValues()
            If Not IsEditMode() Then
                With ScheduleForm
                    'No need now because you set the validation.
                    '.RepeatTimeCheckedComboBoxEdit.EditValue = DateTime.Now.ToLongTimeString
                End With
                FillYearsCombo()
            End If
        End Sub

        Shared Sub SetMisFireCommand(misFire As Cacadore.ITrigger.MisFireCommandsEnum)
            With ScheduleForm
                Select Case misFire
                    Case Cacadore.ITrigger.MisFireCommandsEnum.RunNow
                        .RunImmediatelyCheckEdit.Checked = True
                    Case Cacadore.ITrigger.MisFireCommandsEnum.AskMe
                        .AskMeCheckEdit.Checked = True
                    Case Cacadore.ITrigger.MisFireCommandsEnum.Skip
                        .OKSkipItCheckEdit.Checked = True
                End Select
            End With
        End Sub

        Shared Sub SetHolidays(holidayCronExpression As Quartz.CronExpression)
            FillHolidayControlsForEditing(holidayCronExpression.CronExpressionString)
        End Sub

        Shared Sub SetStatusBarText(text As String)
            If text = vbNullString Then
                ScheduleTriggerForm.ScheduleTriggerFormStatusBarStaticItem.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
            Else
                ScheduleTriggerForm.ScheduleTriggerFormStatusBarStaticItem.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
            End If
            ScheduleTriggerForm.ScheduleTriggerFormStatusBarStaticItem.Caption = text
        End Sub

        Shared Sub FillSimpleControlsForEditing(innerTrg As Quartz.ITrigger, parsedExpression As Balora.CroXGen.SimpleCron)
            Dim simpleTrigger As Quartz.ISimpleTrigger = CType(innerTrg, Quartz.ISimpleTrigger)
            With ScheduleTriggerForm
                .StartDateTimeDateEdit.EditValue = simpleTrigger.StartTimeUtc.ToLocalTime
                If simpleTrigger.EndTimeUtc.HasValue Then .EndDateTimeDateEdit.EditValue = simpleTrigger.EndTimeUtc.Value.ToLocalTime
                If simpleTrigger.RepeatInterval.TotalSeconds = 0 And simpleTrigger.RepeatCount = 0 Then
                    .RunTaskOnceCheckEdit.Checked = True
                Else
                    .RunTaskEveryCheckEdit.Checked = True
                    .TimesCountSpinEdit.EditValue = parsedExpression.n
                    .TimeUnitComboBoxEdit.SelectedIndex = parsedExpression.unit
                    If simpleTrigger.RepeatCount <> SimpleTriggerImpl.RepeatIndefinitely Then
                        .TimesSpinEdit.Value = simpleTrigger.RepeatCount
                    End If
                    If simpleTrigger.RepeatCount > 1 Then
                        .StopAfterCheckEdit.Checked = True
                    End If
                End If
            End With

            Connectors.ScheduleTriggerFormConnector.SetSelectedTab(ScheduleTriggerForm.SimpleScheduleBackstageViewTabItem)
        End Sub

        Shared Sub FillAdvancedControlsForEditing(cronExpression As String)
            Dim parsedExpr = ParseComplexCronExpression(cronExpression)

            With ScheduleTriggerForm
                If parsedExpr.Seconds.Count <> 61 AndAlso parsedExpr.Minutes.Count <> 61 AndAlso parsedExpr.Hours.Count <> 25 Then
                    Dim [time] As New DateTime()
                    Dim timeString = "#" & parsedExpr.Hours.Item(0).ToString & ":" & parsedExpr.Minutes.Item(0).ToString & ":" & parsedExpr.Seconds.Item(0).ToString & "#"
                    DateTime.TryParse(timeString, [time])
                    .RepeatTimeCheckEdit.Checked = True
                    Dim itmIndex As Integer = .RepeatTimeCheckedComboBoxEdit.Properties.Items.Add([time].ToLongTimeString)
                    .RepeatTimeCheckedComboBoxEdit.Properties.Items(itmIndex).CheckState = CheckState.Checked
                End If
            End With

            With ScheduleTriggerForm
                'Why 98, I don't know, All what I know that days has one element equals 98 if no days are selected.
                If parsedExpr.Days.Count < 32 And CInt(parsedExpr.Days.Item(0)) <> 98 Then
                    .RepeatDayOfMonthCheckEdit.Checked = True
                    For Each _day As Integer In parsedExpr.Days
                        .RepeatDayOfMonthCheckedComboBoxEdit.Properties.Items(_day - 1).CheckState = CheckState.Checked
                    Next
                End If
            End With

            With ScheduleTriggerForm
                If parsedExpr.Months.Count < 13 Then
                    .RepeatMonthCheckEdit.Checked = True
                    For Each _month As Integer In parsedExpr.Months
                        .RepeatMonthCheckedComboBoxEdit.Properties.Items(_month - 1).CheckState = CheckState.Checked
                    Next
                End If
            End With

            With ScheduleTriggerForm
                If parsedExpr.HasOrdinal <> -1 Then
                    .OrdinalDayOfWeekCheckEdit.Checked = True
                    .DaysNameComboBoxEdit.SelectedIndex = parsedExpr.OrdinalDayName - 1
                    .OrdinalComboBoxEdit.SelectedIndex = parsedExpr.Ordinal - 1
                End If

                If CInt(parsedExpr.DayOfWeek.Item(0)) <> 98 Then
                    .RepeatDayOfWeekCheckEdit.Checked = True
                    For Each _day As Integer In parsedExpr.DayOfWeek
                        Dim sdfadsf = .RepeatDayOfWeekCheckedComboBoxEdit.Properties.Items
                        .RepeatDayOfWeekCheckedComboBoxEdit.Properties.Items(_day - 1).CheckState = CheckState.Checked
                    Next
                End If
            End With

            With ScheduleTriggerForm
                If parsedExpr.Years.Count < 10 Then 'greater than any number greater than years displayed in years combo.
                    .RepeatYearsCheckEdit.Checked = True
                    For Each _year In parsedExpr.Years
                        Dim itmIndex As Integer = .RepeatYearCheckedComboBoxEdit.Properties.Items.Add(_year)
                        .RepeatYearCheckedComboBoxEdit.Properties.Items(itmIndex).CheckState = CheckState.Checked
                    Next
                End If
            End With

            Connectors.ScheduleTriggerFormConnector.SetSelectedTab(ScheduleTriggerForm.AdvancedScheduleBackstageViewTabItem)
        End Sub

        Shared Sub FillHolidayControlsForEditing(cronExpression As String)
            If cronExpression <> vbNullString Then ScheduleForm.ExcludeHolidaysCheckBox.Checked = True

            Dim parsedExpr = ParseComplexCronExpression(cronExpression)

            With ScheduleTriggerForm
                If parsedExpr.Seconds.Count <> 61 AndAlso parsedExpr.Minutes.Count <> 61 AndAlso parsedExpr.Hours.Count <> 25 Then
                    Dim [time] As New DateTime()
                    Dim timeString = "#" & parsedExpr.Hours.Item(0).ToString & ":" & parsedExpr.Minutes.Item(0).ToString & ":" & parsedExpr.Seconds.Item(0).ToString & "#"
                    DateTime.TryParse(timeString, [time])
                    .HolidayTimeCheckEdit.Checked = True
                    Dim itmIndex As Integer = .HolidayTimeCheckedComboBoxEdit.Properties.Items.Add([time].ToLongTimeString)
                    .HolidayTimeCheckedComboBoxEdit.Properties.Items(itmIndex).CheckState = CheckState.Checked
                End If
            End With

            With ScheduleTriggerForm
                'Why 98, I don't know, All what I know that days has one element equals 98 if no days are selected.
                If parsedExpr.Days.Count < 32 And CInt(parsedExpr.Days.Item(0)) <> 98 Then
                    .HolidayDayOfMonthCheckEdit.Checked = True
                    For Each _day As Integer In parsedExpr.Days
                        .HolidayDayOfMonthCheckedComboBoxEdit.Properties.Items(_day - 1).CheckState = CheckState.Checked
                    Next
                End If
            End With

            With ScheduleTriggerForm
                If parsedExpr.Months.Count < 13 Then
                    .HolidayMonthCheckEdit.Checked = True
                    For Each _month As Integer In parsedExpr.Months
                        .HolidayMonthCheckedComboBoxEdit.Properties.Items(_month - 1).CheckState = CheckState.Checked
                    Next
                End If
            End With

            With ScheduleTriggerForm
                If parsedExpr.HasOrdinal <> -1 Then
                    .HolidayOrdinalDayOfWeekCheckEdit.Checked = True
                    .HolidayDaysNameComboBoxEdit.SelectedIndex = parsedExpr.OrdinalDayName - 1
                    .HolidayOrdinalComboBoxEdit.SelectedIndex = parsedExpr.Ordinal - 1
                End If

                If CInt(parsedExpr.DayOfWeek.Item(0)) <> 98 Then
                    .HolidayDayOfWeekCheckEdit.Checked = True
                    For Each _day As Integer In parsedExpr.DayOfWeek
                        Dim sdfadsf = .HolidayDayOfWeekCheckedComboBoxEdit.Properties.Items
                        .HolidayDayOfWeekCheckedComboBoxEdit.Properties.Items(_day - 1).CheckState = CheckState.Checked
                    Next
                End If
            End With

            With ScheduleTriggerForm
                If parsedExpr.Years.Count < 10 Then 'greater than any number greater than years displayed in years combo.
                    .HolidayYearsCheckEdit.Checked = True
                    For Each _year In parsedExpr.Years
                        Dim itmIndex As Integer = .HolidayYearCheckedComboBoxEdit.Properties.Items.Add(_year)
                        .HolidayYearCheckedComboBoxEdit.Properties.Items(itmIndex).CheckState = CheckState.Checked
                    Next
                End If
            End With
        End Sub

        Shared Sub FillYearsCombo()
            Dim currentYear As Integer = DateTime.Now.Year
            With ScheduleTriggerForm
                .RepeatYearCheckedComboBoxEdit.Properties.Items.AddRange({New DevExpress.XtraEditors.Controls.CheckedListBoxItem(currentYear),
                                                                          New DevExpress.XtraEditors.Controls.CheckedListBoxItem(currentYear + 1),
                                                                          New DevExpress.XtraEditors.Controls.CheckedListBoxItem(currentYear + 2)})
                .RepeatYearCheckedComboBoxEdit.Properties.Items(0).CheckState = CheckState.Checked

                .HolidayYearCheckedComboBoxEdit.Properties.Items.AddRange({New DevExpress.XtraEditors.Controls.CheckedListBoxItem(currentYear),
                                                                                 New DevExpress.XtraEditors.Controls.CheckedListBoxItem(currentYear + 1),
                                                                                 New DevExpress.XtraEditors.Controls.CheckedListBoxItem(currentYear + 2)})
                .HolidayYearCheckedComboBoxEdit.Properties.Items(0).CheckState = CheckState.Checked
            End With
        End Sub

        Shared Sub AddRepeatTime()
            With ScheduleForm
                Dim editValue = .RepeatTimeCheckedComboBoxEdit.EditValue
                If editValue.ToString = vbNullString Then Exit Sub
                If editValue.ToString.Contains(",") Then Exit Sub
                Dim dateTimeValue As DateTime = CDate(editValue)
                Dim timeValue As String = dateTimeValue.ToLongTimeString
                'Exit sub if no value, or all of them are displayed like this 10:00:55,10:30:00
                If Not IsValidRepeatTime(timeValue) Then Exit Sub
                .RepeatTimeCheckedComboBoxEdit.Properties.Items.Add(timeValue, True)
            End With
        End Sub

        Shared Sub AddCalenderRepeatTime()
            With ScheduleForm
                Dim editValue = .HolidayTimeCheckedComboBoxEdit.EditValue
                If editValue.ToString = vbNullString Then Exit Sub
                .HolidayTimeCheckedComboBoxEdit.Properties.Items.Clear()
                Dim dateTimeValue As DateTime = CDate(editValue)
                Dim timeValue As String = dateTimeValue.ToLongTimeString
                'Exit sub if no value, or all of them are displayed like this 10:00:55,10:30:00
                If Not IsValidRepeatTime(timeValue) Then Exit Sub
                .HolidayTimeCheckedComboBoxEdit.Properties.Items.Add(timeValue, True)
            End With
        End Sub
#End Region

#Region "Building Triggers"
        Shared Function GenerateCalenderCronExpressions() As Dictionary(Of String, String)
            Dim times As New SortedSet(Of Date)
            Dim days As New SortedSet(Of Integer)
            Dim years As New SortedSet(Of Integer)
            Dim daysOfWeek As New SortedList(Of WeekDays, String)
            Dim months As New SortedSet(Of Balora.PublicEnums.Months)

            With ScheduleTriggerForm
                'Get times set
                If .HolidayTimeCheckEdit.Checked Then
                    For Each itm As CheckedListBoxItem In .HolidayTimeCheckedComboBoxEdit.Properties.Items
                        If itm.CheckState = CheckState.Checked Then
                            times.Add(CDate(itm.Value))
                        End If
                    Next
                End If

                'Get days set
                If .HolidayDayOfMonthCheckEdit.Checked Then
                    For Each itm As CheckedListBoxItem In .HolidayDayOfMonthCheckedComboBoxEdit.Properties.Items
                        If itm.CheckState = CheckState.Checked Then
                            days.Add(CInt(itm.Value))
                        End If
                    Next
                End If

                'Get months set
                If .HolidayMonthCheckEdit.Checked Then
                    For Each itm As CheckedListBoxItem In .HolidayMonthCheckedComboBoxEdit.Properties.Items
                        If itm.CheckState = CheckState.Checked Then
                            Dim itmIndex = .HolidayMonthCheckedComboBoxEdit.Properties.Items.IndexOf(itm) + 1
                            months.Add(CType(itmIndex, PublicEnums.Months))
                        End If
                    Next
                End If

                'Get day of week set
                If .HolidayDayOfWeekCheckEdit.Checked Then
                    For Each itm As CheckedListBoxItem In .HolidayDayOfWeekCheckedComboBoxEdit.Properties.Items
                        If itm.CheckState = CheckState.Checked Then
                            Dim itmIndex = .HolidayDayOfWeekCheckedComboBoxEdit.Properties.Items.IndexOf(itm) + 1
                            Dim dayOfWeek As WeekDays = CType(itmIndex, WeekDays)
                            daysOfWeek.Add(dayOfWeek, dayOfWeek.ToString)
                        End If
                    Next
                End If

                'Get ordinal Day Of Week set
                If .HolidayOrdinalDayOfWeekCheckEdit.Checked Then
                    'The '#' is used to specify "the nth" XXX weekday of the month. For example, the value of "6#3" or "FRI#3" in the day-of-week field means "the third Friday of the month".
                    Dim dayNumber As Integer = .HolidayDaysNameComboBoxEdit.SelectedIndex + 1
                    Dim ordinalNumber As Integer = .HolidayOrdinalComboBoxEdit.SelectedIndex + 1
                    Dim ordinalValue As String = dayNumber & "#" & ordinalNumber
                    daysOfWeek.Add(WeekDays.Ordinal, ordinalValue)
                End If

                'Get Years set
                If .HolidayYearsCheckEdit.Checked Then
                    For Each itm As CheckedListBoxItem In .HolidayYearCheckedComboBoxEdit.Properties.Items
                        If itm.CheckState = CheckState.Checked Then
                            years.Add(CInt(itm.Value))
                        End If
                    Next
                End If
            End With
            Dim calenderCronExpressions As Dictionary(Of String, String) = CroXGen.GenerateComplexExpressions(times,
                                                                                                              days,
                                                                                                              months,
                                                                                                              daysOfWeek,
                                                                                                              years)
            Return calenderCronExpressions
        End Function

        Shared Function GenerateAdvancedCronExpressions() As Dictionary(Of String, String)
            Dim times As New SortedSet(Of Date)
            Dim days As New SortedSet(Of Integer)
            Dim years As New SortedSet(Of Integer)
            Dim daysOfWeek As New SortedList(Of WeekDays, String)
            Dim months As New SortedSet(Of Balora.PublicEnums.Months)

            With ScheduleTriggerForm
                'Get times set
                If .RepeatTimeCheckEdit.Checked Then
                    For Each itm As CheckedListBoxItem In .RepeatTimeCheckedComboBoxEdit.Properties.Items
                        If itm.CheckState = CheckState.Checked Then
                            times.Add(CDate(itm.Value))
                        End If
                    Next
                End If

                'Get days set
                If .RepeatDayOfMonthCheckEdit.Checked Then
                    For Each itm As CheckedListBoxItem In .RepeatDayOfMonthCheckedComboBoxEdit.Properties.Items
                        If itm.CheckState = CheckState.Checked Then
                            days.Add(CInt(itm.Value))
                        End If
                    Next
                End If

                'Get months set
                If .RepeatMonthCheckEdit.Checked Then
                    For Each itm As CheckedListBoxItem In .RepeatMonthCheckedComboBoxEdit.Properties.Items
                        If itm.CheckState = CheckState.Checked Then
                            Dim itmIndex = .RepeatMonthCheckedComboBoxEdit.Properties.Items.IndexOf(itm) + 1
                            months.Add(CType(itmIndex, PublicEnums.Months))
                        End If
                    Next
                End If

                'Get day of week set
                If .RepeatDayOfWeekCheckEdit.Checked Then
                    For Each itm As CheckedListBoxItem In .RepeatDayOfWeekCheckedComboBoxEdit.Properties.Items
                        If itm.CheckState = CheckState.Checked Then
                            Dim itmIndex = .RepeatDayOfWeekCheckedComboBoxEdit.Properties.Items.IndexOf(itm) + 1
                            Dim dayOfWeek As WeekDays = CType(itmIndex, WeekDays)
                            daysOfWeek.Add(dayOfWeek, dayOfWeek.ToString)
                        End If
                    Next
                End If

                'Get ordinal Day Of Week set
                If .OrdinalDayOfWeekCheckEdit.Checked Then
                    'The '#' is used to specify "the nth" XXX weekday of the month. For example, the value of "6#3" or "FRI#3" in the day-of-week field means "the third Friday of the month".
                    Dim dayNumber As Integer = .DaysNameComboBoxEdit.SelectedIndex + 1
                    Dim ordinalNumber As Integer = .OrdinalComboBoxEdit.SelectedIndex + 1
                    Dim ordinalValue As String = dayNumber & "#" & ordinalNumber
                    daysOfWeek.Add(WeekDays.Ordinal, ordinalValue)
                End If

                'Get Years set
                If .RepeatYearsCheckEdit.Checked Then
                    For Each itm As CheckedListBoxItem In .RepeatYearCheckedComboBoxEdit.Properties.Items
                        If itm.CheckState = CheckState.Checked Then
                            years.Add(CInt(itm.Value))
                        End If
                    Next
                End If
            End With
            Dim cronExpressions As Dictionary(Of String, String) = CroXGen.GenerateComplexExpressions(times,
                                                                                                      days,
                                                                                                      months,
                                                                                                      daysOfWeek,
                                                                                                      years)
            Return cronExpressions
        End Function

        Shared Function CreateSimpleScheduleTrigger(ByRef startDate As DateTime,
                                                    ByRef endDate As DateTime?,
                                                    ByRef timeCount As Integer,
                                                    ByRef timeUnit As TimeUnit,
                                                    ByRef runOnce As Boolean,
                                                    Optional repeatCount As Int32 = 1) As QuartzHelper.QuartzTrigger
            Dim calendarCronExpression As Dictionary(Of String, String) = Nothing
            Dim tskMeta As TaskMeta = SetTaskMetaDates(startDate, endDate)

            tskMeta.TriggerType = Cacadore.ITrigger.TriggerTypeEnum.SimpleTrigger

            Dim cronExpressions As Dictionary(Of String, String) = Nothing
            If Not runOnce Then
                cronExpressions = CroXGen.GenerateEveryNTimeunitExpression(timeCount,
                                                                           timeUnit,
                                                                           tskMeta.StartDateTimeUTC.Value.ToLocalTime)
            ElseIf runOnce Then
                cronExpressions = CroXGen.GenerateRunOnceExpression(tskMeta.StartDateTimeUTC.Value.ToLocalTime)
                tskMeta.RunOnce = True
            End If

            With tskMeta
                If Not runOnce Then .RepeatCount = repeatCount
                .MisFire = CheckMisFireCommand()
                .CronExpression = cronExpressions.First.Key
                .CronExpressionDescription = cronExpressions.First.Value
                'No holiday for Simple Triggers.
                'calendarCronExpression = CheckHolidays()
                'If Not IsNothing(calendarCronExpression) AndAlso calendarCronExpression.Count > 0 Then
                '    .CalendarCronExpression = CheckHolidays().First.Key
                '    .CalendarCronExpressionDescription = CheckHolidays().First.Value
                'End If
            End With

            Dim oldTrg = CheckIfSimpleTriggerExists(startDate, runOnce, timeCount, timeUnit)
            'AhSaEldin 20120918: Added ElseIf and "repeatCount = 0 " to
            'fix this bug https://conderella.com/wi/client/index.php?issue=485
            If Not IsNothing(oldTrg) AndAlso repeatCount = 0 Then
                'This causes the infamous "Calendar Not Found."
                'If IsNothing(calendarCronExpression) Then
                'oldTrg.Calendar = Nothing
                'End If
                Return oldTrg
            ElseIf Not IsNothing(oldTrg) AndAlso repeatCount <> 0 Then
                oldTrg.Delete()
            End If

            Dim qutrg As QuartzHelper.QuartzTrigger = CacadoreHelperObj.AttachScheduledTriggerToTask(ScheduleTriggerForm.InnerTask, tskMeta)
            tskMeta.Dispose()
            tskMeta = Nothing
            Return qutrg
        End Function

        Shared Function CreateAdvancedTriggersList(cronExpressions As Dictionary(Of String, String), tskMeta As TaskMeta) As List(Of QuartzHelper.QuartzTrigger)
            Dim calendarCronExpression As Dictionary(Of String, String)
            Dim triggersList As New List(Of QuartzHelper.QuartzTrigger)
            For Each cronExpr In cronExpressions
                With tskMeta
                    .CronExpression = cronExpr.Key
                    .CronExpressionDescription = cronExpr.Value
                    calendarCronExpression = CheckHolidays()
                    If Not IsNothing(calendarCronExpression) AndAlso calendarCronExpression.Count > 0 Then
                        .CalendarCronExpression = CheckHolidays().First.Key
                        .CalendarCronExpressionDescription = CheckHolidays().First.Value
                    End If
                End With
                Dim oldTrg = CheckIfCronTriggerExists(cronExpr.Key)
                If Not IsNothing(oldTrg) Then
                    If IsNothing(calendarCronExpression) Then
                        oldTrg.Calendar = Nothing
                    End If

                    triggersList.Add(oldTrg)
                Else
                    Dim newTrg As QuartzHelper.QuartzTrigger = CacadoreHelperObj.AttachScheduledTriggerToTask(ScheduleTriggerForm.InnerTask, tskMeta)
                    triggersList.Add(newTrg)
                End If
            Next
            Return triggersList
        End Function

        Shared Function CreateAdvancedTrigger(startDate As DateTime, endDate As DateTime?) As List(Of QuartzHelper.QuartzTrigger)
            Dim cronExpressions As Dictionary(Of String, String) = GenerateAdvancedCronExpressions()
            Dim tskMeta As TaskMeta = SetTaskMetaDates(startDate, endDate)
            With tskMeta
                .TriggerType = Cacadore.ITrigger.TriggerTypeEnum.CronTrigger
                .MisFire = CheckMisFireCommand()
            End With
            Dim triggersList = CreateAdvancedTriggersList(cronExpressions, tskMeta)
            tskMeta.Dispose()
            tskMeta = Nothing
            Return triggersList
        End Function
#End Region

#Region "General Helpers"
        Shared Function SetTaskMetaDates(ByRef startDate As DateTime,
                                     ByRef endDate As DateTime?) As TaskMeta
            Dim tskMeta As New TaskMeta
            With tskMeta
                .StartDateTimeUTC = startDate.ToUniversalTime
                If endDate.HasValue Then .EndDateTimeUTC = endDate.Value.ToUniversalTime
            End With
            Return tskMeta
        End Function

        Shared Sub ProcessKey(keyValue As Integer)
            Select Case keyValue
                Case Keys.Enter
                    OnOK()
                Case Keys.Escape
                    OnCancel()
            End Select
        End Sub

        Shared Sub ResetTabsSelectedProperty()
            With ScheduleTriggerForm
                .SimpleScheduleBackstageViewTabItem.Selected = True
                .AdvancedScheduleBackstageViewTabItem.Selected = False
                .MissFireBackstageViewTabItem.Selected = False
                .HolidaysBackstageViewTabItem.Selected = False
            End With
        End Sub

        Shared Function ParseComplexCronExpression(cronExpression As String) As ParsedComplexCronExpression
            Dim croxExprObj As New Quartz.CronExpression(cronExpression)
            '1:00:00 AM at day(s) 27,28 in the month(s) OCT,NOV during the year(s) 2013,2014 starting from 4/14/2012 7:59:00 PM
            Dim secondsField = croxExprObj.GetType.GetField("seconds", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
            Dim seconds As Quartz.Collection.TreeSet(Of Integer) = secondsField.GetValue(croxExprObj)

            Dim minutesField = croxExprObj.GetType.GetField("minutes", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
            Dim minutes As Quartz.Collection.TreeSet(Of Integer) = minutesField.GetValue(croxExprObj)

            Dim hoursField = croxExprObj.GetType.GetField("hours", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
            Dim hours As Quartz.Collection.TreeSet(Of Integer) = hoursField.GetValue(croxExprObj)

            Dim daysField = croxExprObj.GetType.GetField("daysOfMonth", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
            Dim days As Quartz.Collection.TreeSet(Of Integer) = daysField.GetValue(croxExprObj)

            Dim monthsField = croxExprObj.GetType.GetField("months", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
            Dim months As Quartz.Collection.TreeSet(Of Integer) = monthsField.GetValue(croxExprObj)

            Dim dayOfWeekField = croxExprObj.GetType.GetField("daysOfWeek", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
            Dim dayOfWeek As Quartz.Collection.TreeSet(Of Integer) = dayOfWeekField.GetValue(croxExprObj)

            Dim hasOrdinal As Integer = cronExpression.IndexOf("#")
            Dim ordinalDayName As Integer
            Dim ordinal As Integer
            If hasOrdinal <> -1 Then
                ordinalDayName = CInt(Strings.Mid(cronExpression, hasOrdinal, 1))
                ordinal = CInt(Strings.Mid(cronExpression, hasOrdinal + 2, 1))
            End If

            Dim yearsField = croxExprObj.GetType.GetField("years", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
            Dim years As Quartz.Collection.TreeSet(Of Integer) = yearsField.GetValue(croxExprObj)

            Dim parseExpr As New ParsedComplexCronExpression
            With parseExpr
                .Seconds = seconds
                .Minutes = minutes
                .Hours = hours
                .Days = days
                .Months = months
                .DayOfWeek = dayOfWeek
                .Years = years
                .HasOrdinal = hasOrdinal
                .Ordinal = ordinal
                .OrdinalDayName = ordinalDayName
            End With

            Return parseExpr
        End Function

        ''' <summary>
        ''' Processes the day checked controls.
        ''' </summary>
        ''' <param name="dayCheckControl">The day check control.</param>
        ''' <remarks>Used to check if one of day checkbox controls checked, in order to disable the two other controls of the three</remarks>
        Shared Sub ProcessDayOfWeekCheckedControls(dayCheckControl As DevExpress.XtraEditors.CheckEdit)
            Dim checkedState As Boolean = dayCheckControl.Checked
            Dim currentEnabledState As Boolean = dayCheckControl.Enabled
            ToggleDaysControlsState(Not checkedState, dayCheckControl)
            dayCheckControl.Enabled = currentEnabledState
        End Sub

        Shared Sub ToggleAddTimeButton()
            With ScheduleForm
                Dim RepeatTimeCheckedComboBoxValue As String = .RepeatTimeCheckedComboBoxEdit.EditValue.ToString
                Dim selectedTimesCount As Object = .RepeatTimeCheckedComboBoxEdit.Properties.GetCheckedItems()
                If RepeatTimeCheckedComboBoxValue <> vbNullString AndAlso Not RepeatTimeCheckedComboBoxValue.Contains(",") Then
                    .AddTimeSimpleButton.Enabled = True
                Else
                    .AddTimeSimpleButton.Enabled = False
                End If
            End With
        End Sub

        Shared Sub ToggleDaysControlsState(state As Boolean, dayCheckControl As DevExpress.XtraEditors.CheckEdit)
            With ScheduleForm
                If dayCheckControl.Name.Contains("Holiday") Then
                    .HolidayDayOfMonthCheckEdit.Enabled = state
                    .HolidayDayOfWeekCheckEdit.Enabled = state
                    .HolidayOrdinalDayOfWeekCheckEdit.Enabled = state
                Else
                    .RepeatDayOfMonthCheckEdit.Enabled = state
                    .RepeatDayOfWeekCheckEdit.Enabled = state
                    .OrdinalDayOfWeekCheckEdit.Enabled = state
                End If
            End With
        End Sub
#End Region

#Region "Checker Methods"
        Shared Function IsAdvancedTabSet() As Boolean
            'واحد منهم متعلم عليه، يبقي أددفانسد
            With ScheduleTriggerForm
                If .RepeatTimeCheckEdit.Checked Then Return True
                If .RepeatDayOfMonthCheckEdit.Checked Then Return True
                If .RepeatMonthCheckEdit.Checked Then Return True
                If .RepeatDayOfMonthCheckEdit.Checked Then Return True
                If .OrdinalDayOfWeekCheckEdit.Checked Then Return True
                If .RepeatYearsCheckEdit.Checked Then Return True
            End With
            Return False
        End Function

        Shared Function IsSimpleTabSet() As Boolean
            Return Not IsAdvancedTabSet()
        End Function

        Shared Function IsEditMode() As Boolean
            If ScheduleTriggerForm.Mode.Contains("Edit") Then
                Return True
            Else
                Return False
            End If
        End Function

        Shared Function IsDuplicatedTime(repeatTimeValue As String) As Boolean
            With ScheduleForm
                Dim repeatTimeItems = .RepeatTimeCheckedComboBoxEdit.Properties.Items
                Dim queryResult = From itm In repeatTimeItems Where itm.ToString = repeatTimeValue Select itm
                If queryResult.Count > 0 Then
                    Return True
                Else
                    Return False
                End If
            End With
        End Function

        Shared Function IsValidRepeatTime(repeatTimeValue As String) As Boolean
            With ScheduleForm
                'false if no value, or all of them are displayed like this 10:00:55,10:30:00
                If repeatTimeValue = vbNullString OrElse repeatTimeValue.Contains(",") Then Return False
                If IsDuplicatedTime(repeatTimeValue) Then Return False
            End With
            Return True
        End Function

        Shared Function CheckHolidays() As Dictionary(Of String, String)
            If ScheduleForm.ExcludeHolidaysCheckBox.Checked Then Return GenerateCalenderCronExpressions()
            Return Nothing
        End Function

        Shared Function CheckMisFireCommand() As Cacadore.ITrigger.MisFireCommandsEnum
            If ScheduleForm.RunImmediatelyCheckEdit.Checked Then
                Return Cacadore.ITrigger.MisFireCommandsEnum.RunNow
            ElseIf ScheduleForm.OKSkipItCheckEdit.Checked Then
                Return Cacadore.ITrigger.MisFireCommandsEnum.Skip
            ElseIf ScheduleForm.AskMeCheckEdit.Checked Then
                Return Cacadore.ITrigger.MisFireCommandsEnum.AskMe
            End If
            Return Cacadore.ITrigger.MisFireCommandsEnum.AskMe
        End Function

        Shared Function CheckIfSimpleTriggerExists(startDate As DateTime,
                                                   isRunOnce As Boolean,
                                                   ByRef timeCount As Integer,
                                                   ByRef timeUnit As TimeUnit) As QuartzHelper.QuartzTrigger
            If IsEditMode() Then
                Dim trgId = ScheduleTriggerForm.TriggerId
                Dim trg As QuartzHelper.QuartzTrigger = CType(ScheduleTriggerForm.InnerTask.MyTriggers.Triggers(trgId), QuartzHelper.QuartzTrigger)
                Dim innerTrg As Quartz.ISimpleTrigger = CType(trg.InnerTrigger, Quartz.ISimpleTrigger)
                Dim parsedExpr As Balora.CroXGen.SimpleCron
                Helpers.QuartzHelper.QuartzUtils.ParseSimpleTrigger(trg, parsedExpr)

                If isRunOnce Then
                    If startDate <> innerTrg.StartTimeUtc Then
                        ScheduleTriggerForm.InnerTask.UnAttachTrigger(trgId)
                    End If
                ElseIf timeCount = parsedExpr.n AndAlso timeUnit = parsedExpr.unit Then
                    Return trg
                Else
                    ScheduleTriggerForm.InnerTask.UnAttachTrigger(trgId)
                End If

            Else
                For Each trg In ScheduleTriggerForm.InnerTask.MyTriggers.Triggers.Values
                    Dim parsedExpr As Balora.CroXGen.SimpleCron
                    Helpers.QuartzHelper.QuartzUtils.ParseSimpleTrigger(trg, parsedExpr)
                    Dim tempTrg As Quartz.ITrigger = trg.InnerTrigger
                    Dim tempTrgST As DateTime = tempTrg.StartTimeUtc.DateTime
                    If isRunOnce AndAlso startDate = tempTrgST Then
                        Return trg
                    ElseIf isRunOnce AndAlso startDate <> tempTrgST Then
                        Return Nothing
                    ElseIf timeCount = parsedExpr.n AndAlso
                           timeUnit = parsedExpr.unit AndAlso
                           timeCount <> 0 And
                           timeUnit <> PublicEnums.TimeUnit.Seconds Then
                        'بنتأكد الأول انها مش الحالة الصفرية
                        'TimeCount = 0 and
                        'timeUnit = Seconds (ya3ne 0 bardo)
                        Return trg
                    End If
                Next
            End If
            Return Nothing
        End Function

        'Compare generated cron expression with the old one.
        'Return nothing if not defined, the old trg if both are the same.
        Shared Function CheckIfCronTriggerExists(cronExpression As String) As QuartzHelper.QuartzTrigger
            If IsEditMode() Then
                Dim trgId = ScheduleTriggerForm.TriggerId
                Dim trg As QuartzHelper.QuartzTrigger = CType(ScheduleTriggerForm.InnerTask.MyTriggers.Triggers(trgId), QuartzHelper.QuartzTrigger)
                Dim innerTrg As Quartz.ICronTrigger = CType(trg.InnerTrigger, Quartz.ICronTrigger)
                Dim oldCronExpr As String = innerTrg.CronExpressionString
                'Check old one with the new one
                If oldCronExpr.Contains(cronExpression) Then
                    Return trg
                Else
                    ScheduleTriggerForm.InnerTask.UnAttachTrigger(trgId)
                End If
            Else
                'Adding Mode
                Dim oldTrg = Cacadore.Commands.IsTriggerAttachedByCronExpression(ScheduleTriggerForm.InnerTask, cronExpression)
                If Not IsNothing(oldTrg) Then
                    Return oldTrg
                End If
            End If
            Return Nothing
        End Function
#End Region
#End Region
    End Class
End Namespace

