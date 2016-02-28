Imports DevExpress.XtraEditors.DXErrorProvider
Imports Cacadu.Connectors.ScheduleTriggerFormConnector
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Controls

Namespace UI
    Friend Class ScheduleTriggerForm

#Region "Fields"
        Private _innerTask As Cacadore.Task
        Private _myTaskValue As Cacadore.Task
        Private isFormLoaded As Boolean = False
#End Region

#Region "Fields"
        Property Mode As String
        Property TriggerId As String
        Property MyTrigger As Helpers.QuartzHelper.QuartzTrigger
#End Region

#Region "Constructors"
        Public Sub New()
            ' This call is required by the designer.
            InitializeComponent()
            ' Add any initialization after the InitializeComponent() call.
            ScheduleForm = Me
        End Sub
#End Region

#Region "Properties"
        Public ReadOnly Property MyTaskId() As String
            Get
                Return UI.TriggersForm.MyTaskId
            End Get
        End Property

        Public ReadOnly Property InnerTask() As Cacadore.Task
            Get
                Return _innerTask
            End Get
        End Property
#End Region

#Region "Events Handlers"
#Region "Form"
        Private Sub ScheduleTrigger_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
            _innerTask = CacadoreToTectonic.GetTaskById(MyTaskId)

            If Not isFormLoaded Then
                LoadScheduleTriggerForm()
                isFormLoaded = True
            End If

            If Not IsEditMode() Then
                Me.StartDateTimeDateEdit.EditValue = DateTime.Now
                EndDateTimeDateEdit.Properties.MinValue = DateTime.Now
                EndDateTimeDateEdit.EditValue = Nothing

            Else

                MyTrigger = CType(InnerTask.MyTriggers.Triggers(TriggerId), Helpers.QuartzHelper.QuartzTrigger)
                Me.StartDateTimeDateEdit.EditValue = MyTrigger.StartTimeUTC.Value.ToLocalTime

                If MyTrigger.EndTimeUTC.HasValue Then
                    EndDateTimeDateEdit.EditValue = MyTrigger.EndTimeUTC.Value.ToLocalTime
                Else
                    EndDateTimeDateEdit.Properties.MinValue = DateTime.Now
                    EndDateTimeDateEdit.EditValue = Nothing
                End If
            End If

            SetStatusBarText("")
            SetValidationRules()
            MoveCommonControlsBetweenTabs(ScheduleTriggerBackstage.SelectedTab.Name)
        End Sub

        Private Sub ScheduleTriggerForm_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
            Mode = vbNullString
            TriggerId = vbNullString
        End Sub

        Private Sub ScheduleTriggerForm_KeyUp(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
            ProcessKey(e.KeyValue)
        End Sub
#End Region

#Region "Controls"
        Private Sub ExcludeHolidaysCheckBox_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles ExcludeHolidaysCheckBox.CheckedChanged
            OnExcludeHolidaysCheckBox_CheckedChanged()
        End Sub

        Private Sub ScheduleTriggerBackstage_SelectedTabChanged(sender As System.Object, e As DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs) Handles ScheduleTriggerBackstage.SelectedTabChanged
            'To avoid an before loading the form.
            If Me.IsHandleCreated Then
                MoveCommonControlsBetweenTabs(e.Item.Name)
                If e.Item.Name = "AdvancedScheduleBackstageViewTabItem" Then
                    EndDateTimeDateEdit.Enabled = True
                ElseIf e.Item.Name = "SimpleScheduleBackstageViewTabItem" Then
                    If RunTaskOnceCheckEdit.Checked Then
                        EndDateTimeDateEdit.Enabled = False
                    Else
                        EndDateTimeDateEdit.Enabled = True
                    End If
                End If

            End If
        End Sub

        Private Sub RepeatTimeCheckedComboBoxEdit_EditValueChanged(sender As System.Object, e As System.EventArgs) Handles RepeatTimeCheckedComboBoxEdit.EditValueChanged
            'ToggleAddTimeButton()
        End Sub

        Private Sub AddTimeButton_Click(sender As System.Object, e As System.EventArgs) Handles AddTimeSimpleButton.Click
            AddRepeatTime()
        End Sub

        Private Sub RunTaskOnceCheckEdit_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RunTaskOnceCheckEdit.CheckedChanged
            If RunTaskOnceCheckEdit.Checked AndAlso IsSimpleTabSet() Then
                EndDateTimeDateEdit.Enabled = False
            Else
                EndDateTimeDateEdit.Enabled = True
            End If
        End Sub

        Private Sub AddTimeHolidaySimpleButton_Click(sender As System.Object, e As System.EventArgs) Handles HolidayAddTimeSimpleButton.Click
            AddCalenderRepeatTime()
        End Sub

        Private Sub OKButton_Click(sender As System.Object, e As System.EventArgs) Handles OKSimpleButton.Click
            Dim validationResult As Boolean = ScheduleTriggerFormDxValidationProvider.Validate()
            If validationResult Then OnOK()
        End Sub

        Private Sub CancelSimpleButton_Click(sender As System.Object, e As System.EventArgs) Handles CancelSimpleButton.Click
            OnCancel()
        End Sub
#End Region

#Region "Day Of Week Enable Checks"
        Private Sub RepeatDayOfMonthCheckEdit_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RepeatDayOfMonthCheckEdit.CheckedChanged
            ProcessDayOfWeekCheckedControls(RepeatDayOfMonthCheckEdit)
        End Sub

        Private Sub RepeatDayOfWeekCheckEdit_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RepeatDayOfWeekCheckEdit.CheckedChanged
            ProcessDayOfWeekCheckedControls(RepeatDayOfWeekCheckEdit)
        End Sub

        Private Sub OrdinalDayOfWeekCheckEdit_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles OrdinalDayOfWeekCheckEdit.CheckedChanged
            ProcessDayOfWeekCheckedControls(OrdinalDayOfWeekCheckEdit)
        End Sub

        Private Sub HolidayDayOfMonthCheckEdit_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles HolidayDayOfMonthCheckEdit.CheckedChanged
            ProcessDayOfWeekCheckedControls(HolidayDayOfMonthCheckEdit)
        End Sub

        Private Sub HolidayDayOfWeekCheckEdit_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles HolidayDayOfWeekCheckEdit.CheckedChanged
            ProcessDayOfWeekCheckedControls(HolidayDayOfWeekCheckEdit)
        End Sub

        Private Sub HolidayOrdinalDayOfWeekCheckEdit_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles HolidayOrdinalDayOfWeekCheckEdit.CheckedChanged
            ProcessDayOfWeekCheckedControls(HolidayOrdinalDayOfWeekCheckEdit)
        End Sub
#End Region

#Region "Validation Rules"
        Sub SetValidationRules()
            With ScheduleForm
                .ScheduleTriggerFormDxValidationProvider.SetValidationRule(EndDateTimeDateEdit,
                                                                           New BlankComboCustomValidation With {.ErrorText = "End date must be greater than start date.",
                                                                                                                .ErrorType = ErrorType.Critical})

                .ScheduleTriggerFormDxValidationProvider.SetValidationRule(RepeatDayOfMonthCheckedComboBoxEdit,
                                                                           New BlankComboCustomValidation With {.ErrorText = "Select repeat day of month or uncheck Repeat Time checkbox.",
                                                                                                                .ErrorType = ErrorType.Critical})


                .ScheduleTriggerFormDxValidationProvider.SetValidationRule(RepeatTimeCheckedComboBoxEdit,
                                                                           New BlankComboCustomValidation With {.ErrorText = "Write your desired time then click Add Time button",
                                                                                                                .ErrorType = ErrorType.Critical})

                .ScheduleTriggerFormDxValidationProvider.SetValidationRule(RepeatMonthCheckedComboBoxEdit,
                                                                           New BlankComboCustomValidation With {.ErrorText = "Select repeat month or uncheck Repeat Month checkbox.",
                                                                                                                .ErrorType = ErrorType.Critical})

                .ScheduleTriggerFormDxValidationProvider.SetValidationRule(RepeatDayOfWeekCheckedComboBoxEdit,
                                                                           New BlankComboCustomValidation With {.ErrorText = "Select repeat day of week or uncheck Repeat Day Of Week checkbox.",
                                                                                                                .ErrorType = ErrorType.Critical})

                .ScheduleTriggerFormDxValidationProvider.SetValidationRule(HolidayDayOfMonthCheckedComboBoxEdit,
                                                                           New BlankComboCustomValidation With {.ErrorText = "Select repeat day of month or uncheck Repeat Time checkbox.",
                                                                                                                .ErrorType = ErrorType.Critical})


                .ScheduleTriggerFormDxValidationProvider.SetValidationRule(HolidayTimeCheckedComboBoxEdit,
                                                                           New BlankComboCustomValidation With {.ErrorText = "Write your desired time then click Add Time button",
                                                                                                                .ErrorType = ErrorType.Critical})

                .ScheduleTriggerFormDxValidationProvider.SetValidationRule(HolidayMonthCheckedComboBoxEdit,
                                                                           New BlankComboCustomValidation With {.ErrorText = "Select repeat month or uncheck Repeat Month checkbox.",
                                                                                                                .ErrorType = ErrorType.Critical})

                .ScheduleTriggerFormDxValidationProvider.SetValidationRule(HolidayDayOfWeekCheckedComboBoxEdit,
                                                                           New BlankComboCustomValidation With {.ErrorText = "Select repeat day of week or uncheck Repeat Day Of Week checkbox.",
                                                                                                                .ErrorType = ErrorType.Critical})
            End With
        End Sub

        'More info on custom validation
        'http://documentation.devexpress.com/#WindowsForms/clsDevExpressXtraEditorsDXErrorProviderDXValidationProvidertopic
        Friend Class BlankComboCustomValidation
            Inherits ValidationRule

            Public Overrides Function Validate(control As System.Windows.Forms.Control, value As Object) As Boolean
                With ScheduleForm
                    Select Case control.Name
                        Case "EndDateTimeDateEdit"
                            If Not IsNothing(value) AndAlso CDate(.EndDateTimeDateEdit.EditValue) <= CDate(.StartDateTimeDateEdit.EditValue) Then
                                Return False
                            End If

                        Case "RepeatDayOfMonthCheckedComboBoxEdit"
                            If .RepeatDayOfMonthCheckEdit.Checked Then
                                If value.ToString = vbNullString Then
                                    Connectors.ScheduleTriggerFormConnector.SetSelectedTab(ScheduleTriggerForm.AdvancedScheduleBackstageViewTabItem)
                                    Return False
                                End If
                            End If

                        Case "RepeatTimeCheckedComboBoxEdit"
                            If .RepeatTimeCheckEdit.Checked Then
                                Dim noItemChecked As Boolean = True
                                'We make check first that there is only one item checked in Times combo box.
                                'because of the nature of this control that may have a value but it is not checked.
                                For Each itm As CheckedListBoxItem In .RepeatTimeCheckedComboBoxEdit.Properties.Items
                                    If itm.CheckState = CheckState.Checked Then noItemChecked = False
                                Next
                                If noItemChecked Then
                                    Connectors.ScheduleTriggerFormConnector.SetSelectedTab(ScheduleTriggerForm.AdvancedScheduleBackstageViewTabItem)
                                    Return False
                                End If
                            End If

                        Case "RepeatMonthCheckedComboBoxEdit"
                            If .RepeatMonthCheckEdit.Checked Then
                                If value.ToString = vbNullString Then
                                    Connectors.ScheduleTriggerFormConnector.SetSelectedTab(ScheduleTriggerForm.AdvancedScheduleBackstageViewTabItem)
                                    Return False
                                End If
                            End If

                        Case "RepeatDayOfWeekCheckedComboBoxEdit"
                            If .RepeatDayOfWeekCheckEdit.Checked Then
                                If value.ToString = vbNullString Then
                                    Connectors.ScheduleTriggerFormConnector.SetSelectedTab(ScheduleTriggerForm.AdvancedScheduleBackstageViewTabItem)
                                    Return False
                                End If
                            End If

                        Case "HolidayTimeCheckedComboBoxEdit"
                            If .HolidayTimeCheckEdit.Checked Then
                                Dim noItemChecked As Boolean = True
                                'We make check first that there is only one item checked in Times combo box.
                                'because of the nature of this control that may have a value but it is not checked.
                                For Each itm As CheckedListBoxItem In .HolidayTimeCheckedComboBoxEdit.Properties.Items
                                    If itm.CheckState = CheckState.Checked Then noItemChecked = False
                                Next
                                If noItemChecked Then
                                    Connectors.ScheduleTriggerFormConnector.SetSelectedTab(ScheduleTriggerForm.HolidaysBackstageViewTabItem)
                                    Return False
                                End If
                            End If

                        Case "HolidayDayOfMonthCheckedComboBoxEdit"
                            If .HolidayDayOfMonthCheckEdit.Checked Then
                                If value.ToString = vbNullString Then
                                    Connectors.ScheduleTriggerFormConnector.SetSelectedTab(ScheduleTriggerForm.HolidaysBackstageViewTabItem)
                                    Return False
                                End If
                            End If

                        Case "HolidayMonthCheckedComboBoxEdit"
                            If .HolidayMonthCheckEdit.Checked Then
                                If value.ToString = vbNullString Then
                                    Connectors.ScheduleTriggerFormConnector.SetSelectedTab(ScheduleTriggerForm.HolidaysBackstageViewTabItem)
                                    Return False
                                End If
                            End If

                        Case "HolidayDayOfWeekCheckedComboBoxEdit"
                            If .HolidayDayOfWeekCheckEdit.Checked Then
                                If value.ToString = vbNullString Then
                                    Connectors.ScheduleTriggerFormConnector.SetSelectedTab(ScheduleTriggerForm.HolidaysBackstageViewTabItem)
                                    Return False
                                End If
                            End If

                    End Select
                End With
                Return True
            End Function

        End Class
#End Region
#End Region
    End Class
End Namespace
