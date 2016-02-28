Imports Cacadore
Imports Cacadore.Actions.System
Imports DevExpress.XtraEditors.Controls
Imports Cacadore.Actions.System.ShutdownPC
Imports DevExpress.XtraEditors.DXErrorProvider
Imports Cacadu.Connectors.CommonActionConnector
Imports Cacadu.Connectors.ShutdownPCFormConnector

Namespace UI
    Friend Class ShutdownPCForm

#Region "Properties"
        Property TargetTask As Task
        Property IsLoaded As Boolean = False
        Property IsEditMode As Boolean = False
        Property CreatedAction As New ShutdownPC
#End Region

#Region "Form Events"
        Private Sub ActionForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
            SetValidationRules()
            SetIconToCacaduIcon(Me)
            SetSelectedBackStageTab()
            LoadRunAnotherTaskComboBoxEdit(RunAnotherTaskComboBoxEdit)
            LoadPossibleErrors(CreatedAction, PossibleErrorsListBoxControl)
            BeforeSpinEdit.DataBindings.Add("Enabled", BeforePeriodCheckEdit, "Checked")
            If IsEditMode Then
                FillControlsForEdit()
            Else
                CheckActionInputs()
            End If
            IsLoaded = True
        End Sub

        Private Sub StartApplicationForm_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
            'Dispose rather than close to avoid a pre-loaded controls.
            If Not Me.IsDisposed Then Me.Dispose()
        End Sub
#End Region

#Region "Backstage Events"
        Private Sub ActionFormBackstageViewControl_SelectedTabChanged(sender As System.Object, e As DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs) Handles ActionFormBackstageViewControl.SelectedTabChanged
            Select Case e.Item.Name
                Case "SimpleBackstageViewTabItem"
                    MoveCommonControlsToTab(SimpleBackstageViewClientControl)
                Case "AdvancedBackstageViewTabItem"
                    MoveCommonControlsToTab(AdvancedBackstageViewClientControl)
            End Select
        End Sub
#End Region

#Region "Inputs Controls Events"
        Private Sub CommentsMemoEdit_MouseDown(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles CommentsMemoEdit.MouseDown
            TweakCommentsTextBox(CommentsMemoEdit)
        End Sub

        Private Sub PossibleErrorsListBoxControl_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles PossibleErrorsListBoxControl.SelectedIndexChanged
            If PossibleErrorsListBoxControl.SelectedItems.Count = 0 Then
                OnErrorsRadioGroup.Enabled = False
            Else
                OnErrorsRadioGroup.Enabled = True
            End If
            Dim exceptionHandler = CreatedAction.ExceptionsManager.GetSelectedExceptionHandler(PossibleErrorsListBoxControl.SelectedValue.ToString)
            SetErrorsRadioGroupSettings(exceptionHandler,
                                        OnErrorsRadioGroup,
                                        RetryActionSpinEdit,
                                        RunAnotherTaskComboBoxEdit)
        End Sub

        Private Sub OnErrorsRadioGroup_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles OnErrorsRadioGroup.SelectedIndexChanged, RunAnotherTaskComboBoxEdit.EditValueChanged, RetryActionSpinEdit.EditValueChanged
            If Me.IsHandleCreated Then
                Select Case sender.GetType.Name
                    Case "SpinEdit"
                        OnErrorsRadioGroup.SelectedIndex = 1
                    Case "ComboBoxEdit"
                        If Not IsNothing(RunAnotherTaskComboBoxEdit.EditValue) Then OnErrorsRadioGroup.SelectedIndex = 3
                End Select
                If IsLoaded Then
                    SaveErrorsSettings(Me,
                                       PossibleErrorsListBoxControl,
                                       CreatedAction,
                                       OnErrorsRadioGroup,
                                       RetryActionSpinEdit,
                                       RunAnotherTaskComboBoxEdit)
                End If
            End If
        End Sub

        Private Sub WaitSpecificPeriodSpinEdit_EditValueChanging(sender As Object, e As DevExpress.XtraEditors.Controls.ChangingEventArgs) Handles WaitSpecificPeriodSpinEdit.EditValueChanging
            WaitSpecificPeriodCheckEdit.CheckState = CheckState.Checked
        End Sub

        Private Sub ShutdownTypeComboBoxEdit_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ShutdownTypeComboBoxEdit.SelectedIndexChanged
            If ShutdownTypeComboBoxEdit.SelectedIndex >= 0 Then
                ShutdownPCOKSimpleButton.Enabled = True
            Else
                ShutdownPCOKSimpleButton.Enabled = False
            End If
        End Sub
#End Region

#Region "Buttons Events"
        Private Sub OKButton_Click(sender As System.Object, e As System.EventArgs) Handles ShutdownPCOKSimpleButton.Click
            If Not DxValidationProvider.Validate() Then Exit Sub

            Dim actionInputs As ControlsData = CollectFormInputs()

            CreateAction(actionInputs)

            SaveCreatedAction(CreatedAction, TargetTask)

            Me.Dispose()
        End Sub

        Private Sub CancelButton_Click(sender As System.Object, e As System.EventArgs) Handles CancelSimpleButton.Click
            If Not Me.IsDisposed Then Me.Dispose()
        End Sub

        Private Sub SetValidationRules()
            Me.DxValidationProvider.SetValidationRule(RunAnotherTaskComboBoxEdit,
                                                      New BlankComboCustomValidation With {.ErrorText = "Select alternative task name.",
                                                      .ErrorType = ErrorType.Critical})

            Me.DxValidationProvider.SetValidationRule(BeforeSpinEdit,
                                                      New BlankComboCustomValidation With {.ErrorText = "Set a period before execution.",
                                                      .ErrorType = ErrorType.Critical})
        End Sub
#End Region

        Friend Class BlankComboCustomValidation
            Inherits ValidationRule

            Public Overrides Function Validate(control As System.Windows.Forms.Control, value As Object) As Boolean
                Select Case control.Name
                    Case "RunAnotherTaskComboBoxEdit"
                        If IsNothing(value) And LastFacebookNotificationActionForm.OnErrorsRadioGroup.SelectedIndex = 3 Then
                            Return False
                        End If
                    Case "BeforeSpinEdit"
                        If IsNothing(value) Then
                            Return False
                        End If
                End Select
                Return True
            End Function
        End Class

    End Class
End Namespace
