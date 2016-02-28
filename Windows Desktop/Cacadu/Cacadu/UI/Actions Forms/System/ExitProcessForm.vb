Imports Cacadore
Imports Cacadore.Actions.System
Imports DevExpress.XtraEditors.Controls
Imports Cacadore.Actions.System.ExitProcess
Imports DevExpress.XtraEditors.DXErrorProvider
Imports Cacadu.Connectors.CommonActionConnector
Imports Cacadu.Connectors.ExitProcessFormConnector

Namespace UI
    Friend Class ExitProcessActionForm

#Region "Properties"
        Property TargetTask As Task
        Property IsLoaded As Boolean = False
        Property IsEditMode As Boolean = False
        Property CreatedAction As New ExitProcess
#End Region

#Region "Form Events"
        Private Sub ActionForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
            Me.DxValidationProvider.SetValidationRule(RunAnotherTaskComboBoxEdit,
                                                      New BlankComboCustomValidation With {.ErrorText = "Select alternative task name.",
                                                      .ErrorType = ErrorType.Critical})
            SetIconToCacaduIcon(Me)
            LoadMRUFileName()
            LoadRunAnotherTaskComboBoxEdit(RunAnotherTaskComboBoxEdit)
            LoadPossibleErrors(CreatedAction,
                               PossibleErrorsListBoxControl)
            ActionsSettingBackstageViewTabItem.Selected = True
            If IsEditMode Then
                FillControlsForEdit()
            Else
                CheckActionInputs()
            End If
            IsLoaded = True
        End Sub

        Private Sub ExitProcessForm_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
            'Dispose rather than close to avoid a pre-loaded controls.
            If Not Me.IsDisposed Then Me.Dispose()
        End Sub
#End Region

#Region "Inputs Controls Events"
        Private Sub FileNameMruEdit_Enter(sender As System.Object, e As System.EventArgs) Handles ProcessNameMruEdit.Enter
            LoadMRUFileName()
        End Sub

        Private Sub FileNameMruEdit_EditValueChanged(sender As System.Object, e As System.EventArgs) Handles ProcessNameMruEdit.EditValueChanged
            ToggleFileNameBoxEnable()
        End Sub

        Private Sub FileNameMruEdit_ButtonClick(sender As System.Object, e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs) Handles ProcessNameMruEdit.ButtonClick
            If e.Button.Index = 1 Then LoadMRUFileName()
        End Sub

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
                If IsLoaded Then SaveErrorsSettings(Me,
                                                   PossibleErrorsListBoxControl,
                                                   CreatedAction,
                                                   OnErrorsRadioGroup,
                                                   RetryActionSpinEdit,
                                                   RunAnotherTaskComboBoxEdit)
            End If
        End Sub

        Private Sub WaitSpecificPeriodSpinEdit_EditValueChanging(sender As Object, e As DevExpress.XtraEditors.Controls.ChangingEventArgs) Handles WaitSpecificPeriodSpinEdit.EditValueChanging
            WaitSpecificPeriodCheckEdit.CheckState = CheckState.Checked
        End Sub
#End Region

#Region "Buttons Events"
        Private Sub OKExitProcessSimpleButton_Click(sender As System.Object, e As System.EventArgs) Handles OKExitProcessSimpleButton.Click
            If Not DxValidationProvider.Validate() Then Exit Sub

            Dim actionInputs As ControlsData = CollectFormInputs()
            CreateAction(actionInputs)
            SaveCreatedAction(CreatedAction, TargetTask)
            Me.Dispose()
        End Sub

        Private Sub CancelButton_Click(sender As System.Object, e As System.EventArgs) Handles CancelSimpleButton.Click
            If Not Me.IsDisposed Then Me.Dispose()
        End Sub
#End Region

        Friend Class BlankComboCustomValidation
            Inherits ValidationRule

            Public Overrides Function Validate(control As System.Windows.Forms.Control, value As Object) As Boolean
                Select Case control.Name
                    Case "RunAnotherTaskComboBoxEdit"
                        If IsNothing(value) And ExitProcessActionForm.OnErrorsRadioGroup.SelectedIndex = 3 Then
                            Return False
                        End If
                End Select
                Return True
            End Function
        End Class
    End Class
End Namespace
