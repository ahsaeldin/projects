Imports Cacadu.UI
Imports Cacadore.Actions.System
Imports DevExpress.XtraBars.Ribbon
Imports Cacadore.Actions.System.ShutdownPC
Imports Cacadu.Connectors.CommonActionConnector

Namespace Connectors
    Friend Class ShutdownPCFormConnector

        Structure ControlsData
            Public Comments As String
            Public IsEnabled As Boolean
            Public ShutdownType As Balora.ShutdownTypeEnum
            Public IsForced As Boolean
            Public ConfirmBeforeExecution As Boolean
            Public HasPeriodBeforeExecution As Boolean
            Public PeriodBeforeExecution As Integer
            Public AllowCancel As Boolean
        End Structure

        Shared Sub SetSelectedBackStageTab()
            With ShutdownPCForm
                If .IsEditMode AndAlso IsAdvancedAction(.CreatedAction) Then
                    MoveCommonControlsToTab(.AdvancedBackstageViewClientControl)
                    Connectors.CommonActionConnector.SetSelectedTab(.ActionFormBackstageViewControl,
                                                                    .AdvancedBackstageViewTabItem)
                Else
                    Connectors.CommonActionConnector.SetSelectedTab(.ActionFormBackstageViewControl,
                                                                    .SimpleBackstageViewTabItem)
                End If
            End With
        End Sub

        Shared Function IsAdvancedAction(action As ShutdownPC) As Boolean
            Dim actionInputs As ShutdownPCInputs = CType(action.Inputs, ShutdownPCInputs)
            If IsNothing(actionInputs) Then Return False
            With actionInputs
                If (.IsForced) OrElse
                   (.AllowCancel) OrElse
                   (.PeriodBeforeShutdown > 0) OrElse
                   ShutdownPCForm.CommentsMemoEdit.EditValue.ToString = "" AndAlso ShutdownPCForm.CommentsMemoEdit.EditValue.ToString = "Comments..." Then
                    Return True
                End If
            End With
            Return False
        End Function

        Shared Sub MoveCommonControlsToTab(tab As BackstageViewClientControl)
            With ShutdownPCForm
                Dim isAlreadyHaveCommonControls As Boolean = tab.Controls.Contains(.ShutdownTypeComboBoxEdit)
                If Not isAlreadyHaveCommonControls Then
                    tab.Controls.Add(.ShutdownTypeComboBoxEdit)
                    tab.Controls.Add(.ShutdownTypeLabelControl)
                End If
            End With
        End Sub

        Shared Sub CheckActionInputs()
            With ShutdownPCForm
                If IsNothing(.CreatedAction.Inputs) Then
                    .CreatedAction.Inputs = New ShutdownPCInputs()
                End If
            End With
        End Sub

        Shared Function ConvertShutdownEnumToShutdownCombo(shutdownEnum As Balora.ShutdownTypeEnum) As Integer
            Dim shutdownType As String = shutdownEnum.ToString.Replace("Forced", "")
            Return ShutdownPCForm.ShutdownTypeComboBoxEdit.Properties.Items.IndexOf(shutdownType)
        End Function

        Shared Function ConvertShutdownComboToShutdownEnum() As Balora.ShutdownTypeEnum
            Select Case ShutdownPCForm.ShutdownTypeComboBoxEdit.SelectedItem.ToString
                Case "LogOff"
                    Return Balora.ShutdownTypeEnum.LogOff
                Case "Hibernate"
                    Return Balora.ShutdownTypeEnum.Hibernate
                Case "Restart"
                    Return Balora.ShutdownTypeEnum.Restart
                Case "Shutdown"
                    Return Balora.ShutdownTypeEnum.Shutdown
                Case "Sleep"
                    Return Balora.ShutdownTypeEnum.Sleep
                Case Else
                    Return Nothing
            End Select
        End Function

        Shared Sub FillControlsForEdit()
            With ShutdownPCForm
                If .IsEditMode Then
                    Dim input As ShutdownPCInputs = CType(.CreatedAction.Inputs, ShutdownPCInputs)

                    .ShutdownTypeComboBoxEdit.SelectedIndex = ConvertShutdownEnumToShutdownCombo(input.ShutdownType)
                    .IsForcedCheckEdit.Checked = input.IsForced
                    .ConfirmBeforeCheckEdit.Checked = input.ConfirmBeforeExecute
                    .AllowCancelCheckEdit.Checked = input.AllowCancel

                    If input.PeriodBeforeShutdown > 0 Then
                        .BeforePeriodCheckEdit.Checked = True
                        .BeforeSpinEdit.EditValue = input.PeriodBeforeShutdown
                    Else
                        .BeforePeriodCheckEdit.Checked = False
                    End If

                    If Not IsNothing(.CreatedAction.Comments) AndAlso .CreatedAction.Comments <> "N/A" Then .CommentsMemoEdit.EditValue = .CreatedAction.Comments

                    LoadErrorSettings(StartApplicationForm,
                                     .PossibleErrorsListBoxControl,
                                     .CreatedAction,
                                     .OnErrorsRadioGroup,
                                     .RetryActionSpinEdit,
                                     .RunAnotherTaskComboBoxEdit)

                    .ActionEnabledCheckEdit.EditValue = .CreatedAction.Enabled
                End If
            End With
        End Sub

        Shared Function CollectFormInputs() As ControlsData
            With ShutdownPCForm
                Dim actionInputs As New ControlsData
                actionInputs.ShutdownType = ConvertShutdownComboToShutdownEnum()
                actionInputs.IsForced = .IsForcedCheckEdit.Checked
                actionInputs.AllowCancel = .AllowCancelCheckEdit.Checked
                actionInputs.ConfirmBeforeExecution = .ConfirmBeforeCheckEdit.Checked
                actionInputs.HasPeriodBeforeExecution = .BeforePeriodCheckEdit.Checked
                If .BeforePeriodCheckEdit.Checked AndAlso Not IsNothing(.BeforeSpinEdit.EditValue) Then
                    actionInputs.PeriodBeforeExecution = CInt(.BeforeSpinEdit.EditValue)
                Else
                    actionInputs.PeriodBeforeExecution = 0
                End If
                If Not IsNothing(.CommentsMemoEdit.EditValue) Then actionInputs.Comments = .CommentsMemoEdit.EditValue.ToString
                actionInputs.IsEnabled = .ActionEnabledCheckEdit.Checked
                Return actionInputs
            End With
        End Function

        Shared Sub CreateAction(_controlsData As ControlsData)
            CheckActionInputs()
            Dim actionInputs As ShutdownPCInputs

            With ShutdownPCForm
                actionInputs = CType(.CreatedAction.Inputs, ShutdownPCInputs)
                .CreatedAction.Comments = _controlsData.Comments
                .CreatedAction.Enabled = _controlsData.IsEnabled 'Now CreateAction is ready, just one more thing.
            End With

            With actionInputs
                .ShutdownType = _controlsData.ShutdownType
                .IsForced = _controlsData.IsForced
                .ConfirmBeforeExecute = _controlsData.ConfirmBeforeExecution
                .AllowCancel = _controlsData.AllowCancel
                .PeriodBeforeShutdown = _controlsData.PeriodBeforeExecution
            End With
        End Sub
    End Class
End Namespace

