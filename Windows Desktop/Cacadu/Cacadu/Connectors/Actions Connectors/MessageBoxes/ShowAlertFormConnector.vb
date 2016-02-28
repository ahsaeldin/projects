Imports Balora
Imports Cacadu.UI
Imports DevExpress.XtraEditors
Imports DevExpress.XtraBars.Ribbon
Imports Cacadore.Actions.MessageBoxes
Imports DevExpress.XtraEditors.Controls
Imports Cacadu.Connectors.CommonActionConnector
Imports Cacadore.Actions.MessageBoxes.ShowAlert

Namespace Connectors
    Public Class ShowAlertFormConnector

        Structure ControlsData
            Public IsActionEnabled As Boolean
            Public AlertMessage As String
            Public Comments As String
            Public AlertLocation As ShowAlertsInputs.AlertLocationEnum
            Public Delay As Decimal
        End Structure

        Shared Sub MoveCommonControlsToTab(tab As BackstageViewClientControl)
            With ShowAlertActionForm
                Dim isAlreadyHaveCommonControls As Boolean = tab.Controls.Contains(.AlertMessageMemoEdit)
                If Not isAlreadyHaveCommonControls Then
                    tab.Controls.Add(.AlertMessageMemoEdit)
                    tab.Controls.Add(.AlertMessageLabelControl)
                    tab.Controls.Add(.ActionEnabledCheckEdit)
                End If
            End With
        End Sub

        Shared Sub SetSelectedBackStageTab()
            With ShowAlertActionForm
                If .IsEditMode And IsAdvancedAction(.CreatedAction) Then
                    MoveCommonControlsToTab(.AdvancedBackstageViewClientControl)
                    Connectors.CommonActionConnector.SetSelectedTab(.ActionFormBackstageViewControl,
                                                                    .AdvancedBackstageViewTabItem)
                Else
                    Connectors.CommonActionConnector.SetSelectedTab(.ActionFormBackstageViewControl,
                                                                    .SimpleBackstageViewTabItem)
                End If
            End With
        End Sub

        Shared Function IsAdvancedAction(action As ShowAlert) As Boolean
            Dim actionInputs As ShowAlertsInputs = CType(action.Inputs, ShowAlertsInputs)
            If IsNothing(actionInputs) Then Return False
            With actionInputs
                If .AlertLocation <> ShowAlertsInputs.AlertLocationEnum.BottomLeft OrElse
                   (action.Comments <> "" AndAlso action.Comments <> "Comments...") Then
                    Return True
                End If
            End With
            Return False
        End Function

        Shared Sub FillControlsForEdit()
            With ShowAlertActionForm
                If .IsEditMode Then
                    Dim input As ShowAlertsInputs = CType(.CreatedAction.Inputs, ShowAlertsInputs)

                    .AlertMessageMemoEdit.EditValue = input.Message
                    .ActionEnabledCheckEdit.EditValue = .CreatedAction.Enabled

                    If IsAdvancedAction(.CreatedAction) Then
                        .CommentsMemoEdit.EditValue = .CreatedAction.Comments
                        .AlertFormLocationComboBoxEdit.SelectedIndex = input.AlertLocation
                    End If

                    If Not IsNothing(.CreatedAction.Comments) AndAlso .CreatedAction.Comments <> "N/A" Then .CommentsMemoEdit.EditValue = .CreatedAction.Comments

                    SetDelaySettings(input.WaitForExitPeriod,
                                     .ContinueImmediatelyCheckEdit,
                                     .WaitLoadOnlyCheckEdit,
                                     .WaitSpecificPeriodCheckEdit,
                                     .WaitAppExitCheckEdit,
                                     .WaitSpecificPeriodSpinEdit)

                    LoadErrorSettings(StartApplicationForm,
                                     .PossibleErrorsListBoxControl,
                                     .CreatedAction,
                                     .OnErrorsRadioGroup,
                                     .RetryActionSpinEdit,
                                     .RunAnotherTaskComboBoxEdit)
                End If
            End With
        End Sub

        Shared Sub CheckActionInputs()
            With ShowAlertActionForm
                If IsNothing(.CreatedAction.Inputs) Then
                    If Not IsNothing(.AlertMessageMemoEdit.EditValue) Then
                        .CreatedAction.Inputs = New ShowAlertsInputs(.AlertMessageMemoEdit.EditValue.ToString)
                    Else
                        .CreatedAction.Inputs = New ShowAlertsInputs("No alert message specifed.")
                    End If
                End If
            End With
        End Sub

        Shared Function CollectFormInputs() As ControlsData
            With ShowAlertActionForm
                Dim actionInputs As New ControlsData
                actionInputs.AlertLocation = CType(.AlertFormLocationComboBoxEdit.SelectedIndex, ShowAlertsInputs.AlertLocationEnum)
                actionInputs.AlertMessage = .AlertMessageMemoEdit.EditValue.ToString
                actionInputs.Comments = .CommentsMemoEdit.EditValue.ToString
                actionInputs.IsActionEnabled = .ActionEnabledCheckEdit.Checked
                actionInputs.Delay = GetDelaySettings(.ContinueImmediatelyCheckEdit,
                                                      .WaitSpecificPeriodCheckEdit,
                                                      .WaitAppExitCheckEdit,
                                                      .WaitLoadOnlyCheckEdit,
                                                      .WaitSpecificPeriodSpinEdit)
                Return actionInputs
            End With
        End Function

        Shared Sub CreateAction(_controlsData As ControlsData)
            CheckActionInputs()
            Dim actionInputs As ShowAlertsInputs

            With ShowAlertActionForm
                actionInputs = CType(.CreatedAction.Inputs, ShowAlertsInputs)
                .CreatedAction.Comments = _controlsData.Comments
                .CreatedAction.Enabled = _controlsData.IsActionEnabled 'Now CreateAction is ready, just one more thing.
            End With

            With actionInputs
                .AlertLocation = _controlsData.AlertLocation
                .Message = _controlsData.AlertMessage
                If _controlsData.Delay >= 0 Then
                    .WaitForExitPeriod = _controlsData.Delay * 1000
                Else
                    .WaitForExitPeriod = _controlsData.Delay
                End If
            End With
        End Sub
    End Class
End Namespace
