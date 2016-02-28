Imports Cacadu.UI
Imports Cacadu.Connectors.CommonActionConnector
Imports Actions.FacebookActions.LastFacebookNotification
Imports DevExpress.XtraEditors

Namespace Connectors
    Public Class LastFacebookNotificationActionConnector

        Structure ControlsData
            Public Delay As Decimal
            Public Comments As String
            Public SaveOutputTo As String
            Public IsActionEnabled As Boolean
            Public LoginWithDifferentAccount As Boolean
        End Structure

        Shared Sub CheckActionInputs()
            With LastFacebookNotificationActionForm
                If IsNothing(.CreatedAction.Inputs) Then
                    .CreatedAction.Inputs = New LastFacebookNotificationInputs
                End If
            End With
        End Sub

        Shared Sub FillControlsForEdit()
            With LastFacebookNotificationActionForm
                If .IsEditMode Then
                    Dim input As LastFacebookNotificationInputs = CType(.CreatedAction.Inputs, LastFacebookNotificationInputs)
                    Dim output As LastFacebookNotificationOutputs = CType(.CreatedAction.Outputs, LastFacebookNotificationOutputs)

                    .LoginWithADifferentAccountCheckEdit.Checked = input.LoginInAnotherAccount

                    LoadOutputVariable(output.SaveToMember, .OutputMruEdit)

                    If Not IsNothing(.CreatedAction.Comments) AndAlso .CreatedAction.Comments <> "N/A" Then .CommentsMemoEdit.EditValue = .CreatedAction.Comments

                    SetDelaySettings(input.WaitForExitPeriod,
                                     .ContinueImmediatelyCheckEdit,
                                     .WaitForLoginCheckEdit,
                                     .WaitSpecificPeriodCheckEdit,
                                     .WaitForNotificationReceviedCheckEdit,
                                     .WaitSpecificPeriodSpinEdit)

                    LoadErrorSettings(LastFacebookNotificationActionForm,
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
            With LastFacebookNotificationActionForm
                Dim actionInputs As New ControlsData
                actionInputs.Comments = .CommentsMemoEdit.EditValue.ToString
                actionInputs.IsActionEnabled = .ActionEnabledCheckEdit.Checked
                actionInputs.Delay = GetDelaySettings(.ContinueImmediatelyCheckEdit,
                                                      .WaitSpecificPeriodCheckEdit,
                                                      .WaitForNotificationReceviedCheckEdit,
                                                      .WaitForLoginCheckEdit,
                                                      .WaitSpecificPeriodSpinEdit)
                actionInputs.LoginWithDifferentAccount = .LoginWithADifferentAccountCheckEdit.Checked

                actionInputs.SaveOutputTo = CheckOutputVariable(.OutputMruEdit)
                Return actionInputs
            End With
        End Function

        Shared Sub CreateAction(_controlsData As ControlsData)
            CheckActionInputs()
            Dim actionInputs As LastFacebookNotificationInputs
            Dim actionOutputs As LastFacebookNotificationOutputs

            With LastFacebookNotificationActionForm
                actionInputs = CType(.CreatedAction.Inputs, LastFacebookNotificationInputs)
                actionOutputs = CType(.CreatedAction.Outputs, LastFacebookNotificationOutputs)
                .CreatedAction.Comments = _controlsData.Comments
                .CreatedAction.Enabled = _controlsData.IsActionEnabled 'Now CreateAction is ready, just one more thing.
            End With

            With actionInputs
                .LoginInAnotherAccount = _controlsData.LoginWithDifferentAccount
                If _controlsData.Delay >= 0 Then
                    .WaitForExitPeriod = _controlsData.Delay * 1000
                Else
                    .WaitForExitPeriod = _controlsData.Delay
                End If
            End With

            If Not IsNothing(_controlsData.SaveOutputTo) Then
                actionOutputs.SaveToMember = _controlsData.SaveOutputTo
                actionOutputs.ParentAction = LastFacebookNotificationActionForm.CreatedAction
            End If

        End Sub
    End Class
End Namespace