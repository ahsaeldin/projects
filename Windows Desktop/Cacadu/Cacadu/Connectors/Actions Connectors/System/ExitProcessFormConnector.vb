Imports Cacadu.UI
Imports Cacadore.Actions.System.ExitProcess
Imports Cacadu.Connectors.CommonActionConnector

Namespace Connectors
    Friend Class ExitProcessFormConnector

        Structure ControlsData
            Public IsActionEnabled As Boolean
            Public FileName As String
            Public Comments As String
            Public Delay As Decimal
        End Structure

        Shared Sub LoadMRUFileName()
            With ExitProcessActionForm
                .ProcessNameMruEdit.Properties.Items.Clear()
                LoadVariables(.ProcessNameMruEdit, .TargetTask)
                LoadGlobalMembers(.ProcessNameMruEdit)
                LoadProcessNames(.ProcessNameMruEdit, .TargetTask)
            End With
        End Sub

        Shared Sub FillControlsForEdit()
            With ExitProcessActionForm
                If .IsEditMode Then
                    Dim input As ExitProcessInputs = CType(.CreatedAction.Inputs, ExitProcessInputs)

                    .ProcessNameMruEdit.EditValue = input.ProcessName
                    .ActionEnabledCheckEdit.EditValue = .CreatedAction.Enabled

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
            With ExitProcessActionForm
                If IsNothing(.CreatedAction.Inputs) Then
                    If Not IsNothing(.ProcessNameMruEdit.EditValue) Then
                        .CreatedAction.Inputs = New ExitProcessInputs(.ProcessNameMruEdit.EditValue.ToString)
                    Else
                        .CreatedAction.Inputs = New ExitProcessInputs("plz set file name.")
                    End If
                End If
            End With
        End Sub

        Shared Sub ToggleFileNameBoxEnable()
            With ExitProcessActionForm
                If Not IsNothing(.ProcessNameMruEdit.EditValue) AndAlso .ProcessNameMruEdit.EditValue.ToString <> "" Then
                    .OKExitProcessSimpleButton.Enabled = True
                Else
                    .OKExitProcessSimpleButton.Enabled = False
                End If
            End With
        End Sub

#Region "On OK Button"
        Shared Function CollectFormInputs() As ControlsData
            Dim actionInputs As New ControlsData
            With ExitProcessActionForm
                actionInputs.IsActionEnabled = .ActionEnabledCheckEdit.Checked
                If Not IsNothing(.ProcessNameMruEdit.EditValue) Then actionInputs.FileName = .ProcessNameMruEdit.EditValue.ToString
                If Not IsNothing(.CommentsMemoEdit.EditValue) Then actionInputs.Comments = .CommentsMemoEdit.EditValue.ToString
                actionInputs.Delay = GetDelaySettings(.ContinueImmediatelyCheckEdit,
                                                      .WaitSpecificPeriodCheckEdit,
                                                      .WaitAppExitCheckEdit,
                                                      .WaitLoadOnlyCheckEdit,
                                                      .WaitSpecificPeriodSpinEdit)
            End With
            Return actionInputs
        End Function

        Shared Sub CreateAction(_controlsData As ControlsData)
            CheckActionInputs()
            With ExitProcessActionForm
                Dim actionInputs As ExitProcessInputs = CType(.CreatedAction.Inputs, ExitProcessInputs)
                With actionInputs
                    .ProcessName = _controlsData.FileName
                    If _controlsData.Delay >= 0 Then
                        .WaitForExitPeriod = _controlsData.Delay * 1000
                    Else
                        .WaitForExitPeriod = _controlsData.Delay
                    End If
                End With
                .CreatedAction.Comments = _controlsData.Comments
                .CreatedAction.Enabled = _controlsData.IsActionEnabled
            End With
        End Sub
#End Region
    End Class
End Namespace

