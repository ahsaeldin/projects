Imports Balora
Imports Cacadu.UI
Imports DevExpress.XtraEditors
Imports DevExpress.XtraBars.Ribbon
Imports Cacadore.Actions.MessageBoxes
Imports DevExpress.XtraEditors.Controls
Imports Cacadu.Connectors.CommonActionConnector
Imports Cacadore.Actions.MessageBoxes.ShowMessageBox

Namespace Connectors
    Public Class ShowMessageBoxFormConnector

        Structure ControlsData
            Public IsActionEnabled As Boolean
            Public Message As String
            Public Title As String
            Public MessageType As MsgBoxStyle
            Public Comments As String
            Public Delay As Decimal
        End Structure

        Shared Sub MoveCommonControlsToTab(tab As BackstageViewClientControl)
            With ShowMessageBoxForm
                Dim isAlreadyHaveCommonControls As Boolean = tab.Controls.Contains(.MessageBoxMemoEdit)
                If Not isAlreadyHaveCommonControls Then
                    tab.Controls.Add(.MessageBoxMemoEdit)
                    tab.Controls.Add(.MessageBoxLabelControl)
                    tab.Controls.Add(.ActionEnabledCheckEdit)
                End If
            End With
        End Sub

        Shared Sub SetSelectedBackStageTab()
            With ShowMessageBoxForm
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

        Shared Function IsAdvancedAction(action As ShowMessageBox) As Boolean
            Dim actionInputs As ShowMessageBoxInputs = CType(action.Inputs, ShowMessageBoxInputs)
            If IsNothing(actionInputs) Then Return False
            With actionInputs
                If .MessageBoxType <> MsgBoxStyle.Critical OrElse
                   (action.Comments <> "" AndAlso action.Comments <> "Comments...") Then
                    Return True
                End If
            End With
            Return False
        End Function

        Shared Sub FillControlsForEdit()
            With ShowMessageBoxForm
                If .IsEditMode Then
                    Dim input As ShowMessageBoxInputs = CType(.CreatedAction.Inputs, ShowMessageBoxInputs)

                    .MessageBoxMemoEdit.EditValue = input.Message
                    .ActionEnabledCheckEdit.EditValue = .CreatedAction.Enabled

                    If IsAdvancedAction(.CreatedAction) Then
                        .CommentsMemoEdit.EditValue = .CreatedAction.Comments
                        .MessageBoxTitleTextEdit.EditValue = input.Title

                        Select Case input.MessageBoxType
                            Case MsgBoxStyle.Critical
                                .MessageBoxTypeComboBoxEdit.SelectedIndex = 0
                            Case MsgBoxStyle.Question
                                .MessageBoxTypeComboBoxEdit.SelectedIndex = 1
                            Case MsgBoxStyle.Exclamation
                                .MessageBoxTypeComboBoxEdit.SelectedIndex = 2
                            Case MsgBoxStyle.Information
                                .MessageBoxTypeComboBoxEdit.SelectedIndex = 3
                        End Select

                    End If

                    If Not IsNothing(.CreatedAction.Comments) AndAlso .CreatedAction.Comments <> "N/A" Then .CommentsMemoEdit.EditValue = .CreatedAction.Comments

                    SetDelaySettings(input.WaitForExitPeriod,
                                     .ContinueImmediatelyCheckEdit,
                                     .WaitLoadOnlyCheckEdit,
                                     .WaitSpecificPeriodCheckEdit,
                                     .WaitMessageBoxExitCheckEdit,
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
            With ShowMessageBoxForm
                If IsNothing(.CreatedAction.Inputs) Then
                    If Not IsNothing(.MessageBoxMemoEdit.EditValue) Then
                        .CreatedAction.Inputs = New ShowMessageBoxInputs(.MessageBoxMemoEdit.EditValue.ToString)
                    Else
                        .CreatedAction.Inputs = New ShowMessageBoxInputs("No alert message specifed.")
                    End If
                End If
            End With
        End Sub

        Shared Function CollectFormInputs() As ControlsData
            With ShowMessageBoxForm
                Dim actionInputs As New ControlsData
                If Not IsNothing(.MessageBoxTitleTextEdit.EditValue) Then actionInputs.Title = .MessageBoxTitleTextEdit.EditValue.ToString
                Select Case .MessageBoxTypeComboBoxEdit.SelectedIndex
                    Case 0
                        actionInputs.MessageType = MsgBoxStyle.Critical
                    Case 0
                        actionInputs.MessageType = MsgBoxStyle.Question
                    Case 0
                        actionInputs.MessageType = MsgBoxStyle.Exclamation
                    Case 0
                        actionInputs.MessageType = MsgBoxStyle.Information
                End Select
                actionInputs.MessageType = CType(.MessageBoxTypeComboBoxEdit.SelectedIndex, MsgBoxStyle)
                If Not IsNothing(.MessageBoxMemoEdit.EditValue) Then actionInputs.Message = .MessageBoxMemoEdit.EditValue.ToString
                If Not IsNothing(.CommentsMemoEdit.EditValue) Then actionInputs.Comments = .CommentsMemoEdit.EditValue.ToString
                actionInputs.IsActionEnabled = .ActionEnabledCheckEdit.Checked
                actionInputs.Delay = GetDelaySettings(.ContinueImmediatelyCheckEdit,
                                                      .WaitSpecificPeriodCheckEdit,
                                                      .WaitMessageBoxExitCheckEdit,
                                                      .WaitLoadOnlyCheckEdit,
                                                      .WaitSpecificPeriodSpinEdit)
                Return actionInputs
            End With
        End Function

        Shared Sub CreateAction(_controlsData As ControlsData)
            CheckActionInputs()
            Dim actionInputs As ShowMessageBoxInputs

            With ShowMessageBoxForm
                actionInputs = CType(.CreatedAction.Inputs, ShowMessageBoxInputs)
                .CreatedAction.Comments = _controlsData.Comments
                .CreatedAction.Enabled = _controlsData.IsActionEnabled 'Now CreateAction is ready, just one more thing.
            End With

            With actionInputs
                .Title = _controlsData.Title
                .Message = _controlsData.Message
                .MessageBoxType = _controlsData.MessageType
                If _controlsData.Delay >= 0 Then
                    .WaitForExitPeriod = _controlsData.Delay * 1000
                Else
                    .WaitForExitPeriod = _controlsData.Delay
                End If
            End With
        End Sub
    End Class
End Namespace
