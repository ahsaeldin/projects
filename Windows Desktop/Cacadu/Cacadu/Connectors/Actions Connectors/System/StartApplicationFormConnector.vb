Imports Balora
Imports Cacadu.UI
Imports DevExpress.XtraEditors
Imports Cacadore.Actions.System
Imports DevExpress.XtraBars.Ribbon
Imports DevExpress.XtraEditors.Controls
Imports Cacadu.Connectors.CommonActionConnector
Imports Cacadore.Actions.System.StartApplication

Namespace Connectors
    Friend Class StartApplicationFormConnector

        Structure ControlsData
            Public IsActionEnabled As Boolean
            Public IsSimpleSet As Boolean
            Public FileName As String
            Public Verb As String
            Public Arguments As String
            Public WorkingDirectory As String
            Public WindowStyle As ProcessWindowStyle
            Public Comments As String
            Public Delay As Decimal
        End Structure

#Region "BackstageView"
        Shared Function IsSimpleSet() As Boolean
            Return Not IsAdvancedSet()
        End Function

        Shared Function IsAdvancedSet() As Boolean
            With StartApplicationForm
                Dim verbsSelectedIndex = .VerbsComboBoxEdit.SelectedIndex
                Dim argumentsTextEditValue = .ArgumentsTextEdit.EditValue
                Dim workingDirectoryTextEditValue = .WorkingDirectoryTextEdit.EditValue
                Dim windowStyleComboBoxEditSelectedIndex = .WindowStyleComboBoxEdit.SelectedIndex
                Dim commentsMemoEditValue = .CommentsMemoEdit.EditValue
                If verbsSelectedIndex = -1 AndAlso
                    argumentsTextEditValue Is Nothing AndAlso
                    workingDirectoryTextEditValue Is Nothing AndAlso
                    windowStyleComboBoxEditSelectedIndex = -1 AndAlso
                    commentsMemoEditValue.ToString = "Comments..." Then
                    Return False
                End If
            End With
            Return True
        End Function

        Shared Function IsSimpleAction(action As StartApplication) As Boolean
            Return Not IsAdvancedAction(action)
        End Function

        Shared Function IsAdvancedAction(action As StartApplication) As Boolean
            Dim actionInputs As StartApplicationInputs = CType(action.Inputs, StartApplicationInputs)
            If IsNothing(actionInputs) Then Return False
            With actionInputs
                If (Not IsNothing(.Verb) AndAlso .Verb <> "") OrElse
                   (Not IsNothing(.Arguments) AndAlso .Arguments <> "") OrElse
                   (Not IsNothing(.WorkingDirectory) AndAlso .WorkingDirectory <> "") Then
                    Return True
                End If
            End With
            Return False
        End Function
#End Region

#Region "MruEdit Helpers"
        Shared Sub ProcessMruEditButton(e As DevExpress.XtraEditors.Controls.ButtonPressedEventArgs)
            Dim IsEllipsedButton As Boolean = e.Button.Tag.ToString.Contains("ShowOpen")
            With StartApplicationForm
                If IsEllipsedButton Then
                    Dim res As DialogResult = .OpenFileDialog.ShowDialog()
                    If res = Windows.Forms.DialogResult.OK Then
                        .FileNameMruEdit.EditValue = .OpenFileDialog.FileName
                        MoveStartingPointToTheEnd()
                    End If
                ElseIf e.Button.Index = 2 Then
                    LoadVariablesAndConstantsAndEnviromentsVars(.FileNameMruEdit, .TargetTask)
                End If
            End With
        End Sub

        Shared Sub MoveStartingPointToTheEnd()
            With StartApplicationForm
                .FileNameMruEdit.SelectionStart = .FileNameMruEdit.EditValue.ToString.Length
            End With
        End Sub
#End Region

#Region "Inputs Controls Helpers"
        Shared Sub MoveCommonControlsToTab(tab As BackstageViewClientControl)
            With StartApplicationForm
                Dim isAlreadyHaveCommonControls As Boolean = tab.Controls.Contains(.FileNameMruEdit)
                If Not isAlreadyHaveCommonControls Then
                    tab.Controls.Add(.FileNameMruEdit)
                    tab.Controls.Add(.FileNameLabelControl)
                    tab.Controls.Add(.ActionEnabledCheckEdit)
                End If
            End With
        End Sub

        Shared Sub CheckActionInputs()
            With StartApplicationForm
                If IsNothing(.CreatedAction.Inputs) Then
                    If Not IsNothing(.FileNameMruEdit.EditValue) Then
                        .CreatedAction.Inputs = New StartApplicationInputs(.FileNameMruEdit.EditValue.ToString)
                    Else
                        .CreatedAction.Inputs = New StartApplicationInputs("plz set file name.")
                    End If
                End If
            End With
        End Sub

        Shared Sub SetFileName()
            With StartApplicationForm
                Dim fileName As String = vbNullString
                If Not IsNothing(.FileNameMruEdit.EditValue) Then fileName = .FileNameMruEdit.EditValue.ToString
                If IsNothing(.CreatedAction.Inputs) Then
                    .CreatedAction.Inputs = New StartApplicationInputs(fileName)
                End If
            End With
        End Sub

        Shared Sub LoadVerbs()
            CheckActionInputs()
            With StartApplicationForm
                .VerbsComboBoxEdit.Properties.Items.Clear()
                .VerbsComboBoxEdit.Properties.Items.AddRange(.CreatedAction.Verbs)
            End With
        End Sub

        Shared Sub SetWorkingDirectory(dirPath As String)
            CheckActionInputs()
            With StartApplicationForm
                Dim inputs As StartApplicationInputs = CType(.CreatedAction.Inputs, StartApplicationInputs)
                inputs.WorkingDirectory = dirPath
            End With
        End Sub

        Shared Sub FillControlsForEdit()
            With StartApplicationForm
                If .IsEditMode Then
                    Dim input As StartApplicationInputs = CType(.CreatedAction.Inputs, StartApplicationInputs)

                    .FileNameMruEdit.EditValue = input.FileName
                    .ActionEnabledCheckEdit.EditValue = .CreatedAction.Enabled

                    If IsAdvancedAction(.CreatedAction) Then
                        LoadVerbs()
                        .VerbsComboBoxEdit.SelectedIndex = .VerbsComboBoxEdit.Properties.Items.IndexOf(input.Verb)
                        .ArgumentsTextEdit.EditValue = input.Arguments
                        .WorkingDirectoryTextEdit.EditValue = input.WorkingDirectory
                        .WindowStyleComboBoxEdit.SelectedIndex = .WindowStyleComboBoxEdit.Properties.Items.IndexOf(input.WindowStyle)
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

        Shared Sub SetSelectedBackStageTab()
            With StartApplicationForm
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
#End Region

#Region "Creating Action Helpers"
        Shared Sub CreateAction(actionInputs As ControlsData)
            If IsSimpleSet() Then
                CreateSimpleStartApplicationAction(actionInputs)
            Else
                CreateAdvancedStartApplicationAction(actionInputs)
            End If
            'اوعي تنسي تكوبي دة لما تعمل اكشن جديد.
            StartApplicationForm.CreatedAction.Enabled = actionInputs.IsActionEnabled 'Now CreateAction is ready, just one more thing.
        End Sub

        Shared Function CollectFormInputs() As ControlsData
            Dim actionInputs As New ControlsData
            With StartApplicationForm
                actionInputs.IsSimpleSet = IsSimpleSet()
                actionInputs.IsActionEnabled = .ActionEnabledCheckEdit.Checked
                If Not IsNothing(.FileNameMruEdit.EditValue) Then actionInputs.FileName = .FileNameMruEdit.EditValue.ToString
                If Not IsNothing(.VerbsComboBoxEdit.EditValue) Then actionInputs.Verb = .VerbsComboBoxEdit.EditValue.ToString
                If Not IsNothing(.ArgumentsTextEdit.EditValue) Then actionInputs.Arguments = .ArgumentsTextEdit.EditValue.ToString
                If Not IsNothing(.CommentsMemoEdit.EditValue) Then actionInputs.Comments = .CommentsMemoEdit.EditValue.ToString
                actionInputs.WindowStyle = CType(.WindowStyleComboBoxEdit.SelectedIndex, ProcessWindowStyle)
                actionInputs.Delay = GetDelaySettings(.ContinueImmediatelyCheckEdit,
                                                      .WaitSpecificPeriodCheckEdit,
                                                      .WaitAppExitCheckEdit,
                                                      .WaitLoadOnlyCheckEdit,
                                                      .WaitSpecificPeriodSpinEdit)
            End With
            Return actionInputs
        End Function

        Shared Sub CreateSimpleStartApplicationAction(_controlsData As ControlsData)
            CheckActionInputs()
            Dim actionInputs As StartApplicationInputs

            With StartApplicationForm
                actionInputs = CType(.CreatedAction.Inputs, StartApplicationInputs)
                .CreatedAction.Comments = _controlsData.Comments
            End With

            With actionInputs
                .FileName = _controlsData.FileName
                If _controlsData.Delay >= 0 Then
                    .WaitForExitPeriod = _controlsData.Delay * 1000
                Else
                    .WaitForExitPeriod = _controlsData.Delay
                End If
            End With
        End Sub

        Shared Sub CreateAdvancedStartApplicationAction(_controlsData As ControlsData)
            CheckActionInputs()
            Dim actionInputs As StartApplicationInputs

            With StartApplicationForm
                actionInputs = CType(.CreatedAction.Inputs, StartApplicationInputs)
                .CreatedAction.Comments = _controlsData.Comments
            End With

            With actionInputs
                .FileName = _controlsData.FileName
                If _controlsData.Delay >= 0 Then
                    .WaitForExitPeriod = _controlsData.Delay * 1000
                Else
                    .WaitForExitPeriod = _controlsData.Delay
                End If
                .Verb = _controlsData.Verb
                .Arguments = _controlsData.Arguments
                .WindowStyle = _controlsData.WindowStyle
            End With
        End Sub
#End Region
    End Class
End Namespace

