Option Strict Off

#Region "Imports"
Imports Cacadore
Imports Cacadu.UI
Imports DevExpress.XtraEditors
Imports DevExpress.XtraBars.Ribbon
Imports DevExpress.XtraEditors.Controls
Imports Balora

#End Region

Namespace Connectors
    Friend Class CommonActionConnector

#Region "Fields"
        Private Shared alertnativeTasksDic As Dictionary(Of Integer, String)
#End Region

#Region "Methods"
        Shared Sub SaveErrorsSettings(actionForm As XtraForm,
                                     possibleErrorsControl As ListBoxControl,
                                     CreatedAction As Cacadore.Action,
                                     onErrorsRadioGroup As RadioGroup,
                                     retryActionSpinEdit As SpinEdit,
                                     runAnotherTaskComboBoxEdit As ComboBoxEdit)
            With actionForm
                Dim exceptionHandler As ActionExceptionHandler = CreatedAction.ExceptionsManager.GetSelectedExceptionHandler(possibleErrorsControl.SelectedItem)
                Dim selectedRadioItem = onErrorsRadioGroup.Properties.Items(onErrorsRadioGroup.SelectedIndex)

                If selectedRadioItem.Description.Contains("Alert me") Then
                    exceptionHandler.AlertOnError = True
                ElseIf selectedRadioItem.Description.Contains("Retry Action") Then
                    exceptionHandler.RetryAction = True
                    exceptionHandler.RetryTimes = CInt(retryActionSpinEdit.EditValue)
                ElseIf selectedRadioItem.Description.Contains("Ignore") Then
                    exceptionHandler.ContinueOnError = True
                ElseIf selectedRadioItem.Description.Contains("Run Another") Then
                    exceptionHandler.AlternativeTaskId = alertnativeTasksDic.Values(runAnotherTaskComboBoxEdit.SelectedIndex)
                    exceptionHandler.RunAnotherTaskOnError = True
                ElseIf selectedRadioItem.Description.Contains("Quit Task") Then
                    exceptionHandler.QuitOnError = True
                End If
            End With
        End Sub

        Shared Sub LoadErrorSettings(actionForm As XtraForm,
                                    possibleErrorsControl As ListBoxControl,
                                    createdAction As Cacadore.Action,
                                    onErrorsRadioGroup As RadioGroup,
                                    retryActionSpinEdit As SpinEdit,
                                    runAnotherTaskComboBoxEdit As ComboBoxEdit)
            For Each possibleException In createdAction.ExceptionsManager.PossibleExceptions
                Dim errorString As String = possibleException.Key
                Dim selectedIndex As Integer = 1

                For Each itm In possibleErrorsControl.Items
                    If itm = errorString Then
                        selectedIndex = possibleErrorsControl.Items.IndexOf(itm)
                    End If
                Next

                If selectedIndex <> -1 Then
                    possibleErrorsControl.SelectedIndex = selectedIndex
                    SetErrorsRadioGroupSettings(possibleException.Value,
                                                onErrorsRadioGroup,
                                                retryActionSpinEdit,
                                                runAnotherTaskComboBoxEdit)
                End If
            Next
        End Sub

        Shared Sub SetErrorsRadioGroupSettings(possibleException As ActionExceptionHandler,
                                               onErrorsRadioGroup As RadioGroup,
                                               retryActionSpinEdit As SpinEdit,
                                               runAnotherTaskComboBoxEdit As ComboBoxEdit)
            If possibleException.AlertOnError Then
                onErrorsRadioGroup.SelectedIndex = 0
                retryActionSpinEdit.EditValue = 1
                runAnotherTaskComboBoxEdit.SelectedItem = Nothing
            ElseIf possibleException.RetryAction Then
                onErrorsRadioGroup.SelectedIndex = 1
                retryActionSpinEdit.EditValue = possibleException.RetryTimes
            ElseIf possibleException.ContinueOnError Then
                onErrorsRadioGroup.SelectedIndex = 2
                retryActionSpinEdit.EditValue = 1
                runAnotherTaskComboBoxEdit.SelectedItem = Nothing
            ElseIf possibleException.RunAnotherTaskOnError Then
                onErrorsRadioGroup.SelectedIndex = 3
                Dim tskId As String = possibleException.AlternativeTaskId
                Dim taskName As String = CacadoreToTectonic.GetTaskName(tskId)
                runAnotherTaskComboBoxEdit.SelectedItem = taskName
            ElseIf possibleException.QuitOnError Then
                onErrorsRadioGroup.SelectedIndex = 4
                retryActionSpinEdit.EditValue = 1
                runAnotherTaskComboBoxEdit.SelectedItem = Nothing
            End If
        End Sub

        Shared Sub LoadRunAnotherTaskComboBoxEdit(runAnotherTaskComboBoxEdit As ComboBoxEdit)
            Dim allTasksDic = CacadoreToTectonic.GetAllRegisteredTasks()
            alertnativeTasksDic = New Dictionary(Of Integer, String)()
            For Each tsk In allTasksDic.Values
                Dim index As Integer = runAnotherTaskComboBoxEdit.Properties.Items.Add(tsk.Properties.TaskName)
                alertnativeTasksDic.Add(index, tsk.TaskId)
            Next
        End Sub

        Shared Sub LoadPossibleErrors(createdAction As Cacadore.Action,
                                      possibleErrorsCheckedListBoxControl As ListBoxControl)
            Dim itemsArray = createdAction.ExceptionsManager.PossibleExceptionsDescriptions.ToArray
            For Each itm In itemsArray
                'Set the default value for every possible error. This value represent a selected 
                'radio in OnErrorsRadioGroup
                possibleErrorsCheckedListBoxControl.Items.Add(itm)
            Next
        End Sub

        Shared Sub SetSelectedTab(actionFormBackstage As DevExpress.XtraBars.Ribbon.BackstageViewControl,
                                  tab As DevExpress.XtraBars.Ribbon.BackstageViewTabItem)
            actionFormBackstage.SelectedTab = tab
        End Sub

        Shared Function CheckOutputVariable(OutputMruEdit As MRUEdit) As String
            Dim variable As String = vbNullString
            If Not IsNothing(OutputMruEdit.EditValue) Then
                Dim isGlobalVariable As Boolean = Cacadore.Commands.IsRegisteredGlobalVariable(OutputMruEdit.EditValue)
                Dim isTaskVariable As Boolean = Cacadore.Commands.IsRegisteredTaskVariable(LastFacebookNotificationActionForm.TargetTask.TaskId,
                                                                                           OutputMruEdit.EditValue)
                If isGlobalVariable OrElse isTaskVariable Then variable = OutputMruEdit.EditValue
            End If
            Return variable
        End Function

        Shared Sub LoadOutputVariable(ByRef saveToMember As String, ByRef OutputMruEdit As MRUEdit)
            If saveToMember = vbNullString Then Exit Sub
            Dim isGlobalVariable As Boolean = Cacadore.Commands.IsRegisteredGlobalVariable(saveToMember)
            Dim isTaskVariable As Boolean = Cacadore.Commands.IsRegisteredTaskVariable(LastFacebookNotificationActionForm.TargetTask.TaskId, saveToMember)
            If isGlobalVariable OrElse isTaskVariable Then OutputMruEdit.EditValue = saveToMember
        End Sub

        Shared Sub TweakCommentsTextBox(commentsMemoEdit As DevExpress.XtraEditors.MemoEdit)
            If Not IsNothing(commentsMemoEdit.EditValue) Then
                If commentsMemoEdit.EditValue.ToString = "Comments..." Then
                    commentsMemoEdit.EditValue = ""
                End If
                commentsMemoEdit.ForeColor = Color.Black
            End If
        End Sub

        Shared Sub LoadVariables(mrueditControl As MRUEdit, targetTask As Cacadore.Task)
            Dim taskVariables = CacadoreToTectonic.GetTaskVariables(targetTask)
            mrueditControl.Properties.Items.AddRange(taskVariables.InnerDictionary.Keys)
        End Sub

        Shared Sub LoadGlobalMembers(mrueditControl As MRUEdit, Optional byBassConstants As Boolean = False)
            If Not byBassConstants Then
                Dim globalConstants = CacadoreToTectonic.GetGlobals(Cacadore.IGlobals.GlobalType.Constant)
                mrueditControl.Properties.Items.AddRange(globalConstants.Keys)
            End If
            Dim globalVariables = CacadoreToTectonic.GetGlobals(Cacadore.IGlobals.GlobalType.Variable)
            mrueditControl.Properties.Items.AddRange(globalVariables.Keys)
        End Sub

        Shared Sub LoadEnviromentsVars(mrueditControl As MRUEdit)
            Dim enviromentVariablePaths As SortedList(Of String, String) = PathsHelper.GetExecutablesNamesAndPathsInPathEnvVarFolders()
            mrueditControl.Properties.Items.AddRange(enviromentVariablePaths.Keys.ToArray)
        End Sub

        Shared Sub LoadProcessNames(mrueditControl As MRUEdit, targetTask As Cacadore.Task)
            Dim processNamesList As List(Of String) = Cacadore.Actions.System.ExitProcess.GetRunningProcessesNames()

            mrueditControl.Properties.Items.AddRange(processNamesList)
        End Sub

        Shared Sub LoadVariablesAndConstantsAndEnviromentsVars(mrueditControl As MRUEdit,
                                                               ByVal targetTask As Cacadore.Task,
                                                               Optional byPassEnviromentVariables As Boolean = False,
                                                               Optional byPassConstants As Boolean = False)
            LoadVariables(mrueditControl, targetTask)
            LoadGlobalMembers(mrueditControl, byPassConstants)
            If Not byPassEnviromentVariables Then
                LoadEnviromentsVars(mrueditControl)
            End If
        End Sub

        Shared Sub SetDelaySettings(delayValue As Decimal,
                                    continueImmediatelyCheckEdit As CheckEdit,
                                    WaitLoadOnlyCheckEdit As CheckEdit,
                                    waitSpecificPeriodCheckEdit As CheckEdit,
                                    waitAppExitCheckEdit As CheckEdit,
                                    waitSpecificPeriodSpinEdit As SpinEdit)
            Select Case delayValue
                Case Is > 1
                    waitSpecificPeriodCheckEdit.Checked = True
                    waitSpecificPeriodSpinEdit.EditValue = delayValue / 1000
                Case 0
                    waitAppExitCheckEdit.Checked = True
                Case -1
                    continueImmediatelyCheckEdit.Checked = True
                Case -2
                    WaitLoadOnlyCheckEdit.Checked = True
            End Select
        End Sub

        Shared Function GetDelaySettings(continueImmediatelyCheckEdit As CheckEdit,
                                         waitSpecificPeriodCheckEdit As CheckEdit,
                                         waitAppExitCheckEdit As CheckEdit,
                                         waitLoadOnlyCheckEdit As CheckEdit,
                                         waitSpecificPeriodSpinEdit As SpinEdit) As Decimal
            If waitSpecificPeriodCheckEdit.Checked Then
                Return CLng(waitSpecificPeriodSpinEdit.EditValue)
            ElseIf continueImmediatelyCheckEdit.Checked Then
                'Return CDec(continueImmediatelyCheckEdit.Tag) 
                Return -1
            ElseIf waitLoadOnlyCheckEdit.Checked Then
                'Return CLng(waitLoadOnlyCheckEdit.Tag)
                Return -2
            ElseIf waitAppExitCheckEdit.Checked Then
                Return CLng(waitAppExitCheckEdit.Tag)
            Else
                Return 1
            End If
        End Function

        Shared Sub SaveCreatedAction([action] As Cacadore.Action, targetTask As Cacadore.Task)
            Dim fullDescription As String = [action].GetFullDescription
            Dim icoIndex As Integer = Connectors.MainFormConnector.GetActionIconIndex([action].GetType.Name)

            'Now attach/update a real action for the first time. 30/04/2012 12:24 AM 
            If [action].ActionId = vbNullString Then
                Dim actionId As String = targetTask.AttachAction([action])
                If actionId <> "" Then TaskFlowPresenter.AddActionListItem([action], fullDescription, icoIndex)
            Else
                targetTask.UpdateAction([action])
                Dim actionId As String = [action].ActionId
                Dim actionIndex As Integer = TaskFlowPresenter.GetSelectedListItemIndex()
                If actionId <> "" Then TaskFlowPresenter.EditActionListItem(actionIndex, fullDescription, icoIndex)
            End If
        End Sub
#End Region
    End Class
End Namespace

