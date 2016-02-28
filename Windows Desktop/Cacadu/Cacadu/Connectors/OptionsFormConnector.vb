Imports Cacadu.UI
Namespace Connectors
    Friend NotInheritable Class OptionsFormConnector

        Private Sub New()
            ': "Because type 'OptionsFormConnector' contains only 
            'static' ('Shared' in Visual Basic) members, add a 
            'default private constructor to prevent the compiler 
            'from adding a default public constructor."
            'http://msdn.microsoft.com/library/ms182169(VS.100).aspx 
        End Sub

        Shared Sub LoadOptionForm()
            With OptionsForm
                .MaximizeBox = False
                .MinimizeBox = False
                SetIconToCacaduIcon(OptionsForm)
                LoadOptions()
            End With
        End Sub

        Shared Sub LoadOptions()
            LoadCacaduEnabledOption()
            LoadErrorReporterOption()
            LoadStartupOption()
            LoadDisableSplashScreenOption()
            LoadSendToSystrayOption()
            LoadAutomaticallyStartTriggersOption()
            LoadAskMeFirstBeforeResumingTriggersOption()
            LoadAutomaticallyCloseAlertsOption()
            loadDefaultAlertsPosition()
        End Sub

        Shared Sub SaveOptions()
            SetCacaduEnabledOption()
            SetErrorReporterOption()
            SetStartupOption()
            SetDisabledSplashScreenOption()
            SetOnClosingRadioGroupOption()
            SetAutomaticallyStartTriggersOption()
            SetAskMeFirstBeforeResumingTriggersOption()
            SetAutomaticallyCloseAlertsOption()
            SetDefaultAlertsPosition()
            My.Settings.Save()
        End Sub

#Region "Loading Helpers"
        Shared Sub LoadCacaduEnabledOption()
            With OptionsForm
                If Cacadore.Settings.Enabled Then
                    .CacaduEnabledCheckEdit.Checked = True
                Else
                    .CacaduEnabledCheckEdit.Checked = False
                End If
            End With
        End Sub

        Shared Sub LoadErrorReporterOption()
            OptionsForm.SendErrorReportsCheckEdit.Checked = My.Settings.CheckBoxAutoSendState
        End Sub

        Shared Sub LoadStartupOption()
            With OptionsForm
                .WindowStartupCheckEdit.Checked = My.Settings.ST
                If My.Settings.ST AndAlso Not IsShortcutInPlace(System.Environment.GetFolderPath(Environment.SpecialFolder.Startup), "Cacadu") Then
                    CreateShortcut("Cacadu",
                                   System.Environment.GetFolderPath(Environment.SpecialFolder.Startup),
                                   Application.ExecutablePath,
                                   Application.StartupPath,
                                   Application.ExecutablePath,
                                   0)
                End If
            End With
        End Sub

        Shared Sub LoadDisableSplashScreenOption()
            OptionsForm.DisableSplashScreenCheckEdit.Checked = Not My.Settings.SS
        End Sub

        Shared Sub LoadSendToSystrayOption()
            With OptionsForm
                If My.Settings.STST Then
                    .OnClosingRadioGroup.SelectedIndex = 0
                Else
                    .OnClosingRadioGroup.SelectedIndex = 1
                End If
            End With
        End Sub

        Shared Sub LoadAutomaticallyStartTriggersOption()
            OptionsForm.AutoStartTriggersCheckEdit.Checked = My.Settings.AST
        End Sub

        Shared Sub LoadAskMeFirstBeforeResumingTriggersOption()
            OptionsForm.AskMeFirstBeforeResumingTriggersCheckEdit.Checked = My.Settings.AMFBRTAS
        End Sub

        Private Shared Sub LoadAutomaticallyCloseAlertsOption()
            OptionsForm.AutoCloseAlertsCheckEdit.Checked = My.Settings.ACA
        End Sub

        Private Shared Sub loadDefaultAlertsPosition()
            OptionsForm.DefaultpositionForCacaduAlertComboBoxEdit.SelectedIndex = My.Settings.DPFCA
        End Sub
#End Region

#Region "Settings Helpers"
        Shared Sub SetCacaduEnabledOption()
            If Cacadore.Settings.Enabled <> OptionsForm.CacaduEnabledCheckEdit.Checked Then MainFormConnector.ToggleCacaduEnable()
        End Sub

        Shared Sub SetErrorReporterOption()
            My.Settings.CheckBoxAutoSendState = OptionsForm.SendErrorReportsCheckEdit.Checked
        End Sub

        Shared Sub SetStartupOption()
            My.Settings.ST = OptionsForm.WindowStartupCheckEdit.Checked
            If My.Settings.ST Then
                CreateShortcut("Cacadu",
                               System.Environment.GetFolderPath(Environment.SpecialFolder.Startup),
                               Application.ExecutablePath,
                               Application.StartupPath,
                               Application.ExecutablePath,
                               0)
            Else
                DeleteStartupShortcut()
            End If
        End Sub

        Shared Sub SetDisabledSplashScreenOption()
            My.Settings.SS = Not OptionsForm.DisableSplashScreenCheckEdit.Checked
        End Sub

        Shared Sub SetOnClosingRadioGroupOption()
            If OptionsForm.OnClosingRadioGroup.SelectedIndex = 0 Then
                My.Settings.STST = True
            Else
                My.Settings.STST = False
            End If
        End Sub

        Private Shared Sub SetAutomaticallyStartTriggersOption()
            My.Settings.AST = OptionsForm.AutoStartTriggersCheckEdit.Checked
        End Sub

        Private Shared Sub SetAskMeFirstBeforeResumingTriggersOption()
            My.Settings.AMFBRTAS = OptionsForm.AskMeFirstBeforeResumingTriggersCheckEdit.Checked
        End Sub

        Private Shared Sub SetAutomaticallyCloseAlertsOption()
            My.Settings.ACA = OptionsForm.AutoCloseAlertsCheckEdit.Checked
        End Sub

        Private Shared Sub SetDefaultAlertsPosition()
            My.Settings.DPFCA = OptionsForm.DefaultpositionForCacaduAlertComboBoxEdit.SelectedIndex
        End Sub
#End Region

#Region "Legacy code need to be re-written"
        Shared Function IsShortcutInPlace(ByVal creationDir As String, ByVal shortcutName As String) As Boolean
            Dim IsShortcutExists As Boolean = IO.File.Exists(creationDir & "\" & shortcutName & ".lnk")
            Return IsShortcutExists
        End Function

        Shared Function CreateShortcut(ByVal shortcutName As String,
                                       ByVal creationDir As String,
                                       ByVal targetFullpath As String,
                                       ByVal workingDir As String,
                                       ByVal iconFile As String,
                                       ByVal iconNumber As Integer,
                                       Optional ByVal IsFailed As Boolean = False) As Boolean
            'How to use:-
            'CreateShortCut("Perfect Macro Recorder", System.Environment.GetFolderPath(Environment.SpecialFolder.Startup), Application.ExecutablePath, Application.StartupPath, Application.ExecutablePath, 0)
            Try
                If Not IO.Directory.Exists(creationDir) Then
                    Dim retVal As MsgBoxResult = MsgBox(creationDir & " doesn't exist. Do you wish to create it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo)
                    If retVal = DialogResult.Yes Then
                        IO.Directory.CreateDirectory(creationDir)
                    Else
                        Return False
                    End If
                End If

                Dim IsShortcutExists As Boolean = IsShortcutInPlace(creationDir, shortcutName)

                '//escape creating the shortcut if it is already created.
                If IsShortcutExists Then
                    Return False
                End If

                Dim wShell As New IWshRuntimeLibrary.WshShell
                Dim shortCut As IWshRuntimeLibrary.IWshShortcut
                shortCut = CType(wShell.CreateShortcut(creationDir & "\" & shortcutName & ".lnk"), IWshRuntimeLibrary.IWshShortcut)
                shortCut.TargetPath = targetFullpath
                shortCut.WindowStyle = 1
                shortCut.Arguments = "min"
                shortCut.Description = shortcutName
                shortCut.WorkingDirectory = workingDir
                shortCut.IconLocation = iconFile & ", " & iconNumber
                shortCut.Save()

                '//shortpth == shortcut path, saving the shortcut path to the registry in order to use directly when deleteing the shortcut using DeleteShortcut
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Cacadu\opt\", "shortpth", creationDir & "\" & shortcutName & ".lnk")

                Return True

            Catch ex As System.Exception
                Balora.Alerter.REP("Error creating startup shortcut.", ex, True, False)
                IsFailed = True
                Return False
            End Try

        End Function

        Shared Function DeleteStartupShortcut() As Boolean
            Try

                Dim IsShortcutExists As Boolean
                Dim StartupShortcutPath As String

                StartupShortcutPath = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Cacadu\opt\", "shortpth", "").ToString

                If StartupShortcutPath = "" Then Return False

                IsShortcutExists = IO.File.Exists(StartupShortcutPath)

                If Not IsShortcutExists Then Return False

                Try
                    IO.File.Delete(StartupShortcutPath)
                Catch ex As System.Exception
                    Balora.Alerter.REP("Error deleting startup shortcut.", ex, True, False)
                End Try

                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Cacadu\opt\", "shortpth", "")

                IsShortcutExists = IO.File.Exists(StartupShortcutPath)

                If Not IsShortcutExists Then '//Check if delation successed
                    Return True
                Else
                    Return False
                End If

            Catch ex As Exception
                Balora.Alerter.REP("Error executing DeleteStartupShortcut method", ex, True, False)
                Return False
            End Try
        End Function
#End Region
    End Class
End Namespace

