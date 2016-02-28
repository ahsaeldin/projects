Imports CuteMacro.Balora

Friend Class CMOptionsForm

    '//19 Nov 2008: 
    '//ana mabsoot awy eny 3rft 7d zy Hagar.
    '//Ya try maw9'ow3 nene da hay7'9 3ly eh, ana 8rft..!
    '//ana 7ass isA 2n elbernamg da ha7geeb flos 7elwa isA..!

    '//20 Nov 2008: How to reuse the HotKey Functionalty
    '1. Add TextBox_KeyDown & KeyUp & Key Press
    '2. Add TxtBxsKyUp to KeyUp & ProcessHotKeyTxtBox to KeyDown & e.KeyChar = "" to KeyPress
    '3. Add RegHotKey to OkButton_Click for ex: IsReged = RegHotKey(txtStopRec, &HBFF2&)
    '4. Add RegMyHotKeys to MainForm_Load for ex : RegMyHotKeys("txtStrRec", &HBFF1&) 
    '5. Wait for and process the HotKey in WndProc
    '6. LoadHotKeys in OptionsForm_Load

    Private Sub OptionsForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadHotKeys(txtStrRec)
        LoadHotKeys(txtStopRec)
        LoadHotKeys(txtStrPly)
        LoadHotKeys(txtStopPly)

        'chkMouMov == chkMM
        'chkAllMouActs = chkAM
        'chkKeyActs = chkKA
        chkMouMov.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt", "chkMM", 1)
        chkAllMouActs.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkAM", 1)
        chkKeyActs.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkKA", 1)

        '//AR == auto record upon start up
        chkAutoRecord.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkAR", 0)

		'If is_x2342 = "True" Then
		'//DF == defualt Folder
		chkDefFold.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkDF", 0)
		'End If

		'//txtDFP = txtDefFoldPath.text
		lblPath.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "txtDFP", My.Computer.FileSystem.SpecialDirectories.Desktop)

		'//txtDFN = txtDefMacName.Text
		txtDefMacName.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "txtDFN", "default.pmac")

		'//chkTips = chkTips.Checked
		chkTips.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkTips", 1)

		'//chkStartup == chkStartup.Checked
		chkStartup.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkStartup", 0)

		chkNotify.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkNotify", 1)

		chkCloseCM.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkCloseCM", 0)

		'//txtRT  == txtRepeatTimes
		txtRepeatTimes.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "txtRT", 1)

		'//spdsld ==  PlaSpdSlider.Value 
		PlaSpdSlider.Value = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "spdsld", 2)

		ChkSmartPlayback.Checked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "ChkSmartPlayback", 1)

		If chkAllMouActs.Checked Then
			chkMouMov.Checked = True
			chkMouMov.Enabled = False
		End If

    End Sub

    Private Sub LoadHotKeys(ByVal txtBox As TextBox)
        txtBox.Text = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & txtBox.Name, "HotKey", "")
    End Sub

    Private Sub txtStrRec_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtStrRec.KeyDown
        ProcessHotKeyTxtBox(e, txtStrRec)
    End Sub

    Private Sub txtStrRec_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtStrRec.KeyPress
        '//handle the case of doubleing the char like this "aA" or "bB",,, note : it appears in char(s) without modifires only
        e.KeyChar = ""
    End Sub

    Private Sub txtStrRec_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtStrRec.KeyUp
        TxtBxsKyUp(e, txtStrRec)
    End Sub

    Private Sub txtStopRec_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtStopRec.KeyDown
        ProcessHotKeyTxtBox(e, txtStopRec)
    End Sub

    Private Sub txtStopRec_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtStopRec.KeyPress
        '//handle the case of doubleing the char like this "aA" or "bB",,, note : it appears in char(s) without modifires only
        e.KeyChar = ""
    End Sub

    Private Sub txtStopRec_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtStopRec.KeyUp
        TxtBxsKyUp(e, txtStopRec)
    End Sub

    Private Sub txtStrPly_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtStrPly.KeyDown
        ProcessHotKeyTxtBox(e, txtStrPly)
    End Sub

    Private Sub txtStrPly_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtStrPly.KeyPress
        '//handle the case of doubleing the char like this "aA" or "bB",,, note : it appears in char(s) without modifires only
        e.KeyChar = ""
    End Sub

    Private Sub txtStrPly_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtStrPly.KeyUp
        TxtBxsKyUp(e, txtStrPly)
    End Sub

    Private Sub txtStopPly_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtStopPly.KeyDown
        ProcessHotKeyTxtBox(e, txtStopPly)
    End Sub

    Private Sub txtStopPly_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtStopPly.KeyPress
        '//handle the case of doubleing the char like this "aA" or "bB",,, note : it appears in char(s) without modifires only
        e.KeyChar = ""
    End Sub

    Private Sub txtStopPly_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtStopPly.KeyUp
        TxtBxsKyUp(e, txtStopPly)
    End Sub

    Private Sub OkButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click

        Dim IsReged1 As Boolean, IsReged2 As Boolean, IsReged3 As Boolean, IsReged4 As Boolean

        ValidatetxtDefMacName()

        'chkMouMov == chkMM
        'chkAllMouActs = chkAM
        'chkKeyActs = chkKA
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkMM", chkMouMov.Checked)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkAM", chkAllMouActs.Checked)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkKA", chkKeyActs.Checked)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkAR", chkAutoRecord.Checked)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkDF", chkDefFold.Checked)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "txtDFP", lblPath.Text)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "txtDFN", txtDefMacName.Text)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkTips", chkTips.Checked)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkStartup", chkStartup.Checked)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkNotify", chkNotify.Checked)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkCloseCM", chkCloseCM.Checked)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "txtRT", txtRepeatTimes.Text)

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "spdsld", PlaSpdSlider.Value)

        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "ChkSmartPlayback", ChkSmartPlayback.Checked)

        '//spdsld ==  PlaSpdSlider.Value 
        PlaSpdSlider.Value = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "spdsld", 2)

        If chkStartup.Checked Then
            CreateShortCut("Perfect Macro Recorder", System.Environment.GetFolderPath(Environment.SpecialFolder.Startup), Application.ExecutablePath, Application.StartupPath, Application.ExecutablePath, 0)
        Else
            DeleteStartupShortcut()
        End If

        IsReged1 = RegHotKey(txtStrRec, &HBFF1&)
        IsReged2 = RegHotKey(txtStopRec, &HBFF2&)
        IsReged3 = RegHotKey(txtStrPly, &HBFF3&)
        IsReged4 = RegHotKey(txtStopPly, &HBFF4&)

        If IsReged1 And IsReged2 And IsReged3 And IsReged4 Then

            '//Just delete the temp key since no need to it Now
            ClearReg()

            Me.Close()

        End If

    End Sub

    Private Sub CancelButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        Me.Close()
    End Sub

    Private Sub OptionsForm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ClearReg()
    End Sub

    Private Sub ClearReg()
        Try
            My.Computer.Registry.CurrentUser.DeleteSubKeyTree("Software\CuteMacro\TK") '//Just delete the temp key since no need to it Now
        Catch ex As Exception '//The exception araises if there are no keys 'no need to process it'
            'WriteToErrorLog("can't delete Temp Keys from Reg", ex.StackTrace, "Error")
        End Try
    End Sub

    Private Sub ProcessHotKeyTxtBox(ByVal e As System.Windows.Forms.KeyEventArgs, ByRef TextBox As TextBox)
        Dim Modifiers As Long
        Dim ShortCutKey As String = ""
        If e.Modifiers = Keys.None Then

            'If e.KeyCode = Keys.F12 Or e.KeyCode = Keys.Escape Or e.KeyCode = Keys.Back Then

            '//coz Windows NT4 and Windows 2000/XP: The F12 key is reserved 
            '//for use by the debugger at all times, so it should not be registered as a hot key. 
            '//Even when you are not debugging an application, F12 is reserved in case a kernel-mode 
            '//debugger or a just-in-time debugger is resident.  (qtd. from MSDN=>RegisterHotKey Function) 

            '//Keys.Escape in order to give the user a way to disable the key
            If e.KeyCode = Keys.Escape Or e.KeyCode = Keys.Back Then

                TextBox.Text = ""

            End If

            'End If

            ShortCutKey = e.KeyCode.ToString

            TextBox.Tag = "" '//clear the tag to avoid confliction inf key up

        ElseIf e.KeyCode = Keys.ControlKey Or e.KeyCode = Keys.Menu Or e.KeyCode = Keys.ShiftKey Then

            ShortCutKey = e.Modifiers.ToString & " + "

            TextBox.Tag = "none" '//to check below in the key up if the char is only a ctrl or shift or alt 

        Else

            ShortCutKey = e.Modifiers.ToString & " + " & e.KeyCode.ToString

            TextBox.Tag = "" '//clear the tag to avoid confliction inf key up

        End If

        If e.KeyCode <> Keys.Escape And e.KeyCode <> Keys.Back Then

            TextBox.Text = ShortCutKey

        End If

        TextBox.SelectionStart = Len(TextBox.Text)

        Select Case e.Modifiers

            Case Keys.Alt
                Modifiers = MOD_ALT

            Case Keys.Shift
                Modifiers = MOD_SHIFT

            Case Keys.Control
                Modifiers = MOD_CONTROL

            Case Keys.Shift + Keys.Control
                Modifiers = MOD_SHIFT + MOD_CONTROL

            Case Keys.Shift + Keys.Alt
                Modifiers = MOD_SHIFT + MOD_ALT

            Case Keys.Alt + Keys.Control
                Modifiers = MOD_ALT + MOD_CONTROL

            Case Keys.Shift + Keys.Alt + Keys.Control
                Modifiers = MOD_SHIFT + MOD_ALT + MOD_CONTROL

        End Select

        '//Tk stands for Temp Keys 'coz we 'll delete them
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\TK\" & TextBox.Name, "m", Str(Modifiers))
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\TK\" & TextBox.Name, "k", Str(e.KeyCode))
        '//"e.m" is a special key that holds the e.modifiers value that we'll use in PlayMacro sub to check aginst HotKeys to execute them
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\TK\" & TextBox.Name, "e.m", Str(e.Modifiers))

    End Sub

    Private Sub TxtBxsKyUp(ByVal e As System.Windows.Forms.KeyEventArgs, ByRef TextBox As TextBox)

        '//handle the case of pressing only the modifire keys ( Ctrl + ) or (Ctrl + Alt + )
        '//so if there is no normal key like (a or b or c ) with the modifire then no key will appear in the text box
        If TextBox.Tag = "none" Then

            TextBox.Text = ""

        End If

    End Sub

    Private Function RegHotKey(ByVal txtbox As TextBox, ByVal KeyId As Integer) As Boolean

        Dim Res As Boolean
        Dim SKeyCode As Long = 0 '//Shortcut KeyCode
        Dim Modifiers As Long = 0
        Dim eModifiers As Long '//Holds the actual value of e.Modifiers that we can use to check hotkeys in PlayMacro Sub

        SKeyCode = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\TK\" & txtbox.Name, "k", 0)
        Modifiers = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\TK\" & txtbox.Name, "m", 0)
        '//"e.m" is a special key that holds the e.modifiers value that we'll use in PlayMacro sub to check aginst HotKeys to execute them
        eModifiers = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\TK\" & txtbox.Name, "e.m", 0)

        If SKeyCode = Keys.Escape Or SKeyCode = Keys.Back And Modifiers = 0 Then

            UnregisterHotKey(MainForm.Handle.ToInt32, KeyId) '//to disable the shortcut key so it not get triggerd

            Try

                '//delete the key so it doesn't gt loaded again
                My.Computer.Registry.CurrentUser.DeleteSubKeyTree("Software\CuteMacro\HK\" & txtbox.Name)

                RegHotKey = True

                Exit Function

            Catch ex As Exception '//The exception araises if there are no keys 'no need to process it'

                'WriteToErrorLog("can't delete Temp Keys from Reg", ex.StackTrace, "Error")

            End Try

        End If

        '//Focus with me please, there are cases when the user select a key that already registered by windows
        '//or any other application like Ctrl + ESC; in a case like this and for unknown reason, the registerhotkey API
        '//causes the windows to not to send textbox_keydown message to the textbox
        '//what the problem then?!
        '//this is causes the hot key to appear like this "Ctrl +" in Ctrl + ESC case 
        '//and problem with that is "Ctrl +" will trigger the hotkey and this doesn't make sense, hense i will pass
        '//any key if i found that it matched this form "modifier +".
        If SKeyCode = Keys.ControlKey Or SKeyCode = Keys.ShiftKey Or SKeyCode = Keys.Menu Then

            RegHotKey = True

            Exit Function

        End If

        If SKeyCode <> 0 Then

            Res = RegisterHotKey(MainForm.Handle.ToInt32, KeyId, Modifiers, SKeyCode)

            If Not Res Then

                MsgBox("Error: re-assign the " & """" & txtbox.Text & """" & " hotkey", MsgBoxStyle.Critical, "Perfect Macro Recorder")
                RegHotKey = False
                Exit Function

            End If

            'HK : stands for Hot Keys ; and "HK" key is where we save the keys 
            'however TK stands fro Temp Key; where we temporialy save the keys until we process under ok button or cancel button
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & txtbox.Name, "HotKey", txtbox.Text)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & txtbox.Name, "k", Str(SKeyCode))
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & txtbox.Name, "m", Str(Modifiers))
            '//"e.m" is a special key that holds the e.modifiers value that we'll use in PlayMacro sub to check aginst HotKeys to execute them
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & txtbox.Name, "e.m", Str(eModifiers))

        End If

        RegHotKey = True

    End Function

    Private Sub chkAllMouActs_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAllMouActs.CheckStateChanged
        If chkAllMouActs.Checked Then
            chkMouMov.Enabled = False
            chkMouMov.Checked = True
        Else
            chkMouMov.Enabled = True
        End If

        CheckAllRecOptions()
    End Sub

    Private Sub chkMouMov_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkMouMov.CheckStateChanged
        CheckAllRecOptions()
    End Sub

    Private Sub chkKeyActs_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkKeyActs.CheckStateChanged
        CheckAllRecOptions()
    End Sub

    Private Sub CheckAllRecOptions()

        Dim Message As String

        Message = "You cannot record a new macro if all the recording options are disabled."

        If Not chkMouMov.Checked And Not chkAllMouActs.Checked And Not chkKeyActs.Checked Then

            MsgBox(Message, MsgBoxStyle.Information, "Perfect Macro Recorder")

        End If

    End Sub

    Private Sub chkDefFold_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkDefFold.CheckStateChanged

        If chkDefFold.Checked = True Then

            'If _license.IsTrial Then
			'If is_x2342 = "False" Then
			'MsgBox("You cannot use this option in the evaluation version.", MsgBoxStyle.Information, "Perfect Macro Recorder")
			'chkDefFold.Checked = False
			'Exit Sub
			'End If

			txtDefMacName.Enabled = True
			'txtDefFoldPath.Enabled = True
			btnDefFoldPath.Enabled = True

		ElseIf chkDefFold.Checked = False Then

			txtDefMacName.Enabled = False
			'txtDefFoldPath.Enabled = False
			btnDefFoldPath.Enabled = False

		End If

	End Sub

    Private Sub btnDefFoldPath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDefFoldPath.Click

        FoldBrowDia.Description = "Select a default folder to save your recorded macros to"
        FoldBrowDia.RootFolder = Environment.SpecialFolder.Desktop

        FoldBrowDia.ShowDialog()

        If FoldBrowDia.SelectedPath <> "" Then
            lblPath.Text = FoldBrowDia.SelectedPath
        End If


    End Sub

    Private Function ValidatetxtDefMacName() As Boolean

        Try
            If chkDefFold.Checked Then

                If txtDefMacName.Text = "" Then
                    MsgBox("You must specify a default macro name.", MsgBoxStyle.Critical, "Perfect Macro Recorder")
                    txtDefMacName.Focus()
                    Exit Function
                End If

                If ValidateTextBoxEntry(txtDefMacName.Text) Then

                    MsgBox("A macro file name cannot contain any of the following characters:" & vbNewLine & _
                            "\ / : * ?  < > | " & """", MsgBoxStyle.Critical, "Perfect Macro Recorder")

                    txtDefMacName.HideSelection = False '// to make the selections appears to the user
                    txtDefMacName.Focus() '// to allow the user to edit txtDefMacName directly without focusing it first 
                    txtDefMacName.SelectAll()
                    'or you can use this
                    'txtDefMacName.SelectionLength = txtDefMacName.TextLength
                    Exit Function

                End If

                If Not txtDefMacName.ToString.EndsWith(".pmac") Then
                    txtDefMacName.AppendText(".pmac")
                End If

            End If

            Return True

        Catch ex As Exception

            Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error Validating txtDefMacName")
            Return False

        End Try

    End Function

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Sub txtDefMacName_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDefMacName.Validating
        '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

        '//http://msdn.microsoft.com/en-us/library/system.windows.forms.control.validating.aspx

        'When you change the focus by using the keyboard (TAB, SHIFT+TAB, and so on), 
        'by calling the Select or SelectNextControl methods, or by setting the ContainerControl..::.
        'ActiveControl property to the current form, focus events occur in the following order: 

        'Enter
        'GotFocus
        'Leave
        'Validating
        'Validated
        'LostFocus

        'When you change the focus by using the mouse or by calling the Focus method, focus events 
        'occur in the following order: 

        'Enter
        'GotFocus
        'LostFocus
        'Leave
        'Validating
        'Validated

        ValidatetxtDefMacName()

    End Sub

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Function CreateShortCut(ByVal shortcutName As String, ByVal creationDir As String, ByVal targetFullpath As String, ByVal workingDir As String, ByVal iconFile As String, ByVal iconNumber As Integer, Optional ByVal IsFailed As Boolean = False) As Boolean
        '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

        '//09/09/2009 : added IsFailed parameter for unit testing

        'CreateShortCut("Perfect Macro Recorder", System.Environment.GetFolderPath(Environment.SpecialFolder.Startup), Application.ExecutablePath, Application.StartupPath, Application.ExecutablePath, 0)

        Try
            If Not IO.Directory.Exists(creationDir) Then
                Dim retVal As DialogResult = MsgBox(creationDir & " doesn't exist. Do you wish to create it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo)
                If retVal = DialogResult.Yes Then
                    IO.Directory.CreateDirectory(creationDir)
                Else
                    Return False
                End If
            End If

            Dim IsShortcutExists As Boolean = IO.File.Exists(creationDir & "\" & shortcutName & ".lnk")

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
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "shortpth", creationDir & "\" & shortcutName & ".lnk")

            Return True

        Catch ex As System.Exception
            Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error creating startup shortcut.")
            IsFailed = True
            Return False
        End Try

    End Function

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Public Function DeleteStartupShortcut() As Boolean
        '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
        Try

            Dim IsShortcutExists As Boolean
            Dim StartupShortcutPath As String

            StartupShortcutPath = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "shortpth", "")

            If StartupShortcutPath = "" Then Exit Function

            IsShortcutExists = IO.File.Exists(StartupShortcutPath)

            If Not IsShortcutExists Then Exit Function

            Try
                IO.File.Delete(StartupShortcutPath)
            Catch ex As System.Exception
                Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error deleting startup shortcut.")
            End Try

            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "shortpth", "")

            IsShortcutExists = IO.File.Exists(StartupShortcutPath)

            If Not IsShortcutExists Then '//Check if delation successed
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error executing DeleteStartupShortcut method")
            Return False
        End Try
    End Function

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Sub chkCloseCM_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkCloseCM.CheckStateChanged
        '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
        If chkCloseCM.Checked Then
            txtRepeatTimes.Enabled = False
        Else
            txtRepeatTimes.Enabled = True
        End If
    End Sub

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Sub txtRepeatTimes_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtRepeatTimes.Validating
        '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
        IsOnlyNumeric(txtRepeatTimes.Text, txtRepeatTimes)

        If Val(txtRepeatTimes.Text) <= 0 Then ' <= zero isn't allowed
            txtRepeatTimes.Text = 1
        End If
    End Sub

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Private Sub txtRepeatTimes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRepeatTimes.TextChanged
        '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
        If Len(txtRepeatTimes.Text) >= 10 Then

            txtRepeatTimes.Text = Strings.Left(txtRepeatTimes.Text, 9)
            txtRepeatTimes.SelectionStart = Len(txtRepeatTimes.Text)

        End If
    End Sub

    Private Sub BtnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnReset.Click
        resetToDefHotKeys("txtStrRec", &HBFF1&, "Control + R", Int(82), Int(2)) ' &HBFF1& is a unique id for RegisterHotKey API function, you can assign any but must be uniqe
        resetToDefHotKeys("txtStopRec", &HBFF2&, "Control + W", Int(87), Int(2))
        resetToDefHotKeys("txtStrPly", &HBFF3&, "Control + P", Int(80), Int(2))
        resetToDefHotKeys("txtStopPly", &HBFF4&, "Control + B", Int(66), Int(2))

        LoadHotKeys(txtStrRec)
        LoadHotKeys(txtStopRec)
        LoadHotKeys(txtStrPly)
        LoadHotKeys(txtStopPly)
    End Sub

    Private Sub resetToDefHotKeys(ByVal txtbxName As String, ByVal KeyId As Integer, ByVal DefaultHotKey As String _
                           , ByVal DefaultKey As String, ByVal DefaultModifier As String)
        'HK : stands for Hot Keys ; and "HK" key is where we save the keys 
        'however TK stands fro Temp Key; where we temporialy save the keys until we process under ok button or cancel button
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & txtbxName, "HotKey", DefaultHotKey)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & txtbxName, "k", DefaultKey)
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & txtbxName, "m", DefaultModifier)
        '//"e.m" is a special key that holds the e.modifiers value that we'll use in PlayMacro sub to check aginst HotKeys to execute them
        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & txtbxName, "e.m", Str(131072)) '131072 == Keys.Control
    End Sub

End Class