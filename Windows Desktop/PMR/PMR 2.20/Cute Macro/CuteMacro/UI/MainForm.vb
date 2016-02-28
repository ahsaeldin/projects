'____oooooooooo_______oooooooo
'___ooooooooooooo___ooooooooooooo
'___oooooooooooooo_oooooooooooooo
'___ooooooooooooooooooooooooooooo
'____ooooooooooooooooooooooooooo
'_____ooooooooooooooooooooooooo
'______ooooooooooooooooooooooo
'_________oooooooooooooooooo
'___________ooooooooooooo
'_____________ooooooooo
'______________oooooo
'_______________oooo
'_______________ooo
'______________oo
'_____________o
'///////////////////////////////////////////////////////////////
'"In the middle of difficulty lies opportunity"~Albert Einstein'
'///////////////////////////////////////////////////////////////

Option Infer On
Option Explicit On

#Region "Imported Libraries"
Imports System.IO
Imports System.Text
Imports CuteMacro.Balora
Imports System.Threading
Imports System.Environment
Imports System.Diagnostics
Imports MouseKeyboardLibrary
Imports System.Reflection
#End Region

Friend Class MainForm

#Region "Fields"

	Dim SerializeTheList As Boolean = True
	Dim DeSerializeTheList As Boolean = True

	Dim MacroFilePath As String	'//must be assigned before every use of PlayMacro sub
	Dim m_events As New ArrayList
	Dim IsPlayBack As Boolean = False
	Dim lastTimeRecorded As Integer = 0
	Dim WithEvents mouseHook As New MouseHook
	Dim WithEvents keyboardHook As New KeyboardHook

	'//added next three lines to create a way for checking the hotkeys in PlaybackTheList method
	Dim IsTheListPlaybacked As Boolean = False
	Dim StopPlaybackingTheList As Boolean = False
	Dim ResumePlaybackingTheList As Boolean = False
	Dim SuspendPlaybackingTheList As Boolean = False

	Dim ForeGroundWindowSwitch As Boolean = True '//enable/disable adding foreground in the list

	Dim PlaybackThread As Thread

#End Region	'Global Variables in MainFrom

	Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)

		If CMOptionsForm.Visible = False Then
			'//don't process keys if optionfrom is visible to avoid trigger the hotkey if the userpress it in the txtboxs

			If m.Msg = WM_HOTKEY Then

				Dim e As System.EventArgs = Nothing
				Dim sender As System.Object = Nothing

				Select Case m.WParam

					Case &HBFF1&

						'chkMouMov == chkMM
						'chkAllMouActs = chkAM
						'chkKeyActs = chkKA
						'Dim MouseMoveOpt As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt", "chkMM", 1)
						'Dim AllMouseMovementsOpt As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkAM", 1)
						'Dim KeyboardActsOpt As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkKA", 1)

						If mouseHook.IsPaused Or keyboardHook.IsPaused Then
							ResRec()
						ElseIf mouseHook.IsStarted Or keyboardHook.IsStarted Then
							Me.TrayIcon.Icon = My.Resources.Pause_Record_Icon
							TrayIcon.Text = "Recording Paused"
							mouseHook.PauseRecord()
							keyboardHook.PauseRecord()
							'//changed next line to use the form tag rather than the button text in order to exclude the button.
							'StopRecordingForm.ButStopRecording.Text = "&Resume Record"
							StopRecordingForm.Tag = "&Resume Record"
							StopRecordingForm.ButStopRecording.Image = My.Resources.Pause_Record_Icon.ToBitmap
						ElseIf Not mouseHook.IsStarted And Not keyboardHook.IsStarted Then
							btnRecord_Click(sender, e)
						End If

					Case &HBFF2&
						BtnStop_Click(sender, e)
						ShowMe()
					Case &HBFF3&

						Static IsSuspended As Boolean

						If IsPlayBack And IsSuspended Then

							ShowPlaybackAtSystray()

							IsSuspended = False

							Try
								'//to check later in PlaybackTheList method in order 
								'//to process the hotkeys whenever we call PlaybackTheList
								ResumePlaybackingTheList = True
								SuspendPlaybackingTheList = False
								''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
								If Not IsTheListPlaybacked Then	'//if IsTheListPlaybacked = true then
									'لم يعد يستخدم
									'PlaybackThread.Resume() '// no need to process this line
								End If

							Catch ex As Exception
								Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error resume the Playback")
							End Try

							'Me.Opacity = 0

						ElseIf IsPlayBack Then

							TrayIcon.Icon = My.Resources.PausePlayback
							TrayIcon.Text = "Playback Paused"

							Try
								'//to check later in PlaybackTheList method in order 
								'//to process the hotkeys whenever we call PlaybackTheList
								SuspendPlaybackingTheList = True
								ResumePlaybackingTheList = False
								''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
								If Not IsTheListPlaybacked Then	'//if IsTheListPlaybacked = true then
									'لم يعد يستخدم
									'PlaybackThread.Suspend() '// no need to process this line
								End If
							Catch ex As Exception
								Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error suspending Playback")
							End Try

							IsSuspended = True

							'If Me.Opacity = 0 Then Me.Opacity = 100

						Else
							'لأنه لم يعد يستخدم هذا الزر علي الإطلاق ,, ويستخدم مكانه زر
							'التشغيل الجديد ,, لذا وجب التغيير
							'BtnPlayback_Click(sender, e)
							ToolStripRun_Click(sender, e)
						End If

					Case &HBFF4&

						RestoreSystrayIcon()

						'//Story : I clicked a pmac file to playback it but i stoped the playback after a seconds
						'//.. then i clicked the system tray icon to display the CM .. ohhhh .. just one more step 
						'// to my idea,, i opened a CM file then play it full .. nothing playbacked and balloon
						'//tool tip just appears telling me that playbacking has been finished !!

						'//Why : the first playback trigger Playmacro method because it is a command line
						'// playback .. and if the user just hit the stop playback hot key (Ctrl + b),
						'//then "StopPlaybackingTheList" var will be assigned to true .. but wait a second
						'//when you just play full any macro you call PlaybackTheList method that will found that'//
						'//StopPlaybackingTheList = True hence nothing playbacked .

						'//The Sol : Just add "StopPlaybackingTheList = True" in the next if statment 
						'//beacause
						If IsTheListPlaybacked Then
							'//why we add the next line : to check later in PlaybackTheList method in order 
							'//to process the hotkeys whenever we call PlaybackTheList
							StopPlaybackingTheList = True
							''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
						End If

						Try
							If Not IsTheListPlaybacked And IsPlayBack Then '//if IsTheListPlaybacked = true and
								'//there's already playbacking then...
								PlaybackThread.Abort() '// no need to process this line
							End If
						Catch ex As Exception
							Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error terminating Playback")
						End Try

						If IsPlayBack Then

							'//Story: press Ctrl + b "stop playback hotkey" and the balloon tooltip will appear even if there is no playback
							'//The Sol: calling "NotifyAtPlybckEnding" in an if statment that checks if there is a playback
							NotifyAtPlybckEnding()

						End If

						'//after PlaybackThread.Abort(), we must call the last lines in PlayMacro
						BlockInput(False)
						IsPlayBack = False
						TmrStop.Enabled = False

						'If Me.Opacity = 0 Then Me.Opacity = 100

				End Select

			ElseIf m.Msg = &H115 Then
				' WM_VSCROLL
				'MsgBox("asdfasdf")
			ElseIf m.Msg = WM_COPYDATA Then

				'تم تحييد هذا الجزء بعد إستخدام منطق الكوبي بيسيت في النقل السريع بين البرنامجين
				'Try
				'    'get the standard message structure from lparam
				'    Dim CD As IntProcCom.COPYDATASTRUCT = m.GetLParam(GetType(IntProcCom.COPYDATASTRUCT))
				'    'setup byte array
				'    Dim B(CD.cbData) As Byte
				'    'copy data from memory into array
				'    Runtime.InteropServices.Marshal.Copy(CD.lpData, B, 0, CD.cbData)

				'    '//GetString returns useless char at the end of the path that must be trimed
				'    MacroFilePath = System.Text.Encoding.Default.GetString(B)

				'    Dim WhereExtenstion As Integer = Strings.InStr(MacroFilePath, ".pmac")

				'    MacroFilePath = MacroFilePath.Substring(0, WhereExtenstion + 4)

				'    'empty array
				'    Erase B

				'Catch ex As Exception
				'    WriteToErrorLog(ex.Message, ex.StackTrace, "Error extracting the path and running the file.")
				'End Try

				''set message result to 'true', meaning message handled
				'm.Result = New IntPtr(1)
				'Me.WindowState = FormWindowState.Minimized

				''StartPlayBackThread()
				'LoadListViewFromCSV(MacroFilePath, lstviwMacEdit, False)
				'PlaybackTheList(False)

			End If

		End If

		MyBase.WndProc(m)

	End Sub

	Private Function RegMyHotKeys(ByVal txtbxName As String, ByVal KeyId As Integer, ByVal DefaultHotKey As String _
	  , ByVal DefaultKey As String, ByVal DefaultModifier As String) As Boolean

		Try
			'RegMyHotKeys("txtStopRec", &HBFF2&, "Control + W", Int(87), Int(2))

			Dim Res As Boolean

			Dim SKeyCode As Long = 0 '//Shortcut KeyCode
			Dim Modifiers As Long = 0
			Dim IsFirstRun As Long = 0

			'//FR stands for First Run
			IsFirstRun = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro", "FR", 0)

			If IsFirstRun = 0 Then

				'HK : stands for Hot Keys ; and "HK" key is where we save the keys 
				'however TK stands fro Temp Key; where we temporialy save the keys until we process under ok button or cancel button
				My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & txtbxName, "HotKey", DefaultHotKey)
				My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & txtbxName, "k", DefaultKey)
				My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & txtbxName, "m", DefaultModifier)
				'//"e.m" is a special key that holds the e.modifiers value that we'll use in PlayMacro sub to check aginst HotKeys to execute them
				My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & txtbxName, "e.m", Str(131072)) '131072 == Keys.Control

			End If

			'//SKeyCode .. Specifies the virtual-key code of the hot key.
			SKeyCode = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & txtbxName, "k", 0)
			Modifiers = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & txtbxName, "m", 0)

			If SKeyCode <> 0 Then

				Res = RegisterHotKey(Me.Handle.ToInt32, KeyId, Modifiers, SKeyCode)

				If Not Res Then
					'//WriteToErrorLog("MainFrom::RegMyHotKeys=>RegisterHotKey Failed", "ex.StackTrace", "Error")
					Util.WriteToErrorLog("RegisterHotKey failed and return 0", "", "RegisterHotKey Error")
					Return False
				Else
					Return True
				End If

			End If

		Catch ex As Exception

			Util.WriteToErrorLog(ex.Message, ex.StackTrace, "RegisterHotKey Error")

			Return False

		End Try


	End Function

	Private Sub btnRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRecord.Click

		StartRecording()

	End Sub

	Private Function StartRecording() As Boolean

		Try

			'//Before any Exit Sub statment you add here, you must add "IsbtnRecordIsAlreadyPressed = False"

			'//The idea behind "IsbtnRecordIsAlreadyPressed" is to exit this sub if 
			'//record hotkey pressed many times where ShowSave dialog is still displayed coz it will raise an exception
			Static IsbtnRecordIsAlreadyPressed As Boolean
			If IsbtnRecordIsAlreadyPressed Then Exit Function
			IsbtnRecordIsAlreadyPressed = True

			If IsPlayBack Then
				'' WriteToErrorLog("cannot Record a new macro while play bacing another one.", "", "")
				IsbtnRecordIsAlreadyPressed = False
				Exit Function
			End If

			'chkMouMov == chkMM
			'chkAllMouActs = chkAM
			'chkKeyActs = chkKA
			Dim IsMouMovEna As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt", "chkMM", 1)
			Dim IsAllMouActsEna As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkAM", 1)
			Dim IsKeyActsEna As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkKA", 1)

			If Not IsMouMovEna And Not IsAllMouActsEna And Not IsKeyActsEna Then

				Dim Message As String

				Message = "You cannot record a new macro because all the Recording Options are disabled. " & _
				"You can enable the Recording Options in the " & "Perfect Macro Recorder" & "'s Options"

				MsgBox(Message, MsgBoxStyle.Information, "Perfect Macro Recorder")

				IsbtnRecordIsAlreadyPressed = False

				Exit Function

			End If

			'If _license.IsUnlocked Or Debugger.IsAttached Then
			'If is_x2342 = "True" Then
			Dim DefaultFilePath As String = IsMacroSavedToDefFolds()
			Dim RecordedFilePath As String
			If DefaultFilePath = "" Then
				RecordedFilePath = ShowSaveDialog("Perfect Macro Recorder" & " Files (*.pmac)|*.pmac")
				If RecordedFilePath = "" Then IsbtnRecordIsAlreadyPressed = False : Exit Function
			Else
				RecordedFilePath = DefaultFilePath
			End If
			'End If

			Me.WindowState = FormWindowState.Minimized

			StopRecordingForm.Show()

			m_events.Clear()

			lastTimeRecorded = Environment.TickCount

			'chkMouMov == chkMM
			'chkAllMouActs = chkAM
			'chkKeyActs = chkKA
			Dim MouseMoveOpt As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt", "chkMM", 1)
			Dim AllMouseMovementsOpt As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkAM", 1)
			Dim KeyboardActsOpt As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkKA", 1)

			TrayIcon.Icon = My.Resources.Recording
			TrayIcon.Text = "Recording..."

			Dim IsTipsAllowed As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkTips", 1)

			If IsTipsAllowed Then
				Dim stopRecordingHotkey As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & "txtStopRec", "HotKey", "")
				Dim pause_resume_RecordingHotKey As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & "txtStrRec", "HotKey", "")
				Dim balloonMessage As String = _
				"Press " & stopRecordingHotkey & " to stop recording," & vbNewLine & _
				"Press " & pause_resume_RecordingHotKey & " to pause/resume recording"

				TrayIcon.ShowBalloonTip(1000, "Perfect Macro Recorder" & " is now recording.", balloonMessage, ToolTipIcon.Info)
			End If

			If AllMouseMovementsOpt Or MouseMoveOpt Then
				mouseHook.Start()
			End If

			If KeyboardActsOpt Then
				keyboardHook.Start()
			End If

			TmrStop.Enabled = True

			IsbtnRecordIsAlreadyPressed = False

			StartRecording = True

		Catch ex As Exception

			StartRecording = False

			Util.WriteToErrorLog(ex.Message, ex.StackTrace, "error in executing StartRecording Method")

		End Try


	End Function

	Private Sub BtnStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnStop.Click

		StopMacro()

		TmrStop.Enabled = False

	End Sub

	Private Sub RestoreSystrayIcon()
		Dim resources As New System.Resources.ResourceManager(GetType(MainForm))
		Me.TrayIcon.Icon = CType(resources.GetObject("TrayIcon.Icon"), System.Drawing.Icon)
		TrayIcon.Text = "Perfect Macro Recorder"
	End Sub

	Friend Sub StopMacro()

		'//the next 2 lines assume that you record both mouse and keyboard events
		'If Not mouseHook.IsStarted Then Exit Sub
		'If Not keyboardHook.IsStarted Then Exit Sub

		If Not mouseHook.IsStarted And Not keyboardHook.IsStarted Then Exit Sub

		RestoreSystrayIcon()

		mouseHook.Stop()

		keyboardHook.Stop()

		StopRecordingForm.Close()

		Dim DefaultFilePath As String = IsMacroSavedToDefFolds()

		Dim RecordedFilePath As String = ""

		'If Debugger.IsAttached Or _license.IsUnlocked Then
		'If is_x2342 = "True" Then
		If DefaultFilePath = "" Then
			RecordedFilePath = ShowSave.FileName
		Else
			RecordedFilePath = DefaultFilePath
		End If
		'End If

		Dim filename As String = My.Computer.FileSystem.GetTempFileName()

		SaveSerializedObject(filename, m_events)

		'يعتبر هذا هو الإستدعاء الوحيد لهذه الوظيفة
		'والفكرة هنا هو حفظ الماكرو علي الأرض فقط
		'لتحميله إلي الليست عن طريق السطر التالي
		LoadMacList(filename)

		'اكتشفت اليوم مشكلة رهيبة ..
		' أن الوظائف التي اعتدت استخدامها لحفظ الماكرو لاتكفل لي حفظ الاضافات المستقبلية كعنصر الإنتظار مثلا كما ينبغي .. علي سبيل المثال .. عنصر الإنتظار غير موجود في ملف الماكرو .. بل يتم استخراجه من العنصر الذي يليه .. ولهذا فقط استبدلت وظائف الحفظ العادية مثل 
		'SaveSerializedObject()
		'SaveListView()
		'بوظيفة()
		'SaveListViewToCSV()
		'مما يكفل لي حفظ الليست كما هي .. وفي نفس الوقت غيرت في وظيفة
		'PlaybackTheList()
		'        و()
		'        IWannaSleep()
		'بما يكفل التعامل مع عنصر الإنتظار كعنصر منفصل تماما وبالتالي لا يعتمد علي العنصر الذي يليه مما يسهل علي إضافة ما شئت من العناصر فيما بعد بدون القلق من عنصر الإنتظار 

		'If Debugger.IsAttached Or _license.IsUnlocked Then
		'If is_x2342 = "True" Then
		SaveListViewToCSV(lstviwMacEdit, RecordedFilePath)
		'End If

		'If _license.IsTrial Then
		If is_x2342 = "False" Then
			Me.Text = "Perfect Macro Recorder" & " - Trial Version"
			'ElseIf _license.IsUnlocked Or Debugger.IsAttached Then
		ElseIf is_x2342 = "True" Then
			Me.Text = "Perfect Macro Recorder" & " - " & Path.GetFileName(RecordedFilePath)
		End If

	End Sub

	Private Sub BtnPlayback_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPlayback.Click

		'//Before any Exit Sub statment you add here, you must add "IsBtnPlaybackIsAlreadyPressed = False"

		'//The idea behind "IsBtnPlaybackIsAlreadyPressed" is to exit this sub if 
		'//record hotkey pressed many times where ShowSave dialog is still displayed coz it will raise an exception
		Static IsBtnPlaybackIsAlreadyPressed As Boolean
		If IsBtnPlaybackIsAlreadyPressed Then Exit Sub
		IsBtnPlaybackIsAlreadyPressed = True

		Dim FileName As String = ShowOpenDialog("Select a macro file to play back")

		If FileName = "" Then IsBtnPlaybackIsAlreadyPressed = False : Exit Sub

		MacroFilePath = ShowOpen.FileName

		Me.WindowState = FormWindowState.Minimized

		StartPlayBackThread()

		IsBtnPlaybackIsAlreadyPressed = False

	End Sub

	'//From Now, PlayMacro method won't be used and 
	'//PlaybackTheList will be used instead, in other words..
	'//the only way to playback will be through the list view only.. 
	'//but there is one place in which the playmacro will be called ,, 
	'//in MainForm_Shown to call a macro at startup
	Private Sub PlayMacro(ByVal FilePath As String)

		'من الآن هذه الوظيفة لاتستخدم علي الإظلاق

		IsTheListPlaybacked = False	'//Used in WndProc to avoid calling the thread lines

		FilePath = MacroFilePath

		'Dim SKeyCode As Long = 0 '//Shortcut KeyCode
		'Dim eModifiers As Long '//Holds the actual value of e.Modifiers that we can use to check hotkeys in PlayMacro Sub

		'SKeyCode = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & "txtStrPly", "k", 80)
		'//"e.m" is a special key that holds the e.modifiers value that we'll use in PlayMacro sub to check aginst HotKeys to execute them
		'eModifiers = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & "txtStrPly", "e.m", 131072)

		Dim macroEvent As MacroEvent

		Dim StoredEvents As New ArrayList

		Dim IsFailed As Boolean

		StoredEvents = RestoreSerializedObject(FilePath, IsFailed)

		If IsFailed Then

			MsgBox("Perfect Macro Recorder" & " could not playback this macro because the file has been damaged.", MsgBoxStyle.Information, "Perfect Macro Recorder")

			Exit Sub

		End If

		''''''BlockInput(True)

		TmrStop.Enabled = True '//Using this time to terminate Play back if the user wishes to 

		IsPlayBack = True

		Dim RepeatCount As Int32 = GetPlyBckRepeatCount()

		Dim PlayBackSpeed As Integer = GetPlaybackSpeed()

		Dim i As Int32 = 1

		For i = 1 To RepeatCount

			For Each macroEvent In StoredEvents

				If Not IsPlayBack Then '//Terminate Playback whenever IsPlayback set to Flase in TmrStop

					Exit For '//It is just a way to stop play back by a hot key from within a timer

				End If

				If PlayBackSpeed = 2 Then
					Thread.Sleep(macroEvent.Time_SinceLastEvent)
				Else
					Thread.Sleep(PlayBackSpeed)
				End If

				My.Application.DoEvents()

				Select Case macroEvent.Macro_Event_Type

					Case MacroEventType.MouseMove

						MouseSimulator.X = macroEvent.ClonedMouseEventArgs.X

						MouseSimulator.Y = macroEvent.ClonedMouseEventArgs.Y

					Case MacroEventType.MouseDown

						MouseSimulator.MouseDown(macroEvent.ClonedMouseEventArgs.Button)

					Case MacroEventType.MouseUp

						MouseSimulator.MouseUp(macroEvent.ClonedMouseEventArgs.Button)

					Case MacroEventType.KeyDown

						KeyboardSimulator.KeyDown(macroEvent.ClonedKeyEventArgs.KeyCode)

					Case MacroEventType.KeyUp

						KeyboardSimulator.KeyUp(macroEvent.ClonedKeyEventArgs.KeyCode)

				End Select

			Next

		Next

		BlockInput(False)

		IsPlayBack = False

		TmrStop.Enabled = False

		'If Me.Opacity = 0 Then Me.Opacity = 100

		Dim IsAppTobeClosedAftPlyback As Boolean

		IsAppTobeClosedAftPlyback = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkCloseCM", 0)

		If IsAppTobeClosedAftPlyback Then

			'//because Cross-thread operation not valid: Control 'MainForm' accessed from 
			'//a thread other than the thread it was created on. Hence i cannot call Me.Dispose() to close the app
			'//so i will use application.exit
			Application.Exit()

		End If

		NotifyAtPlybckEnding()

	End Sub

	Private Function getForeWndAttr(ByRef foreWndRect As ForeGroundWindowAttr) As Boolean

		Static OldHwnd As IntPtr
		Dim ForeWinRect As RECT
		Dim WindowTextLength As Long
		Dim foreWndHandle As IntPtr = GetForegroundWindow()

		If OldHwnd <> foreWndHandle Then

			WindowTextLength = GetWindowTextLength(foreWndHandle)
			Dim WindowText As String = Space(WindowTextLength + 1)
			Dim getwindowtextResult As Integer = GetWindowText(foreWndHandle, WindowText, WindowTextLength + 1)

			GetWindowRect(foreWndHandle, ForeWinRect)
			foreWndRect.WinTitle = WindowText.Trim
			foreWndRect.Left = ForeWinRect.Left
			foreWndRect.Top = ForeWinRect.Top
			foreWndRect.Right = ForeWinRect.Right
			foreWndRect.Bottom = ForeWinRect.Bottom

			'28/09/2010
			'Dim ForeProcessName As String = "" 'FoxSmr
			'Dim ForeProcessFilePath As String = "" 'FoxSmr
			'الحصول علي اسم برنامج النافذة ومساره لكي يحفظ مع الماكرو وفي الليست
			'كي يتسني لي فتحها اذا لم تكون بالجوار ومن ثم يزيد ذكاء البرنامج
			'Dim foreProcess As Process = GetProcessPathByHandle(foreWndHandle, WindowText, ForeProcessName, ForeProcessFilePath) 'FoxSmr
			'foreWndRect.ProcessName = ForeProcessName 'FoxSmr
			'foreWndRect.ProcessFilePath = ForeProcessFilePath 'FoxSmr

		End If

		OldHwnd = foreWndHandle

		If WindowTextLength = 0 Then
			Return False
		Else
			Return True
		End If

	End Function

	Private Sub AddMouseEvent(ByVal EventType As MacroEventType, ByVal e As System.Windows.Forms.MouseEventArgs)

		Dim foreRes As Boolean = False
		Dim ForeGroundWindowAttr As New ForeGroundWindowAttr

		foreRes = getForeWndAttr(ForeGroundWindowAttr)

		If foreRes And ForeGroundWindowSwitch Then
			m_events.Add(New MacroEvent(MacroEventType.ForeGroundWindow, e, 0, ForeGroundWindowAttr))
		End If

		m_events.Add(New MacroEvent(EventType, e, Environment.TickCount - lastTimeRecorded))

		lastTimeRecorded = Environment.TickCount

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub AddKeyboardEvent(ByVal EventType As MacroEventType, ByVal e As System.Windows.Forms.KeyEventArgs)
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim foreRes As Boolean = False
		Dim ForeGroundWindowAttr As New ForeGroundWindowAttr

		foreRes = getForeWndAttr(ForeGroundWindowAttr)

		If foreRes And ForeGroundWindowSwitch Then
			m_events.Add(New MacroEvent(MacroEventType.ForeGroundWindow, e, 0, ForeGroundWindowAttr))
		End If

		m_events.Add(New MacroEvent(EventType, e, Environment.TickCount - lastTimeRecorded))

		lastTimeRecorded = Environment.TickCount

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub mouseHook_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles mouseHook.MouseWheel
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		AddMouseEvent(MacroEventType.MouseWheel, e)
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub mouseHook_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles mouseHook.MouseMove
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		AddMouseEvent(MacroEventType.MouseMove, e)

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub mouseHook_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles mouseHook.MouseDown
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim AllMouseMovementsAllowed As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkAM", 1)

		If AllMouseMovementsAllowed Then
			AddMouseEvent(MacroEventType.MouseDown, e)
		End If

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub mouseHook_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles mouseHook.MouseUp
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim AllMouseMovementsAllowed As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkAM", 1)

		If AllMouseMovementsAllowed Then
			AddMouseEvent(MacroEventType.MouseUp, e)
		End If

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub keyboardHook_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles keyboardHook.KeyDown
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		AddKeyboardEvent(MacroEventType.KeyDown, e)

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub keyboardHook_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles keyboardHook.KeyUp
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		AddKeyboardEvent(MacroEventType.KeyUp, e)

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub MainForm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		mouseHook.Stop()
		keyboardHook.Stop()

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub mainform_load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		'Automatically remove mklextp from GAC
		If Debugger.IsAttached Then
			Dim EntServ As New System.EnterpriseServices.Internal.Publish
			EntServ.GacRemove("C:\Windows\System32\PMR\mklextp.dll")
		End If

		If My.Settings.companyname <> "Perfection Tools" Then
			MsgBox("Trial period expired. Thank you!")
			End
		End If

		Try
			If My.Settings.gac = "nothing" Then
				If Not Balora.Util.IsAssemplyInGAC("mklextp.dll") Then
					Dim libPath As String = Application.StartupPath & "\" & "mklextp.dll"
					Balora.Util.AddAssemplyToGAC(Application.StartupPath & "\" & "mklextp.dll")
					My.Settings.gac = "ok"
				End If
			End If
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try

		Balora.Settings.LogErrors = True '//This is default but to remember you of this property
		Balora.Settings.LogFileName = "Perfect-Macro-Recorder-Error-Log.txt"
		Balora.Settings.ShowLogWindow = True

		TrayIcon.Visible = True	'//just display the TrayIcon if not an executable macro file

		Dim ArgsCount As Integer = My.Application.CommandLineArgs.Count

		Try

			'//if there is no command line parameter then ArgumentOutOfRangeException will be raised
			If ArgsCount > 0 Then

				If My.Application.CommandLineArgs.Item(0).Contains(".pmac") Then

					Me.Opacity = 0 '//click a .pmac file to playback it then 
					'//the mainform will be appears at a glance then disappear ..
					'//to avoid this we add Me.Opacity = 0

				End If

			End If

		Catch ex As ArgumentOutOfRangeException

			Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error hiding the form")

		End Try

		Me.AllowDrop = True

		RegAllMyHotKeys()

		Dim IsFirstRun As Long

		'//FR stands for First Run
		IsFirstRun = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro", "FR", 0)

		If IsFirstRun = 0 Then

			My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\CuteMacro", "FR", Str(1))	'//Close the First Run Flag

			Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\CuteMacro\opt")

		End If

		If ArgsCount = 0 Then '// means no playback from a file in order to not to start playback and record

			'//Auto Start Recording Option
			Dim AutoStartRecord As Boolean

			AutoStartRecord = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkAR", 0)

			If AutoStartRecord Then

				Dim SenderObj As System.Object = Nothing
				Dim eEventArgs As System.EventArgs = Nothing

				btnRecord_Click(SenderObj, eEventArgs)

			End If

		End If

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub RecordSystemTrayMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RecordSystemTrayMenu.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim btnRecord_Click_Sender As Object = Nothing
		Dim btnRecord_Click_e As System.EventArgs = Nothing
		btnRecord_Click(btnRecord_Click_Sender, btnRecord_Click_e)

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub PlayBackSystemTrayMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlayBackSystemTrayMenu.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim BtnPlayback_Sender As Object = Nothing
		Dim BtnPlayback_e As System.EventArgs = Nothing
		ToolStripRun_Click(BtnPlayback_Sender, BtnPlayback_e)

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub OptionsSystemTrayMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OptionsSystemTrayMenu.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		CMOptionsForm.ShowDialog()

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub AboutSystemTrayMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutSystemTrayMenu.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		AboutBox.ShowDialog(Me)
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub ExitSystemTrayMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitSystemTrayMenu.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		mouseHook.Stop()
		keyboardHook.Stop()
		Me.Dispose()
		'//Me.Close()
		'//Application.Exit()
		'//Dispose vs Close http://www.groupsrv.com/dotnet/about6208.html
		'//Me.Close vs application.exit vs me.dispose http://discuss.joelonsoftware.com/default.asp?dotnet.12.402651.5

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub ButOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButOptions.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		CMOptionsForm.ShowDialog()

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub TmrStop_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TmrStop.Tick
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		If GetAsyncKeyState(VK_MENU) And GetAsyncKeyState(VK_CONTROL) And GetAsyncKeyState(VK_DELETE) Then

			BlockInput(False)

			If IsPlayBack Then

				IsPlayBack = False

				MsgBox("Playback Terminated", MsgBoxStyle.Information, "Perfect Macro Recorder")

			ElseIf mouseHook.IsStarted And keyboardHook.IsStarted Then

				StopMacro()

				MsgBox("Recording has been Terminated", MsgBoxStyle.Information, "Perfect Macro Recorder")

			End If

			'Delete after mimic to the new executable macro.
			'ElseIf IsAnExecutableMacroFile Then

			'    If GetAsyncKeyState(VK_CONTROL) And GetAsyncKeyState(VK_SHIFT) And GetAsyncKeyState(VK_F2) Then

			'        If IsPlayBack Then

			'            IsPlayBack = False

			'            MsgBox("Playback Terminated", MsgBoxStyle.Information, "Perfect Macro Recorder")

			'        End If

			'    End If

		End If

		'//added "just in case" to make sure that this timer in disabled whenever there are no need to use it
		If IsPlayBack = False And keyboardHook.IsStarted = False And mouseHook.IsStarted = False Then

			TmrStop.Enabled = False

		End If

	End Sub

	Friend Sub ResRec()

		'//return to the default icon
		Me.TrayIcon.Icon = My.Resources.Recording
		TrayIcon.Text = "Recording..."
		'Dim resources As New System.Resources.ResourceManager(GetType(MainForm))
		'Me.TrayIcon.Icon = CType(Resources.GetObject("TrayIcon.Icon"), System.Drawing.Icon)

		'//ResRec = Resume Record : just collecting all the lines that resume record  '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'//in order to use in it's place in WndProc and also to call it from within 
		'//StopRecordingForm.ButStopRecording_Click()and that's why i declared it as a public Sub.
		'//P.S. i called it ResRec to differ it from the ResumeRecord in GlobalHook.cs

		lastTimeRecorded = Environment.TickCount '//To delete the time gab between Pausing and recording 
		mouseHook.ResumeRecord()
		keyboardHook.ResumeRecord()
		'//changed next line to use the form tag rather than the button text in order to exclude the button.
		'StopRecordingForm.ButStopRecording.Text = "&Stop Recording"
		StopRecordingForm.Tag = "&Stop Recording"
		StopRecordingForm.ButStopRecording.Image = My.Resources.StopRecording.ToBitmap
	End Sub

	Private Sub StartPlayBackThread(Optional ByVal PlayList As Boolean = False, Optional ByVal SelectedItemsOnly As Boolean = False)

		Try

			If PlayList Then
				PlaybackThread = New Thread(AddressOf PlaybackTheList)
			Else

				'11/10/2009
				'الآن اخبرك رسميا أن وظيفة
				'PlayMacro
				'لم تعد تستخدم علي الإطلاق نظراً لأن الوظيفة
				'StartPlayBackThread
				'لم تعد تستدعي مع البارامتر الأول لها وهو 
				'False
				'مما يعني أن السطر التالي لن يتم تنفيذه أبدا
				'وقد جاء كل هذا بعد قرار الكوماندا الكبير أحمد سعد
				'بتعديل منطق الحفظ والبلاي باك في البرنامج بعد أزمة العاشر من أكتوبر 2009 الشهيرة
				'لمعرفة المزيد عن الأزمة ابحث عن كلمة العاشر في أرجاء الكود العامر


				'//From Now, PlayMacro method won't be used and 
				'//PlaybackTheList will be used instead, in other words..
				'//the only way to playback will be through the list view only.. 
				'//but there is one place in which the playmacro will be called ,, 
				'//in MainForm_Shown to call a macro at startup
				PlaybackThread = New Thread(AddressOf PlayMacro)
			End If

			PlaybackThread.IsBackground = True
			PlaybackThread.Start()

		Catch ex As Exception

			Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error Starting Playback")

		End Try

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Function IsMacroSavedToDefFolds() As String
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim FullPath As String
		Dim DefFoldPath As String
		Dim DefFileName As String
		Dim IsDefFoldChecked As Boolean

		'//DF == defualt Folder
		IsDefFoldChecked = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkDF", 0)

		If IsDefFoldChecked = True Then

			'//txtDFN = Options.txtDefMacName.Text
			DefFileName = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "txtDFN", "default.pmac")

			'//txtDFP = Options.lblPath.text
			DefFoldPath = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "txtDFP", My.Computer.FileSystem.CurrentDirectory())

			If DefFoldPath.EndsWith("\") Then
				FullPath = DefFoldPath & DefFileName
			Else
				FullPath = DefFoldPath & "\" & DefFileName
			End If

			If Not My.Computer.FileSystem.DirectoryExists(DefFoldPath) Then
				My.Computer.FileSystem.CreateDirectory(DefFoldPath)	'//create the folder to avoid error when saving the macro using SaveSerializedObject method
			End If

		Else

			FullPath = ""

		End If

		IsMacroSavedToDefFolds = FullPath

	End Function

	Private Sub TrayIcon_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TrayIcon.MouseDoubleClick
		ShowMe()
	End Sub

	Private Sub HideMe()
		ShowWindow(Me.Handle, 0)
	End Sub

	Friend Sub ShowMe()
		If Me.Opacity = 0 Then Me.Opacity = 100
		Me.Show()
		Me.WindowState = FormWindowState.Normal
	End Sub

	Private Sub NotifyAtPlybckEnding()

		Dim IsUsrTobeNotifiedAftPlyback As Boolean
		IsUsrTobeNotifiedAftPlyback = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkNotify", 1)

		If IsUsrTobeNotifiedAftPlyback Then
			TrayIcon.ShowBalloonTip(2000, "Perfect Macro Recorder", "Playback finish.", ToolTipIcon.Info)
		End If

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Function GetPlyBckRepeatCount() As Int32
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'// okay look, for the Int32 data type, it ranges in value from -2,147,483,648 through 2,147,483,647, 
		'//2,147,483,647 is 10 digits number so i will max digits number in txtrepeatcount is 9 to avoid a certin overflow in 
		'//PlayMacro method .
		GetPlyBckRepeatCount = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "txtRT", 1)

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Function GetPlaybackSpeed(Optional ByVal ExeSpeed As Integer = 0) As Integer
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim SpeedSlider As Integer

		'//spdsld ==  PlaSpdSlider.Value 
		SpeedSlider = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "spdsld", 2)

		Select Case SpeedSlider

			Case 0
				GetPlaybackSpeed = 100 '//Very, Slow 100 passed as a millsecond timeout to thread.sleep in playmacro
			Case 1
				GetPlaybackSpeed = 50 '//Slow, 100 passed as a millsecond timeout to thread.sleep in playmacro
			Case 2
				GetPlaybackSpeed = 2 '//Normal, checked in playmacro, and if GetPlaybackSpeed = 2 then pass the normal value to thread.sleep
			Case 3
				GetPlaybackSpeed = 5 '//Fast, 5 passed as a millsecond timeout to thread.sleep in playmacro
			Case 4
				GetPlaybackSpeed = 0 '//Very Fast, 0 passed as a millsecond timeout to thread.sleep in playmacro

		End Select

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub RegAllMyHotKeys()
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		RegMyHotKeys("txtStrRec", &HBFF1&, "Control + R", Int(82), Int(2)) ' &HBFF1& is a unique id for RegisterHotKey API function, you can assign any but must be uniqe
		RegMyHotKeys("txtStopRec", &HBFF2&, "Control + W", Int(87), Int(2))
		RegMyHotKeys("txtStrPly", &HBFF3&, "Control + P", Int(80), Int(2))
		RegMyHotKeys("txtStopPly", &HBFF4&, "Control + B", Int(66), Int(2))

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub MainForm_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		If is_x2342 = "False" Then
			Me.Text = "Perfect Macro Recorder" & " - Trial Version"
		End If
		'Dim IsRunAtStartUp As Boolean

		'IsRunAtStartUp = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkStartup", 0)

		Dim ArgsCount As Integer = My.Application.CommandLineArgs.Count

		If ArgsCount > 0 Then

			'If IsRunAtStartUp Then

			'  Me.Hide()

			'End If

			If My.Application.CommandLineArgs.Item(0) = "min" Then

				HideMe()

			Else '// What is the meaning of "Else" here, ha?!!!! .. r u .... nuts ?
				'// Story : run Perfect Macro Recorder while we add any dump command line like
				'// Perfect Macro Recorder.exe asdfasdfw3jksfkjasdjf ..... and guess what the next lines will
				'//Beep executed ..

				'// The Sol : use the contains method to check if it contains a path or something
				If My.Application.CommandLineArgs.Item(0).Contains(".pmac") Then

					HideMe()
					MacroFilePath = My.Application.CommandLineArgs.Item(0)
					'تم تعديل السطر التالي بعد أزمة العاشر من أكتوبر سنة 2009 الشهيرة 
					'عندما اكتشفت ان عنصر الإنتظار سوف يخلق أزمة رهيبة سواء في أسلوب حفظه 
					'أو استرجاعه أو إعادة تشغيله من جديد ومن ثم فقد استبدلت منطق البرنامج
					'بمنطق الإعتماد علي حفظ الليست فيو كما هو بدون أسلوب التسلسل للماكرو 
					'الشهير ثم إعادة تحميل هذا الملف .. مع ملاحظة إنني انوي بالفعل حفظ الليست فيو غدا 
					'عن طريق التسلسل
					'StartPlayBackThread(True)
					LoadListViewFromCSV(MacroFilePath, lstviwMacEdit, False)
					PlaybackTheList(False)


				End If

			End If

		End If

	End Sub

	Private Sub HideScriptListHorzScrollBars()
		ShowScrollBar(lstviwMacEdit.Handle.ToInt32, 0, False)
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Friend Sub LoadMacList(ByVal MacFilePath As String, Optional ByVal Merge As Boolean = False)
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim keyItm As ListViewItem
		Dim lstItm As ListViewItem
		Dim waitIem As ListViewItem

		Dim macroEvent As MacroEvent

		Dim StoredEvents As New ArrayList

		Dim IsFailed As Boolean

		StoredEvents = RestoreSerializedObject(MacFilePath, IsFailed)

		If IsFailed Then

			If StoredEvents.Count = 0 Then
				MsgBox("Nothing to load; the macro file is empty. " & vbNewLine & "Check the recording options." _
				, MsgBoxStyle.Critical, "Perfect Macro Recorder")
			Else
				MsgBox("Cannot load the macro to the list," + vbNewLine + "the macro file might be corrupted" _
				, MsgBoxStyle.Critical, "Perfect Macro Recorder")
			End If

			Exit Sub

		End If

		If Not Merge Then '//Merge parameter to decide if i will load a new file to the list or note with
			'//with an old file or not

			lstviwMacEdit.Items.Clear()

		End If

		frmProgress.Text = "Preparing Macro Script..."
		frmProgress.StartPosition = FormStartPosition.CenterScreen
		frmProgress.ProgBar.Maximum = StoredEvents.Count
		frmProgress.Show()

		lstviwMacEdit.BeginUpdate()

		Dim Progress As Long

		For Each macroEvent In StoredEvents

			Progress = Progress + 1

			frmProgress.ProgBar.Value = Progress

			If macroEvent.Time_SinceLastEvent <> 0 Then

				waitIem = lstviwMacEdit.Items.Add("Wait")

				waitIem.ForeColor = GetItemColor(waitIem) 'Color.Red

				waitIem.SubItems.Add(macroEvent.Time_SinceLastEvent)

			End If

			Select Case macroEvent.Macro_Event_Type

				Case MacroEventType.ForeGroundWindow

					'a switch to disable/enable smart playback
					If ForeGroundWindowSwitch Then
						lstItm = lstviwMacEdit.Items.Add("Background Window")
						'lstItm.ForeColor = Color.Gray
						lstItm.ForeColor = GetItemColor(lstItm)

						lstItm.SubItems.Add(macroEvent.ForeWinStruct.WinTitle)
						lstItm.SubItems.Add(macroEvent.ForeWinStruct.Left)
						lstItm.SubItems.Add(macroEvent.ForeWinStruct.Top)
						lstItm.SubItems.Add(macroEvent.ForeWinStruct.Right)
						lstItm.SubItems.Add(macroEvent.ForeWinStruct.Bottom)

						lstItm.SubItems.Add(0)
						lstItm.SubItems.Add(False)
						'//Default value for "wait" or delay .. Default Value is 1000 milliseconds
						'lstItm.SubItems.Add(1000) 'FoxSmr
						'lstItm.SubItems.Add(True) 'FoxSmr

						'28/09/2010
						'هنا نضع مسار برنامج النافذة واسم هذا البرنامج .. لاحقاً سأبحث عنه إن لم يكن موجوداً
						'lstItm.SubItems.Add(macroEvent.ForeWinStruct.ProcessName) 'FoxSmr
						'lstItm.SubItems.Add(macroEvent.ForeWinStruct.ProcessFilePath) 'FoxSmr
					End If

				Case MacroEventType.MouseWheel

					lstItm = lstviwMacEdit.Items.Add("MouseWheel")

					'lstItm.ForeColor = Color.DeepPink
					lstItm.ForeColor = GetItemColor(lstItm)

					lstItm.SubItems.Add(macroEvent.ClonedMouseEventArgs.Delta)
					'//gomla morkaba .. adding a new subitems then assign The X and Y to the Tag value 
					'//in order to use later in ToolStripSveMac_Click
					lstItm.SubItems.Add("X = " & macroEvent.ClonedMouseEventArgs.X).Tag = macroEvent.ClonedMouseEventArgs.X
					lstItm.SubItems.Add("Y = " & macroEvent.ClonedMouseEventArgs.Y).Tag = macroEvent.ClonedMouseEventArgs.Y
					lstItm.SubItems.Add(macroEvent.Time_SinceLastEvent)
					lstItm.SubItems.Add(macroEvent.ClonedMouseEventArgs.Button)
					lstItm.SubItems.Add(macroEvent.ClonedMouseEventArgs.Clicks)

				Case MacroEventType.MouseMove

					lstItm = lstviwMacEdit.Items.Add("MouseMove")

					'lstItm.ForeColor = Color.Blue
					lstItm.ForeColor = GetItemColor(lstItm)

					'//gomla morkaba .. adding a new subitems then assign The X and Y to the Tag value 
					'//in order to use later in ToolStripSveMac_Click
					lstItm.SubItems.Add("X = " & macroEvent.ClonedMouseEventArgs.X).Tag = macroEvent.ClonedMouseEventArgs.X
					lstItm.SubItems.Add("Y = " & macroEvent.ClonedMouseEventArgs.Y).Tag = macroEvent.ClonedMouseEventArgs.Y

					'//Button Clicks Delta : Storing Wait, Button, Clicks & Delta in the column
					'// in order to use in "ToolStripSveMac_Click" ( i.e. Save the edited Macro )
					lstItm.SubItems.Add(macroEvent.Time_SinceLastEvent)
					lstItm.SubItems.Add(macroEvent.ClonedMouseEventArgs.Button)
					lstItm.SubItems.Add(macroEvent.ClonedMouseEventArgs.Clicks)
					lstItm.SubItems.Add(macroEvent.ClonedMouseEventArgs.Delta)

					'@@@@lstItm.Tag = macroEvent.Time_SinceLastEvent

				Case MacroEventType.MouseDown

					Dim ItemText As String = ""

					Select Case macroEvent.ClonedMouseEventArgs.Button
						Case Windows.Forms.MouseButtons.Left
							ItemText = "LeftMouseDown"
						Case Windows.Forms.MouseButtons.Middle
							ItemText = "MiddleMouseDown"
						Case Windows.Forms.MouseButtons.Right
							ItemText = "RightMouseDown"
					End Select

					lstItm = lstviwMacEdit.Items.Add(ItemText)

					'lstItm.ForeColor = Color.BlueViolet
					lstItm.ForeColor = GetItemColor(lstItm)

					'//gomla morkaba .. adding a new subitems then assign The X and Y to the Tag value 
					'//in order to use later in ToolStripSveMac_Click
					lstItm.SubItems.Add("X = " & macroEvent.ClonedMouseEventArgs.X).Tag = macroEvent.ClonedMouseEventArgs.X
					lstItm.SubItems.Add("Y = " & macroEvent.ClonedMouseEventArgs.Y).Tag = macroEvent.ClonedMouseEventArgs.Y

					'//Button Clicks Delta : Storing Wait, Button, Clicks & Delta in the column
					'// in order to use in "ToolStripSveMac_Click" ( i.e. Save the edited Macro )
					lstItm.SubItems.Add(macroEvent.Time_SinceLastEvent)
					lstItm.SubItems.Add(macroEvent.ClonedMouseEventArgs.Button)
					lstItm.SubItems.Add(macroEvent.ClonedMouseEventArgs.Clicks)
					lstItm.SubItems.Add(macroEvent.ClonedMouseEventArgs.Delta)

					'@@@@lstItm.Tag = macroEvent.Time_SinceLastEvent

				Case MacroEventType.MouseUp

					Dim ItemText As String = ""

					Select Case macroEvent.ClonedMouseEventArgs.Button
						Case Windows.Forms.MouseButtons.Left
							ItemText = "LeftMouseUp"
						Case Windows.Forms.MouseButtons.Middle
							ItemText = "MiddleMouseUp"
						Case Windows.Forms.MouseButtons.Right
							ItemText = "RightMouseUp"
					End Select

					lstItm = lstviwMacEdit.Items.Add(ItemText)

					'lstItm.ForeColor = Color.DarkBlue
					lstItm.ForeColor = GetItemColor(lstItm)

					'//gomla morkaba .. adding a new subitems then assign The X and Y to the Tag value 
					'//in order to use later in ToolStripSveMac_Click
					lstItm.SubItems.Add("X = " & macroEvent.ClonedMouseEventArgs.X).Tag = macroEvent.ClonedMouseEventArgs.X
					lstItm.SubItems.Add("Y = " & macroEvent.ClonedMouseEventArgs.Y).Tag = macroEvent.ClonedMouseEventArgs.Y

					'//Button Clicks Delta : Storing Wait, Button, Clicks & Delta in the column
					'// in order to use in "ToolStripSveMac_Click" ( i.e. Save the edited Macro )
					lstItm.SubItems.Add(macroEvent.Time_SinceLastEvent)
					lstItm.SubItems.Add(macroEvent.ClonedMouseEventArgs.Button)
					lstItm.SubItems.Add(macroEvent.ClonedMouseEventArgs.Clicks)
					lstItm.SubItems.Add(macroEvent.ClonedMouseEventArgs.Delta)

					'@@@@lstItm.Tag = macroEvent.Time_SinceLastEvent

				Case MacroEventType.KeyDown

					keyItm = lstviwMacEdit.Items.Add("KeyDown")

					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.KeyCode.ToString)

					'keyItm.ForeColor = Color.OliveDrab
					keyItm.ForeColor = GetItemColor(keyItm)

					'//Button Clicks Delta : Storing Wait, KeyEvents.Alt -- KeyEvents.Control 
					'//-- KeyEvents.Handled -- KeyEvents.KeyData -- KeyEvents.KeyValue -- KeyEvents.Modifiers 
					'//-- KeyEvents.Shift -- KeyEvents.SuppressKeyPress in the column
					'// in order to use in "ToolStripSveMac_Click" ( i.e. Save the edited Macro )
					keyItm.SubItems.Add(macroEvent.Time_SinceLastEvent)
					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.Alt)
					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.Control)
					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.Handled)
					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.KeyData)
					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.KeyValue)
					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.Modifiers)
					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.Shift)
					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.SuppressKeyPress)

					'@@@@@keyItm.Tag = macroEvent.Time_SinceLastEvent

				Case MacroEventType.KeyUp

					keyItm = lstviwMacEdit.Items.Add("KeyUp")

					'keyItm.ForeColor = Color.Purple
					keyItm.ForeColor = GetItemColor(keyItm)

					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.KeyCode.ToString)

					'//Button Clicks Delta : Storing Wait, KeyEvents.Alt -- KeyEvents.Control 
					'//-- KeyEvents.Handled -- KeyEvents.KeyData -- KeyEvents.KeyValue -- KeyEvents.Modifiers 
					'//-- KeyEvents.Shift -- KeyEvents.SuppressKeyPress in the column
					'// in order to use in "ToolStripSveMac_Click" ( i.e. Save the edited Macro )
					keyItm.SubItems.Add(macroEvent.Time_SinceLastEvent)
					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.Alt)
					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.Control)
					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.Handled)
					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.KeyData)
					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.KeyValue)
					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.Modifiers)
					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.Shift)
					keyItm.SubItems.Add(macroEvent.ClonedKeyEventArgs.SuppressKeyPress)

					'@@@@@keyItm.Tag = macroEvent.Time_SinceLastEvent

			End Select

		Next

		frmProgress.Dispose()

		lstviwMacEdit.EndUpdate()

		'منح التركيز لإول عنصر في القائمة , حتي يتم إضافة أي إضافة من إضافات التحرير بعده مباشرة
		'في حالة ما إذا لم يختار المستخدم عنصر ليتم الإدراج بعده
		lstviwMacEdit.Select()
		lstviwMacEdit.Items(0).Selected = True

		HideScriptListHorzScrollBars()

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub ShowCuteMacroToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowCuteMacroToolStripMenuItem.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		ShowMe()

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub toolstripRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles toolstripRecord.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim be As System.EventArgs = Nothing
		Dim bsender As System.Object = Nothing

		btnRecord_Click(bsender, be)

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub toolstripStopRecord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles toolstripStopRecord.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Me.StopMacro()
		Me.ShowMe()

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub toolstripPlyBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles toolstripPlyBack.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim be As System.EventArgs = Nothing
		Dim bsender As System.Object = Nothing

		BtnPlayback_Click(bsender, be)

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub toolstripOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles toolstripOptions.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim be As System.EventArgs = Nothing
		Dim bsender As System.Object = Nothing

		ButOptions_Click(bsender, be)

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub ToolStripEditMac_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripOpenMacro.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'If Not Debugger.IsAttached Then
		'	If is_x2342 = "False" Then
		'		MsgBox("You cannot load a macro in the evaluation version.", MsgBoxStyle.Information, "Perfect Macro Recorder")
		'		Exit Sub
		'	End If
		'	If is_x2342 = "False" Then
		'		Exit Sub
		'	End If
		'End If

		Dim DiaFileName As String = ShowOpenDialog("Select a macro file to edit")

		If DiaFileName = "" Then Exit Sub

		Dim FileNames() As String = ShowOpen.FileNames

		If FileNames.Length >= 1 Then

			Dim FileName As String

			Dim Res As Result = Nothing

			If FileNames.Length > 1 Then
				'Before you make the merge you must clear the list otherwise
				'the content of the list will be playbacked beside those files you add
				lstviwMacEdit.Items.Clear()
			End If

			For Each FileName In FileNames

				Dim FileExt As String = Strings.Right(FileName, 4)

				If FileNames.Length = 1 Then '//merge only if more than one file 
					If FileExt = ".exe" Then
						Res = PlaybackTheExtractedMacFile("", FileName, Merge:=False)
					ElseIf FileExt = "pmac" Then
						'تم تعديل السطر التالي بعد أزمة العاشر من أكتوبر سنة 2009 الشهيرة 
						'عندما اكتشفت ان عنصر الإنتظار سوف يخلق أزمة رهيبة سواء في أسلوب حفظه 
						'أو استرجاعه أو إعادة تشغيله من جديد ومن ثم فقد استبدلت منطق البرنامج
						'بمنطق الإعتماد علي حفظ الليست فيو كما هو بدون أسلوب التسلسل للماكرو 
						'الشهير ثم إعادة تحميل هذا الملف .. مع ملاحظة إنني انوي بالفعل حفظ الليست فيو غدا 
						'عن طريق التسلسل
						LoadListViewFromCSV(FileName, lstviwMacEdit, False)
						'LoadMacList(FileName, Merge:=False)
					End If
				Else
					If FileExt = ".exe" Then
						Res = PlaybackTheExtractedMacFile("", FileName, Merge:=True)
					ElseIf FileExt = "pmac" Then
						LoadListViewFromCSV(FileName, lstviwMacEdit, True)
					End If
				End If '//if we did not add this if stat. then evert click to edit button won't clear the list
				'// from perivous files

			Next

			If Res.Successed = False Then
				If Res.Message <> "" Then
					MsgBox(Res.Message, MsgBoxStyle.Critical, Me.Text)
				End If
			End If

		End If



	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub MainForm_DragEnter(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles MyBase.DragEnter
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		If e.Data.GetDataPresent(DataFormats.FileDrop) Then
			e.Effect = DragDropEffects.Copy
		Else
			e.Effect = DragDropEffects.None
		End If

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub MainForm_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles MyBase.DragDrop
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'If is_x2342 = "False" Then
		'MsgBox("Drag and Drop macro files is disabled in the evaluation version. ", MsgBoxStyle.Information, "Perfect Macro Recorder")
		'Exit Sub
		'End If

		Dim FilePath As String = ""

		If e.Data.GetDataPresent(DataFormats.FileDrop) Then

			Dim filePaths As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())

			'For Each fileLoc As String In filePaths

			'' Code to read the contents of the text file

			'If File.Exists(fileLoc) Then

			'Using tr As TextReader = New StreamReader(fileLoc)

			'FilePath = fileLoc

			'End Using

			'End If

			'Next fileLoc

			lstviwMacEdit.Items.Clear()	'//if we do not do this, the user can drag the same file
			'// more than once. he can drag then drag it and load the list with the same file twice
			'// or more .

			For Each fileLoc As String In filePaths

				' Code to read the contents of the text file

				If File.Exists(fileLoc) Then

					Dim Res As Result = Nothing
					Dim FileExt As String = Strings.Right(fileLoc, 4)
					If filePaths.Length = 1 Then '//merge only if more than one file 
						If FileExt = ".exe" Then
							Res = PlaybackTheExtractedMacFile("", fileLoc, Merge:=False)
							If Res.Successed = False Then
								MsgBox(Res.Message, MsgBoxStyle.Critical, "Perfect Macro Recorder")
							End If
						ElseIf FileExt = "pmac" Then
							'تم تعديل السطر التالي بعد أزمة العاشر من أكتوبر سنة 2009 الشهيرة 
							'عندما اكتشفت ان عنصر الإنتظار سوف يخلق أزمة رهيبة سواء في أسلوب حفظه 
							'أو استرجاعه أو إعادة تشغيله من جديد ومن ثم فقد استبدلت منطق البرنامج
							'بمنطق الإعتماد علي حفظ الليست فيو كما هو بدون أسلوب التسلسل للماكرو 
							'الشهير ثم إعادة تحميل هذا الملف .. مع ملاحظة إنني انوي بالفعل حفظ الليست فيو غدا 
							'عن طريق التسلسل
							'LoadMacList(fileLoc, Merge:=False)
							LoadListViewFromCSV(fileLoc, lstviwMacEdit, False)
						End If
					Else
						If FileExt = ".exe" Then
							Res = PlaybackTheExtractedMacFile("", fileLoc, Merge:=True)
							If Res.Successed = False Then
								MsgBox(Res.Message, MsgBoxStyle.Critical, "Perfect Macro Recorder")
							End If
						ElseIf FileExt = "pmac" Then
							'LoadMacList(fileLoc, Merge:=True)
							LoadListViewFromCSV(fileLoc, lstviwMacEdit, True)
						End If
					End If '//if we did not add this if stat. then evert click to edit button won't clear the list
					'// from perivous files

					'LoadMacList(fileLoc, Merge:=True)

				End If

			Next fileLoc

		End If

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub ToolStripSveMac_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripSveMac.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'Dim IsTrialVer As Boolean
		'Dim IsUnLocked As Boolean

		'If is_x2342 = "False" Then
		'	IsTrialVer = True
		'	IsUnLocked = False
		'ElseIf is_x2342 = "True" Then
		'	IsTrialVer = False
		'	IsUnLocked = True
		'End If

		'If Not Debugger.IsAttached Then
		'	If IsTrialVer Then
		'		MsgBox("You cannot save macro in the evaluation version.", MsgBoxStyle.Information, "Perfect Macro Recorder")
		'		Exit Sub
		'	End If
		'	If Not IsUnLocked Then Exit Sub
		'End If
		'If is_x2342 = "True" Then
		Dim Path As String = ShowSaveDialog("Perfect Macro Recorder" & " Files (*.pmac)|*.pmac")
		If Path = "" Then Exit Sub
		SaveListViewToCSV(lstviwMacEdit, Path)
		'End If
		'تبعا للمنطق الجديد , لم نعد نستخدم أو نستدعي هذه الدالة في إستخدام حيوي
		'وقد تم هذا فيما يعرف بأزمة العاشر من أكتوبر لعام 2009 عندما إكتشف ان
		'عنصر الإنتظار سوف يسبب أزمة بسبب لغظ كبير يمكن في المنطق الخاص به .. يتراوح
		'فيما بين إعتماده علي العنصر الذي يليه وبين عدم فاعليته من الأساس في البلاي باك
		'SaveListView()

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Friend Function ShowSaveDialog(ByVal Filter As String) As String
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		ShowSave.Title = "Save Macro File As"

		ShowSave.Filter = Filter '"Perfect Macro Recorder" & " Files (*.pmac)|*.pmac"

		ShowSave.AddExtension = True

		ShowSave.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop

		ShowSave.AutoUpgradeEnabled = True

		Dim res As DialogResult = ShowSave.ShowDialog()

		If res = Windows.Forms.DialogResult.OK Then
			ShowSaveDialog = ShowSave.FileName
		ElseIf res = Windows.Forms.DialogResult.Cancel Then
			ShowSave.FileName = ""
		End If

		Return ShowSave.FileName

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Function addMouseEveToStoredList(ByVal i As Integer) As MouseEventArgs
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'//Button Clicks Delta : Storing Wait, Button, Clicks & Delta in the column
		'// in order to use in "ToolStripSveMac_Click" ( i.e. Save the edited Macro )

		Dim eMouse As New System.Windows.Forms.MouseEventArgs( _
		lstviwMacEdit.Items.Item(i).SubItems(4).Text, _
		lstviwMacEdit.Items.Item(i).SubItems(5).Text, _
		lstviwMacEdit.Items.Item(i).SubItems(1).Tag, _
		lstviwMacEdit.Items.Item(i).SubItems(2).Tag, _
		lstviwMacEdit.Items.Item(i).SubItems(6).Text)

		addMouseEveToStoredList = eMouse

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Function addKeyEveToStoredList(ByVal i As Integer) As KeyEventArgs
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'//Button Clicks Delta : Storing Wait 2 , KeyEvents.Alt 3 -- KeyEvents.Control 4
		'//-- KeyEvents.Handled 5 -- KeyEvents.KeyData 6 -- KeyEvents.KeyValue 7 -- 
		'// KeyEvents.Modifiers(8) -- KeyEvents.Shift 9 -- KeyEvents.SuppressKeyPress 10 in the column
		'// in order to use in "ToolStripSveMac_Click" ( i.e. Save the edited Macro )

		Dim eKey As New KeyEventArgs(lstviwMacEdit.Items.Item(i).SubItems(7).Text)

		eKey.SuppressKeyPress = lstviwMacEdit.Items.Item(i).SubItems(10).Text
		eKey.Handled = lstviwMacEdit.Items.Item(i).SubItems(5).Text
		addKeyEveToStoredList = eKey

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Friend Function SaveListView(Optional ByVal Path As String = "") As Boolean
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'//my twits : 20090912 : hehehehehehe ;)
		'//أن يكون الاختبار ذكياً قدر ما تستطيع .. وذلك بأن يركز علي اختبار 
		'//قلب الوظيفة الرئيسي التي يختبرها وليس مجرد تنفيذ متتالي لسطورها 

		Dim listStoredevents As New ArrayList

		Dim SavedFilePath As String

		If Path <> "" Then
			SavedFilePath = Path
		Else
			SavedFilePath = ShowSaveDialog("Perfect Macro Recorder" & " Files (*.pmac)|*.pmac")
		End If

		If SavedFilePath = "" Then Exit Function

		For i = 0 To lstviwMacEdit.Items.Count - 1

			Select Case lstviwMacEdit.Items.Item(i).Text

				Case "MouseMove"

					Dim emouse As MouseEventArgs = addMouseEveToStoredList(i)

					listStoredevents.Add(New MacroEvent(MacroEventType.MouseMove, emouse, lstviwMacEdit.Items.Item(i).SubItems(3).Text))

				Case "LeftMouseDown", "MiddleMouseDown", "RightMouseDown"

					Dim emouse As MouseEventArgs = addMouseEveToStoredList(i)

					listStoredevents.Add(New MacroEvent(MacroEventType.MouseDown, emouse, lstviwMacEdit.Items.Item(i).SubItems(3).Text))

				Case "LeftMouseUp", "MiddleMouseUp", "RightMouseUp"

					Dim emouse As MouseEventArgs = addMouseEveToStoredList(i)

					listStoredevents.Add(New MacroEvent(MacroEventType.MouseUp, emouse, lstviwMacEdit.Items.Item(i).SubItems(3).Text))

				Case "KeyDown"

					Dim ekey As KeyEventArgs = addKeyEveToStoredList(i)

					listStoredevents.Add(New MacroEvent(MacroEventType.KeyDown, ekey, lstviwMacEdit.Items.Item(i).SubItems(2).Text))

				Case "KeyUp"

					Dim ekey As KeyEventArgs = addKeyEveToStoredList(i)

					listStoredevents.Add(New MacroEvent(MacroEventType.KeyUp, ekey, lstviwMacEdit.Items.Item(i).SubItems(2).Text))

			End Select

		Next

		SaveListView = SaveSerializedObject(SavedFilePath, listStoredevents)

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub ToolStripNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripNew.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		lstviwMacEdit.Items.Clear()
		PasteToolStripButton.Enabled = False
		If Clipboard.ContainsData("CMFormat") Then
			Clipboard.Clear()
		End If

	End Sub

	Private Sub SuspendMeSomeWhile(ByVal Delay As Integer)
		Dim OldTickCount As Integer = TickCount

		Dim DelayValue As Integer

		Do Until DelayValue >= Delay

			DelayValue = TickCount - OldTickCount
			Application.DoEvents()

		Loop
	End Sub

	Private Sub IWannaSleep(ByVal index As Integer, ByVal WaitItemIndex As Integer, Optional ByVal ExeSpeed As Integer = 2)

		Dim PlayBackSpeed As Integer

		PlayBackSpeed = GetPlaybackSpeed()

		If PlayBackSpeed = 2 Then '//normal playback

			Dim Delay As Integer = lstviwMacEdit.Items.Item(index).SubItems(1).Text
			SuspendMeSomeWhile(Delay)
			'اكتشفت انه لسبب ما لاحاجة لوجود السطر التالي أصلا .. إذ أن البرنامج يطبق فترات 
			'التأخير من نفسه ! ولا أعلم حقيقة كيف .. ولكن علي اية حال انا احتاج بالفعل 
			'للإستغناء عن السطر القادم ,, حتي يتسني للمستخدم التعامل مع البرنامج أثناء فترات 
			'التأخير الطوييييييييييلة .. مثل 50000000
			'لأنني وجدت أن البرنامج كله يقف عن العمل ولا تعمل حتي الهوكيز وبالتالي .. هي مشكلة فعلاً
			'وقد استغنيت عن السطر القادم بالوظيفة أعلاه
			'Thread.Sleep(lstviwMacEdit.Items.Item(index).SubItems(1).Text)

			'If lstviwMacEdit.Items.Item(index - 1).Text = "Wait" And index <> 0 Then
			'    Thread.Sleep(lstviwMacEdit.Items.Item(index - 1).SubItems(1).Text)
			'End If

			'11/10/2009
			'تم تعديل هذه الوظيفة بما يلائم جعل عنصر الإنتظار عنصرا قائماً بحد ذاته 
			'وليس مجرد إنعكاس للعنصر الذي يليه , وبالتالي يمكنني إظافة عناصر إنتظار
			'تعمل بغذ النظر عن العنصر الذي يليها
			'Thread.Sleep(lstviwMacEdit.Items.Item(index).SubItems(WaitItemIndex).Text)

		Else

			Thread.Sleep(PlayBackSpeed)

		End If

		My.Application.DoEvents()

	End Sub

	Private Sub ShowDesktop()
		keybd_event(VK_LWIN, 0, 0, 0)
		keybd_event(77, 0, 0, 0)
		keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
	End Sub

	Private Sub ActivateBackWindow(ByVal lstItem As ListViewItem, ByVal WndHandle As IntPtr)
		Dim WinTop As Integer = lstItem.SubItems(3).Text
		Dim WinLeft As Integer = lstItem.SubItems(2).Text
		Dim NewWidth As Integer = lstItem.SubItems(4).Text - WinLeft
		Dim NewHight As Integer = lstItem.SubItems(5).Text - WinTop

		ShowWindow(WndHandle, 1) '//restore the window from taskbar it it was there

		Dim Res As Integer = SetForegroundWindow(WndHandle)	'// activate

		SetWindowPos(WndHandle, HWND_NOTOPMOST, WinLeft, _
		WinTop, NewWidth, NewHight, SWP_SHOWWINDOW)	'//adjust
	End Sub

	Private Sub ShowPlaybackAtSystray()
		TrayIcon.Icon = My.Resources.Playback
		TrayIcon.Text = "Playback..."
	End Sub

	Private Function IsBackGroundWindowDisplayed(ByVal windowText As String) As IntPtr
		Return FindWindow(vbNullString, windowText)
	End Function

	Private Sub SearchAndShowBackGroundWindow(ByVal lstItem As ListViewItem)

		Dim WndHandle As IntPtr = IsBackGroundWindowDisplayed(lstItem.SubItems(1).Text)
		If WndHandle <> 0 Then
			If lstItem.SubItems(1).Text.Contains("Program Manager") Then
				ShowDesktop()
			Else '//any window else
				ActivateBackWindow(lstItem, WndHandle)
			End If
		Else
			'بمعني ماذا سأفعل ما إذا كانت النافذة غير موجودة بالفعل؟
			'أتجاهل كل الخيارات الخاصة بها
			'أو
			'إنتظر مدة معلومة
			If lstItem.SubItems(7).Text = "True" Then
				'النوم مؤقتا فربما تظهر النافذة
				SuspendMeSomeWhile(lstItem.SubItems(6).Text)
				'Thread.Sleep(lstItem.SubItems(6).Text)
				'trying again:
				Dim WndHandle2 As IntPtr = IsBackGroundWindowDisplayed(lstItem.SubItems(1).Text)
				If WndHandle2 <> 0 Then
					ActivateBackWindow(lstItem, WndHandle2)
				Else
					'28/09/2010~Omitted~"if the window isn't there ignore all the actions realted to it""
					'بتاريخ 28/09/2010 تم الغوص لمستوي جديد من ذكاء هذا التطبيق
					'لن يتم تجاهل النافذة في حالة عدم وجودها .. ستيم البحث عنها
					'ومن ثم فتحها وتشغليها وهذه هي البداية يا رج
					Try
						'Process.Start(lstItem.SubItems(9).Text) 'FoxSmr
					Catch ex As Exception
						'Dim runningProcessError As String
						'runningProcessError = "Cannot starting a background window titled " &
						'"""" & lstItem.SubItems(1).Text & """" & " in which the path to its process is " & """" &
						'.SubItems(9).Text.Trim & """"
						'WriteToErrorLog(ex.Message, ex.StackTrace, runningProcessError)
					End Try
					'28/09/2010
					'مر ما يقرب علي العامين منذ بدء هذا التطبيق في 16 اكتوبر 2008
					'ويكفي ان تعمل نبذة عن الاحداث التي عاصرها من أول محاولات السعي مع
					'ني ني ,, مروراً بهاجر .. وانتهاءاً بنهي ومريم
					'مريم الطفلة الجميلة التي دخلت حياتي لتضيف إليها جديدًا كل يوم ..
					'كم أحب تلك الوغدة الصغيرة
					'لا, لم احقق حلمي بعد في مال كثير وفير من هذا العمل
					'أو افكار قوية تحقق الحلم الأعظم
					'ربنا يسهل
				End If

			End If
		End If

	End Sub

	Private Function PlaybackTheList(ByVal SelectionOnly As Boolean, Optional ByVal ExeSettings As String() = Nothing) As Boolean

		Dim isSmartPlayback As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "ChkSmartPlayback", 1)

		'If isSmartPlayback Then 'FoxSmr
		'Dim backgroundAppsNamesAndPathes As Hashtable = collectAppsPaths() 'Get Process Names And Paths from macro list
		'Dim table As IDictionaryEnumerator = backgroundAppsNamesAndPathes.GetEnumerator 'Now check if it is running 
		'While table.MoveNext
		'Dim proc() As Process
		'proc = isProcessRunning(table.Key.ToString)
		'If proc.Length = 0 Then 'process is not running
		'Try
		'Dim myproc As Process = Process.Start(table.Value.ToString)
		'Shell(table.Value.ToString, AppWinStyle.MinimizedFocus)
		'Catch ex As Exception
		' runningProcessError As String
		'runningProcessError = "Cannot starting a background window titled " &
		'"""" & lstItem.SubItems(1).Text & """" & " in which the path to its process is " & """" &
		'lstItem.SubItems(9).Text.Trim & """"
		'WriteToErrorLog(ex.Message, ex.StackTrace, runningProcessError)
		'End Try
		'Else 'process is running
		'Then skip it; nothing to do man ;)
		'End If
		'End While
		'End If

		'RepeatTime = ExeSettings(0)
		'PlaybackSpeed = ExeSettings(1)
		'chkNotify.Checked = ExeSettings(2)
		'chkCompileSelected.Checked = ExeSettings(3)
		'chkEnableEditing.Checked = ExeSettings(4)

		SuspendPlaybackingTheList = False '//just to make sure that we clear that global var (Grrrr, hate glob.)

		Me.WindowState = FormWindowState.Minimized

		IsTheListPlaybacked = True '//Used in WndProc to avoid calling the thread lines

		TmrStop.Enabled = True '//Using this time to terminate Play back if the user wishes to 

		IsPlayBack = True

		Dim RepeatCount As Int32
		Dim PlayBackSpeed As Integer
		RepeatCount = GetPlyBckRepeatCount()
		PlayBackSpeed = GetPlaybackSpeed()

		ShowPlaybackAtSystray()

		Dim IsTipsAllowed As Boolean = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkTips", 1)

		If IsTipsAllowed Then
			Dim stopPlaybackHotkey As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & "txtStopPly", "HotKey", "")
			Dim pause_resume_PlaybackHotKey As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\HK\" & "txtStrPly", "HotKey", "")
			Dim balloonMessage As String = _
			"Press " & stopPlaybackHotkey & " to stop playback," & vbNewLine & _
			"Press " & pause_resume_PlaybackHotKey & " to pause/resume playback"

			TrayIcon.ShowBalloonTip(500, "Perfect Macro Recorder is now playbacking.", balloonMessage, ToolTipIcon.Info)
		End If

		For g = 1 To RepeatCount

			Dim ListItems As Object

			Dim lstItem As ListViewItem

			If SelectionOnly Then

				ListItems = New ListView.SelectedListViewItemCollection(lstviwMacEdit)
			Else

				ListItems = New ListView.ListViewItemCollection(lstviwMacEdit)

			End If

			For Each lstItem In ListItems

				'//Trace Though items -- Now you can trace through lines to detect which line 24/02/2010 "after publish"
				lstviwMacEdit.SelectedItems.Clear()
				lstItem.Selected = True
				lstItem.EnsureVisible()

				''''''''''''''''''''''''''''''''''''''''''''''''''''' 
				If StopPlaybackingTheList Then '//to process the "stop playback hotkey"

					StopPlaybackingTheList = False

					Exit For

				End If

				If SuspendPlaybackingTheList Then

					Do While SuspendPlaybackingTheList

						Application.DoEvents()

					Loop

					SuspendPlaybackingTheList = False

				End If
				'''''''''''''''''''''''''''''''''''''''''''''''''''''

				If Not IsPlayBack Then '//Terminate Play back whenever IsPlayback set to Flase in TmrStop

					Exit For '//It is just a way to stop play back by a hot key from within a timer

				End If

				Select Case lstItem.Text

					Case "Wait"

						IWannaSleep(lstItem.Index, 0, PlayBackSpeed)

					Case "Background Window"

						If isSmartPlayback Then
							SearchAndShowBackGroundWindow(lstItem)
						End If

					Case "MouseWheel"

						MouseSimulator.MouseWheel(lstItem.SubItems(1).Text)	'//Pass Deleta
						MouseSimulator.X = lstviwMacEdit.Items.Item(lstItem.Index).SubItems(2).Tag
						MouseSimulator.Y = lstviwMacEdit.Items.Item(lstItem.Index).SubItems(3).Tag

					Case "MouseMove"

						'IWannaSleep(lstItem.Index, 3, PlayBackSpeed)

						MouseSimulator.X = lstviwMacEdit.Items.Item(lstItem.Index).SubItems(1).Tag
						MouseSimulator.Y = lstviwMacEdit.Items.Item(lstItem.Index).SubItems(2).Tag

					Case "LeftMouseDown", "MiddleMouseDown", "RightMouseDown"

						'IWannaSleep(lstItem.Index, 3, PlayBackSpeed)

						If lstviwMacEdit.Items.Item(lstItem.Index).SubItems(4).Text = 2097152 Then
							MouseSimulator.MouseDown(MouseButton.Right)
						ElseIf lstviwMacEdit.Items.Item(lstItem.Index).SubItems(4).Text = 1048576 Then
							MouseSimulator.MouseDown(MouseButton.Left)
						End If

					Case "LeftMouseUp", "MiddleMouseUp", "RightMouseUp"

						'IWannaSleep(lstItem.Index, 3, PlayBackSpeed)

						If lstviwMacEdit.Items.Item(lstItem.Index).SubItems(4).Text = 2097152 Then
							MouseSimulator.MouseUp(MouseButton.Right)
						ElseIf lstviwMacEdit.Items.Item(lstItem.Index).SubItems(4).Text = 1048576 Then
							MouseSimulator.MouseUp(MouseButton.Left)
						End If

					Case "KeyDown"

						If lstviwMacEdit.Items.Count = -1 Then
							Debug.Print("catch you")
						End If

						'IWannaSleep(lstItem.Index, 2, PlayBackSpeed)

						KeyboardSimulator.KeyDown(lstviwMacEdit.Items.Item(lstItem.Index).SubItems(7).Text)

					Case "KeyUp"

						If lstItem.Index = -1 Then
							Debug.Print("catch you")
						End If

						'IWannaSleep(lstItem.Index, 2, PlayBackSpeed)

						KeyboardSimulator.KeyUp(lstviwMacEdit.Items.Item(lstItem.Index).SubItems(7).Text)

				End Select

			Next

		Next

		RestoreSystrayIcon()

		IsTheListPlaybacked = False	'//Used in WndProc to avoid calling the thread lines

		IsPlayBack = False

		TmrStop.Enabled = False

		Dim IsAppTobeClosedAftPlyback As Boolean

		IsAppTobeClosedAftPlyback = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\CuteMacro\opt\", "chkCloseCM", 0)

		If IsAppTobeClosedAftPlyback Then

			'//because Cross-thread operation not valid: Control 'MainForm' accessed from 
			'//a thread other than the thread it was created on. Hence i cannot call Me.Dispose() to close the app
			'//so i will use application.exit
			Application.Exit()

		End If

		'chkNotify.Checked = ExeSettings(2)

		NotifyAtPlybckEnding()

		Return True

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub ToolStripRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripRun.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		If lstviwMacEdit.Items.Count < 1 Then Exit Sub

		If lstviwMacEdit.SelectedItems.Count <= 1 Then
			PlaybackTheList(False)
		ElseIf lstviwMacEdit.SelectedItems.Count > 1 Then
			PlaybackTheList(True)
		End If

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Function ShowOpenDialog(ByVal DialogTitle As String) As String
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Try

			ShowOpen.Title = DialogTitle ' "Select a macro file to edit" or "Select a macro file to play back"

			ShowOpen.Filter = "Perfect Macro Recorder" & " Files (*.pmac)|*.pmac|Executable " & "Perfect Macro Recorder" & " Files (*.exe)|*.exe"

			ShowOpen.AddExtension = True

			ShowOpen.AutoUpgradeEnabled = True

			If DialogTitle = "Select a macro file to edit" Then
				ShowOpen.Multiselect = True
			End If

			ShowOpen.FileName = ""

			ShowOpen.ShowDialog()

			ShowOpenDialog = ShowOpen.FileName

		Catch ex As Exception

			ShowOpenDialog = "there is a n exception"

		End Try



	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub lstviwMacEdit_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lstviwMacEdit.MouseClick
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim ClickedListViewItem As ListViewHitTestInfo

		ClickedListViewItem = lstviwMacEdit.HitTest(e.X, e.Y)

		'الضغط علي زر ال
		'Alt
		'مع زر الماوس الأيسر ,, يختار
		'كل عناصر النوع الواحد
		If e.Button = Windows.Forms.MouseButtons.Left And GetAsyncKeyState(VK_MENU) Then
			Select Case lstviwMacEdit.SelectedItems.Item(0).Text
				Case "Background Window"
					SelectSpecifiedItems("Background Window")
				Case "Wait"
					SelectSpecifiedItems("Wait")
				Case "MouseWheel"
					SelectSpecifiedItems("MouseWheel")
				Case "MouseMove"
					SelectSpecifiedItems("MouseMove")
				Case "LeftMouseUp", "MiddleMouseUp", "RightMouseUp", "LeftMouseDown", "MiddleMouseDown", "RightMouseDown"
					SelectSpecifiedItems(lstviwMacEdit.SelectedItems.Item(0).Text)
				Case "KeyUp"
					SelectSpecifiedItems("KeyUp")
				Case "KeyDown"
					SelectSpecifiedItems("KeyDown")
			End Select
		End If

		If e.Button = Windows.Forms.MouseButtons.Right Then

			If Not ClickedListViewItem Is Nothing Then
				Dim ItemLocation As Point
				ItemLocation.X = e.X : ItemLocation.Y = e.Y
				'///Prevent Multipal MoveUp & MoveDown
				If lstviwMacEdit.SelectedItems.Count > 1 Then
					lstviewStripMoveUp.Enabled = False
					lstviewStripMoveDown.Enabled = False
				Else
					lstviewStripMoveUp.Enabled = True
					lstviewStripMoveDown.Enabled = True
				End If

				If lstviwMacEdit.SelectedItems.Count > 1 Then

					ToolStripEditCommand.Visible = False
					InsertToolStripMenuItem.Visible = False

					'لقد قررت إن أظهر ال
					'BulkeditTimeouts
					'في ال
					'Right Click menu on the list
					'فقط اذا ما كانت كل العناصر من نوع الإنتظار
					Dim SelItemS_Type As String = "" 'byref one
					GetSelectedItemsType(SelItemS_Type)

					Select Case SelItemS_Type
						Case "KeyUp", "KeyDown", "Wait"
							ToolStripEditCommand.Visible = True
					End Select

				ElseIf lstviwMacEdit.SelectedItems.Count = 1 Then

					ToolStripEditCommand.Visible = True
					InsertToolStripMenuItem.Visible = True

					'لقد قررت إن أظهر ال
					'BulkeditTimeouts
					'في ال
					'Right Click menu on the list
					'فقط اذا ما كانت كل العناصر من نوع الإنتظار
					'ولن أظهره ما اذا كان عنصر واحد هو المختار
					'حتي لو كان عنصر إنتظار
					'BulkeditTimeouts.Visible = False
				End If

				'// we add "lstviwMacEdit" as the reference Point so we can show the popup menu at the item 
				EditRightClickList.Show(lstviwMacEdit, ItemLocation)

			End If

		End If

	End Sub

	Friend Function GetSelectedItemsType(ByRef ItemType As String) As Boolean

		'لو كان عدد العناصر المختارة أكبر من واحد .. فهذه الوظيفة يمكنها معرفة هذه العنصر
		'في حالة ما إذا كان عنصر واحد .. أما لو كان خليط من العناصر فسوف تكون قيمتها
		'False المرجعية
		'ولو كان الخليط متجانس , فسوف ترجع فيمة 
		'True
		'وسترجع أيضا نوع العنصر في البارامتر الوحيد لها

		Dim anyItemFromSelectedItems As ListViewItem = lstviwMacEdit.SelectedItems.Item(0)

		Dim lstItem As ListViewItem
		For Each lstItem In lstviwMacEdit.SelectedItems
			If lstItem.Text <> anyItemFromSelectedItems.Text Then
				GetSelectedItemsType = False
				Exit Function
			End If
		Next
		ItemType = anyItemFromSelectedItems.Text
		GetSelectedItemsType = True
	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub DeleteListItems()
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		frmProgress.Tag = "Deleting"
		frmProgress.ShowDialog()
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub lstviewStripDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstviewStripDelete.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		DeleteListItems()
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub lstviwMacEdit_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lstviwMacEdit.KeyUp
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		If e.KeyValue = Keys.Delete Then DeleteListItems()

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub MoveUpListViewItem()
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim SelItem As ListViewItem
		Dim tmpItem As ListViewItem

		'//Prevent Multipal MoveUp & MoveDown
		If lstviwMacEdit.SelectedItems.Count > 1 Then Exit Sub

		For Each SelItem In lstviwMacEdit.SelectedItems
			Dim SelectedItemIndex As Integer = SelItem.Index
			If SelectedItemIndex = 0 Then Exit Sub '//Exit if first item
			tmpItem = lstviwMacEdit.Items(SelectedItemIndex - 1).Clone()
			lstviwMacEdit.Items(SelectedItemIndex - 1) = lstviwMacEdit.Items(SelectedItemIndex).Clone
			lstviwMacEdit.Items(SelectedItemIndex) = tmpItem.Clone
			lstviwMacEdit.Items(SelectedItemIndex - 1).Selected = True
		Next

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub MoveDownListViewItem()
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim SelItem As ListViewItem
		Dim tmpItem As ListViewItem

		'//Prevent Multipal MoveUp & MoveDown
		If lstviwMacEdit.SelectedItems.Count > 1 Then Exit Sub

		For Each SelItem In lstviwMacEdit.SelectedItems
			Dim SelectedItemIndex As Integer = SelItem.Index
			If SelectedItemIndex = lstviwMacEdit.Items.Count + 1 Then Exit Sub '//Exit if last item

			If lstviwMacEdit.Items.Count = SelectedItemIndex + 1 Then
				'هذه الحالة غير موجودة ومن الممكن تجاهلها كلية عند طريق الآتي
				'If lstviwMacEdit.Items.Count <> SelectedItemIndex + 1 Then
				'ولكني آثرت أن أكودها هكذا للتذكرة .. ان هذه الشائبة تحدث عند
				'تنزيل عنصر للأسفل فيما بعد اخر عنصر .. ولأنه لايوجد أي عناصر بعد
				'اخر عنصر فسيحدث هنا أخطاء كثيرة يا معلم فحواها أن الانديكس المطلوب
				'غير موجود.
				'Debug.Print(lstviwMacEdit.Items.Count)
			Else
				tmpItem = lstviwMacEdit.Items(SelectedItemIndex + 1).Clone()
				lstviwMacEdit.Items(SelectedItemIndex + 1) = lstviwMacEdit.Items(SelectedItemIndex).Clone
				lstviwMacEdit.Items(SelectedItemIndex) = tmpItem.Clone
				lstviwMacEdit.Items(SelectedItemIndex + 1).Selected = True
			End If
		Next

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub lstviewStripMoveUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstviewStripMoveUp.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		MoveUpListViewItem()

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub lstviewStripMoveDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstviewStripMoveDown.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		MoveDownListViewItem()

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub lstviwMacEdit_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lstviwMacEdit.KeyDown
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		If lstviwMacEdit.Items.Count <> 0 Then
			If e.KeyCode = Keys.Right Then
				e.SuppressKeyPress = True
			ElseIf e.KeyValue = Keys.Up And e.Control = True Then
				MoveUpListViewItem()
			ElseIf e.KeyValue = Keys.Down And e.Control = True Then
				MoveDownListViewItem()
			ElseIf e.KeyValue = Keys.A And e.Control = True Then
				SelectAllListViewItems(True)
			ElseIf e.Control = True And e.KeyValue = Keys.C Then
				CopySelectedRowItemsInListView(lstviwMacEdit, "*")
			ElseIf e.Control = True And e.KeyValue = Keys.X Then
				Dim ce As EventArgs = Nothing
				Dim cSender As Object = Nothing
				CutToolStripButton_Click(cSender, ce)
			ElseIf e.Control = True And e.KeyValue = Keys.V Then
				PasteAfter()
			ElseIf e.KeyValue = Keys.Enter Then
				selectItemToEdit(lstviwMacEdit.SelectedItems(0))
			End If
		End If
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub SelectSpecifiedItems(ByVal ItemType As String)
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim lstItem As ListViewItem
		For Each lstItem In lstviwMacEdit.Items
			If lstItem.Text.Contains(ItemType) Then
				lstItem.Selected = True
				lstItem.Checked = True
			End If
		Next

	End Sub
	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Friend Sub CopySelectedRowItemsInListView(ByRef lvwView As ListView, ByVal sSeparator As String)
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'==============================================================================
		' Purpose      :  Copy selected ListView Rows
		' Inputs       :  ListView, sSeparator = " ", or vbTab
		' Returns      :  Copies selected Listview Items and SubItems to Windows clipboard
		' Calls        :  CopySelectedRowItemsInListView(ListView, sSeparator)
		' Use in       :  Any Procedure 
		' Date/Time    :  Monday, August 17, 2009 11:16:21 PM
		'==============================================================================

		Dim Item As ListViewItem
		Dim sItem As String = ""
		Dim sItems As String = ""

		Try
			With lvwView
				For Each Item In .SelectedItems
					' sItem = Item.Text & "!%@" ' For Item
					Dim SubItem As ListViewItem.ListViewSubItem
					For Each SubItem In Item.SubItems
						sItem &= SubItem.Text & sSeparator
					Next
					sItems &= sItem & "#%&"	'//The Separator between Items is "#%&"
					sItem = ""
				Next
			End With

			Dim CallingMethodName As String = ""

			CallingMethodName = GetCallingMethod(2)

			If CallingMethodName = "lstviewStripCut_Click" Or CallingMethodName = "CutToolStripButton_Click" Then
				RemoveSelectedItemsInTheList()
			End If

			If sItems.Length > 0 Then
				CopyRow(sItems)
			End If

		Catch ex As Exception
			MessageBox.Show("Error while coping items", "Copy Error", MessageBoxButtons.OK)
		End Try

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub lstviewStripCut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstviewStripCut.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		CopySelectedRowItemsInListView(lstviwMacEdit, "*")

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub lstviewStripCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstviewStripCopy.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		CopySelectedRowItemsInListView(lstviwMacEdit, "*")

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub PasteAfterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteAfterToolStripMenuItem.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		PasteAfter()

	End Sub

	Private Sub PasteAfter()

		Dim ItemIndex As Integer

		Select Case lstviwMacEdit.SelectedIndices.Count
			Case 1
				ItemIndex = lstviwMacEdit.SelectedIndices.Item(0)
			Case Is > 1
				ItemIndex = lstviwMacEdit.SelectedIndices.Item(lstviwMacEdit.SelectedIndices.Count - 1)
		End Select

		Paste(ItemIndex + 1)

	End Sub
	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub PasteBeforeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteBeforeToolStripMenuItem.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim ItemIndex As Integer = lstviwMacEdit.SelectedIndices.Item(0)

		Paste(ItemIndex)

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Friend Sub CopyRow(ByVal text As String)
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		' Creates a new data format.
		Dim CMFormat As DataFormats.Format = DataFormats.GetFormat("CMFormat")

		' Creates a new object and store it in a DataObject using myFormat 
		' as the type of format. 
		Dim CopiedText As String = text
		Dim DataObject As New DataObject(CMFormat.Name, CopiedText)

		' Copies myObject into the clipboard.
		Clipboard.SetDataObject(DataObject)

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Function GetCopiedListViewItemsInClipboard() As String
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		' Creates a new data format.
		Dim CMFormat As DataFormats.Format = DataFormats.GetFormat("CMFormat")

		' Performs some processing steps.
		' Retrieves the data from the clipboard.
		Dim RetrievedItemsObject As IDataObject = Clipboard.GetDataObject()

		' Converts the IDataObject type to MyNewObject type. 
		Dim CopiedItems As String = CType(RetrievedItemsObject.GetData(CMFormat.Name), String)

		GetCopiedListViewItemsInClipboard = CopiedItems

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub ReplaceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReplaceToolStripMenuItem.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Paste(, True)

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub RemoveSelectedItemsInTheList()
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim RemovedItem As ListViewItem

		lstviwMacEdit.BeginUpdate()

		For Each RemovedItem In lstviwMacEdit.SelectedItems
			RemovedItem.Remove()
		Next

		lstviwMacEdit.EndUpdate()

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub Paste(Optional ByVal Position As Integer = 1, Optional ByVal Replace As Boolean = False)
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		If Not Clipboard.ContainsData("CMFormat") Then Exit Sub

		Dim CopiedItems As String = GetCopiedListViewItemsInClipboard()

		Dim DontStopTheLoop As Boolean = True

		Do While DontStopTheLoop

			Dim RowNumber As Integer

			Dim WhereThesSeparator As Integer = CopiedItems.IndexOf("#%&")

			Dim Row As String = CopiedItems.Substring(0, WhereThesSeparator)

			PasteTheCopiedRowToTheList(Row, Position + RowNumber, Replace)

			RowNumber = RowNumber + 1 '//if i do not update the Position, the bulk copied items will be pasted
			'// reverted "123 321 "

			CopiedItems = CopiedItems.Substring(WhereThesSeparator + 3)

			If CopiedItems = "" Then

				DontStopTheLoop = False

			End If

		Loop

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub PasteTheCopiedRowToTheList(ByVal Row As String, ByVal Position As Integer, Optional ByVal Replace As Boolean = False)
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'================================================================================
		' Purpose      :  Add The Copied (or cut)Row To The List
		' Inputs       :  The Listview Row
		' Returns      :  Copies Checked Listview Items and SubItems to Windows clipboard
		' Calls        :  CopyCheckedListViewRows(ListView, sSeparator)
		' Use in       :  Any Procedure 
		' Date/Time    :  Sunday, October 19, 2003 1:16:04 AM
		'================================================================================

		Dim Item = ""

		Dim WhereThesItem = Row.IndexOf("*")

		Item = Row.Substring(0, WhereThesItem)

		Dim LstItm As New ListViewItem

		Dim ItemIndex As Integer = lstviwMacEdit.SelectedIndices.Item(0)

		If Replace = False Then

			LstItm = lstviwMacEdit.Items.Insert(Position, Item)

			LstItm.ForeColor = GetItemColor(LstItm)

		Else

			If lstviwMacEdit.SelectedIndices.Count = 1 Then	'//One selected Row to replace

				PrepareTheRow(ItemIndex, Item, LstItm)

			ElseIf lstviwMacEdit.SelectedIndices.Count > 1 Then	'//More than a row to replace

				PrepareTheRow(ItemIndex, Item, LstItm)

				Dim RemovedItem As ListViewItem

				For Each RemovedItem In lstviwMacEdit.SelectedItems
					Static c As Integer : c = c + 1
					If c >= 2 Then '//Skip removing the first item
						RemovedItem.Remove()
					End If
				Next

			End If

		End If

		Row = Row.Substring(WhereThesItem + 1)

		Dim EndOfRow = True

		Do While EndOfRow

			Dim WhereTheSeparator = Row.IndexOf("*")

			Dim SubItem = Row.Substring(0, WhereTheSeparator)

			Row = Row.Substring(WhereTheSeparator + 1)

			If Not Replace Then
				LstItm.SubItems.Add(SubItem)
			Else
				lstviwMacEdit.Items(ItemIndex).SubItems.Add(SubItem)
			End If

			If Row = "" Then

				EndOfRow = False

			End If

		Loop

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub PrepareTheRow(ByVal ItemIndex As Integer, ByVal Item As String, ByRef LstItm As ListViewItem)
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'==============================================================================
		' Purpose      :  Prepare The Row to receive the copied value
		' Inputs       :  ItemIndex to be replaced, Item.text to be inserted instead and LstItm to be replaced
		' Returns      :  -
		' Calls        :  -
		' Use in       :  AddTheCopiedRowToTheList
		' Date/Time    :  Thursday, August 20, 2009 2:25:00 PM
		'==============================================================================

		lstviwMacEdit.Items(ItemIndex).SubItems.Clear()
		lstviwMacEdit.Items(ItemIndex).Text = Item
		LstItm = lstviwMacEdit.Items(ItemIndex)
		LstItm.ForeColor = GetItemColor(LstItm)

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub lstviwMacEdit_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lstviwMacEdit.MouseUp
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		If Clipboard.ContainsData("CMFormat") Then
			lstviewStripPaste.Enabled = True
		Else
			lstviewStripPaste.Enabled = False
		End If

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub RunSelectionToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunSelectionToolStripMenuItem.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		If lstviwMacEdit.Items.Count < 1 Then Exit Sub

		PlaybackTheList(True)

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub RunTheMacroToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunTheMacroToolStripMenuItem.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		If lstviwMacEdit.Items.Count < 1 Then Exit Sub

		PlaybackTheList(False)

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub SelectAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		SelectAllListViewItems(True)
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub DeselectAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeselectAllToolStripMenuItem.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		SelectAllListViewItems(False)
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub SelectAllListViewItems(ByVal SelectItems As Boolean)
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'SelectItems = True ==>  lstItem.Selected = true
		Dim lstItem As ListViewItem
		For Each lstItem In lstviwMacEdit.Items
			lstItem.Selected = SelectItems
		Next

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub CutToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutToolStripButton.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		CopySelectedRowItemsInListView(lstviwMacEdit, "*")
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub CopyToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripButton.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		CopySelectedRowItemsInListView(lstviwMacEdit, "*")
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub PasteToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripButton.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		PasteAfter()
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub TmrGenLoop_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TmrGenLoop.Tick
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		If Me.WindowState = FormWindowState.Minimized Then
			HideMe()
		End If

		'//el Timer da 3ashan elmaham elmo3tada elly t7tag 2le tkraar
		Try
			If Clipboard.ContainsData("CMFormat") Then
				PasteToolStripButton.Enabled = True
			Else
				PasteToolStripButton.Enabled = False
			End If
		Catch ex As Exception
			Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error getting data from clipboard")
			PasteToolStripButton.Enabled = False
		End Try

		If lstviwMacEdit.Items.Count = 0 Then
			ToolStripRun.Enabled = False
			PlayBackSystemTrayMenu.Enabled = False
			lstviewStripRun.Enabled = False
			ToolStripCompile.Enabled = False
			ToolStripSveMac.Enabled = False
		Else
			ToolStripRun.Enabled = True
			PlayBackSystemTrayMenu.Enabled = True
			lstviewStripRun.Enabled = True
			ToolStripSveMac.Enabled = True
			ToolStripCompile.Enabled = True
		End If

		Dim PMACFilePath As String = IntProcCom.PasteDataFromClipboard("Nanno&Ahmed")

		If PMACFilePath <> "" Then
			Clipboard.Clear()
			Me.WindowState = FormWindowState.Minimized
			LoadListViewFromCSV(PMACFilePath, lstviwMacEdit, False)
			PlaybackTheList(False)
		End If

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub ToolStripDelItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDelItems.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		DeleteListItems()
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub lstviewStripRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstviewStripRun.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		If lstviwMacEdit.Items.Count < 1 Then Exit Sub

		If lstviwMacEdit.SelectedItems.Count <= 1 Then
			PlaybackTheList(False)
		ElseIf lstviwMacEdit.SelectedItems.Count > 1 Then
			PlaybackTheList(True)
		End If

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub ToolStripCompile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripCompile.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * 
		CompileForm.ShowDialog()
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Friend Function PlaybackTheExtractedMacFile(Optional ByVal UnitTestingPath As String = "", Optional ByVal ExtractAndLoadListOnly As String = "", Optional ByVal Merge As Boolean = False) As Result
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		If Not Merge Then '//Merge parameter to decide if i will load a new file to the list or note with
			'//with an old file or not
			lstviwMacEdit.Items.Clear()
		End If

		' Read the macro file embedded in the executable macro.
		Dim macroFileObj As Object = xResources.ReadEmbeddedResourceItemAtRuntime("exeMacroResources", "MacroFile", Assembly.LoadFrom(ExtractAndLoadListOnly))
		Dim macFileBytes As Byte() = CType(macroFileObj, Byte())
		Try
			Dim BinFormatter As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
			Dim ms As New System.IO.MemoryStream(macFileBytes)
			lstviwMacEdit.Items.AddRange(BinFormatter.Deserialize(ms).ToArray(GetType(ListViewItem)))
			ms.Dispose()
		Catch ex As Exception
			MessageBox.Show("Cannot load this executable macro because the file has been damaged.")
			Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Cannot load this executable macro because the file has been damaged.")
		End Try


		'ExtractAndLoadListOnly parameter contains a exe file that i want to extract and load only
		'its hidden macro to edit this one
		Dim MethodResult As Result = Nothing

		'Extract and load only ya man
		If ExtractAndLoadListOnly <> "" Then
			MethodResult.Successed = True : MethodResult.Message = "extract and load only"
			Return MethodResult
		End If

		If PlaybackTheList(False) Then
			MethodResult.Successed = True
		Else
			MethodResult.Successed = False
		End If

		Return MethodResult

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Friend Function SaveListViewToCSV(ByVal listview As ListView, ByVal SavedFilePath As String, Optional ByVal SelectedItemsOnly As Boolean = False, Optional ByVal IsExe As Boolean = False, Optional ByVal isLocked As Boolean = False) As Result
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'SerializeTheList is a switch to enable/disable serialization for the listview
		If SerializeTheList And Not IsExe Then
			'ثلاثة شروط
			'1. المفتاح
			'2. ليس هذا ملفاً تنفيذياً
			'3. لسنا هنا بصدد حفظ الماكرو كملف تنفيذي
			SerializeListView(SavedFilePath, lstviwMacEdit)
			GoTo Finisheze
		End If

		Dim MyResult As Result = Nothing
		Dim FileName As String = SavedFilePath

		Dim currentLine As String = String.Empty
		Dim csvFileContents As New System.Text.StringBuilder

		If listview.Items.Count = 0 Then
			MyResult.Successed = False
			MyResult.Message = ""
			Return MyResult
		End If
		Dim ListItems As Object

		If SelectedItemsOnly Then
			ListItems = New ListView.SelectedListViewItemCollection(lstviwMacEdit)
		Else
			ListItems = New ListView.ListViewItemCollection(lstviwMacEdit)
		End If

		For Each item As ListViewItem In ListItems

			For Each subitem As ListViewItem.ListViewSubItem In item.SubItems
				currentLine &= (String.Format("""{0}"",", subitem.Text))
			Next

			'Remove trailing comma
			Dim NewLine As String = currentLine.Substring(0, currentLine.Length - 1)

			csvFileContents.AppendLine(NewLine)
			currentLine = String.Empty

		Next

		'Create the file.
		Dim sw As New System.IO.StreamWriter(FileName)
		Try
			sw.WriteLine(csvFileContents.ToString)
			sw.Flush()
			sw.Dispose()
		Catch ex As Exception
			Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error writing the temp CSV mac file")
			MyResult.Successed = False
			Return MyResult
		End Try

Finisheze:

		MyResult.Successed = True
		MyResult.Message = ""
		Return MyResult

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Friend Function LoadListViewFromCSV(ByVal CSVFilePath As String, ByVal listview As ListView, Optional ByVal Merge As Boolean = False) As Result
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim MyResult As Result = Nothing

		If Not Merge Then '//Merge parameter to decide if i will load a new file to the list or note with
			'//with an old file or not
			lstviwMacEdit.Items.Clear()
		End If

		Try

			Dim IsTheOpenedFilesIsExe As Boolean = False

			If GetCallingMethod(2) = "PlaybackTheExtractedMacFile" Then
				IsTheOpenedFilesIsExe = True
			End If

			'SerializeTheList is a switch to enable/disable serialization for the listview
			If DeSerializeTheList And Not IsTheOpenedFilesIsExe Then
				'فقط في حالة توافر ثلاثة شروط :
				'الأول : المفتاح
				'DeSerializeTheList
				'مفتوح
				'ثانياً : ليس ملفاً تنفيذاً يُجري تنفيذ الماكرو خاصته
				'ثالثا : ليس ملفاً تنفيذيا بصدد التشغيل الآن .. وسوف نعرف ذلك
				'من الوظيفة التي استدعت الوظيفة الحالية
				'لو كانت
				'PlaybackTheExtractedMacFile
				'فأذا هذا ملف تنفيذي
				DeserializeListView(CSVFilePath, listview)
				GoTo Finisheze
			End If

			Status_LoadingLabel.Text = "Loading..." : Application.DoEvents()

			'Read the file
			Dim sw As New System.IO.StreamReader(CSVFilePath)

			Do While sw.Peek > -1

				Debug.Print(frmProgress.ProgBar.Value)
				If frmProgress.ProgBar.Value = frmProgress.ProgBar.Maximum Then
					frmProgress.ProgBar.Maximum = frmProgress.ProgBar.Maximum * 2
				End If
				frmProgress.ProgBar.Value += 1

				Dim NewLine As String = sw.ReadLine()
				Dim Values() As String = Split(NewLine, ",")
				Dim NewItem As ListViewItem = Nothing

				Dim index As Integer = 0
				For Each Str As String In Values

					index = index + 1

					'This line causes the execuable macros to take long time in loading in the list
					'listview.BeginUpdate()
					If index = 1 Then
						Dim NewItemText As String = Values(0).Replace("""", "")
						If NewItemText = "Background Window" And ForeGroundWindowSwitch = False Then
							GoTo NextItem
						End If
						NewItem = listview.Items.Add(Values(0).Replace("""", ""))
						NewItem.ForeColor = GetItemColor(NewItem)
					Else
						Dim subItem As ListViewItem.ListViewSubItem
						subItem = NewItem.SubItems.Add(Str.Replace("""", "")) '//delete the double quotes

						'You may delete this if block if you want to reuse the method
						If subItem.Text.Contains("Y") Then
							subItem.Tag = subItem.Text.Replace("Y = ", "") '//extract the Y and assign to Tag, use later in playback
							'subItem.ForeColor = Color.Blue
						ElseIf subItem.Text.Contains("X") Then
							subItem.Tag = subItem.Text.Replace("X = ", "") '//extract the X and assign to Tag, use later in playback
							'subItem.ForeColor = Color.Blue
						ElseIf subItem.Text.Contains("") Then
						End If
					End If
					'listview.EndUpdate()
				Next
NextItem:
			Loop
			Status_LoadingLabel.Text = ""
			sw.Dispose()
Finisheze:
			TagMouseSubItems() '//add mouse x & y to mouse subitem tag otherwise the mosue won't move

			'منح التركيز لإول عنصر في القائمة , حتي يتم إضافة أي إضافة من إضافات التحرير بعده مباشرة
			'في حالة ما إذا لم يختار المستخدم عنصر ليتم الإدراج بعده
			lstviwMacEdit.Select()
			lstviwMacEdit.Items(0).Selected = True

			MyResult.Successed = True

			Me.Text = "Perfect Macro Recorder" & " - " & Path.GetFileName(CSVFilePath)

			HideScriptListHorzScrollBars()

			Return MyResult

		Catch ex As Exception
			Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error loading the temp CSV macro file.")
			MyResult.Successed = False
			Return MyResult
		End Try

	End Function

	Private Sub TagMouseSubItems()
		For Each lstItem As ListViewItem In lstviwMacEdit.Items
			For Each subItemos As ListViewItem.ListViewSubItem In lstItem.SubItems
				With subItemos
					'You may delete this if block if you want to reuse the method
					If .Text.Contains("Y") Then
						.Tag = .Text.Replace("Y = ", "") '//extract the Y and assign to Tag, use later in playback
						'subItem.ForeColor = Color.Blue
					ElseIf .Text.Contains("X") Then
						.Tag = .Text.Replace("X = ", "") '//extract the X and assign to Tag, use later in playback
						'subItem.ForeColor = Color.Blue
					End If
				End With
			Next
		Next
	End Sub

	Private Sub CompileToExeFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CompileToExeFileToolStripMenuItem.Click
		CompileForm.ShowDialog()
	End Sub

	Private Sub ToolStripAddDelay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripAddDelay.Click
		AddWaitEvent()
	End Sub

	Private Sub AddWaitEvent()

		Dim DiaRes As DialogResult = FrmDelay.ShowDialog()
		If DiaRes = Windows.Forms.DialogResult.Cancel Then Exit Sub

		Dim ItemIndex As Integer
		Dim lstItem As ListViewItem = Nothing
		Dim DelayValue As Long = FrmDelay.txtDelayValue.Text

		If lstviwMacEdit.Items.Count > 0 Then 'لو كان فيه ماكرو معروض في القائمة

			Select Case lstviwMacEdit.SelectedIndices.Count
				Case 1
					ItemIndex = lstviwMacEdit.SelectedIndices.Item(0)
				Case Is > 1
					ItemIndex = lstviwMacEdit.SelectedIndices.Item(lstviwMacEdit.SelectedIndices.Count - 1)
			End Select

			lstItem = lstviwMacEdit.Items.Insert(ItemIndex + 1, "Wait")
			lstItem.SubItems.Add(DelayValue)
			lstItem.SubItems.Add(0)	'//This value for "wait" for window feature

		ElseIf lstviwMacEdit.Items.Count = 0 Then 'لو مافيش ماكرو معروض

			lstItem = lstviwMacEdit.Items.Add("Wait")
			lstItem.SubItems.Add(DelayValue)
			lstItem.SubItems.Add(0)	'//This value for "wait" for window feature

		End If

		lstItem.ForeColor = GetItemColor(lstItem)

		'العنصر المضاف لازم يتعمل لأمه هاي لايت مش العنصر اللي قبله
		'علي الأقل عشان اليوزر يعرف يميز ميتين أمه في الليست
		FocusThisItem(lstItem)

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub addBackGroundItem()
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		FrmBackGround.Tag = "" '//just to ensure that we provoke the adding lines not the editing ones in frmbackground 
		FrmBackGround.ShowDialog()
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Friend Function GetItemColor(ByRef NewItem As ListViewItem) As Color
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Select Case NewItem.Text
			Case "ForeGroundWindow", "Background Window"
				GetItemColor = Color.Gray
			Case "MouseWheel"
				GetItemColor = Color.DeepPink
			Case "MouseMove"
				GetItemColor = Color.Blue
			Case "Wait"
				GetItemColor = Color.Red
			Case "LeftMouseUp"
				GetItemColor = Color.DarkBlue
			Case "MiddleMouseUp"
				GetItemColor = Color.Crimson
			Case "RightMouseUp"
				GetItemColor = Color.DodgerBlue
			Case "LeftMouseDown"
				GetItemColor = Color.Brown
			Case "MiddleMouseDown"
				GetItemColor = Color.Firebrick
			Case "RightMouseDown"
				GetItemColor = Color.DarkSalmon
			Case "KeyUp"
				GetItemColor = Color.Purple
			Case "KeyDown"
				GetItemColor = Color.OliveDrab
		End Select

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Friend Sub FocusThisItem(ByRef lstItem As ListViewItem)
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		'العنصر المضاف لازم يتعمل لأمه هاي لايت مش العنصر اللي قبله
		'علي الأقل عشان اليوزر يعرف يميز ميتين أمه في الليست
		lstviwMacEdit.SelectedItems.Clear()
		lstviwMacEdit.Select()
		lstItem.Selected = True
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub editWaitItem(ByVal Index As Integer)
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'ملحوظة هامة نوعاً ما >>>
		'بعد أزمة العاشر من اكتوبر سنة 2009 الشهيرة التي أجبرتني علي تغيير
		'منطق الإعتماد علي التسلسل في حفظ ملف البرنامج .. أصبح قيمة الإنتظار
		'في أي عنصر من القائمة غير ذا جدوي أو قيمة في البلاي باك .. عدا عنصر
		'الإنتظار نفسه بالطبع .. ومن ثم فلا فائدة حقيقة ترجي من تغيير قيمة
		'للعنصر التالي لعنصر الإنتظار هذه إلا فقط في إتزان المعادلة أمام عينيك ..
		'بالطبع أنت تعرف أن كل عنصر يحتوي علي قيمة إنتظار يسبقه عنصر إنتظار
		'بالفعل يحتوي علي نفس القيمة ولكن عند البلاي باك لايعتد إلا بقيمة عنصر الأنتظار
		'نفسه وأبضا يتم تنفيذ كل عنصر بمعزل عن اخوته
		'شكرا

		'15/10/2009
		'نفسي أرجع اكتب تاني زي الأول والله في مدونتي اهئ اهئ اهئ

		With FrmDelay
			.txtDelayValue.Text = lstviwMacEdit.Items(Index).SubItems(1).Text
			.txtDelayValue.HideSelection = False '// to make the selections appears to the user
			.txtDelayValue.Focus() '// to allow the user to edit txtDefMacName directly without focusing it first 
			.txtDelayValue.SelectAll()
		End With

		Dim DiaRes As DialogResult = FrmDelay.ShowDialog()
		If DiaRes = Windows.Forms.DialogResult.Cancel Then Exit Sub

		Dim lstItem As ListViewItem

		For Each lstItem In lstviwMacEdit.SelectedItems
			lstviwMacEdit.Items(lstItem.Index).SubItems(1).Text _
			= FrmDelay.txtDelayValue.Text
		Next

		If lstviwMacEdit.SelectedItems.Count > 1 Then lstviwMacEdit.SelectedItems.Clear()

	End Sub

	Private Sub lstviwMacEdit_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lstviwMacEdit.MouseDoubleClick

		Dim ClickedListViewItem As ListViewHitTestInfo
		ClickedListViewItem = lstviwMacEdit.HitTest(e.X, e.Y)

		selectItemToEdit(ClickedListViewItem.Item)

	End Sub

	Private Sub selectItemToEdit(ByVal item As ListViewItem)

		Select Case item.Text
			Case "Wait"
				'ملحوظة هامة نوعاً ما >>>
				'بعد أزمة العاشر من اكتوبر سنة 2009 الشهيرة التي أجبرتني علي تغيير
				'منطق الإعتماد علي التسلسل في حفظ ملف البرنامج .. أصبح قيمة الإنتظار
				'في أي عنصر من القائمة غير ذا جدوي أو قيمة في البلاي باك .. عدا عنصر
				'الإنتظار نفسه بالطبع .. ومن ثم فلا فائدة حقيقة ترجي من تغيير قيمة
				'للعنصر التالي لعنصر الإنتظار هذه إلا فقط في إتزان المعادلة أمام عينيك ..
				'بالطبع أنت تعرف أن كل عنصر يحتوي علي قيمة إنتظار يسبقه عنصر إنتظار
				'بالفعل يحتوي علي نفس القيمة ولكن عند البلاي باك لايعتد إلا بقيمة عنصر الأنتظار
				'نفسه وأبضا يتم تنفيذ كل عنصر بمعزل عن اخوته
				'شكرا

				'15/10/2009
				'نفسي أرجع اكتب تاني زي الأول والله في مدونتي اهئ اهئ اهئ
				editWaitItem(item.Index)
			Case "Background Window"
				editBackGroundWin(item)
			Case "LeftMouseUp", "MiddleMouseUp", "RightMouseUp", "MouseWheel", "MouseMove", _
			  "LeftMouseDown", "MiddleMouseDown", "RightMouseDown"
				editMouseItem(item.Index)
			Case "KeyUp", "KeyDown"
				editKeyStrokeItem(item)
		End Select

	End Sub

	Private Sub editKeyStrokeItem(ByRef item As ListViewItem)
		With FrmKeyStroke
			.Tag = "edit" '//عشان اعرف أميز نوع الإدخال , هل هو إيديت والا إضافة جديدة
			.ComKeyAction.SelectedItem = item.Text
			.ComKeys.SelectedItem = item.SubItems(1).Text
			.ShowDialog()
		End With
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub editBackGroundWin(ByRef item As ListViewItem)
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		With FrmBackGround
			.Tag = "edit" '//عشان اعرف أميز نوع الإدخال , هل هو إيديت والا إضافة جديدة
			.TxtTitle.Text = item.SubItems(1).Text
			.mstLeft.Text = item.SubItems(2).Text
			.mstTop.Text = item.SubItems(3).Text
			.mstHeight.Text = item.SubItems(5).Text - item.SubItems(3).Text
			.mstWidth.Text = item.SubItems(4).Text - item.SubItems(2).Text
			.TxtWait.Text = item.SubItems(6).Text
			.ChkWaitForWindow.Checked = item.SubItems(7).Text
			'28/09/2010
			'.LabelProcessName.Text = item.SubItems(8).Text 'FoxSmr 
			'.LabelProcessPath.Text = item.SubItems(9).Text 'FoxSmr 
			.ShowDialog()
		End With
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub editMouseItem(ByVal Index As Integer)
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim selectedIndex As Integer

		frmMouse.lblDesc.Text = "Use this window to edit a mouse action from the list. Select the mouse action type and the coordinates of the cursor associated with the action."

		Select Case lstviwMacEdit.Items(Index).SubItems(0).Text
			Case "LeftMouseDown"
				selectedIndex = 0
			Case "LeftMouseUp"
				selectedIndex = 1
			Case "MiddleMouseDown"
				selectedIndex = 2
			Case "MiddleMouseUp"
				selectedIndex = 3
			Case "MouseMove"
				selectedIndex = 4
			Case "MouseWheel"
				selectedIndex = 5
			Case "RightMouseDown"
				selectedIndex = 6
			Case "RightMouseUp"
				selectedIndex = 7
		End Select

		With frmMouse
			.comMouseEvents.SelectedIndex = selectedIndex
			If selectedIndex = 5 Then 'لو كان عجلة الماوس
				.mstWheelValue.Text = lstviwMacEdit.Items(Index).SubItems(1).Text
				.mstX.Text = lstviwMacEdit.Items(Index).SubItems(2).Tag
				.mstY.Text = lstviwMacEdit.Items(Index).SubItems(3).Tag
			Else ' لو كان أي حاجة غير كدة
				.mstX.Text = lstviwMacEdit.Items(Index).SubItems(1).Tag
				.mstY.Text = lstviwMacEdit.Items(Index).SubItems(2).Tag
			End If
			.CaptureStarted = False
			.ShowDialog()
		End With

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub ToolStripEditCommand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripEditCommand.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim selectedListItem As ListViewItem = lstviwMacEdit.SelectedItems(0)
		Dim selectedListItemIndex = selectedListItem.Index

		Select Case lstviwMacEdit.SelectedItems(0).Text
			Case "Wait"
				editWaitItem(selectedListItemIndex)
			Case "LeftMouseUp", "MiddleMouseUp", "RightMouseUp", "MouseWheel", "MouseMove", _
			 "LeftMouseDown", "MiddleMouseDown", "RightMouseDown"
				editMouseItem(selectedListItemIndex)
			Case "Background Window"
				editBackGroundWin(selectedListItem)
			Case "KeyUp", "KeyDown"
				editKeyStrokeItem(selectedListItem)
		End Select

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub DelayToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DelayToolStripMenuItem.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		AddWaitEvent()
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub MouseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MouseToolStripMenuItem.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		ToolStripAddMouseEve_Click(sender, e)
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub KeyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KeyToolStripMenuItem.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		FrmKeyStroke.ShowDialog()
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub ToolStripAddKeys_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripAddKeys.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *      
		FrmKeyStroke.ShowDialog()
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub ToolStripAddMouseEve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripAddMouseEve.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		With frmMouse
			.comMouseEvents.SelectedIndex = 0
			.mstWheelValue.HideSelection = False '// to make the selections appears to the user
			.mstWheelValue.Focus() '// to allow the user to edit txtRepeatTimes directly without focusing it first 
			.mstWheelValue.SelectAll()
			.lblDesc.Text = "Use this window to add a mouse action to the list. Select the mouse action type and the coordinates of the cursor associated with the action."
			.ShowDialog()
		End With

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub ToolStripAddMore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripAddMore.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		MsgBox("FocusThisItem")
		MsgBox("lstItem.ForeColor = MainForm.GetItemColor(lstItem)")
		frmMoreItems.ShowDialog()
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub StopToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StopToolStripMenuItem.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		StopMacro()
		ShowMe()
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Private Sub ToolStripBackGround_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripBackGround.Click
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		addBackGroundItem()
	End Sub

	Private Sub BackgroundWindowTitleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackgroundWindowTitleToolStripMenuItem.Click
		addBackGroundItem()
	End Sub

	Private Sub WebPageToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WebPageToolStripMenuItem1.Click
		Dim webAddress As String = "http://www.perfectiontools.com/perfect_macro_recorder.html"
		Process.Start(webAddress)
	End Sub

	Private Sub toolstripHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles toolstripHelp.Click

		'Dim helpFilePath As String = Application.StartupPath & "\" & "Help.chm"
		Try
			Process.Start("http://www.perfect-macro-recorder.com/online-help/online-help.html")
		Catch ex As Exception
			MsgBox(ex.Message)
		End Try

	End Sub

	Private Sub lstviwMacEdit_Layout(ByVal sender As Object, ByVal e As System.Windows.Forms.LayoutEventArgs) Handles lstviwMacEdit.Layout
		'//to hide the horz bar of the listview when the user scroll down the list
		HideScriptListHorzScrollBars()
	End Sub

	Private Function collectAppsPaths() As Hashtable  'FoxSmr
		'Get App Names and paths for every  "Background Window" item
		Dim processesTable As Hashtable = New Hashtable
		For Each lstItem In lstviwMacEdit.Items
			If lstItem.Text = "Background Window" Then
				If Not processesTable.ContainsKey(lstItem.SubItems(8).Text) Then
					Dim processName As String = lstItem.SubItems(8).Text
					Dim processPath As String = lstItem.SubItems(9).Text
					'Debug.Print(processName & " " & processPath)
					processesTable.Add(processName, processPath)
				End If
			End If
		Next
		Return processesTable
	End Function

	Private Function isProcessRunning(ByVal processName As String) As Process()	'FoxSmr
		Return Process.GetProcessesByName(processName)
	End Function

	Private Sub ToolStripButtonRegister_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButtonRegister.Click
		IntelliProtectorService.IntelliProtector.ShowRegistrationWindow()
		If Not GetLicStatus() Then
			is_x2342 = "False"
		Else
			is_x2342 = "True"
		End If
	End Sub


End Class