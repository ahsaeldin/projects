Option Explicit On

Imports System.IO
Imports CuteMacro.Balora
Imports System.Resources
Imports System.Collections.ObjectModel
Imports System.Runtime.Serialization.Formatters.Binary

Module utils

	Public is_x2342 As String
	' deleted in 21/11/2010 coz i will use /define to define is_x2342
	' as a global #const so i can use conditional compiliation #if 
	' statements ... you can switch it from advanced compiler options.
	'Public is_x2342 As Boolean = False 'switch for trial/unlocked version ,,, if true then unlocked
	Public Function GetCallingMethod(ByVal FrameIndex As Integer) As String
		' Frame 0 = Current Method Name
		' Frame 1 = Calling Method Name
		' Frame 2 = The Calling method of the previous one
		Dim StackTrace As New StackTrace
		GetCallingMethod = StackTrace.GetFrame(FrameIndex).GetMethod().Name
	End Function

	Friend Function GetLicStatus() As Boolean

		If IntelliProtectorService.IntelliProtector.IsSoftwareRegistered() Then
			MainForm.ToolStripButtonRegister.Visible = False
			Return True
		Else
			MainForm.ToolStripButtonRegister.Visible = True
			Return False
		End If

	End Function


	Public Function ForEachAtArray(ByVal KeyName As Keys) As Integer
		'//Sample code in how to use Array.ForEach
		'Pass key Name,Retreive key value
		Dim KeyEnumNames() As String = System.Enum.GetNames(GetType(System.Windows.Forms.Keys))
		Dim KeyEnumValues() As Keys = System.Enum.GetValues(GetType(System.Windows.Forms.Keys))
		Array.ForEach(KeyEnumValues, New Action(Of Keys)(AddressOf ProcessNames))
	End Function

	Public Sub ProcessNames(ByVal KeyName As String)
		'//Sample code in how to use Array.ForEach using ForEachAtArray
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Public Function IsChildWindow(ByVal Hwnd As IntPtr) As Integer
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		Return GetParent(Hwnd) '// if zero, the window is parent otherwise is child !
	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Friend Sub SerializeListView(ByVal FilePath As String, ByVal lstView As ListView)
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		Try
			Dim BinFormatter As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
			Dim FS As New System.IO.FileStream(FilePath, IO.FileMode.Create)
			BinFormatter.Serialize(FS, New ArrayList(lstView.Items))
			FS.Flush()
			FS.Dispose()
		Catch ex As Exception
			Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error during saving the macro file.")
		End Try
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Friend Sub DeserializeListView(ByVal FilePath As String, ByVal lstView As ListView)
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		Try
			Dim BinFormatter As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
			Dim FS As New System.IO.FileStream(FilePath, IO.FileMode.Open)
			'//MainForm.Text = "Loading..." '//a way to inform the user about the loading state while loading a macro
			MainForm.Status_LoadingLabel.Text = "Loading..."
			Application.DoEvents() '//we added that "doevents" to make sure the status bar changed before Deserializing
			lstView.Items.AddRange(BinFormatter.Deserialize(FS).ToArray(GetType(ListViewItem)))
			If lstView.Items.Count = 0 Then
				MsgBox("Nothing to load; the macro file is empty." _
				, MsgBoxStyle.Critical, "Perfect Macro Recorder")
			End If
			MainForm.Status_LoadingLabel.Text = ""
			FS.Dispose()
		Catch ex As Exception
			Util.WriteToErrorLog("The macro file is corrupted", ex.StackTrace, "Error : Restoring a the mac file.")
			MsgBox("Perfect Macro Recorder" & " could not load this macro because the file has been damaged.", MsgBoxStyle.Information, "Perfect Macro Recorder")
		End Try
	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Public Function SaveSerializedObject(ByVal FilePath As String, ByVal objArr As ArrayList) As Boolean
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Try

			Dim bFormatter As New BinaryFormatter

			Dim fStream As New FileStream(FilePath, FileMode.Create)

			bFormatter.Serialize(fStream, objArr)

			fStream.Close()

			Return True

		Catch ex As Exception

			Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error during Serializing")
			Return False

		End Try

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Public Function RestoreSerializedObject(ByVal FilePath As String, Optional ByRef IsFailed As Boolean = False) As ArrayList
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim RestoredEventsList As ArrayList = Nothing '//assign to Nothing to avoid the null reference exception warning

		Dim fStream = New FileStream(FilePath, FileMode.Open)

		Dim BinFormatter As BinaryFormatter = New BinaryFormatter

		Try

			RestoredEventsList = CType(BinFormatter.Deserialize(fStream), ArrayList)

			'//Validate The File aginst corruption
			Dim FirstItemType As System.Type = RestoredEventsList.Item(0).GetType()

			If FirstItemType.Name <> "MacroEvent" Then

				Throw New System.Exception("An exception has occurred.")

			End If

		Catch ex As Exception

			Util.WriteToErrorLog("The mac file is corrupted", ex.StackTrace, "Error : Restoring a serialized object.")

			IsFailed = True

		End Try

		fStream.Close()

		RestoreSerializedObject = RestoredEventsList

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Public Function ValidateTextBoxEntry(ByVal txt As String) As Boolean
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'//1. Check aginst the seven prohbited keys \ / : * ? " < > | 
		Dim Result As Boolean

		Result = txt.Contains("\")
		If Result Then ValidateTextBoxEntry = True
		Result = txt.Contains("/")
		If Result Then ValidateTextBoxEntry = True
		Result = txt.Contains(":")
		If Result Then ValidateTextBoxEntry = True
		Result = txt.Contains("*")
		If Result Then ValidateTextBoxEntry = True
		Result = txt.Contains("?")
		If Result Then ValidateTextBoxEntry = True
		Result = txt.Contains("""")
		If Result Then ValidateTextBoxEntry = True
		Result = txt.Contains("<")
		If Result Then ValidateTextBoxEntry = True
		Result = txt.Contains(">")
		If Result Then ValidateTextBoxEntry = True
		Result = txt.Contains("|") '//also you can write itas Keys.Oem5 "Keys.Oem5 = |"
		If Result Then ValidateTextBoxEntry = True

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Public Function AssociateMyIcon(ByVal Ext As String, ByVal AppKeyNameInReg As String, ByVal FileDescription As String) As Boolean
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'Ext : The extension you want to assign to your app.
		'AppKeyNameInReg : The App Key Name in the Reg.
		'FileDescription : File Description that appears in properties window.

		Dim AppPath As String = System.Windows.Forms.Application.ExecutablePath

		Try

			My.Computer.Registry.ClassesRoot.CreateSubKey("." & Ext).SetValue _
			("", AppKeyNameInReg, Microsoft.Win32.RegistryValueKind.String)

			My.Computer.Registry.ClassesRoot.CreateSubKey(AppKeyNameInReg).SetValue _
			("", FileDescription, Microsoft.Win32.RegistryValueKind.String)

			My.Computer.Registry.ClassesRoot.CreateSubKey(AppKeyNameInReg & "\DefaultIcon").SetValue _
			("", AppPath & ",0", Microsoft.Win32.RegistryValueKind.String)

			My.Computer.Registry.ClassesRoot.CreateSubKey(AppKeyNameInReg & "\shell\open\command").SetValue _
			("", AppPath & " ""%l"" ", Microsoft.Win32.RegistryValueKind.String)

			My.Computer.Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\." & Ext).SetValue _
			("Application", AppPath, Microsoft.Win32.RegistryValueKind.String)

			Return True

		Catch ex As Exception

			Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error writing associate the icon")

			Return False

		End Try

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Public Sub CheckInstanceOfApp()
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'//not used yet but useful

		Dim appProc() As Process

		Dim strModName, strProcName As String
		strModName = Process.GetCurrentProcess.MainModule.ModuleName
		strProcName = System.IO.Path.GetFileNameWithoutExtension(strModName)

		appProc = Process.GetProcessesByName(strProcName)
		If appProc.Length > 1 Then
			MessageBox.Show("There is an instance of this application running.")
		Else
			MessageBox.Show("There are no other instances running.")
		End If

	End Sub

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Public Function IsAnotherInstanceRunning() As Boolean
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim ProcessCounts As Integer = Process.GetProcessesByName _
		 (Process.GetCurrentProcess.ProcessName).Length

		If ProcessCounts > 1 Then
			IsAnotherInstanceRunning = True
		End If

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Public Function FindAStringOnAFile(ByVal FilePath As String) As ReadOnlyCollection(Of String)
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'//Search a file for a string
		Dim Files As ReadOnlyCollection(Of String)
		Files = My.Computer.FileSystem.FindInFiles(FilePath, "cpPTECWMextPMRSMakASTCMRAHI#$#$%", True, FileIO.SearchOption.SearchAllSubDirectories, "*.exe")

		FindAStringOnAFile = Files

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Public Function IsFileContainsSrting(ByVal SearchWord As String, ByVal FileName As String) As Integer
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		'//Also Search a file for a string
		Dim FileContents As String = String.Empty
		Dim FileStreamReader As StreamReader = Nothing
		Dim IsWordFound As Boolean = False

		Try
			FileStreamReader = New StreamReader(FileName)
			FileContents = FileStreamReader.ReadToEnd()
		Catch ex As Exception
			IsWordFound = False
		Finally
			If Not (FileStreamReader Is Nothing) Then
				FileStreamReader.Close()
			End If
		End Try

		Dim index As Integer = FileContents.IndexOf(SearchWord)
		If index > 0 Then
			'IndexOf returns -1 of not found and 0 if the textVal is empty
			IsWordFound = True
		End If

		Return index

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Public Function WritePMACFileToRec() As Boolean
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		Dim FileName As String
		FileName = My.Computer.FileSystem.GetTempFileName()

		Dim IsMacroSaved As Boolean
		IsMacroSaved = MainForm.SaveListView(FileName)

		' Code to Create a Resource. 
		Dim strString As Byte()
		Dim rsw As ResourceWriter

		Dim fileContents As Byte()
		fileContents = My.Computer.FileSystem.ReadAllBytes(FileName)

		' strString is the string that will be added as a resource.
		strString = fileContents

		'Creates a resource writer instance to write to MyResource.resources.
		rsw = New ResourceWriter("MyResource.resources")

		'Adds the string to the resource.
		' "MyText" is the name that the string is identified as in the resource.
		rsw.AddResource("MyText", strString)

		rsw.Close()

		Return True

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Public Function ReadMyPMAC() As Boolean
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		' Code to retrieve the information from the resource. 
		Dim myString As Byte()
		Dim rm As ResourceManager

		' Create a Resource Manager instance.
		rm = ResourceManager.CreateFileBasedResourceManager("MyResource", ".", Nothing)

		' Retrieves the string from MyResource.
		myString = rm.GetObject("MyText")

		My.Computer.FileSystem.WriteAllBytes("C:\Output.pmac", myString, False)

	End Function

	'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
	Public Function IsOnlyNumeric(ByVal text As String, ByVal txtbox As TextBox) As Boolean
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

		If Not IsNumeric(text) Then

			MsgBox("You must enter numbers only.", MsgBoxStyle.Critical, "Perfect Macro Recorder")

			txtbox.HideSelection = False '// to make the selections appears to the user
			txtbox.Focus() '// to allow the user to edit txtRepeatTimes directly without focusing it first 
			txtbox.SelectAll()

			Return False

		End If

		Return True

	End Function

	Friend Sub validateMyTxtBoxes(ByVal txtbox As TextBox)
		Dim Res As Boolean = IsOnlyNumeric(txtbox.Text, txtbox)

		If Not Res Then
			txtbox.Text = 0
			txtbox.SelectionStart = Len(txtbox.Text)
		End If

		If txtbox.Text = 0 Then	'//converts 0000000 to 0
			txtbox.Text = 0
		End If
	End Sub

	'28/09/2010 FoxSmr 
	Public Function GetProcessPathByHandle(ByVal WndHandle As IntPtr,
				ByVal WindowText As String,
				ByRef ProcessName As String,
				ByRef ProcessFilePath As String) As Process
		Dim oResult As Process = Nothing

		If WndHandle <> IntPtr.Zero Then

			Dim iProcessID As IntPtr
			Dim Result As IntPtr = GetWindowThreadProcessId(WndHandle, iProcessID)
			Dim foreWindowProcess As Process = Nothing

			Try
				foreWindowProcess = Process.GetProcessById(iProcessID.ToInt32)
				ProcessName = foreWindowProcess.ProcessName
				ProcessFilePath = foreWindowProcess.MainModule.FileName
			Catch ex As Exception
				'Make sure that we clear ProcessName & ProcessFilePath to ignore them while playbacking the macro.
				'And the error itself logged to process later.
				ProcessName = ""
				ProcessFilePath = ""
				Dim errorMessageForGettingProcessName As String = "Error getting process name and process file path for the foreground window titled " &
				"""" & WindowText & """"
				Util.WriteToErrorLog(ex.Message, ex.StackTrace, errorMessageForGettingProcessName)
			End Try

			oResult = foreWindowProcess

		End If

		Return oResult

	End Function
End Module
