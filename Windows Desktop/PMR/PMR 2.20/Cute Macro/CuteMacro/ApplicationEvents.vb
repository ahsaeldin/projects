Imports System.IO
Imports System.Reflection

Namespace My

	' The following events are available for MyApplication:
	' 
	' Startup: Raised when the application starts, before the startup form is created.
	' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
	' UnhandledException: Raised if the application encounters an unhandled exception.
	' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
	' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
	Partial Friend Class MyApplication

	End Class

	' The following events are available for MyApplication:
	' 
	' Startup: Raised when the application starts, before the startup form is created.
	' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
	' UnhandledException: Raised if the application encounters an unhandled exception.
	' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
	' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.

	'<System.Reflection.ObfuscationAttribute(Feature:="renaming", Exclude:=True, ApplyToMembers:=True)>
	Partial Friend Class MyApplication
		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		'<System.Reflection.ObfuscationAttribute(Feature:="renaming", Exclude:=True, ApplyToMembers:=True)>
		Private Sub MyApplication_Startup(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) Handles Me.Startup
			'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * 

			IntelliProtectorService.IntelliProtector.Init()

			'Search the project for IntelliProtectorService
			If Not GetLicStatus() Then
				is_x2342 = "False"
			Else
				is_x2342 = "True"
			End If
			'فقط ترجم لو كانت نسخة تجريبية
			'If is_x2342 = "False" Then
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'It is so simple protection model.
			'first i change if first run my checking My.Settings.eo
			'if it  is == "PerfectMacroRecorder" then it is first run
			'then i get the end date and encrypte it using jeff atwood lib
			'then save it to My.Settings.eo...
			'Next time when run we check for My.Settings.eo ...
			'then get the string and dycrtpy it then compare dates to 
			'get remaining days and if off then we change  My.Settings.companyname
			'to "Perfection Tools Software" (heheheh) to check below so if he 
			'back to the past and change the history to back.it won't work

			'pure value of My.Settings.eo is "PerfectMacroRecorder" that means this is the first run
			'My.Settings.eo ' "PerfectMacroRecorder"

			'نوع من التمويه .. والفائدة هنا انني سأغير قيمتها بالفعل
			'بالأسفل عندما يحين الميعاد وبهذا حتي لو غير التاريخ فلن
			'يجدي ذلك نفعاً

			'If My.Settings.companyname <> "Perfection Tools" Then
			'MsgBox("Trial period expired. Thank you!")
			'End
			'End If

			' lq  ................ "Windows Executable"

			'If My.Settings.eo = "PerfectMacroRecorder" Then	' first time to run

			'	Dim sd As Date = DateTime.Today	'start date
			'	Dim dateInter As DateInterval = DateInterval.Day
			'	Dim ed As Long = DateAdd(dateInter, 15, sd).ToBinary

			'	'encrypte end date and save to My.Settings.eo
			'	Dim sym As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
			'	Dim key As New Encryption.Data("aSdKiUjYhb7jh6h9jh5hIO8JY6Hij9Jyh6hj9jY5FJHoiPOl.LjhnFeDWEcbK97864@#$%!^%")
			'	Dim encryptedData As Encryption.Data
			'	encryptedData = sym.Encrypt(New Encryption.Data(ed), key)
			'	Dim base64EncryptedString As String = encryptedData.ToBase64
			'	My.Settings.eo = base64EncryptedString

			'	sd = Nothing
			'	dateInter = Nothing
			'	ed = Nothing
			'	sym = Nothing
			'	key = Nothing
			'	base64EncryptedString = Nothing

			'Else

			'	'decrypt data
			'	Dim sym As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
			'	Dim key As New Encryption.Data("aSdKiUjYhb7jh6h9jh5hIO8JY6Hij9Jyh6hj9jY5FJHoiPOl.LjhnFeDWEcbK97864@#$%!^%")
			'	Dim encryptedData As New Encryption.Data
			'	encryptedData.Base64 = My.Settings.eo
			'	Dim decryptedData As Encryption.Data
			'	decryptedData = sym.Decrypt(encryptedData, key)

			'	Dim ed As Date = DateTime.FromBinary(CLng(decryptedData.ToString))
			'	Dim delta As Long = DateDiff(DateInterval.Day, DateTime.Today, ed)

			'	key = Nothing
			'	decryptedData = Nothing

			'	If delta > 15 Or delta <= 0 Then

			'		My.Settings.companyname = "Perfection Tools Software" 'hehehehehehehehehehe .. simple change 
			'		My.Settings.Save()
			'		MsgBox("Trial period expired. Thank you!")
			'		End

			'	End If
			'End If
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'End If

			'Kathlen and ElGog reports me that PMR V2.0 doesn't work under vista ,,, after debugging i discovered that
			'Vista SP1 generates this exception "System.TypeInitializationException was unhandled" at calling 
			'of the following methods here in this method
			'IsFileContainsSrting_3shanVista()
			'IsAnotherInstanceRunning_3shanVista()
			'CopyDataToClipboard_3shanVista()
			'and you must know that these methods already exists in Intercom and Util but i just copied them here and rename them to avoid this problem

			'Dim IsDllsInstalled As Boolean = CheckGACForMyDllex()
			'If Not IsDllsInstalled Then
			'    Dim res1 As Boolean = False
			'    Dim res2 As Boolean = False
			'    res1 = InstallDLLs("mklextp.dll")
			'    'res2 = InstallDLLs("DeployLX_Licensing_v3")
			'    IsDllsInstalled = False
			'    IsDllsInstalled = CheckGACForMyDllex() '//recheck again
			'    If Not IsDllsInstalled Then
			'        MsgBox("Error While loading, please reinstall Perfect Macro Recorder", MsgBoxStyle.Critical, "Perfect Macro Recorder")
			'        End
			'    End If
			'End If

			'ilmerge /target:winexe /out:SelfContainedProgram.exe Perfect Macro Recorder.exe mklextp.dll Interop.IWshRuntimeLibrary.dll

			'''''''Check if an executable macro File
			Dim myMod As System.Reflection.Module = [Assembly].GetExecutingAssembly().GetModules()(0)
			'Console.WriteLine("Module Name is " + myMod.Name)
			'Console.WriteLine("Module FullyQualifiedName is " + myMod.FullyQualifiedName)
			'Console.WriteLine("Module ScopeName is " + myMod.ScopeName)

			'Dim Files As ReadOnlyCollection(Of String) = FindAStringOnAFile(FileSystem.CurDir())

			Dim CIntProcCom As New IntProcCom '//an instance of my IntProcCom to send WM_COPYDATA with the
			'//path to the other instance of me.

			Dim ArgsCount As Integer = My.Application.CommandLineArgs.Count

			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'if the user clicks a CM file while an already instance of CM is running then
			'i will send the path of the file to that instance and close me

			'//end if the user clicks Perfect Macro Recorder.exe while it's runnings already
			If IsAnotherInstanceRunning_3shanVista() And ArgsCount = 0 Then End

			'//if there's command lines
			If ArgsCount > 0 Then

				'//if the command line is a path to pmac file (the clicks a cm file to run)
				If My.Application.CommandLineArgs.Item(0).Contains(".pmac") Then

					'//if there's a running instance of me
					If IsAnotherInstanceRunning() Then

						'//make sure the path for a .pmac file
						If My.Application.CommandLineArgs.Item(0).Contains(".pmac") Then
							'Dim SendingPathMessageResult As Integer
							'//Send the path to the running instance of me and END me
							'SendingPathMessageResult = CIntProcCom.SendStringToAprocess(My.Application.CommandLineArgs.Item(0))
							CopyDataToClipboard_3shanVista("Nanno&Ahmed", My.Application.CommandLineArgs.Item(0))
							'If SendingPathMessageResult = 0 Then
							'WriteToErrorLog("SendingPathMessageResult returns 0", "", "Error : Can't send path")
							'End If
							End	'//Kill me
						End If

					End If

				End If

			End If
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			Dim Ext As String = "pmac" '//The extenstion you want to assign to your app.
			Dim AppKeyNameInReg As String = "Perfect Macro Recorder" '//The App Key Name in the Reg.
			Dim FileDescription As String = "Perfect Macro Recorder" & " File" '//File Description that appears in Properites window.

			Dim IsAssociated As Boolean = GetSetting("CuteMacro", "cm", "FR", 0)

			If Not IsAssociated Then
				Dim Result As Boolean = AssociateMyIcon(Ext, AppKeyNameInReg, FileDescription)
				SaveSetting("CuteMacro", "cm", "FR", 1)
			End If

		End Sub

		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		Public Sub CheckForExistingInstance()
			'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

			'Get number of processes of you program
			If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1 Then

				'End

			End If

		End Sub

		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		Public Function IsFileContainsSrting_3shanVista(ByVal SearchWord As String, ByVal FileName As String) As Integer
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
		Public Function IsAnotherInstanceRunning_3shanVista() As Boolean
			'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

			Dim ProcessCounts As Integer = Process.GetProcessesByName _
			 (Process.GetCurrentProcess.ProcessName).Length

			If ProcessCounts > 1 Then
				IsAnotherInstanceRunning_3shanVista = True
			End If

		End Function

		'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		Public Shared Sub CopyDataToClipboard_3shanVista(ByVal format As String, ByVal str As String)
			'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

			'الوظيفة دي ووظيفة
			'PasteDataFromClipboard
			'عشان النسخ واللصق من وإلي الكيب بورد بتنسيقات معينة
			'مما يعنني أولا علي ارسال رسائل بين حالات البرنامج المختلفة
			'لتفادي استخدام
			'WM_COPYDATA 
			'في وظيفة
			'MyApplication_Startup
			'التي تسخدم في إرسال مسار ملف السي ماك إلي النسخة التي تعمل
			'فعلياً من البرنامج

			' Creates a new data format.
			Dim CDataFormat As DataFormats.Format = DataFormats.GetFormat(format)

			' Creates a new object and store it in a DataObject using myFormat 
			' as the type of format. 
			Dim CopiedText As String = str
			Dim DataObject As New DataObject(CDataFormat.Name, CopiedText)

			'لازم تاني بارامتر يبقي ترو عشان تقدر تنقل محتويات الكليب بوورد بين برنامجين
			Clipboard.SetDataObject(DataObject, True)

		End Sub

	End Class

End Namespace