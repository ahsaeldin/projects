'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /	
'CacaduGloby: Contains all global variables for the whole project.
'\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /\ /	
Imports Balora
Imports Quartz
Imports TecDAL
Imports Helpers
Imports Cacadore
Imports Cacadu.UI
Imports Quartz.Impl
Imports Cacadu.Connectors
Imports Cacadu.UI.MainForm
Imports AppLimit.NetSparkle
Imports Balora.ErrorReporter
Imports DevExpress.XtraBars.Alerter
Imports DevExpress.XtraSplashScreen
Imports Microsoft.VisualBasic.Devices
Imports Cacadu.Connectors.MainFormConnector
Imports Cacadore.Actions.MessageBoxes.ShowAlert.ShowAlertsInputs
Imports System.Threading

Module CacaduGloby
#Region "Global Variables"
    Friend CacadoreHelperObj As CacadoreHelper
    Friend WithEvents SparkleAutoUpdater As Sparkle
    Friend ErrorReportingUI As Balora.ErrorReporter.ErrorReportingForm
    Friend Property pendingHistoryEntries As New Queue(Of Cacadore.HistoryEntry)
    'AhSaElDin 20120102: disabled for now, as CacaSer is disabled.
    'Friend WithEvents NamedPipesClient As Balora.Ipc.NamedPipes.PipeClient
#End Region

#Region "Global Functions"
#Region "Preparation Functions"
    Friend Sub ShowSplashScreen()
        SplashScreenManager.ShowForm(GetType(SplashScreen1), True, True)
    End Sub

    Friend Sub CloseSplashScreen()
        SplashScreenManager.CloseForm(False)
        SplashScreen1.Dispose()
        SplashScreen1 = Nothing
    End Sub

    Friend Sub ShowWaitForm(Optional waitFormDescription As String = "Loading...")
        Const maxDescriptionLength = 60
        If waitFormDescription.Length > maxDescriptionLength Then waitFormDescription = "Loading..."
        SplashScreenManager.ShowForm(GetType(WaitForm1), True, True)
        SplashScreenManager.Default.SetWaitFormDescription(waitFormDescription)
    End Sub

    Friend Sub CloseWaitForm()
        SplashScreenManager.CloseForm(False)
    End Sub

    Friend Sub SetIconToCacaduIcon(frm As Form)
        frm.Icon = My.Resources.Cacadu_Icon
    End Sub

    ''' <summary>
    ''' Common settings of Balora like.
    ''' </summary>
    Friend Sub PrepareBaloraForCacadu()
        '++woooooooooooooowz, this is my awesome check details
        '+The Story:-
        'Simply I want Balora or Cacadore to run only via Cacadu, in other words...
        'I don't want a developer to simply import and call them. The idea here
        'is to check the equity of a md5 value of a GUID specified here by me
        'And Cacadu will be the only who knows this vlaue and pass its hash to me.
        Balora.Settings.CAG = Balora.Serializer.BinarySerializer.MD5SerializedObject("{38D4C0D0-8E45-40E6-B40A-67AAFFE264E1}")
        'Track for any event triggered by Balora.
        AddHandler Balora.Alerter.NewErrorOccur, AddressOf CacaduGloby.Alerter_NewErrorOccur
        AddHandler Balora.Alerter.WaitForm_Load, AddressOf CacaduGloby.Alerter_WaitForm_Load
        AddHandler Balora.Alerter.WaitForm_Close, AddressOf CacaduGloby.Alerter_WaitForm_Close
        
        'AhSaElDin 20120102: disabled for now, as CacaSer is disabled.
        'NamedPipesClient = New Balora.Ipc.NamedPipes.PipeClient(My.Resources.NPN, IO.Pipes.PipeOptions.Asynchronous)
        Balora.Settings.LogErrors = True
#If DEBUG Then
        Balora.Settings.ShowLogWindow = True
#End If
        Balora.Settings.LogFileName = "CacaduErrorLog.txt"
    End Sub

    ''' <summary>
    ''' Common settings of Cacadore
    ''' </summary>
    Friend Sub PrepareCacadoreForCacadu()
        '++woooooooooooooowz, this is my awesome check details
        '+The Story:-
        'Simply I want Balora or Cacadore to run only via Cacadu, in other words...
        'I don't want a developer to simply import and call them. The idea here
        'is to check the equity of a md5 value of a GUID specified here by me
        'And Cacadu will be the only who knows this vlaue and pass its hash to me.
        Cacadore.Settings.CAG = Balora.Serializer.BinarySerializer.MD5SerializedObject("{0A70981F-B7D2-4E25-8379-8EBAB7BC9A96}")
        '1.Crud & ICrud is defined and implemented in Cacadore.
        '2.Cacadore classes invoke CRUD operations using Delegates.
        '3.Delegates execuate corresponding methods in an instance of Crud
        '4.Now the instance of CRUD execute the corresponding CRUD method in
        'the shared CrudObject property of settings class.
        '5.So that's it, all you need to do from Cacadu is to set CrudObject to an ICrud
        'class the implements the CRUD functions.
        Cacadore.Settings.CrudObject = New TecDAL.TypedDataSetCrud
    End Sub

    Friend Sub InitGlobalVariables()
        CacadoreHelperObj = New CacadoreHelper
    End Sub

    Friend Sub PrepareCacadu()
        'AhSaElDin 20120102: CacaSer disabled for now. 
        'ConnectToCacaSer()
        InitGlobalVariables()
        ''''''''''''''''''''''''''''''''''
        ''To reset FR key setting.
        'My.Settings.FR = "1"
        'My.Settings.Save()
        'If First run for Cacadu...
        'Balora.Hodhod.STOW("Checking if first run for Cacadu...")
        If My.Settings.FR = "1" Then
            'Balora.Hodhod.STOW("First run for Cacadu, do first run work...")
            DoFirstRunWork()
            'It is not first run now...
            My.Settings.FR = "0"
            My.Settings.Save()
        End If
        ''''''''''''''''''''''''''''''''''
        PrepareSystrayIcon()
        Skins.RegisterSkins()
        Skins.ChangeSkin(My.Settings.SN)
        Skins.FillThemesCombo()
        PrepareSuccessiveChecks()
        PrepareQuartz()
        AddAlertFormEventsHandlers()
    End Sub

    Private Sub PrepareQuartz()
        Dim sf As ISchedulerFactory = New StdSchedulerFactory
        Cacadore.Settings.Scheduler = sf.GetScheduler()
    End Sub

    Friend Sub ShutdownQuartz()
        Dim sch As IScheduler = CType(Cacadore.Settings.Scheduler, IScheduler)
        sch.Shutdown()
        sch = Nothing
    End Sub

    Private Sub PrepareSystrayIcon()
        UI.MainForm.SystrayIcon.Icon = My.Resources.systary_icon
    End Sub

    Friend Sub DoFirstRunWork()
        'Backup Tectonic @ first run.
        FileCopy(TypedDataSetCrud.TectonicPath, TypedDataSetCrud.TectonicBackupPath)
    End Sub

    Friend Sub WriteInStatusBar(state As String, color As Color, Optional reportProgress As Integer = 0)
        With MainForm.LeftStatusLabelBarStaticItem
            .Caption = state
            .ItemAppearance.Normal.ForeColor = color
        End With
        With MainForm.ProgressBarInStatusBarRepositoryItem
            If reportProgress > 0 Then
                .Maximum = 100
                MainForm.StatusbarLeftContainer.Visibility = DevExpress.XtraBars.BarItemVisibility.Always
                MainForm.StatusbarLeftContainer.EditValue = reportProgress
            Else
                MainForm.StatusbarLeftContainer.Visibility = DevExpress.XtraBars.BarItemVisibility.Never
            End If
        End With
    End Sub

    Friend Sub CloseCacadu()
        RemoveHandler TecDAL.TypedDataSetCrud.TectonicDataSetObj.tasks.tasksRowChanged, AddressOf tasksRowChanged
        RemoveHandler MainForm.TasksListDataGridView.DataError, AddressOf TasksListDataGridView_DataError
        RemoveHandler Balora.Alerter.NewErrorOccur, AddressOf CacaduGloby.Alerter_NewErrorOccur
        RemoveHandler Balora.Alerter.WaitForm_Load, AddressOf CacaduGloby.Alerter_WaitForm_Load
        RemoveHandler Balora.Alerter.WaitForm_Close, AddressOf CacaduGloby.Alerter_WaitForm_Close
        RemoveHandler Balora.Alerter.AlertForm_Load, AddressOf AlertControllor.Alerter_AlertFormLoad
        RemoveHandler Balora.Alerter.AlertForm_MessageSent, AddressOf AlertControllor.BaloraAlerter_MessageReceived
        _typedDataSetCrud.Dispose()
        _typedDataSetCrud = Nothing
        MainForm.Close()
    End Sub
#End Region

#Region "Common Cacadu Tasks"
    Friend Sub ConnectToCacaSer()
        If Not Servnst.IsServiceInstalled(My.Resources.CSN) Then
            MsgBox(My.Resources.CSINI, MsgBoxStyle.Information)
            Exit Sub
        End If
        If Not Servnst.GetServiceStatus(My.Resources.CSN) = ServiceProcess.ServiceControllerStatus.Running Then
            MsgBox(My.Resources.SCSM, MsgBoxStyle.Information)
            Exit Sub
        End If
        'AhSaElDin 20120102: disabled for now, as CacaSer is disabled.
        'NamedPipesClient.Connect()
    End Sub

    Friend Sub CheckForUpdates(ByVal versionInfoUrl As String, ByVal appIcon As Icon)
        'Read about how to upload a new update in the top of Inno file.

        'A must read:
        'NetSparkle makes sure that updates file doesn't altered or tampered with, using one of two things:
        '   *.Include a DSA signature of the SHA-1 hash of your published update file.

        '   *.Deliver your update and your appcast through SSL.

        'I have tried the first, but I think it will be time consuming since you need to add the DSA 
        'signature after every update. Guess what, you can't update using inno files using this method
        'since you need to embed the public key (see http://netsparkle.codeplex.com/wikipage?title=How%20to%20generate%20a%20DSA%20signature). 
        'So I have choosed the second method. but how can I offer the download over SSL, the answer is via Dropbox "cool". 
        'I will have 2 links one for Cacadu.exe and the other for versioninfo.xml
        'For more info about sparkle read http://netsparkle.codeplex.com/documentation

        SparkleAutoUpdater = New AppLimit.NetSparkle.Sparkle(versionInfoUrl) With {.ApplicationWindowIcon = appIcon}
#If CONFIG = "SparkleDiagnosticWindow" Then
        'Uncomment if you want to diagnostic.
        SparkleAutoUpdater.ShowDiagnosticWindow = True
#End If
        SparkleAutoUpdater.ShowDiagnosticWindow = True
        If Not Balora.Util.CheckInternetConnection(True) Then
            ShowAlertForm("",
                          "Cacadu cannot check for updates, please check your connection.",
                          AlertLocation,
                          "",
                          False, My.Resources.Cacadu_Image_46x46)
        Else
            'SparkleAutoUpdater.StartLoop(doInitialCheck:=True)
            SparkleAutoUpdater.StartLoop(doInitialCheck:=True, forceInitialCheck:=True, checkFrequency:=New TimeSpan(1, 0, 0))
        End If
    End Sub

    Friend Sub SendExceptionToConderella(ByVal e As Exception, Optional ByVal isUnHandledException As Boolean = False)
        Dim errorreporterobj As New ErrorReporter
        errorreporterobj.IsUnHandledException = isUnHandledException
        ErrorReportingUI = New Balora.ErrorReporter.ErrorReportingForm
        Dim errreportinguiconnector As New ErrorReportingFormConnector(errorreporterobj)
        'if the user write any comments but click "don't send" button in errorreportingui
        'or click the close button, then these comments will appear again when the form get
        'displayed again. to solve this=> i check for every calling for the form if the comments
        'text box texts isn't == "if you wish, you can add your comments here." and change it to
        '"add your comments here...".
        If ErrorReportingUI.CommentsTextBox.Text.Contains("if you wish, you can add your comments here.") Then
            ErrorReportingUI.CommentsTextBox.Text = "add your comments here..."
        End If
        errreportinguiconnector.ShowErrorReportingDialog(e)
    End Sub

    Friend Function SendMessageToCacaSer(message As Object) As Boolean
        'My.Resources.CSN == CacaSer
        'AhSaElDin 20120102: disabled for now, as CacaSer is disabled.
        'Return Ipc.NamedPipes.SendMessageHelper(NamedPipesClient, My.Resources.CSN, "Cacadu", message)
        Return Nothing
    End Function

    Friend Sub TryCutTectonicLink()
        'للأسف فيه مكان تاني لسه بيبقي متصل بعد استدعائها
        Watcher.StopAllWatchedTask()
        SuccessiveChecks._switch = False
        ShutdownQuartz()
        Cacadore.Settings.CrudObject = Nothing
        TypedDataSetCrud.CloseConnection()
        _typedDataSetCrud.Dispose()
        _typedDataSetCrud = Nothing
    End Sub

#Region "Export & Import Tectonic"
    Friend Function ExportTectonic() As Boolean
        Dim encryptedDatabaseFilePath As String = vbNullString
        With MainForm.SaveFileDialog
            .Title = "Export Cacadu Tasks..."
            .Filter = "Cacadu backup files (*.cadat)|*.cadat"
            Dim dialogResult = .ShowDialog()
            If dialogResult = Windows.Forms.DialogResult.OK Then encryptedDatabaseFilePath = .FileName
        End With
        If encryptedDatabaseFilePath = "" Then Return False

        Dim fileContents As Byte() = Nothing
        Dim tempTectonicPath As String = Balora.PathsHelper.GetCommonApplicationData() & "Cacadu\" & "tectonicBE"

        Try
            FileCopy(TecDAL.TypedDataSetCrud.TectonicPath, tempTectonicPath)
            fileContents = My.Computer.FileSystem.ReadAllBytes(tempTectonicPath)
        Catch ex As Exception
            ShowAlertForm("",
                          "error reading Cacadu tasks.",
                          AlertLocation,
                          "",
                          False,
                          My.Resources.Cacadu_Image_46x46)
            Return False
        End Try

        Try
            My.Computer.FileSystem.WriteAllText(encryptedDatabaseFilePath, EncryptTectonic(fileContents).ToBase64(), False)
        Catch ex As Exception
            ShowAlertForm("", "Error writing export file.", AlertLocation, "", False, My.Resources.Cacadu_Image_46x46)
            Return False
        End Try

        IO.File.Delete(tempTectonicPath)
        Return True
    End Function

    Friend Function EncryptTectonic(fileContents As Byte()) As Encryption.Data
        Dim encryptedData As Encryption.Data
        Dim sym As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
        Dim key As New Encryption.Data("aSdKiUjYhb7jh6h9jh5hIO8JY6Hij9Jyh6hj9jY5FJHoiPOl.LjhnFeDWEcbK97864#$@#$@#$@#$@#DFGSDfg@#$%!^%")
        encryptedData = sym.Encrypt(New Encryption.Data(fileContents), key)
        Return encryptedData
    End Function

    Friend Function ImportTectonic() As Boolean
        Dim encryptedDatabaseFilePath As String = vbNullString
        With MainForm.OpenFileDialog
            .Title = "Export Cacadu Tasks..."
            .Filter = "Cacadu backup files (*.cadat)|*.cadat"
            Dim dialogResult = .ShowDialog()
            If dialogResult = Windows.Forms.DialogResult.OK Then encryptedDatabaseFilePath = .FileName
        End With
        If encryptedDatabaseFilePath = "" Then Return False

        Dim fileContents As String = Nothing

        Try
            fileContents = My.Computer.FileSystem.ReadAllText(encryptedDatabaseFilePath)
        Catch ex As Exception
            ShowAlertForm("", "error reading Cacadu tasks.", AlertLocation, "", False, My.Resources.Cacadu_Image_46x46)
            Return False
        End Try

        Dim result = MsgBox("All current tasks will be overwritten with the imported tasks," & vbNewLine &
                            "Are you sure you want to proceed? ",
                            CType(MsgBoxStyle.YesNo + MsgBoxStyle.Information, MsgBoxStyle), "Cacadu")

        If result = MsgBoxResult.Yes Then
            Try
                Dim tempTectonicPath As String = Balora.PathsHelper.GetCommonApplicationData() & "Cacadu\" & "tectonicIF"
                My.Computer.FileSystem.WriteAllBytes(tempTectonicPath, DecryptTectonic(fileContents).Bytes, False)
            Catch ex As Exception
                ShowAlertForm("", "Error importing Cacadu tasks.", AlertLocation, "", False, My.Resources.Cacadu_Image_46x46)
                Return False
            End Try
            MsgBox("Please, restart Cacadu now.", MsgBoxStyle.Information, "Cacadu")
        End If

        Return True
    End Function

    Friend Function DecryptTectonic(fileContents As String) As Encryption.Data
        Dim sym As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
        Dim key As New Encryption.Data("aSdKiUjYhb7jh6h9jh5hIO8JY6Hij9Jyh6hj9jY5FJHoiPOl.LjhnFeDWEcbK97864#$@#$@#$@#$@#DFGSDfg@#$%!^%")
        Dim encryptedData As New Encryption.Data
        encryptedData.Base64 = fileContents.ToString
        Dim decryptedData As Encryption.Data = sym.Decrypt(encryptedData, key)
        Return decryptedData
    End Function

    Friend Sub CheckForImportFile()
        Dim tectonicPath As String = Balora.PathsHelper.GetCommonApplicationData() & "Cacadu\" & "tectonic"
        Dim tempTectonicPath As String = Balora.PathsHelper.GetCommonApplicationData() & "Cacadu\" & "tectonicIF"
        If IO.File.Exists(tempTectonicPath) Then
            IO.File.Delete(tectonicPath)
            FileCopy(tempTectonicPath, tectonicPath)
            IO.File.Delete(tempTectonicPath)
        End If
    End Sub
#End Region
#End Region
#End Region
#Region "Global Events"
#Region "Balora Events"
    <Conditional("DEBUG")>
    Friend Sub UpdateBottleneck(state As Object)
        If Not BottleneckChecker.IsShown Then BottleneckChecker.Show()
        Dim args() As String = CType(state, String())
        For Each row As DataGridViewRow In BottleneckChecker.DataGridView.Rows
            If row.Cells(0).Value.ToString.Contains(args(0)) Then
                Dim callsCount As Integer = CInt(row.Cells(1).Value) + 1
                Dim totalCallsTime As Double = CDbl(row.Cells(3).Value) + CDbl(args(1))
                row.Cells(1).Value = callsCount
                row.Cells(2).Value = args(1)
                row.Cells(3).Value = totalCallsTime
                Exit Sub
            End If
        Next
        BottleneckChecker.DataGridView.Rows.Add(args(0), 1, args(1), CDbl(args(1)))
    End Sub

    'Any events raised from Balora and Cacadu
    Friend Sub Alerter_NewErrorOccur(ByVal e As System.Exception, description As String, broadCast As Boolean)

        'If boradCast and exception is just not "Nothing"
        'Exception sent "nothing" to this event only when you use RaiseUpEvent like this
        ' Balora.Alerter.RaiseErrorUp("any value here or null", Nothing, True)
        If broadCast And Not e Is Nothing Then
            SendExceptionToConderella(e)
        End If
        If description <> "" Then
            ShowAlertForm("", description, AlertLocation, "", False, My.Resources.Cacadu_Image_46x46)
#If DEBUG Then
        Else
            Balora.Hodhod.WriteToErrorLog(e.Message, "", "Alerter_NewErrorOccur::Error has no description")
#End If
        End If
    End Sub

    Private Sub Alerter_WaitForm_Load(description As String)
        ShowWaitForm(description)
    End Sub

    Private Sub Alerter_WaitForm_Close()
        CloseWaitForm()
    End Sub
#End Region
    'AhSaElDin 20120102: disabled for now, as CacaSer is disabled.
    '#Region "NamedPipes Events"
    '    Private Sub NampedPipesClient_Connected(sender As Object, e As System.EventArgs) Handles NamedPipesClient.Connected
    '    End Sub
    '    Private Sub NampedPipesClient_MessageReceived(sender As Object, e As Balora.Ipc.NamedPipes.MessageEveArgs) Handles NamedPipesClient.MessageReceived
    '        Util.STOW("Cacadu received a message from " & My.Resources.CSN & ".")
    '        Dim message As Object = Ipc.NamedPipes.ProcessReceivedMessage(e)
    '        Select Case message.GetType
    '            Case GetType(TaskData.StringMessages)
    '                ProcessStringMessage(CType(message, TaskData.StringMessages))
    '            Case Else
    '                ProcessObjectMessage(message)
    '        End Select
    '    End Sub
    '    Private Sub ProcessStringMessage(msg As TaskData.StringMessages)
    '        Select Case msg
    '            'Case CacaSer.StringMessages.mald 'for ex:
    '        End Select
    '    End Sub
    '    Private Sub ProcessObjectMessage(msg As Object)
    '    End Sub
    '#End Region
#End Region
End Module
