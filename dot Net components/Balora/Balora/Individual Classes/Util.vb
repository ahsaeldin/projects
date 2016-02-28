'Balora v1.00  .. 06/11/2010
'Balora v2.00  .. 16/02/2011

' Contains several functions to assist Balora in many usual tasks.

Option Strict Off
Option Explicit On

#Region "Imported NameSpaces"

Imports System.IO
Imports System.Net
Imports System.Drawing
Imports System.Reflection
Imports Balora.PathsHelper
Imports System.Windows.Forms
Imports System.Collections.ObjectModel

#End Region

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'                                                   Util Module                                                     
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

#Region "Structures"
'This Structure is more than awesome. It enables me to tell...
'the calling method of the method that return "Result" a message...
'and the if succeeded 
Friend Structure Result
    Dim Message As String
    Dim Successed As Boolean
End Structure

''' <summary>
''' Structure created by me to save loaded assembly name and location
''' </summary>
''' <remarks></remarks>
Public Structure LoadedAssembly
    Dim AssemplyName As String
    Dim AssemplyLocation As String
End Structure
#End Region

Public Class Util
    Inherits BaloraBase
    Implements IDisposable

#Region "Fields"
    Private Delegate Sub SetValueDlg(obj As [Object], val As [Object], index As [Object]())
#End Region

#Region "Functions"
    Public Shared Function IsAssemplyInGAC(ByVal assemplyName As String) As Boolean
        IsAssemplyInGAC = False
        Dim GACDirPath As String = ""
        Dim WinDirPath As String = Environment.GetEnvironmentVariable("windir")
        CheckBackSlashAtEnd(WinDirPath)
        GACDirPath = WinDirPath & "assembly"
        For Each directory As String In FileIO.FileSystem.GetDirectories(GACDirPath, FileIO.SearchOption.SearchAllSubDirectories)
            If directory.Contains(assemplyName) Then
                IsAssemplyInGAC = True
            End If
        Next
        Return IsAssemplyInGAC
    End Function

    Public Shared Sub AddAssemplyToGAC(ByVal assemblyPath As String)
        'This function needs a reference to System.EnterpriseSerivces, so 
        'Add this reference before uncommenting it.
        'Try
        '    Dim EntServ As New System.EnterpriseServices.Internal.Publish
        '    EntServ.GacInstall(assemblyPath)
        'Catch ex As Exception
        'If Debugger.IsAttached Then Debugger.Break()
        '    Alerter.RaiseErrorUp("Error while adding assembly to GAC.", ex, True)
        'End Try
    End Sub

    Public Shared Sub RemoveAssemplyFromGAC(ByVal assemblyPath As String)
        'This function needs a reference to System.EnterpriseSerivces, so 
        'Add this reference before uncommenting it.
        ''Note: http://stackoverflow.com/questions/45729/what-path-should-i-pass-as-an-assemblypath-parameter-to-the-publish-gacremove-fun
        'Try
        '    Dim EntServ As New System.EnterpriseServices.Internal.Publish
        '    EntServ.GacRemove(assemblyPath)
        'Catch ex As Exception
        'If Debugger.IsAttached Then Debugger.Break()
        'Alerter.RaiseErrorUp("Error while removing assembly from GAC.", ex, True)
        'End Try
    End Sub

    Public Shared Function IsAnotherInstanceRunning() As Boolean
        'copy of IsAnotherInstanceRunning_3shanVista in Cute Macro
        Dim ProcessCounts As Integer = Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length
        If ProcessCounts > 1 Then
            IsAnotherInstanceRunning = True
        Else
            IsAnotherInstanceRunning = False
        End If
    End Function

    Public Shared Function IsProcessRunning(ByVal processName As String) As Boolean
        For Each clsProcess As Process In Process.GetProcesses()
            If clsProcess.ProcessName.StartsWith(processName) Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Shared Sub CopyDataToClipboard(ByVal format As String, ByVal str As String)
        'الوظيفة دي ووظيفة
        'PasteDataFromClipboard
        'عشان النسخ واللصق من وإلي الكيب بورد بتنسيقات معينة
        'مما يساعد أولا علي ارسال رسائل بين حالات البرنامج المختلفة
        'لتفادي استخدام
        'WM_COPYDATA 

        ' Creates a new data format.
        Dim CDataFormat As DataFormats.Format = DataFormats.GetFormat(format)

        ' Creates a new object and store it in a DataObject using myFormat 
        ' as the type of format. 
        Dim CopiedText As String = str
        Dim DataObject As New DataObject(CDataFormat.Name, CopiedText)

        'لازم تاني بارامتر يبقي ترو عشان تقدر تنقل محتويات الكليب بوورد بين برنامجين
        Clipboard.SetDataObject(DataObject, True)
    End Sub

    Public Shared Function PasteDataFromClipboard(ByVal format As String) As String
        'الوظيفة دي ووظيفة
        'PasteDataFromClipboard
        'عشان النسخ واللصق من وإلي الكيب بورد بتنسيقات معينة
        'مما يعنني أولا علي ارسال رسائل بين حالات البرنامج المختلفة
        'لتفادي استخدام
        'WM_COPYDATA 

        ' Creates a new data format.
        Dim CMFormat As DataFormats.Format = DataFormats.GetFormat(format)

        ' Performs some processing steps.
        ' Retrieves the data from the clipboard.
        Dim RetrievedItemsObject As IDataObject = Clipboard.GetDataObject()

        ' Converts the IDataObject type to MyNewObject type. 
        Dim CopiedItems As String = CType(RetrievedItemsObject.GetData(CMFormat.Name), String)

        PasteDataFromClipboard = CopiedItems
    End Function

    Public Shared Function GetCallingMethod(ByVal FrameIndex As Integer) As String
        ' Frame 0 = Current Method Name
        ' Frame 1 = Calling Method Name
        ' Frame 2 = The Calling method of the previous one
        Dim StackTrace As New StackTrace
        GetCallingMethod = StackTrace.GetFrame(FrameIndex).GetMethod().Name
    End Function

    Public Shared Function GetExecutingAssemblyName() As String
        'Gets the assembly that contains the code that is currently executing.
        'http://msdn.microsoft.com/query/dev10.query?appId=Dev10IDEF1&l=EN-US&k=k%28SYSTEM.REFLECTION.ASSEMBLY.GETEXECUTINGASSEMBLY%29;k%28TargetFrameworkMoniker-%22.NETFRAMEWORK%2cVERSION%3dV4.0%22%29;k%28DevLang-VB%29&rd=true
        Return Assembly.GetExecutingAssembly.GetName.Name
    End Function

    Public Shared Function GetCallingAssemblyName() As String
        'Returns the Assembly of the method that invoked the currently executing method.
        'http://msdn.microsoft.com/query/dev10.query?appId=Dev10IDEF1&l=EN-US&k=k%28SYSTEM.REFLECTION.ASSEMBLY.GETCALLINGASSEMBLY%29;k%28TargetFrameworkMoniker-%22.NETFRAMEWORK%2cVERSION%3dV4.0%22%29;k%28DevLang-VB%29&rd=true
        Return Assembly.GetCallingAssembly.GetName.Name
    End Function

    ''' <summary>
    ''' Gets the first name of the assembly called.
    ''' </summary><returns></returns>
    Public Shared Function GetFirstAssemblyCalledName() As String
        'Gets the process executable in the default application domain. In other application domains, this is the first executable that was executed by AppDomain.ExecuteAssembly.
        'http://msdn.microsoft.com/query/dev10.query?appId=Dev10IDEF1&l=EN-US&k=k%28SYSTEM.REFLECTION.ASSEMBLY.GETENTRYASSEMBLY%29;k%28TargetFrameworkMoniker-%22.NETFRAMEWORK%2cVERSION%3dV4.0%22%29;k%28DevLang-VB%29&rd=true
        Try
            Dim name As String = Assembly.GetEntryAssembly.GetName.Name
            Return name
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Return "Error gets the first name of the assembly called"
        End Try
    End Function

    Public Shared Function AssociateMyIconAndExtension(ByVal ext As String, ByVal appKeyNameInReg As String, ByVal fileDescription As String) As Boolean
        'Ext : The extension you want to assign to your app.
        'AppKeyNameInReg : The App Key Name in the Reg.
        'FileDescription : File Description that appears in properties window.
        'Example:
        'Dim Ext As String = "pmac" '//The extenstion you want to assign to your app.
        'Dim AppKeyNameInReg As String = "Perfect Macro Recorder" '//The App Key Name in the Reg.
        'Dim FileDescription As String = "Perfect Macro Recorder" & " File" '//File Description that appears in Properites window.
        Dim appPath As String = System.Windows.Forms.Application.ExecutablePath
        Try
            My.Computer.Registry.ClassesRoot.CreateSubKey("." & ext).SetValue _
            ("", appKeyNameInReg, Microsoft.Win32.RegistryValueKind.String)

            My.Computer.Registry.ClassesRoot.CreateSubKey(appKeyNameInReg).SetValue _
            ("", fileDescription, Microsoft.Win32.RegistryValueKind.String)

            My.Computer.Registry.ClassesRoot.CreateSubKey(appKeyNameInReg & "\DefaultIcon").SetValue _
            ("", appPath & ",0", Microsoft.Win32.RegistryValueKind.String)

            My.Computer.Registry.ClassesRoot.CreateSubKey(appKeyNameInReg & "\shell\open\command").SetValue _
            ("", appPath & " ""%l"" ", Microsoft.Win32.RegistryValueKind.String)

            My.Computer.Registry.CurrentUser.CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\." & ext).SetValue _
            ("Application", appPath, Microsoft.Win32.RegistryValueKind.String)
            Return True
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("", ex, True)
            Return False
        End Try
    End Function

    Public Shared Function FindAStringOnAFile(ByVal fileToSearchWithin As String, ByVal containsText As String) As ReadOnlyCollection(Of String)
        'ممكن فيما بعد تطلع بقية البرامتزر بتاعة 
        'FindInFiles
        'وتخليهم في
        'FindAStringOnAFile
        '//Search a file for a string
        Dim Files As ReadOnlyCollection(Of String)
        Files = My.Computer.FileSystem.FindInFiles(fileToSearchWithin, containsText, True, FileIO.SearchOption.SearchAllSubDirectories, "*.exe")
        FindAStringOnAFile = Files
    End Function

    Public Shared Function IsFileContainsSrting(ByVal SearchWord As String, ByVal FileName As String) As Integer
        '//Also Search a file for a string
        Dim FileContents As String = String.Empty
        Dim FileStreamReader As StreamReader = Nothing
        Dim IsWordFound As Boolean = False
        Try
            FileStreamReader = New StreamReader(FileName)
            FileContents = FileStreamReader.ReadToEnd()
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
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

    Public Shared Function GetBuildDate(ByVal filePath As String) As DateTime
        Const PeHeaderOffset As Integer = 60
        Const LinkerTimestampOffset As Integer = 8
        Dim b(2047) As Byte
        Dim s As Stream = Nothing
        Try
            s = New FileStream(filePath, FileMode.Open, FileAccess.Read)
            s.Read(b, 0, 2048)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("Util.GetBuildDate::error opening " & Chr(34) & filePath & Chr(34) & " to get its build date.", ex, True)
        Finally
            If Not s Is Nothing Then s.Close()
        End Try
        Dim i As Integer = BitConverter.ToInt32(b, PeHeaderOffset)
        Dim SecondsSince1970 As Integer = BitConverter.ToInt32(b, i + LinkerTimestampOffset)
        Dim dt As New DateTime(1970, 1, 1, 0, 0, 0)
        dt = dt.AddSeconds(SecondsSince1970)
        dt = dt.AddHours(TimeZone.CurrentTimeZone.GetUtcOffset(dt).Hours)
        Return dt
    End Function

    Public Shared Function GetFullWindowsVersion() As String
        Dim osInfo As System.OperatingSystem = System.Environment.OSVersion
        Dim WindowsVersion As String = My.Computer.Info.OSFullName
        Dim WindowsBuildNumber As String = My.Computer.Info.OSVersion
        Dim ServicePack As String = osInfo.ServicePack
        GetFullWindowsVersion = WindowsVersion & " - Version " &
        WindowsBuildNumber & " [" & ServicePack & "]"
    End Function

    Public Shared Function GetWindowsInfo() As System.OperatingSystem
        Return System.Environment.OSVersion
    End Function

    Public Shared Function GetFullWindowsBuildNumber() As Integer
        Return My.Computer.Info.OSVersion
    End Function

    Public Shared Function GetFullWindowsFullName() As String
        Return My.Computer.Info.OSFullName
    End Function

    Public Shared Function GetAvailablePhysicalMemory() As ULong
        Dim memSize As ULong
        Try
            memSize = My.Computer.Info.AvailablePhysicalMemory / 1024 / 1024
            Return memSize
        Catch ex As System.ComponentModel.Win32Exception
            Alerter.REP("The application cannot obtain the memory status.", ex, True)
            Return "The application cannot obtain the memory status."
        End Try
    End Function

    Public Shared Function GetTotalPhysicalMemory() As ULong
        Dim memSize As ULong
        Try
            memSize = My.Computer.Info.TotalPhysicalMemory / 1024 / 1024
            Return memSize
        Catch ex As System.ComponentModel.Win32Exception
            Alerter.REP("The application cannot obtain the memory status.", ex, True)
            Return "The application cannot obtain the memory status."
        End Try
    End Function

    Public Shared Function GetLoadedAssemblies() As Dictionary(Of String, String)
        'How to call:
        'Dim dic As New Dictionary(Of String, String)
        'dic = Util.GetLoadedAssemplies
        'For Each kvp As KeyValuePair(Of String, String) In dic
        'Console.WriteLine("Key = {0}, Value = {1}", _
        '   kvp.Key, kvp.Value)
        'Next kvp
        Dim assem As Reflection.Assembly
        'Structure created by me to save a loaded assembly name and location.
        Dim loadedAssembly As LoadedAssembly
        Dim assembliesList As New Dictionary(Of String, String)
        For Each assem In AppDomain.CurrentDomain.GetAssemblies
            loadedAssembly.AssemplyName = assem.ToString
            loadedAssembly.AssemplyLocation = assem.Location
            Try
                assembliesList.Add(loadedAssembly.AssemplyName, loadedAssembly.AssemplyLocation)
            Catch ex As ArgumentNullException
                Alerter.REP("Null parameters passed to assempbliesList dictionary.", ex, True)
            Catch ex As ArgumentException
                Alerter.REP("Error passing parameters to assempbliesList dictionary.", ex, True)
            End Try
        Next assem
        Return assembliesList
    End Function

    Public Shared Function GetMyRealIP(ByVal uriToReteriveIp As String) As String
        'Synchronous Get IP address
        'Idea from here: http://stackoverflow.com/questions/1242484/get-real-ip-address-with-vb-net
        'Create a php file and paste this in it:
        '<?php
        'echo $_SERVER['REMOTE_ADDR'];
        '?>
        'save it as curip.php and upload it to your server. 
        'ex:GetMyRealIP("http://www.perfect-macro-recorder.com/getip.php")
        Dim uri_val As New Uri(uriToReteriveIp)
        Dim request As HttpWebRequest = CType(HttpWebRequest.Create(uri_val), HttpWebRequest)
        request.Method = WebRequestMethods.Http.Get
        Dim response As HttpWebResponse = Nothing
        request.Timeout = 25000
        Try
            response = CType(request.GetResponse(), HttpWebResponse)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Return ex.Message
        End Try
        Dim myIP As String
        Using reader As New StreamReader(response.GetResponseStream())
            myIP = reader.ReadToEnd()
            response.Close()
        End Using
        Return myIP
    End Function

    Public Shared Function ShowOpenFileDialog(ByVal title As String) As String
        Using openFileDialogObj As New OpenFileDialog
            With openFileDialogObj
                .Title = title
                .AutoUpgradeEnabled = True
                .FileName = ""
                .Filter = "All files (*.*)|*.*"
                .ShowDialog()
                Return .FileName
            End With
        End Using
    End Function

    'Disable/Enable Buttons and Textboxs in a form.
    Public Shared Sub DisableButtonsAndTextBoxes(ByVal form As Form, Optional ByVal enable As Boolean = False)
        For Each ctrl As Control In form.Controls
            If TypeOf (ctrl) Is Button Or TypeOf (ctrl) Is TextBox Then
                ctrl.Enabled = enable
            End If
            For Each childCtrl As Control In ctrl.Controls

                If TypeOf (childCtrl) Is Button Or TypeOf (childCtrl) Is TextBox Then
                    childCtrl.Enabled = enable
                End If
            Next
        Next
    End Sub

    Public Shared Function CreateObjectDynamically(ByVal typeName As String) As Object
        'ex: CreateObjectDynamically("Cacadu.ErrorReporterForm")
        'Based on turkey's book Page no 435.
        Dim T As Type = Type.GetType(typeName)
        Return Activator.CreateInstance(T)
    End Function

    Public Shared Function GetStackTrace() As ArrayList
        Dim stackArray As New ArrayList
        Dim ST As New StackTrace(True)
        For Each sFrame As StackFrame In ST.GetFrames
            Dim count As Integer
            count += 1
            'Only if count <> 1 to skip counting this function
            If count <> 1 Then
                stackArray.Add(sFrame.GetMethod.Name)
            End If
        Next
        Return stackArray
    End Function

    Public Shared Function GetLastFrameInExceptionStackTarce(ByVal exceptionObj As Exception) As StackFrame
        'Stack created by my to reverse the order of stack trace and get the one we need.
        Dim reverseStack As New Stack(Of StackFrame)
        Dim sF As StackFrame
        Dim trace As New System.Diagnostics.StackTrace(exceptionObj, True)
        For Each ahmed As StackFrame In trace.GetFrames
            reverseStack.Push(ahmed)
        Next
        sF = reverseStack.Pop()
        Return sF
    End Function

    Public Shared Function GetFullExceptionStackTrace(ByVal ex As Exception) As String
        Dim outputStack As New System.Text.StringBuilder
        outputStack.Append(ex.StackTrace)
        Dim innerReferences As Byte = 0 'used to ensure memory does not run out
        'during stack looping
        Dim innerException As Exception = ex.InnerException
        While Not innerException Is Nothing _
            AndAlso innerReferences < 50
            outputStack.Insert(0, innerException.StackTrace)
            innerException = innerException.InnerException
            innerReferences += CByte(1)
        End While
        Return outputStack.ToString
    End Function

    ''' <summary>
    '''  Convert an object to a byte array.
    ''' </summary>
    ''' <param name="obj">Any object</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ObjectToByteArray(ByVal obj As [Object]) As Byte()
        Dim serializingResult As Boolean
        Dim memStream As New MemoryStream()
        serializingResult = Serializer.BinarySerializer.SerializeObject(obj, memStream)
        If serializingResult Then
            Return memStream.ToArray
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Take a textbox then add what you want between text lines.
    ''' </summary>
    ''' <param name="textbox"></param>
    ''' <param name="whatToInsert"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InsertBetweenTextBoxLines(ByVal textbox As TextBox, ByVal whatToInsert As String) As String
        Dim result As String = Nothing
        'I will put a line break between every line.
        Dim commentTextCounter As Integer
        'Create a string array and store the contents of the Lines property.
        Dim tempArray() As String
        tempArray = textbox.Lines
        'Loop through the array and add break lines
        For commentTextCounter = 0 To tempArray.GetUpperBound(0)
            'Add break line.
            If tempArray(commentTextCounter) = "" Then
                result += whatToInsert
            Else ' or add the line itself.
                result += tempArray(commentTextCounter)
            End If
        Next
        Return result
    End Function

    ''' <summary>
    ''' Convert image to a byte array by saving to a memory stream
    ''' then convert memory stream to a byte array using ToArray() method. 
    ''' </summary>
    ''' <param name="img">Is the image to be converted</param>
    ''' <returns>Byte array of the image</returns>
    ''' <remarks></remarks>
    Public Shared Function ImageToByteArray(ByVal img As Image) As Byte()
        Dim memStream As New MemoryStream()
        img.Save(memStream, System.Drawing.Imaging.ImageFormat.Png)
        Return memStream.ToArray()
    End Function

    Public Shared Function ByteArrayToObject(ByVal bArray As Byte()) As Object
        Return Serializer.BinarySerializer.DeserializeObject(bArray)
    End Function

    Public Shared Function ByteArrayToImage(ByVal bArray As Byte()) As Image
        Dim memStream As New MemoryStream(bArray)
        Return Image.FromStream(memStream)
    End Function

    Public Shared Function SetReadOnlyProperty(obj As Object, propertyName As String, newValue As Object) As Boolean
        Try
            Dim p As System.Reflection.PropertyInfo = obj.GetType.GetProperty(
            propertyName, Reflection.BindingFlags.Instance Or Reflection.BindingFlags.Public)
            p.SetValue(obj, newValue, Nothing)
            Return True
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("", ex, True)
            Return False
        End Try
    End Function

    Public Shared Function SetPrivateField(obj As Object, fieldName As String, newValue As Object) As Boolean
        Try
            Dim p As System.Reflection.FieldInfo = obj.GetType.GetField(
            fieldName, Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic)
            p.SetValue(obj, newValue)
            Return True
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("", ex, True)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Sets the control property from within a thread other than the calling thread.
    ''' </summary>
    ''' <param name="ctrl">The CTRL.</param>
    ''' <param name="propName">Name of the prop.</param>
    ''' <param name="val">The val.</param>
    Public Shared Sub SetControlProperty(ctrl As Control, propName As [String], val As [Object])
        'SetControlProperty(Me.txtReceive, "Text", System.Text.Encoding.UTF8.GetString(mess.Payload))
        Dim propInfo As PropertyInfo = ctrl.[GetType]().GetProperty(propName)
        Dim dgtSetValue As [Delegate] = New SetValueDlg(AddressOf propInfo.SetValue)
        'index
        Try
            ctrl.Invoke(dgtSetValue, New [Object](2) {ctrl, val, Nothing})
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            Alerter.REP("", ex, True)
        End Try
    End Sub

    ''' <summary>
    ''' Compares the files.
    ''' </summary>
    ''' <param name="file1">The file1.</param>
    ''' <param name="file2">The file2.</param><returns></returns>
    Public Shared Function CompareFiles(ByVal file1 As String, ByVal file2 As String) As Boolean
        'Set to true if the files are equal; false otherwise
        Dim isFilesAreEqual As Boolean = False
        With My.Computer.FileSystem
            'Ensure that the files are the same length before comparing them line by line
            If .GetFileInfo(file1).Length = .GetFileInfo(file2).Length Then
                Try
                    Using file1Reader As New FileStream(file1, FileMode.Open), file2Reader As New FileStream(file2, FileMode.Open)
                        Dim byte1 As Integer = file1Reader.ReadByte()
                        Dim byte2 As Integer = file2Reader.ReadByte()
                        ' If byte1 or byte2 is a negative value, we have reached the end of the file
                        While byte1 > 0 And byte2 > 0
                            If (byte1 <> byte2) Then
                                isFilesAreEqual = False
                                Exit While
                            Else
                                isFilesAreEqual = True
                            End If
                            'Read the next byte
                            byte1 = file1Reader.ReadByte()
                            byte2 = file2Reader.ReadByte()
                        End While
                    End Using
                Catch ex As Exception
                    Alerter.REP("", ex, True)
                End Try
            End If
        End With
        Return isFilesAreEqual
    End Function

    ''' <summary>
    ''' Invokes a method within assembly.
    ''' </summary>
    ''' <param name="assPath">The assembly path.</param>
    ''' <param name="methodName">Name of the method.</param>
    ''' <param name="parameter">Method parameters.</param><returns></returns>
    Public Shared Function InvokeMethodWithinAssembly(ByVal assPath As String, classOfTheMethod As String, methodName As String, parameter() As Object) As Object
        'How to use:
        'Dim quartzTrigger = InvokeMethodWithinAssembly("D:\Rocknee\KV\Cacadu\Code\Trunk\Cacadu\bin\Debug\BIT.dll", 
        '                                               "CreateTriggerByType",
        '                                               {2})

        assPath = GetCurrentExecutingDirectory() + assPath

        Dim ass As System.Reflection.Assembly

        ass = System.Reflection.Assembly.LoadFrom(assPath)

        Dim classTypeOfTheMethod As Type = ass.GetType(classOfTheMethod) 'ex: classOfTheMethod = "BIT.Helpers.QuartzHelper.QuartzUtils"

        Dim ConsineInfo As MethodInfo = classTypeOfTheMethod.GetMethod(methodName)

        Dim returnVal As Object = ConsineInfo.Invoke(classOfTheMethod, parameter)

        Return returnVal
    End Function

    ''' <summary>
    ''' Give me a byte array and tell me if it is an image then I will give you an object
    ''' </summary>
    Public Shared Function BlobToObject(ByVal bArray As Byte(), Optional ByVal isImage As Boolean = False) As Object
        If isImage Then
            Return Util.ByteArrayToImage(bArray)
        Else
            Return Util.ByteArrayToObject(bArray)
        End If
    End Function

    ''' <summary>
    ''' Convert any object to a blob
    ''' </summary>
    ''' <param name="obj">If image then convert it by saving to a mem stream,
    ''' else convert using serialization</param>
    ''' <returns>Byte array</returns>
    ''' <remarks></remarks>
    Public Shared Function ObjectToBlob(ByVal obj As Object) As Byte()
        If TypeOf obj Is Drawing.Image Then
            Return Util.ImageToByteArray(CType(obj, Drawing.Image))
        Else
            Return Util.ObjectToByteArray(obj)
        End If
        Return Nothing
    End Function

    Public Shared Function LoadImageFromUrl(ByRef url As String, Optional ByVal pb As PictureBox = Nothing) As Image
        Dim request As Net.HttpWebRequest = DirectCast(Net.HttpWebRequest.Create(url), Net.HttpWebRequest)
        Dim response As Net.HttpWebResponse = DirectCast(request.GetResponse, Net.HttpWebResponse)
        Dim img As Image = Image.FromStream(response.GetResponseStream())
        response.Close()
        If Not IsNothing(pb) Then
            pb.SizeMode = PictureBoxSizeMode.StretchImage
            pb.Image = img
        End If
        Return img
    End Function

#Region "Internet Connection Checking"
    Public Shared Function CheckInternetConnection(Optional ping As Boolean = False) As Boolean
        If ping Then Return CheckInternetConnectionByPing()
        Return CheckInternetConnectionByWebRequest()
    End Function

    Public Shared Function CheckInternetConnectionByWebRequest() As Boolean
        ' Replace [url]www.google.com[/url] with a site that is guaranteed to be online 
        Dim objUrl As New System.Uri("https://www.dropbox.com/")
        ' Setup WebRequest 
        Dim objWebReq As System.Net.WebRequest
        objWebReq = System.Net.WebRequest.Create(objUrl)
        Dim objResp As System.Net.WebResponse
        Try ' Attempt to get response and return True 
            objWebReq.Timeout = 5000
            objResp = objWebReq.GetResponse
            objResp.Close()
            objWebReq = Nothing
            Return True
        Catch ex As Exception ' Error, exit and return False 
            objWebReq = Nothing
            Return False
        End Try
    End Function

    Public Shared Function CheckInternetConnectionByPing() As Boolean
        Dim isConnected As Boolean
        Try
            isConnected = My.Computer.Network.Ping("www.google.com", 5000)
        Catch ex As Exception
            isConnected = False
        End Try
        Return isConnected
    End Function
#End Region
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
