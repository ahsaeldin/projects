Imports System.CodeDom
Imports CuteMacro.Balora

Module ExecutableMacroHelperFunctions

    ' The main example in how to call GenerateExecutableMacro of MacroCodeGenerator class inherited from "CodeGen"
    ' Ok, sounds tricky, CodeGen is a class of Balora gives you the ability to generate code.
    ' MacroCodeGenerator is a class hosted in ExecutableMacro.vb and inherited from "CodeGen".
    ' MacroCodeGenerator exposes only one method "GenerateExecutableMacro" as of "16/11/2010".
    ' GenerateExecutableMacro is generate executable macro for if you:-
    ' 1. Pass the methods of the core functionality of the exe macro to it using its coreMethods parameter.
    ' 2. Fill all the required properties of MacroCodeGenerator before calling GenerateExecutableMacro...
    ' And this method "CompileMacroToExE" do all of this. Hope it will be simple, light and straightforward for you. :)
    Friend Function CompileMacroToExE(ByVal compiledFilePath As String, ByVal exeFileSettings As ArrayList) As Result

        Dim compilationResult As New Result 
        Dim processingDir As String = Application.StartupPath & "\pDir"
        Dim isProcessingDirExists As Boolean
        isProcessingDirExists = My.Computer.FileSystem.DirectoryExists(processingDir)
        If isProcessingDirExists Then
            My.Computer.FileSystem.DeleteDirectory(processingDir, FileIO.DeleteDirectoryOption.DeleteAllContents)
            My.Computer.FileSystem.CreateDirectory(processingDir)
        Else
            My.Computer.FileSystem.CreateDirectory(processingDir)
        End If


        '*.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.*
        'Practical way of calling and using GenerateExecutableMacro of MacroCodeGenerator
        '*.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.*
        'This is a practical way in how to create executable macro for further project ;)
        Using executableMacroGen As New MacroCodeGenerator
            With executableMacroGen
                ' Create a temp file for saving the generated code file by CodeGen.
                Dim generatedCodeFilePath As String = processingDir + "\cf.tmp"
                ' Where to save the generated code file.
                .OutputCodeFile = generatedCodeFilePath
                'In 24/12/2010 I deleted unused importes from the next line
                ' Add the import(s) statements to the resulting code file.
                .ImportedNameSpaces.AddRange(New String() {
                                    "System", "System.Drawing",
                                    "System.Windows.Forms",
                                    "Microsoft.VisualBasic",
                                    "System.EnterpriseServices",
                                    "MouseKeyboardLibrary"})
                ' Create a temp file for saving the macro file.
                Dim savedMacFile As String = processingDir + "\em-macro.pmac"
                SerializeListView(savedMacFile, MainForm.lstviwMacEdit)

                ' Add a string to the hashtable that will be passed to Balora generator.
                Dim soso As String = "nono"
                ' Add an icon to the hashtable that will be passed to Balora generator.
                Dim exeSystrayIcon As Icon = MainForm.Icon
                ' Add a macro file in order to add to the hashtable that will be passed to Balora generator.
                Dim customFile As Byte() = My.Computer.FileSystem.ReadAllBytes(savedMacFile)

                Dim mklextpLibBytes As Byte() = My.Computer.FileSystem.ReadAllBytes(Application.StartupPath & "\mklextp.dll")

                Dim systrayIconNameInResources As String = "PMRv2_10_all"
                ' Hashtable of the resource name and the corresponding object in the value
                Dim resourcesHashTable As New Hashtable
                resourcesHashTable.Add(systrayIconNameInResources, exeSystrayIcon)
                resourcesHashTable.Add("RepeatTime", exeFileSettings.Item(0))
                resourcesHashTable.Add("PlaybackSpeed", exeFileSettings.Item(1))
                resourcesHashTable.Add("MacroFile", customFile)
                resourcesHashTable.Add("mx", mklextpLibBytes)


                ' Path to the generated resource file that will contains the injected macro file.
                Dim exeMacroResFilePath As String = processingDir + "\exeMacroResources.resources"
                ' Generate resource file with the injected macro file.
                Dim generatingResourceFileResult As Boolean =
                Balora.xResources.WriteResource(exeMacroResFilePath, resourcesHashTable)
                If generatingResourceFileResult = False Then
                    compilationResult.Successed = False
                    compilationResult.Message = "Error generating macro data."
                    Return compilationResult
                End If

                'L:\PT\Perfection Tools Software\PMR\PMR 2.10\Cute Macro\Code\CM Working Copy\MouseKeyboardLibrary11\MouseKeyboardLibrary\bin\Release\mklextp.dll
                'C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Drawing.dll
                'C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Windows.Forms.dll
                Dim mklextpLibPath As String = Application.StartupPath & "\mklextp.dll"
                '.ReferencedAssemblies.AddRange(New String() {"System.dll", "System.Core.dll", "System.Data.dll",
                '                             "System.Drawing.dll", "System.Windows.Forms.dll", "System.Xml.dll",
                '                             "System.Xml.Linq.dll", "System.EnterpriseServices.dll", mklextpLibPath})
                'In 24/12/2010 I deleted unused references from the next line and pass the full path/
                Dim systemDrawingPath As String = Balora.Util.GetAssemplyPath(New System.Drawing.BufferedGraphicsContext())
                Dim windowsFormsPath As String = Balora.Util.GetAssemplyPath(New System.Windows.Forms.Form())
                Dim enterpriseServicesPath As String = Balora.Util.GetAssemplyPath(New System.EnterpriseServices.Internal.Publish())
                Dim systemCorePath As String = Balora.Util.GetAssemplyPath(New System.Collections.Generic.HashSet(Of String))
                Dim systemNameSpacePath As String = Balora.Util.GetAssemplyPath(New System.Diagnostics.EventLog())

                'we want to pass the full paths here
                .ReferencedAssemblies.AddRange(New String() {systemNameSpacePath,
                                                             systemDrawingPath,
                                                             systemCorePath,
                                                             windowsFormsPath,
                                                             enterpriseServicesPath,
                                                             mklextpLibPath})
                ' Tell Balora generator what code file to compile.
                .InputCodeFile = generatedCodeFilePath
                .EmbeddedResources = exeMacroResFilePath
                .CompiledFilePath = compiledFilePath
                Dim pathToIcon As String = Application.StartupPath & "\PMR.ico"
                .CompiledFileIcon = pathToIcon
                ' If you don't pass icon name in the embedded resource file
                ' adding systray icon igonored .
                .SystrayIconNameInResources = systrayIconNameInResources
                'Note: ToolTip text cannot exceed 63 characters . 
                .ExeSystrayIconTooltip = "Executable macro created by Perfect Macro Recorder."

                ' To add mklextp.dll to GAC, we need to add the responsible 
                ' statements @ the top of new constructor of the main class.
                .NewStatementsToConstructor = AddStatementsToConstrutor()
                ' A way for the user to customize the name of the core method of the exe file.
                .CoreMethodName = "CoreMethod()"
                ' "coreMethodForExeMacro" parameter is the core method for...
                ' the purpose that executable file created...
                ' and it will be created and add by the calling...
                ' object for "GenerateExecutableMacro" method
                Dim methodsArrayList As New ArrayList
                methodsArrayList.Add(CoreMethod)
                methodsArrayList.Add(SuspendExecutionForAlittlePeriod)
                methodsArrayList.Add(CodeDomMethodAddAssemplyToGAC)
                methodsArrayList.Add(CodeDomMethodIsAssemplyInGAC)
                methodsArrayList.Add(CodeDomIsAnotherInstanceRunning)
                methodsArrayList.Add(CodeDomTmrStop_Tick)
                compilationResult = .GenerateExecutableMacro(methodsArrayList)
            End With
        End Using
        '*.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.*
        'End
        '*.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.**.*


        isProcessingDirExists = My.Computer.FileSystem.DirectoryExists(processingDir)
        If isProcessingDirExists Then
            My.Computer.FileSystem.DeleteDirectory(processingDir, FileIO.DeleteDirectoryOption.DeleteAllContents)
        End If

        Return compilationResult

    End Function

    Private Function CoreMethod() As CodeMemberMethod

        Dim macroFileNameInResources As String = "MacroFile"
        Dim embededResourcesFileName As String = "exeMacroResources"
        Dim repeatTime As String = "RepeatTime"
        Dim playbackSpeed As String = "PlaybackSpeed"

        ' Generate an instance of MacroCodeGenerator in order to use helper functions like "CreateMethod".
        Dim executableMacroGen As New MacroCodeGenerator
        ' Generate the core method of the exe macro.
        Dim coreMethodForExeMacro As CodeMemberMethod = executableMacroGen.CreateMethod("CoreMethod")

        ' Adding a statement to CoreMethod that will make an instance of ResourceManager using the resource name stored in
        ' embededResourcesFileName. note this i get fucked to get the name of the resource file and took so long to know
        ' that name and used a tool like .net reflector that leads me at the end to a truth that you must add the name of the
        ' resource without the extension. i.e if the resource name is exeMacroResources.resources then pass exeMacroResources
        ' and if the resource name is WindowsApplication1.Resources.resources like usual then add WindowsApplication1.Resources.
        ' tricky huh?!

        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("If IsAnotherInstanceRunning() Then End"))

        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression(
                              "Dim ResourceManager As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager(" & """" & embededResourcesFileName & """" & ", System.Reflection.Assembly.GetExecutingAssembly())"))

        ' Extract embedded macro file from resources.
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression(
                                      "Dim obj As Object = ResourceManager.GetObject(" & """" & macroFileNameInResources & """" & ")"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Dim macFile As Byte()"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("macFile = CType(obj, Byte())"))
        ' Declare and deserialze a listivew with the macro data.
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Dim macroListView As New ListView"))

        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Try"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression(
                                             "Dim BinFormatter As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Dim ms As New System.IO.MemoryStream(macFile)"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("macroListView.Items.AddRange(BinFormatter.Deserialize(ms).ToArray(GetType(ListViewItem)))"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("ms.Dispose()"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Catch ex As Exception"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("MessageBox.Show(""Cannot load this executable macro because the file has been damaged."")"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("End Try"))

        ' Extract repeat time and playbackspeed.
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression(
                              "obj = ResourceManager.GetObject(" & """" & repeatTime & """" & ")"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Dim repeatTime As String"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("repeatTime = CType(obj, String)"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression(
                      "obj= ResourceManager.GetObject(" & """" & playbackSpeed & """" & ")"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Dim playbackSpeed As Integer"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("playbackSpeed = CType(obj, Integer)"))


        ' Now enter the loop of playbacking.
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Dim g As Integer"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("For g = 1 To repeatTime"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Dim ListItems As Object"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Dim lstItem As ListViewItem"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("ListItems = New ListView.ListViewItemCollection(macroListView)"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("For Each lstItem In ListItems"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Select Case lstItem.Text"))

        'Case "Wait"
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Case ""Wait"""))

        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("If playbackSpeed = 2 Then"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Dim Delay As Integer = macroListView.Items.Item(lstItem.Index).SubItems(1).Text"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("SuspendExecutionForAlittlePeriod(Delay)"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Else"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Select Case playbackSpeed"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Case 0"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("playbackSpeed = 100"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Case 1"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("playbackSpeed = 50"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Case 2"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("playbackSpeed = 2"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Case 3"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("playbackSpeed = 5"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Case 4"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("playbackSpeed = 0"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("End Select"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("System.Threading.Thread.Sleep(playbackSpeed)"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("End If"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression(
                                             "Dim forCallingDoevents As New Microsoft.VisualBasic.ApplicationServices.WindowsFormsApplicationBase()"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("forCallingDoevents.DoEvents()"))

        'Case "Background Window"
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Case ""Background Window"""))

        'SearchAndShowBackGroundWindow(lstItem)

        'Case "MouseWheel"
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Case ""MouseWheel"""))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("MouseSimulator.MouseWheel(lstItem.SubItems(1).Text)"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression(
                                             "Dim X As String = Strings.Replace(macroListView.Items.Item(lstItem.Index).SubItems(2).Text, ""X ="", """")"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression(
                                             "Dim Y As String = Strings.Replace(macroListView.Items.Item(lstItem.Index).SubItems(3).Text, ""Y ="", """")"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("MouseSimulator.X = X"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("MouseSimulator.Y = Y"))

        'Case "MouseMove"
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Case ""MouseMove"""))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression(
                                             "Dim X As String = Strings.Replace(macroListView.Items.Item(lstItem.Index).SubItems(1).Text, ""X ="", """")"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression(
                                             "Dim Y As String = Strings.Replace(macroListView.Items.Item(lstItem.Index).SubItems(2).Text, ""Y ="", """")"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("MouseSimulator.X = X"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("MouseSimulator.Y = Y"))

        'Case "LeftMouseDown", "MiddleMouseDown", "RightMouseDown"
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Case ""LeftMouseDown"", ""MiddleMouseDown"", ""RightMouseDown"""))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("If macroListView.Items.Item(lstItem.Index).SubItems(4).Text = 2097152 Then"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("MouseSimulator.MouseDown(MouseButton.Right)"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("ElseIf macroListView.Items.Item(lstItem.Index).SubItems(4).Text = 1048576 Then"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("MouseSimulator.MouseDown(MouseButton.Left)"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("End If"))

        'Case "LeftMouseUp", "MiddleMouseUp", "RightMouseUp"
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Case ""LeftMouseUp"", ""MiddleMouseUp"", ""RightMouseUp"""))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("If macroListView.Items.Item(lstItem.Index).SubItems(4).Text = 2097152 Then"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("MouseSimulator.MouseUp(MouseButton.Right)"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("ElseIf macroListView.Items.Item(lstItem.Index).SubItems(4).Text = 1048576 Then"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("MouseSimulator.MouseUp(MouseButton.Left)"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("End If"))

        'Case "KeyDown"
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Case ""KeyDown"""))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression(
                                             "KeyboardSimulator.KeyDown(macroListView.Items.Item(lstItem.Index).SubItems(7).Text)"))

        'Case "KeyUp"
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Case ""KeyUp"""))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression(
                                             "KeyboardSimulator.KeyUp(macroListView.Items.Item(lstItem.Index).SubItems(7).Text)"))

        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("End Select"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Next"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Next"))
        '
        '
        '
        '
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("TrayIcon.Visible = False"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("TrayIcon.Dispose()"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("Application.Exit()"))
        coreMethodForExeMacro.Statements.Add(New CodeSnippetExpression("End"))

        executableMacroGen.Dispose()

        Return coreMethodForExeMacro

    End Function

    ' A function used in the above mess 
    Private Function SuspendExecutionForAlittlePeriod() As CodeMemberMethod

        Using executableMacroGen As New MacroCodeGenerator

            Dim delayParameterType As Type = GetType(Integer)
            Dim delayParameter As New CodeParameterDeclarationExpression(delayParameterType, "Delay")

            Dim parameters As New CodeParameterDeclarationExpressionCollection
            parameters.Add(delayParameter)
            SuspendExecutionForAlittlePeriod = executableMacroGen.CreateMethod("SuspendExecutionForAlittlePeriod", parameters)

            SuspendExecutionForAlittlePeriod.Statements.Add(New CodeSnippetExpression("Dim OldTickCount As Integer = Environment.TickCount"))
            SuspendExecutionForAlittlePeriod.Statements.Add(New CodeSnippetExpression("Dim DelayValue As Integer"))
            SuspendExecutionForAlittlePeriod.Statements.Add(New CodeSnippetExpression("Do Until DelayValue >= Delay"))
            SuspendExecutionForAlittlePeriod.Statements.Add(New CodeSnippetExpression("DelayValue = Environment.TickCount - OldTickCount"))
            SuspendExecutionForAlittlePeriod.Statements.Add(New CodeSnippetExpression(
                         "Dim forCallingDoevents As New Microsoft.VisualBasic.ApplicationServices.WindowsFormsApplicationBase()"))
            SuspendExecutionForAlittlePeriod.Statements.Add(New CodeSnippetExpression("forCallingDoevents.DoEvents()"))
            SuspendExecutionForAlittlePeriod.Statements.Add(New CodeSnippetExpression("Loop"))
        End Using

        Return SuspendExecutionForAlittlePeriod

    End Function

    Private Function CodeDomMethodIsAssemplyInGAC() As CodeMemberMethod

        Using executableMacroGen As New MacroCodeGenerator

            Dim assemplyNameParameterType As Type = GetType(String)
            Dim assemplyNameParameter As New CodeParameterDeclarationExpression(assemplyNameParameterType, "assemplyName")
            Dim parameters As New CodeParameterDeclarationExpressionCollection
            parameters.Add(assemplyNameParameter)

            CodeDomMethodIsAssemplyInGAC = executableMacroGen.CreateMethod("IsAssemplyInGAC", parameters)
            Dim returnValue As New CodeTypeReference
            returnValue.BaseType = "Boolean"
            CodeDomMethodIsAssemplyInGAC.ReturnType = returnValue

            CodeDomMethodIsAssemplyInGAC.Statements.Add(New CodeSnippetExpression("IsAssemplyInGAC = False"))
            CodeDomMethodIsAssemplyInGAC.Statements.Add(New CodeSnippetExpression("Dim GACDirPath As String ="""""))
            CodeDomMethodIsAssemplyInGAC.Statements.Add(New CodeSnippetExpression(
                                                        "Dim WinDirPath As String = Environment.GetEnvironmentVariable(""windir"")"))
            CodeDomMethodIsAssemplyInGAC.Statements.Add(New CodeSnippetExpression("If Strings.Right(WinDirPath, 1) <> ""\"" Then"))
            CodeDomMethodIsAssemplyInGAC.Statements.Add(New CodeSnippetExpression("GACDirPath = WinDirPath & ""\assembly"""))
            CodeDomMethodIsAssemplyInGAC.Statements.Add(New CodeSnippetExpression("Else"))
            CodeDomMethodIsAssemplyInGAC.Statements.Add(New CodeSnippetExpression("GACDirPath = WinDirPath & ""assembly"""))
            CodeDomMethodIsAssemplyInGAC.Statements.Add(New CodeSnippetExpression("End If"))
            CodeDomMethodIsAssemplyInGAC.Statements.Add(New CodeSnippetExpression(
                                                        "For Each directory As String In FileIO.FileSystem.GetDirectories(GACDirPath, FileIO.SearchOption.SearchAllSubDirectories)"))
            CodeDomMethodIsAssemplyInGAC.Statements.Add(New CodeSnippetExpression("If directory.Contains(assemplyName) Then"))
            CodeDomMethodIsAssemplyInGAC.Statements.Add(New CodeSnippetExpression("IsAssemplyInGAC = True"))
            CodeDomMethodIsAssemplyInGAC.Statements.Add(New CodeSnippetExpression("End If"))
            CodeDomMethodIsAssemplyInGAC.Statements.Add(New CodeSnippetExpression("Next"))
            CodeDomMethodIsAssemplyInGAC.Statements.Add(New CodeSnippetExpression("Return IsAssemplyInGAC"))

        End Using

        Return CodeDomMethodIsAssemplyInGAC

    End Function

    ' Simply a function will be used in CompileMacroToExE to be
    ' passed to GenerateExecutableMacro using an arraylist.
    Private Function CodeDomMethodAddAssemplyToGAC() As CodeMemberMethod
        Using executableMacroGen As New MacroCodeGenerator
            Dim assemblyPathParameterType As Type = GetType(String)
            Dim assemblyPathParameter As New CodeParameterDeclarationExpression(assemblyPathParameterType, "assemblyPath")
            Dim parameters As New CodeParameterDeclarationExpressionCollection
            parameters.Add(assemblyPathParameter)
            CodeDomMethodAddAssemplyToGAC = executableMacroGen.CreateMethod("AddAssemplyToGAC", parameters)
            CodeDomMethodAddAssemplyToGAC.Statements.Add(New CodeSnippetExpression("Try"))
            CodeDomMethodAddAssemplyToGAC.Statements.Add(New CodeSnippetExpression("Dim EntServ As New System.EnterpriseServices.Internal.Publish"))
            CodeDomMethodAddAssemplyToGAC.Statements.Add(New CodeSnippetExpression("EntServ.GacInstall(assemblyPath)"))
            CodeDomMethodAddAssemplyToGAC.Statements.Add(New CodeSnippetExpression("Catch ex As Exception"))
            CodeDomMethodAddAssemplyToGAC.Statements.Add(New CodeSnippetExpression(
                                                         "MessageBox.Show(""Error while start playback the executable macro"")"))
            CodeDomMethodAddAssemplyToGAC.Statements.Add(New CodeSnippetExpression("End Try"))
        End Using
        Return CodeDomMethodAddAssemplyToGAC
    End Function

    ' Tricky a lot, A statements collection will be added to the top of 
    ' the new constructor that will add the mklextp.dll GAC after extracting
    ' it from the embedded resources.
    Private Function AddStatementsToConstrutor() As CodeStatementCollection

        Dim mkLib As String = "mx"
        Dim embededResourcesFileName As String = "exeMacroResources"

        AddStatementsToConstrutor = New CodeStatementCollection

        AddStatementsToConstrutor.Add(New CodeSnippetExpression("If Not IsAssemplyInGAC(""mklextp"") Then"))
        ' Extract mklextp.dll from resources.
        AddStatementsToConstrutor.Add(New CodeSnippetExpression("obj = ResourceManager.GetObject(" & """" & mkLib & """" & ")"))
        AddStatementsToConstrutor.Add(New CodeSnippetExpression("Dim mklibBytes As  Byte()"))
        AddStatementsToConstrutor.Add(New CodeSnippetExpression("mklibBytes = CType(obj, Byte())"))
        ' Create PMR folder in system folder if not exist then write mklextp.dll to it.
        AddStatementsToConstrutor.Add(New CodeSnippetExpression("Dim processingDir As String = System.Environment.SystemDirectory & ""\PMR"""))
        AddStatementsToConstrutor.Add(New CodeSnippetExpression("Dim isProcessingDirExists As Boolean = My.Computer.FileSystem.DirectoryExists(processingDir)"))
        AddStatementsToConstrutor.Add(New CodeSnippetExpression("If Not isProcessingDirExists Then My.Computer.FileSystem.CreateDirectory(processingDir)"))
        AddStatementsToConstrutor.Add(New CodeSnippetExpression("My.Computer.FileSystem.WriteAllBytes(Environment.SystemDirectory & ""\PMR\"" & ""mklextp.dll"", mklibBytes, False)"))
        ' Add mklextp.dll to GAC
        AddStatementsToConstrutor.Add(New CodeSnippetExpression("AddAssemplyToGAC(Environment.SystemDirectory & ""\PMR\"" & ""mklextp.dll"")"))
        AddStatementsToConstrutor.Add(New CodeSnippetExpression("End If"))

        Return AddStatementsToConstrutor

    End Function

    Public Function CodeDomIsAnotherInstanceRunning() As CodeMemberMethod
        Using executableMacroGen As New MacroCodeGenerator
            CodeDomIsAnotherInstanceRunning = executableMacroGen.CreateMethod("IsAnotherInstanceRunning")
            Dim returnValue As New CodeTypeReference
            returnValue.BaseType = "Boolean"
            CodeDomIsAnotherInstanceRunning.ReturnType = returnValue

            CodeDomIsAnotherInstanceRunning.Statements.Add(New CodeSnippetExpression("Dim ProcessCounts As Integer = System.Diagnostics.Process.GetProcessesByName(System.Diagnostics.Process.GetCurrentProcess.ProcessName).Length"))
            CodeDomIsAnotherInstanceRunning.Statements.Add(New CodeSnippetExpression(
                                          "If ProcessCounts > 1 Then"))
            CodeDomIsAnotherInstanceRunning.Statements.Add(New CodeSnippetExpression(
                                          "IsAnotherInstanceRunning = True"))
            CodeDomIsAnotherInstanceRunning.Statements.Add(New CodeSnippetExpression(
                                                      "End If"))
        End Using
    End Function

    Public Function CodeDomTmrStop_Tick() As CodeMemberMethod
        Using executableMacroGen As New MacroCodeGenerator
            Dim senderParameterType As Type = GetType(Global.System.Object)
            Dim eParameterType As Type = GetType(Global.System.EventArgs)
            Dim senderParameter As New CodeParameterDeclarationExpression(senderParameterType, "sender")
            Dim eParameter As New CodeParameterDeclarationExpression(eParameterType, "e")
            Dim TmrStopParameters As New CodeParameterDeclarationExpressionCollection
            TmrStopParameters.Add(senderParameter)
            TmrStopParameters.Add(eParameter)
            CodeDomTmrStop_Tick = executableMacroGen.CreateMethod("TmrStop_Tick", TmrStopParameters)
            'CTRL + ALT+ DEL
            CodeDomTmrStop_Tick.Statements.Add(New CodeSnippetExpression("If GetAsyncKeyState(&H12) And GetAsyncKeyState(&H11) And GetAsyncKeyState(&H2E) Then"))
            CodeDomTmrStop_Tick.Statements.Add(New CodeSnippetExpression("TmrStop.Enabled = False"))
            CodeDomTmrStop_Tick.Statements.Add(New CodeSnippetExpression("MessageBox.Show(""Playback Terminated"")"))
            CodeDomTmrStop_Tick.Statements.Add(New CodeSnippetExpression("TrayIcon.Visible = False"))
            CodeDomTmrStop_Tick.Statements.Add(New CodeSnippetExpression("TrayIcon.Dispose()"))
            CodeDomTmrStop_Tick.Statements.Add(New CodeSnippetExpression("Application.Exit()"))
            CodeDomTmrStop_Tick.Statements.Add(New CodeSnippetExpression("End"))
            CodeDomTmrStop_Tick.Statements.Add(New CodeSnippetExpression("End If"))

            'Ctrl + Shift + F2
            CodeDomTmrStop_Tick.Statements.Add(New CodeSnippetExpression("If GetAsyncKeyState(&H11) And GetAsyncKeyState(&H10) And GetAsyncKeyState(&H71) Then"))
            CodeDomTmrStop_Tick.Statements.Add(New CodeSnippetExpression("TmrStop.Enabled = False"))
            CodeDomTmrStop_Tick.Statements.Add(New CodeSnippetExpression("MessageBox.Show(""Playback Terminated"")"))
            CodeDomTmrStop_Tick.Statements.Add(New CodeSnippetExpression("TrayIcon.Visible = False"))
            CodeDomTmrStop_Tick.Statements.Add(New CodeSnippetExpression("TrayIcon.Dispose()"))
            CodeDomTmrStop_Tick.Statements.Add(New CodeSnippetExpression("Application.Exit()"))
            CodeDomTmrStop_Tick.Statements.Add(New CodeSnippetExpression("End"))
            CodeDomTmrStop_Tick.Statements.Add(New CodeSnippetExpression("End If"))

        End Using

    End Function

End Module
