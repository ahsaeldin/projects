Option Strict On
Option Explicit On


Imports System.CodeDom
Imports System.ComponentModel

'This NameSpace based on an example provided by MSDN & ondotnet
'http://msdn.microsoft.com/en-us/library/system.codedom.compiler.codedomprovider.aspx
'http://ondotnet.com/pub/a/dotnet/2003/02/03/codedom.html
'http://www.developer.com/net/vb/article.php/2115721/Writing-Code-Generators-with-the-CodeDOM---Part-2.htm

'//**********************************************************************************************//'
'*  *  *  *  *  *  *  MacroCodeGenerator Class *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'//**********************************************************************************************//'
'Generate the internal code of the executable macro file
Friend Class MacroCodeGenerator
    Inherits CodeGen

    'Note: ToolTip text cannot exceed 63 characters . 
    Private exeSystrayIconTooltipValue As String
    Public Property ExeSystrayIconTooltip() As String
        Get
            Return exeSystrayIconTooltipValue
        End Get
        Set(ByVal value As String)
            exeSystrayIconTooltipValue = value
        End Set
    End Property

    ' Pass the name of the systray icon in the embedded resource file.
    ' Note: If you don't pass icon name in the embedded resource file adding systray icon igonored .
    Private systrayIconNameInResourcesValue As String
    Public Property SystrayIconNameInResources() As String
        Get
            Return systrayIconNameInResourcesValue
        End Get
        Set(ByVal value As String)
            systrayIconNameInResourcesValue = value
        End Set
    End Property

    ' A way for the user to customize the name of the core method of the exe file.
    Private coreMethodNameValue As String
    Public Property CoreMethodName() As String
        Get
            Return coreMethodNameValue
        End Get
        Set(ByVal value As String)
            coreMethodNameValue = value
        End Set
    End Property

    ' A perfect example in how to  call this method is in ExecutableMacroHelperFunctions.CompileMacroToExE.
    Public Function GenerateExecutableMacro(ByVal coreMethods As ArrayList) As Result
        'coreMethods CodeMemberMethod is an ArrayList of CodeMemberMethod objects
        'that represents the methods contains what you really want the exe to do.

        ' ArrayList contains all method objects you want to add to the exe class.
        Dim methodsList As New ArrayList
        Dim generatingExecutableMacroResult As New Result

        ' Create "new" or constructor for the Exe Class.
        Dim newConstructor As CodeConstructor = CreateExEClassConstructor()


        ' TODO : Write any method you want to add to the exe here just as exitMenuClickMethod below this line then add to "methodsList".


        ' _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _ _  _  _  _  _  _  _  _  _  _  _ _  _  _  _ _  _
        '/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ / \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ / \/ \/ \/ / \/ \
        ' Start: Add methods to the main class as below.
        '---------------------------------------------------------------------------------------------
        Dim senderParameterType As Type = GetType(Global.System.Object)
        Dim eParameterType As Type = GetType(Global.System.EventArgs)
        Dim senderParameter As New CodeParameterDeclarationExpression(senderParameterType, "sender")
        Dim eParameter As New CodeParameterDeclarationExpression(eParameterType, "e")
        Dim exitMenuParameters As New CodeParameterDeclarationExpressionCollection
        exitMenuParameters.Add(senderParameter)
        exitMenuParameters.Add(eParameter)
        Dim exitMenuClickMethod As CodeMemberMethod = CreateMethod("ExitMenu_Click", exitMenuParameters)

        exitMenuClickMethod.Statements.Add(New CodeSnippetExpression("TrayIcon.Visible = False"))
        exitMenuClickMethod.Statements.Add(New CodeSnippetExpression("TrayIcon.Dispose()"))
        exitMenuClickMethod.Statements.Add(New CodeSnippetExpression("Application.Exit()"))
        methodsList.Add(exitMenuClickMethod)
        ' ""coreMethod" parameter is the core method for...
        ' the purpose that executable file created...
        ' and it will be created and add by the calling...
        ' object for "GenerateExecutableMacro" method
        For Each method As CodeMemberMethod In coreMethods
            methodsList.Add(method)
        Next
        '---------------------------------------------------------------------------------------------
        ' End: Then pass methodsList to BuildBasicWindowsExecutable methods.
        '\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/



        ' _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _  _ _  _  _  _  _  _  _  _  _  _  _ _  _  _  _ _  _
        '/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ / \/ \/ \/ \/ \/ \/ \/ \/ \/ \/ / \/ \/ \/ / \/ \
        ' Start: Add fields to the main class as below.
        '---------------------------------------------------------------------------------------------------------------------------------------------
        Dim trayIconVar As CodeTypeMember = New CodeSnippetTypeMember("Private WithEvents TrayIcon As New System.Windows.Forms.NotifyIcon")
        Dim contextMenuStripVar As CodeTypeMember = New CodeSnippetTypeMember("Private WithEvents CMS As New System.Windows.Forms.ContextMenuStrip")
        Dim toolStripMenuItemVar As CodeTypeMember = New CodeSnippetTypeMember("Private WithEvents ExitMenu As New ToolStripMenuItem(""Exit"")")
        Dim tmrStopVar As CodeTypeMember = New CodeSnippetTypeMember("Private TmrStop As New Timer")
        Dim getAsyncKeyStateVar As CodeTypeMember = New CodeSnippetTypeMember(
                                           "Public Declare Function GetAsyncKeyState Lib ""user32"" (ByVal vKey As Integer) As Integer")

        Dim fieldsArrayList As New ArrayList
        fieldsArrayList.AddRange(New CodeTypeMember() {getAsyncKeyStateVar, trayIconVar, contextMenuStripVar, toolStripMenuItemVar, tmrStopVar})
        '---------------------------------------------------------------------------------------------------------------------------------------------
        ' End: Then pass fieldsArrayList to BuildBasicWindowsExecutable methods.
        '\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/\_/


        ' Now we have constructor, methods list and class fields along with 
        ' Namespace Name "ExeMacro" and "MacroClass" as the main class name
        ' in which has "ApplicationContext" as a base class. 
        Dim compileUnit As CodeCompileUnit = BuildBasicWindowsExecutable("ExeMacro",
                                                                       "MacroClass",
                                                               "ApplicationContext",
                                                                     newConstructor,
                                                                        methodsList,
                                                                    fieldsArrayList)

        ' Generate code file and save it to OutputCodeFile.
        Dim generatingResult As Boolean = Generate(OutputCodeFile, compileUnit)

        If Not generatingResult Then
            generatingExecutableMacroResult.Message = "Error during generating executable macro."
            generatingExecutableMacroResult.Successed = False
            Return generatingExecutableMacroResult
        End If

        ' Start compiling process now!
        Dim compilingResult As Boolean = CompileInputCodeFile()

        If compilingResult = False Then
            generatingExecutableMacroResult.Successed = False
            generatingExecutableMacroResult.Message = "Error compiling macro data."
            Return generatingExecutableMacroResult
        End If

        generatingExecutableMacroResult.Successed = True
        Return generatingExecutableMacroResult

    End Function

    Private Function CreateExEClassConstructor() As CodeConstructor

        ' Get embedded resource file without extension.
        Dim embededResourcesFileName As String = System.IO.Path.GetFileNameWithoutExtension(EmbeddedResources)

        ' Creating class constructor.
        Dim newConstructor As CodeConstructor = AddConstructor(isPublic:=True)

        ' Adding statements to constructor.
        newConstructor.Statements.Add(New CodeSnippetExpression(
                                      "Dim icon As System.Drawing.Icon"))
        newConstructor.Statements.Add(New CodeSnippetExpression(
                                      "Dim ResourceManager As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager(" & """" & embededResourcesFileName & """" & ", System.Reflection.Assembly.GetExecutingAssembly())"))
        If SystrayIconNameInResources <> "" Then
            newConstructor.Statements.Add(New CodeSnippetExpression(
                                          "Dim obj As Object"))

            If Not NewStatementsToConstructor Is Nothing Then
                'Very Important:------
                ' Add statements that will add mklextp.dll to GAC @ the top of the (new) .
                newConstructor.Statements.AddRange(NewStatementsToConstructor)
            End If

            newConstructor.Statements.Add(New CodeSnippetExpression("obj = ResourceManager.GetObject(" & """" & SystrayIconNameInResources & """" & ")"))
            newConstructor.Statements.Add(New CodeSnippetExpression("icon = CType(obj, System.Drawing.Icon)"))
            newConstructor.Statements.Add(New CodeSnippetExpression("TrayIcon.Icon = icon"))
            newConstructor.Statements.Add(New CodeSnippetExpression("TrayIcon.Visible = True"))
            newConstructor.Statements.Add(New CodeSnippetExpression("TrayIcon.Text = " & """" & ExeSystrayIconTooltip & """"))
            newConstructor.Statements.Add(New CodeSnippetExpression("CMS.Items.Add(ExitMenu)"))
            newConstructor.Statements.Add(New CodeSnippetExpression("TrayIcon.ContextMenuStrip = CMS"))
            newConstructor.Statements.Add(New CodeSnippetExpression("AddHandler ExitMenu.Click, AddressOf ExitMenu_Click"))
            newConstructor.Statements.Add(New CodeSnippetExpression("AddHandler TmrStop.Tick, AddressOf TmrStop_Tick"))
            newConstructor.Statements.Add(New CodeSnippetExpression("TmrStop.Enabled = True"))
        End If

        newConstructor.Statements.Add(New CodeSnippetExpression(CoreMethodName))
        Return newConstructor

    End Function
End Class
