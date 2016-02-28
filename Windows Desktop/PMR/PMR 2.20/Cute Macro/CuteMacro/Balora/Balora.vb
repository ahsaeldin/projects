'Balora v0.1  .. 06/11/2010

Option Strict On
Option Explicit On

#Region "Imported NameSpaces"
Imports System.CodeDom.Compiler
Imports System.CodeDom

#End Region

Namespace Balora
    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    '                                                   CodeGen Class                                                   
    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Partial Friend Class CodeGen

#Region "Properties"
        ' You pass import statements to this property like this
        ' exMacroCodeGenerator.ImportedNameSpaces.Add("System")
        Private importedNameSpacesValue As New ArrayList
        Public Property ImportedNameSpaces() As ArrayList
            Get
                Return importedNameSpacesValue
            End Get
            Set(ByVal value As ArrayList)
                importedNameSpacesValue = value
            End Set
        End Property

        ' Path to code file that will be compiled
        Private inputCodeFileValue As String
        Public Property InputCodeFile() As String
            Get
                Return inputCodeFileValue
            End Get
            Set(ByVal value As String)
                inputCodeFileValue = value
            End Set
        End Property

        ' Path to save the generated code file to.
        Private outputCodeFileValue As String
        Public Property OutputCodeFile() As String
            Get
                Return outputCodeFileValue
            End Get
            Set(ByVal value As String)
                outputCodeFileValue = value
            End Set
        End Property

        ' Returned compiling errors.
        Private compilingErrorsValue As CompilerErrorCollection
        Public Property CompilingErrors() As CompilerErrorCollection
            Get
                Return compilingErrorsValue
            End Get
            Set(ByVal Value As CompilerErrorCollection)
                compilingErrorsValue = Value
            End Set
        End Property

        ' Path to save to exe compiled file .
        Private compiledFilePathValue As String
        Public Property CompiledFilePath() As String
            Get
                Return compiledFilePathValue
            End Get
            Set(ByVal Value As String)
                compiledFilePathValue = Value
            End Set
        End Property

        ' You can pass code directly to this property then
        ' call compile
        Private codeValue As String
        Public Property Code() As String
            Get
                Return codeValue
            End Get
            Set(ByVal value As String)
                codeValue = value
            End Set
        End Property

        'Path to assemblies to be reference when compiling like this...
        ' exMacroCodeGenerator.ReferencedAssemblies.Add("System.Data.dll") or this...
        ' exMacroCodeGenerator.ReferencedAssemblies.Add("C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Drawing.dll")
        Private referencedAssembliesValue As New ArrayList
        Public Property ReferencedAssemblies() As ArrayList
            Get
                Return referencedAssembliesValue
            End Get
            Set(ByVal value As ArrayList)
                referencedAssembliesValue = value
            End Set
        End Property

        ' Path to resource file to be embedded with.
        Private embeddedResourcesValue As String
        Public Property EmbeddedResources() As String
            Get
                Return embeddedResourcesValue
            End Get
            Set(ByVal value As String)
                embeddedResourcesValue = value
            End Set
        End Property

        ' Path to icon for the compiled exe file. "Cool Prop"
        Private compiledFileIconValue As String
        Public Property CompiledFileIcon() As String
            Get
                Return compiledFileIconValue
            End Get
            Set(ByVal value As String)
                compiledFileIconValue = value
            End Set
        End Property

        ' Property to add a collection of various statements 
        ' at the top of the constructor (new method).
        Private newStatementsToConstructorValue As CodeStatementCollection
        Public Property NewStatementsToConstructor() As CodeStatementCollection
            Get
                Return newStatementsToConstructorValue
            End Get
            Set(ByVal value As CodeStatementCollection)
                newStatementsToConstructorValue = value
            End Set
        End Property
#End Region

    End Class

    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    '                                                   Resources Class                                                 
    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Partial Friend Class xResources

    End Class

    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    '                                                   MacroCodeGenerator Class                                        
    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Partial Friend Class MacroCodeGenerator
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

    End Class

    Partial Friend Class Util

    End Class
    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    '                                                   Balora Settings Class                                           
    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Partial Friend Class Settings
        ' Settings class for the whole Balora.

        ' Let Balora user customize Error log file name.
        Private Shared logFileNameValue As String = "PMR-ErrorLog.txt"
        Public Shared Property LogFileName() As String
            Get
                Return logFileNameValue
            End Get
            Set(ByVal value As String)
                logFileNameValue = value
            End Set
        End Property

        ' Enable/Disable error logging.
        Private Shared logErrorsValue As Boolean = True
        Public Shared Property LogErrors() As Boolean
            Get
                Return logErrorsValue
            End Get
            Set(ByVal value As Boolean)
                logErrorsValue = value
            End Set
        End Property

        ' Enable/Disable error logger windows.
        Private Shared showLogWindowValue As Boolean = True
        Public Shared Property ShowLogWindow() As Boolean
            Get
                Return showLogWindowValue
            End Get
            Set(ByVal value As Boolean)
                showLogWindowValue = value
            End Set
        End Property

    End Class

End Namespace
