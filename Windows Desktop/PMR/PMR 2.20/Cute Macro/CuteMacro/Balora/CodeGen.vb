'Balora version:0.1  .. 06/11/2010

Option Strict On
Option Explicit On

Imports System.IO
Imports System.CodeDom
Imports System.CodeDom.Compiler

'This CodeGen based on an example provided by MSDN & ondotnet & developer.com
'http://msdn.microsoft.com/en-us/library/system.codedom.compiler.codedomprovider.aspx
'http://ondotnet.com/pub/a/dotnet/2003/02/03/codedom.html
'http://www.developer.com/net/vb/article.php/2115721/Writing-Code-Generators-with-the-CodeDOM---Part-2.htm
Namespace Balora
    '//**************************************************************************************************************//'
    '*  *  *  *   *  *  *  *  *  *  *  *  MacroCodeGenerator Class *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
    '//**************************************************************************************************************//'
    Partial Friend Class CodeGen
        Implements IDisposable

#Region "Methods"

        Public Sub New()
        End Sub

        Public Sub New(ByVal nameSpaces As ArrayList)
            importedNameSpacesValue = nameSpaces
        End Sub

        ' This method add a namespace for you code file 
        ' and also add import statements using namespaces 
        ' in ImportedNameSpaces.
        Protected Friend Function InitializeNameSpaces(ByVal name As String) As CodeNamespace
            'Declare a new namespace called ExeMacro.
            Dim exeMacro As New CodeNamespace(name)
            'Get Namespaces to be imported from ImportedNameSpaces property
            For Each NameSpaceToBeImported As String In ImportedNameSpaces
                exeMacro.Imports.Add(New CodeNamespaceImport(NameSpaceToBeImported))
            Next
            Return exeMacro
        End Function

        ' Declare a variable and return CodeVariableDeclarationStatement 
        ' to add to your desired context.
        Protected Friend Function DeclareVariables(ByVal dataType As System.Type, ByVal name As String) As CodeVariableDeclarationStatement
            ' Get a CodeTypeReference for the Type 
            ' of the variable we are about to
            ' create.   This will allow us not to 
            ' have to get bogged down in the
            ' language specific details of specifying 
            ' data types.
            Dim tr As New CodeTypeReference(dataType)

            ' The CodeVariableDeclarationStatement 
            ' will allow us to not have to
            ' worry about such details as whether 
            ' the Data Type or the variable name
            ' comes first or whether or not a key 
            ' word such  as Dim is required.
            Dim Declaration As New CodeVariableDeclarationStatement(tr, name)

            ' The CodeObjectCreateExpression handles 
            ' all of the details for calling
            ' constructors.   In most cases, this 
            ' will be new, but sometimes it is New.
            ' At any rate, we don't want to  have to 
            ' worry about such details.
            Dim newStatement As New CodeObjectCreateExpression()

            ' Here we specify the object whose 
            ' constructor we want to invoke.
            newStatement.CreateType = tr

            ' Here we specify that variable will be 
            ' initialized by calling its constructor.
            Declaration.InitExpression = newStatement
            Return Declaration
        End Function

        'Add constructor to a class 
        Protected Friend Function AddConstructor(ByVal isPublic As Boolean) As CodeConstructor
            Dim Constructor As CodeConstructor = New CodeConstructor()
            If isPublic Then
                Constructor.Attributes = MemberAttributes.Public
            End If
            Return Constructor
        End Function

        Protected Friend Function CreateClass(ByVal name As String, ByVal baseClass As String, ByVal isPublic As Boolean) As CodeTypeDeclaration
            ' Create a new CodeTypeDeclaration object specifying the name of the class to be created.   
            Dim ctd As New CodeTypeDeclaration(name)

            ' Specify that this CodeType is a class as opposed to an enum or a struct
            ctd.IsClass = True

            If isPublic Then
                ctd.TypeAttributes = Reflection.TypeAttributes.Public
            End If

            ctd.BaseTypes.Add(baseClass)

            ' Specify that this class is public
            ctd.Attributes = MemberAttributes.[Public]

            ' Return our freshly created class
            Return ctd
        End Function

        ' Declare a new code entry point method. "Sub Main if VB.NET"
        Protected Friend Function CreateEntryPointMethod() As CodeEntryPointMethod
            Dim method As New CodeEntryPointMethod()

            ' Specify that this method will be both static and public
            method.Attributes = MemberAttributes.[Public] Or MemberAttributes.[Static]

            ' Return the freshly created method
            Return method
        End Function

        Protected Friend Function CreateMethod(
                                              ByVal methodName As String,
                                              Optional ByVal parameters As CodeParameterDeclarationExpressionCollection = Nothing
                                              ) As CodeMemberMethod
            'Example:
            'Dim testingMethod As CodeMemberMethod = CreatetMethod("MariamAndNoha")
            ' Add the code entry point method to the Members collection of the type.
            'macroClass.Members.Add(testingMethod)

            ' Declare a new CodeEntryPointMethod
            Dim method As New CodeMemberMethod()

            method.Name = methodName

            If Not parameters Is Nothing Then
                method.Parameters.AddRange(parameters)
            End If

            ' Specify that this method will be both static and public
            method.Attributes = MemberAttributes.[Private] 'Or MemberAttributes.[Static]

            ' Return the freshly created method
            Return method
        End Function

        Protected Friend Function AddDisposeMethod() As CodeMemberMethod

            ' Implement the Dispose method
            Dim DisposeMethod As CodeMemberMethod = CreateMethod("Dispose")

            DisposeMethod.Attributes = _
              MemberAttributes.Family Or MemberAttributes.Overloaded _
              Or MemberAttributes.Override

            DisposeMethod.Parameters.Add( _
              New CodeParameterDeclarationExpression( _
                GetType(Boolean), "disposing"))

            Dim Statement As CodeConditionStatement = _
              New CodeConditionStatement()
            Statement.Condition = _
              New CodeArgumentReferenceExpression("disposing")

            ' Ahmed Saad: 09/11/2010: You can extend this method 
            'in the future and de-commenting the following lines 
            'and pass the objects you to dispose .
            'Dim TrueStatement As CodeConditionStatement = _
            '  New CodeConditionStatement()
            'TrueStatement.Condition = _
            '  New CodeBinaryOperatorExpression( _
            '    New CodeArgumentReferenceExpression("MouseHookObject"), _
            '    CodeBinaryOperatorType.IdentityInequality, _
            '    New CodePrimitiveExpression(Nothing))


            'TrueStatement.TrueStatements.Add( _
            '  New CodeMethodInvokeExpression( _
            '    New CodeFieldReferenceExpression(Nothing, _
            '        "MouseHookObject"), "Dispose", New  _
            '                                  CodeExpression() {}))

            'Statement.TrueStatements.Add(TrueStatement)

            DisposeMethod.Statements.Add(Statement)

            DisposeMethod.Statements.Add( _
              New CodeMethodInvokeExpression( _
                  New CodeBaseReferenceExpression(), _
                  "Dispose", New CodeArgumentReferenceExpression() _
                  {New CodeArgumentReferenceExpression( _
                    "disposing")}))

            Return DisposeMethod

        End Function

        Private Function GetLanguageProvider(Optional ByVal selectedLanguage As String = "Visual Basic") As CodeDomProvider
            Select Case selectedLanguage
                Case "CSharp"
                    Return CodeDomProvider.CreateProvider("CSharp")
                Case "Visual Basic"
                    Return New VBCodeProvider
                Case Else
                    Return New VBCodeProvider
            End Select
        End Function

        ' Generate code file based in your settings.
        Public Function Generate(ByVal outputCodeFile As String,
                                 ByVal compileUnit As CodeCompileUnit) As Boolean
            'Dim provider As CodeDomProvider = getLanguageProvider("CSharp")
            Dim provider As VBCodeProvider = CType(GetLanguageProvider(), VBCodeProvider)
            Try
                ' Create an IndentedTextWriter, constructed with a StreamWriter to the source file.
                Dim macroCodeFileIndentedTextWriter As New IndentedTextWriter(New StreamWriter(outputCodeFile, False), "....")
                ' Generate source code using the code generator.
                provider.GenerateCodeFromCompileUnit(compileUnit, macroCodeFileIndentedTextWriter, New CodeGeneratorOptions())
                ' Close the output file.
                macroCodeFileIndentedTextWriter.Close()
            Catch ex As Exception
                Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error during compiling macro.")
                Return False
            End Try
            Return True
        End Function

        ' Compile and generate windows exe file .
        Public Function Compile() As Boolean
            Dim oCodeProvider As VBCodeProvider = New VBCodeProvider
            Dim oCParams As CompilerParameters = New CompilerParameters()
            Dim oCResults As CompilerResults = Nothing
            oCParams.GenerateExecutable = True
            oCParams.OutputAssembly = CompiledFilePath
            Dim compilerOptionsVar As String = "/target:winexe /win32icon:" & Chr(34) & CompiledFileIcon & Chr(34)
            oCParams.CompilerOptions = compilerOptionsVar
            'oCParams.MainClass = "codeguru.com.Form1"
            'Get Assemblies to be referenced from ReferencedAssemblies property
            For Each AssembliesToBeReferenced As String In ReferencedAssemblies
                oCParams.ReferencedAssemblies.Add(AssembliesToBeReferenced)
            Next
            oCParams.EmbeddedResources.Add(EmbeddedResources)
            Try
                oCResults = oCodeProvider.CompileAssemblyFromSource(oCParams, Code)
                ' Check for compile time errors
                If oCResults.Errors.Count <> 0 Then
                    CompilingErrors = oCResults.Errors
                    For Each cmpError As CodeDom.Compiler.CompilerError In CompilingErrors
                        Util.WriteToErrorLog(cmpError.ErrorText & " @ Line Number " & cmpError.Line, "",
                                        "Error during compiling the macro.")
                    Next
                    Throw New Exception("Compile Errors")
                Else
                    Return True
                End If
            Catch ex As Exception
                Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error during compiling")
                Return False
            End Try

        End Function

        Public Function CompileInputCodeFile() As Boolean
            Dim fileContents As String
            fileContents = My.Computer.FileSystem.ReadAllText(InputCodeFile)
            Code = fileContents
            If Compile() Then
                Return True
            End If
        End Function


        Public Function BuildBasicWindowsExecutable(ByVal MainNameSpaceName As String,
                                                    ByVal mainClass As String,
                                                    ByVal baseClass As String,
                                                    ByVal constructor As CodeConstructor,
                                                    ByVal methods As ArrayList,
                                                    ByVal fields As ArrayList) As CodeCompileUnit

            ' An Example of calling:
            ' Dim compileUnit As CodeCompileUnit = BuildBasicWindowsExecutable("ExeMacro",
            '                                                   "MacroClass",
            '                                           "ApplicationContext",
            '                                                 newConstructor,
            '                                                    methodsList,
            '                                                fieldsArrayList)
            ' "ExeMacro"           : NameSpace is called "ExeMacro".
            ' "MacroClass"         : Name of the Main class is "MacroClass".
            ' "ApplicationContext" : Base class of "MacroClass".
            ' newConstructor       : The constructor or "new" method of the class. (Type: CodeConstructor)
            ' methodsList          : An arraylist of CodeMemberMethod objects that represents the class methods.
            ' fieldsArrayList      : An arraylist of CodeTypeMember objects that represents the class fields.

            ' Return value is a CodeCompileUnit object that will be passed to "Generate" to build the whole ...
            ' Graph and generating the code file that will be saved as well to the path stored in OutputCodeFile.

            ' Create a new CodeCompileUnit to contain the program graph.
            Dim compileUnit As New CodeCompileUnit()
            ' Add the new namespace imports statements.
            Dim exeNameSpace As CodeNamespace = InitializeNameSpaces(MainNameSpaceName)
            'Add the new namespace to the compile unit.
            compileUnit.Namespaces.Add(exeNameSpace)
            ' Declare a new class with a name stored in mainClass.
            Dim macroClass As CodeTypeDeclaration = CreateClass(mainClass, baseClass, True)
            ' Add the new type to the namespace type collection.
            exeNameSpace.Types.Add(macroClass)
            ' Declare a new code entry point method. "Sub Main"
            Dim entryPoint As CodeEntryPointMethod = CreateEntryPointMethod()
            ' Pass the class to Run in order to display a Form
            entryPoint.Statements.Add(New CodeSnippetExpression("Application.Run(New " & mainClass & ")"))

            ' Add class fields.
            For Each classField As CodeTypeMember In fields
                classField.Attributes = MemberAttributes.Public
                macroClass.Members.Add(classField)
            Next

            ' Add constructor (new method) using arraylist.
            macroClass.Members.Add(constructor)

            ' Add remaining methods to the class via an arraylist.
            For Each method As CodeMemberMethod In methods
                macroClass.Members.Add(method)
            Next

            ' Add the code entry point method to the Members collection of the type.
            macroClass.Members.Add(entryPoint)
            ' Add Dispose method to the class.
            macroClass.Members.Add(AddDisposeMethod())

            Return compileUnit

        End Function

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
End Namespace
