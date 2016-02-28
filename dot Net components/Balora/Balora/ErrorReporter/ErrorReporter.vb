Imports System.Windows.Forms
Imports System.ComponentModel

'commented next line as no need to avoid obfuscation of ErrorReporter 
'<System.Reflection.ObfuscationAttribute(Feature:="renaming", ApplyToMembers:=True)>

Namespace ErrorReporter
    Public Class ErrorReporter
        Inherits Mailer

#Region "Properties"
        ''' <summary>
        ''' When exception thrown, you pass exception object to this property
        ''' </summary>
        Private exceptionObjValue As Exception
        Public Property ExceptionObject() As Exception
            Get
                Return exceptionObjValue
            End Get
            Set(ByVal value As Exception)
                exceptionObjValue = value
                Dim sF As StackFrame = Util.GetLastFrameInExceptionStackTarce(exceptionObjValue)
                ErrorFileName = sF.GetFileName
                ErrorMethodName = sF.GetMethod().Name
                ErrorLineNumber = sF.GetFileLineNumber()
                ErrorColumnNumber = ErrorColumnNumber
            End Set
        End Property

        ''' <summary>
        ''' Set to true if unhandled exception.
        ''' </summary>
        ''' <remarks></remarks>
        Private isUnHandledExceptionValue As Boolean
        Public Property IsUnHandledException() As Boolean
            Get
                Return isUnHandledExceptionValue
            End Get
            Set(ByVal value As Boolean)
                isUnHandledExceptionValue = value
            End Set
        End Property

        ''' <summary>
        ''' Get or set the name of the file in which error raised.  
        ''' </summary>
        ''' <remarks></remarks>
        Private errorFileNameValue As String
        Public Property ErrorFileName() As String
            Get
                Return errorFileNameValue
            End Get
            Set(ByVal value As String)
                errorFileNameValue = value
            End Set
        End Property

        ''' <summary>
        ''' Get or set the method name in which the error raised from.
        ''' </summary>
        ''' <remarks></remarks>
        Private errorMethodNameValue As String
        Public Property ErrorMethodName() As String
            Get
                Return errorMethodNameValue
            End Get
            Set(ByVal value As String)
                errorMethodNameValue = value
            End Set
        End Property

        ''' <summary>
        ''' Get or set the line number which generated the error.
        ''' </summary>
        ''' <remarks></remarks>
        Private errorLineNumberValue As Integer
        Public Property ErrorLineNumber() As Integer
            Get
                Return errorLineNumberValue
            End Get
            Set(ByVal value As Integer)
                errorLineNumberValue = value
            End Set
        End Property

        ''' <summary>
        ''' Get or set the column number which generated the error.
        ''' </summary>
        ''' <remarks></remarks>
        Private errorColNumberValue As Integer
        Public Property ErrorColumnNumber() As Integer
            Get
                Return errorColNumberValue
            End Get
            Set(ByVal value As Integer)
                errorColNumberValue = value
            End Set
        End Property

        '''<summary>
        ''' Comments added by user and as body of the message.
        ''' </summary>
        Private userCommentsValue As String
        Public Property UserComments() As String
            Get
                Return userCommentsValue
            End Get
            Set(ByVal value As String)
                userCommentsValue = value
            End Set
        End Property

        ''' <summary>
        ''' Returns Exception source like "mscorlib"
        ''' </summary>
        Public ReadOnly Property ExceptionSource() As String
            Get
                Return ExceptionObject.Source
            End Get
        End Property

        ''' <summary>
        ''' Returns full exception type like this System.ArgumentException
        ''' </summary>
        Public ReadOnly Property FullExceptionType() As String
            Get
                Return ExceptionObject.GetType.FullName
            End Get
        End Property

        ''' <summary>
        ''' Returns exception message like "The path is not of a legal form"
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ExceptionMessage() As String
            Get
                Dim exceptionMessageStr As String = Nothing
                If Strings.Right(ExceptionObject.Message, 1) = "." Then
                    exceptionMessageStr = Strings.Left(ExceptionObject.Message, ExceptionObject.Message.Length - 1)
                End If
                Return exceptionMessageStr
            End Get
        End Property

        ''' <summary>
        ''' Returns product version like 1.0.0.0
        ''' </summary>
        Public ReadOnly Property AppVersion() As String
            Get
                Return Application.ProductVersion
            End Get
        End Property

        ''' <summary>
        ''' Returns build date of the assembly
        ''' </summary>
        Public ReadOnly Property BuildDate() As DateTime
            Get
                Return Util.GetBuildDate(Application.ExecutablePath)
            End Get
        End Property

        ''' <summary>
        ''' Returns Windows version like 
        ''' Microsoft Windows 7 Ultimate - Version 6.1.7601.65536 [Service Pack 1]
        ''' </summary>
        Private windowsVersionValue As String
        Public ReadOnly Property FullWindowsVersion() As String
            Get
                Return Util.GetFullWindowsVersion
            End Get
        End Property

        ''' <summary>
        ''' Returns exe name of the app.
        ''' </summary>
        Public ReadOnly Property ExeName() As String
            Get
                Dim exeFileName As String
                Try
                    exeFileName = LCase(System.IO.Path.GetFileName(Application.ExecutablePath))
                    Return exeFileName
                Catch ex As ArgumentException
                    Alerter.REP("Error in the parameter passed to " & Chr(34) & "System.IO.Path.GetFileName" & Chr(34) & " method.", ex, True)
                    Return ""
                End Try
            End Get
        End Property

        ''' <summary>
        ''' Returns exe path of the app.
        ''' </summary>
        Public ReadOnly Property ExePath() As String
            Get
                Return Application.ExecutablePath
            End Get
        End Property

        ''' <summary>
        ''' Returns user IP
        ''' </summary>
        ''' <remarks></remarks>
        Private userRealIpValue As String
        Public Property UserRealIp() As String
            'Call Util.GetMyRealIP(uri) and pass it to UserRealIp 
            'http://www.perfect-macro-recorder.com/getip.php
            Get
                Return userRealIpValue
            End Get
            Set(ByVal value As String)
                userRealIpValue = value
            End Set
        End Property

        Public ReadOnly Property AvailablePhysicalMemory() As ULong
            Get
                Return Util.GetAvailablePhysicalMemory
            End Get
        End Property

        Public ReadOnly Property TotalPhysicalMemory() As ULong
            Get
                Return Util.GetTotalPhysicalMemory
            End Get
        End Property

        ''' <summary>
        ''' Returns Loaded Assemblies 
        ''' </summary>
        ''' <remarks>
        '''   How to call
        '''   Dim errRep As New ErrorReporter
        '''   Dim dic As New Dictionary(Of String, String)
        '''   dic = errRep.LoadedAssemplies
        '''   For Each kvp As KeyValuePair(Of String, String) In dic
        '''       Console.WriteLine("Key = {0}, Value = {1}", _
        '''       kvp.Key, kvp.Value)
        '''   Next kvp
        ''' </remarks>
        Public ReadOnly Property LoadedAssemblies() As Dictionary(Of String, String)

            Get
                Return Util.GetLoadedAssemblies
            End Get

        End Property

        ''' <summary>
        ''' Returns a stack of stack trace
        ''' </summary>
        ''' <remarks>
        '''Get stacktrace string value then 
        '''split into lines then
        '''push them in a stack 
        '''Note: We use stack to list them from bottom to top, {from the end point to start point},
        '''so I can know the line where the exception occurs @ first entry the table of stack trace.  
        ''' </remarks>
        Public ReadOnly Property StackTrace() As Stack(Of String)
            Get
                Dim stackTraceStrings() As String = ExceptionObject.StackTrace.Split(New [Char]() {CChar(vbNewLine)})
                'clear all thing from previous calls.
                Dim stackTraceValue As New Stack(Of String)
                For Each sts As String In stackTraceStrings
                    stackTraceValue.Push(sts)
                Next sts
                Return stackTraceValue
            End Get
        End Property
#End Region

#Region "Methods"
        ''' <summary>
        ''' Send error report by calling base method "Send"
        ''' in the base class "Mailer".
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub Send()
            Clipboard.SetText(ExceptionObject.StackTrace)
            MyBase.Send()
        End Sub
#End Region

    End Class
End Namespace