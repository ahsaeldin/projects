'Balora version:0.1  .. 06/11/2010
' Contains several functions to assist Balora in many usual tasks.

Option Strict On
Option Explicit On

#Region "Imported NameSpaces"
Imports System.IO
Imports System.Reflection

#End Region

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'                                                   Util Module                                                     
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Namespace Balora

#Region "Structures"
    'This Structure is more than awesome. It enables me to tell...
    'the calling method of the method that return "Result" a message...
    'and the if succeeded 
    Public Structure Result
        Dim Message As String
        Dim Successed As Boolean
    End Structure
#End Region

    Partial Friend Class Util

#Region "Functions"

        Public Shared Function WriteToErrorLog(ByVal msg As String, ByVal stkTrace As String, ByVal title As String) As Boolean

            If Balora.Settings.LogErrors = False Then Exit Function
            Dim errorLogFileName As String = Balora.Settings.LogFileName

            Static errorLogWindow As Form
            Static errorListBox As ListBox
            If Balora.Settings.ShowLogWindow Then
                If Debugger.IsAttached Then
                    If errorLogWindow Is Nothing Then
                        errorLogWindow = New Form
                        errorLogWindow.Show()
                        errorLogWindow.TopMost = True
                        errorLogWindow.Top = Screen.PrimaryScreen.Bounds.Height - errorLogWindow.Height
                        errorLogWindow.Left = 0
                        errorLogWindow.Width = Screen.PrimaryScreen.Bounds.Width - 300
                        errorListBox = New ListBox
                        errorListBox.Dock = DockStyle.Fill
                        errorLogWindow.Controls.Add(errorListBox)
                    End If
                End If
            Else
                errorLogWindow = Nothing
                errorListBox = Nothing
            End If

            Try
                'check the file
                Dim fs As FileStream = New FileStream(Application.StartupPath & "\" & errorLogFileName, FileMode.OpenOrCreate, FileAccess.ReadWrite)
                Dim s As StreamWriter = New StreamWriter(fs)
                s.Close()
                fs.Close()
                'log it
                Dim fs1 As FileStream = New FileStream(Application.StartupPath & "\" & errorLogFileName, FileMode.Append, FileAccess.Write)
                Dim s1 As StreamWriter = New StreamWriter(fs1)
                s1.Write("Title: " & title & " ")
                s1.Write("(Date/Time: " & DateTime.Now.ToString() & ")" & vbCrLf)
                s1.Write("" & vbCrLf)
                s1.Write("Error Description: " & msg & vbCrLf)
                s1.Write("" & vbCrLf)
                s1.Write("StackTrace:" & stkTrace & vbCrLf)
                s1.Write("============================================" & vbCrLf)
                s1.Close()
                fs1.Close()
                errorListBox.Items.Add(title & " " & msg)
                'If Debugger.IsAttached Then
                'Process.Start(errorLogFileName)
                'End If
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Shared Function IsAssemplyInGAC(ByVal assemplyName As String) As Boolean
            IsAssemplyInGAC = False
            Dim GACDirPath As String = ""
            Dim WinDirPath As String = Environment.GetEnvironmentVariable("windir")
            If Strings.Right(WinDirPath, 1) <> "\" Then
                GACDirPath = WinDirPath & "\assembly"
            Else
                GACDirPath = WinDirPath & "assembly"
            End If
            For Each directory As String In FileIO.FileSystem.GetDirectories(GACDirPath, FileIO.SearchOption.SearchAllSubDirectories)
                If directory.Contains(assemplyName) Then
                    IsAssemplyInGAC = True
                End If
            Next
            Return IsAssemplyInGAC
        End Function

        Public Shared Sub AddAssemplyToGAC(ByVal assemblyPath As String)
            Try
                Dim EntServ As New System.EnterpriseServices.Internal.Publish
                EntServ.GacInstall(assemblyPath)
            Catch ex As Exception
                Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error while adding assembly to GAC.")
            End Try
        End Sub

        Public Shared Sub RemoveAssemplyFromGAC(ByVal assemblyPath As String)
            'Note: http://stackoverflow.com/questions/45729/what-path-should-i-pass-as-an-assemblypath-parameter-to-the-publish-gacremove-fun
            Try
                Dim EntServ As New System.EnterpriseServices.Internal.Publish
                EntServ.GacRemove(assemblyPath)
            Catch ex As Exception
                Util.WriteToErrorLog(ex.Message, ex.StackTrace, "Error while removing assembly from GAC.")
            End Try
        End Sub

        Public Shared Function IsAnotherInstanceRunning() As Boolean
            'copy of IsAnotherInstanceRunning_3shanVista in Cute Macro
            Dim ProcessCounts As Integer = Process.GetProcessesByName _
                (Process.GetCurrentProcess.ProcessName).Length

            If ProcessCounts > 1 Then
                IsAnotherInstanceRunning = True
            End If
        End Function

        Public Shared Function GetExecutablePath() As String
            Return Application.StartupPath
        End Function

        Public Shared Function GetAssemplyPath(ByVal type As Object) As String
            '(Useful to get the path of references)
            'Pass an object of a type exists in the assembly you want to 
            'get its path, and it will return the location of the assembly.

            'ex : Balora.Util.GetAssemplyPath(New System.Collections.Generic.HashSet(Of String)) 'To get the path for "System.Core.dll”
            Return (type.GetType().Assembly.Location())
        End Function

#End Region

    End Class

End Namespace
