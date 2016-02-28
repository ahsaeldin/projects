Public Class PathsHelper
    Inherits BaloraBase

#Region "Path Variable Methods"

    Public Shared Function GetPathVariableList() As List(Of String)
        Dim _environmentVariables = Environment.GetEnvironmentVariables

        Dim pathVarList As List(Of String) = Nothing

        For Each varDicEnt As DictionaryEntry In _environmentVariables
            If varDicEnt.Key.ToString = "Path" Or varDicEnt.Key.ToString = "path" Then
                pathVarList = Split(varDicEnt.Value.ToString, ";").ToList()
            End If
        Next

        Return pathVarList

    End Function

    ''' <summary>
    ''' Gets a list of executables paths in the folders listed in "Path" environment variable.
    ''' </summary><returns></returns>
    Public Shared Function GetExecutablesPathsListInPathEnVarFolders() As List(Of String)

        Dim pathVarList As List(Of String) = GetPathVariableList()

        Dim executalblesList As New List(Of String)

        For Each _path In pathVarList
            Dim tempList As System.Collections.Generic.IEnumerable(Of String) = Nothing
            Try
                tempList = IO.Directory.GetFiles(_path, "*.exe", IO.SearchOption.TopDirectoryOnly).ToList
                executalblesList.AddRange(tempList)
            Catch ex As System.IO.DirectoryNotFoundException
                Balora.Alerter.REP(ex.Message, ex, False, True)
            End Try
        Next

        Return executalblesList

    End Function

    ''' <summary>
    ''' Gets a sorted list of executables names in the folders 
    ''' listed in "Path" environment variable along with their paths.
    ''' </summary><returns></returns>
    Public Shared Function GetExecutablesNamesAndPathsInPathEnvVarFolders() As SortedList(Of String, String)

        Dim exePathsList As List(Of String) = GetExecutablesPathsListInPathEnVarFolders()

        Dim executalblesNamesList As New SortedList(Of String, String)

        For Each path In exePathsList
            Dim exeName As String = IO.Path.GetFileNameWithoutExtension(path)
            If Not executalblesNamesList.ContainsKey(exeName) Then
                executalblesNamesList.Add(exeName, path)
            End If
        Next

        Return executalblesNamesList

    End Function

    ''' <summary>
    ''' Save a sorted list of executables names in the folders 
    ''' listed in "Path" environment variable along with their paths.
    ''' </summary>
    Public Shared Sub WriteExecutablesListInPathEnVarFoldersToFile(fileName As String)

        'How to use
        'WriteExecutablesListInPathEnVarFoldersToFile("c:\SortedListOfExecutablesNamesInTheFoldersListedInPath.txt")

        Dim executalblesList = GetExecutablesNamesAndPathsInPathEnvVarFolders()

        Dim writer As New IO.StreamWriter(fileName)

        For Each kvp In executalblesList
            writer.WriteLine(kvp.Key)
        Next

        writer.Close()

    End Sub

#End Region

#Region "General Path Methods"

    Public Shared Function GetExecutablePath() As String
        'Note that you have another 2 overloaded methods to get current executing directory. 
        Dim startupPath As String = GetCurrentExecutingDirectory()
        Return startupPath
    End Function

    Public Overloads Shared Function GetCurrentExecutingDirectory() As String
        Dim executingPath As String = AppDomain.CurrentDomain.BaseDirectory
        'Useful in unit testing
        'http://stackoverflow.com/questions/52797/c-how-do-i-get-the-path-of-the-assembly-the-code-is-in
        CheckBackSlashAtEnd(executingPath)
        Return executingPath
    End Function


    Public Overloads Shared Function GetCurrentExecutingDirectory(ByVal obj As Type) As String
        'Useful in unit testing
        'http://stackoverflow.com/questions/52797/c-how-do-i-get-the-path-of-the-assembly-the-code-is-in
        'EX:
        'Friend WithEvents CacaduDatabase As New SQLiteDatabase
        'Dim databasePlace As String = Util.GetCurrentExecutingDirectory(CacaduDatabase.GetType) & "UnitTestingDatabase.db"

        'get the full location of the assembly with obj in it
        Dim fullPath As String = System.Reflection.Assembly.GetAssembly(obj).Location
        'get the folder that's in
        Dim theDirectory As String = IO.Path.GetDirectoryName(fullPath)
        CheckBackSlashAtEnd(theDirectory)
        Return theDirectory
    End Function

    Public Shared Sub CheckBackSlashAtEnd(ByRef stringToBeChecked As String)
        If Not Strings.Right(stringToBeChecked, 1) = "\" Then
            stringToBeChecked = stringToBeChecked + "\"
        End If
    End Sub

    Public Shared Function GetAssemplyPath(ByVal type As Object) As String
        '(Useful to get the path of references)
        'Pass an object of a type exists in the assembly you want to 
        'get its path, and it will return the location of the assembly.

        'ex : Util.GetAssemplyPath(New System.Collections.Generic.HashSet(Of String)) 'To get the path for "System.Core.dll”
        Return (type.GetType().Assembly.Location())
    End Function

    Public Shared Function GetMachineConfigPath() As String
        Return System.Runtime.InteropServices.RuntimeEnvironment.SystemConfigurationFile
    End Function

    ''' <summary>
    ''' Gets the common application data. 
    ''' </summary>
    ''' <returns>
    ''' The directory that serves as a common repository for application-specific data that is used by all users.
    ''' </returns>
    '''  <remarks>
    ''' Check this for more info http://msdn.microsoft.com/en-us/library/system.environment.specialfolder.aspx
    ''' </remarks>
    Public Shared Function GetCommonApplicationData() As String
        Dim path As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)
        CheckBackSlashAtEnd(path)
        Return path
    End Function

    ''' <summary>
    ''' Gets the local application data.
    ''' </summary>
    ''' <returns>
    ''' The directory that serves as a common repository for application-specific data that is used by the current, non-roaming user.
    ''' </returns>
    ''' <remarks>
    ''' Check this for more info http://msdn.microsoft.com/en-us/library/system.environment.specialfolder.aspx
    ''' </remarks>
    Public Shared Function GetLocalApplicationData() As String
        Dim path As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        CheckBackSlashAtEnd(path)
        Return path
    End Function

    Public Shared Function GetDefaultBrowserPath() As String
        Dim key As String = "htmlfile\shell\open\command"
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(key, False)
        ' get default browser path
        Return DirectCast(registryKey.GetValue(Nothing, Nothing), String).Split(""""c)(1)
    End Function
#End Region

End Class
