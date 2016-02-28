'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   TraditionalCrud Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
''Use this class to connect directly to a SQLite database using Datler.

'Imports System.Reflection
'Imports System.Data.Common
'Imports Balora.Datler
'Imports Cacadore
''DatabaseProvider is a global compilation constant for all configuration 
''used to define both System dot connect and sys.data.sqlite in the project, 
''so i can swap using them as i want.
'#If DatabaseProvider = "Devart.Data.Sqlite" Then
'Imports System.Data
'#ElseIf DatabaseProvider = "System.Data.Sqlite" Then
'Imports System.Data.SQLite
'#End If

'    Public Class TraditionalCrud
'        Implements ICrud
'        Implements IDisposable

'#Region "Fields"
'#If DatabaseProvider = "Devart.Data.Sqlite" Then
'        Friend WithEvents CacaduDatabase As New DevartSQLiteDatabase
'#ElseIf DatabaseProvider = "System.Data.Sqlite" Then
'        Friend WithEvents CacaduDatabase As New SQLiteDatabase
'#End If

'#End Region

'#Region "Constructors"
'        Public Sub New()
'            Dim dbConn As DbConnection
'            Dim ConnectionFactory As ConnectionFactory
'#If DatabaseProvider = "Devart.Data.Sqlite" Then
'            ConnectionFactory = New ConnectionFactory() 'Database Type Change: Change: Only one remainings.(3/4)
'#ElseIf DatabaseProvider = "System.Data.Sqlite" Then
'            ConnectionFactory = New ConnectionFactory() 'Database Type Change: Change: Only one remainings.(3/4)
'#End If
'            Dim databasePlace As String = Balora.Util.GetExecutablePath & "tectonic"
'            'Dim databasePlace As String = "G:\Visual Studio Chronicles\VisualStudio 2010\Tools for VS.NET2010\NorthWind\Northwind 'sqlite 3'.db"
'            Try
'                'If you didn't specify DatabaseProviderName  and/or DatabaseType Properties, defaults values used
'                ConnectionFactory.DatabaseProviderName = GlobalEnums.Dbprovider.System_Data_SQLite
'                dbConn = ConnectionFactory.CreateConnection()
'                'dbConn.ConnectionString = "Data Source=" & databasePlace & ";Password=SADFasdfq$%twe%tSAER#WQ$%RTYW$&DRSZDFGSER^W#$%^@$^#45"
'                dbConn.ConnectionString = "Data Source=" & databasePlace & ";"
'                'Here i'm assigning the connection created by factory to CacaduDatabase 
'                CacaduDatabase.DatabaseConnection = CType(dbConn, SQLite.SQLiteConnection) 'Convert to the appropriate connection class Change: Done.(4/4)
'                CacaduDatabase.DatabaseConnection.Open()
'            Catch ex As Exception
'                Balora.Alerter.RaiseErrorUp("", ex, True)
'                dbConn = Nothing
'            End Try
'            'How to perform all this using only ExecuteUsingDataAdapter method
'            'Dim command As String = "SELECT * FROM TableNameHere"
'            'Dim returnedDataAdapter As DbDataAdapter = CacaduDatabase.ExecuteUsingDataAdapter(command)
'        End Sub
'#End Region

'#Region "Methods"
'        Public Function Query(ByVal cacadoreObj As ICrudable, ByVal ParamArray args() As String) As Object Implements ICrud.Query
'            Dim command As String = "SELECT * FROM tasks_tsk"
'            Dim returnedDataAdapter As DbDataAdapter = CacaduDatabase.ExecuteUsingDataAdapter(command)
'            'MainForm.DataGridView1.DataSource = CacaduDatabase.DataSet.Tables("Table") 'OR...
'            'MainForm.DataGridView1.DataSource = CacaduDatabase.DataSet.Tables(0)
'            Return Nothing
'        End Function

'        Public Function Insert(ByVal cacadoreObj As ICrudable, ByVal ParamArray args() As String) As Boolean Implements ICrud.Insert

'            Dim insertParameters As New Dictionary(Of String, Object)
'            Dim propertiesDictionary As New Dictionary(Of String, Object)

'            For Each p As PropertyInfo In cacadoreObj.GetType.GetProperties
'                propertiesDictionary.Add(p.Name, p.GetValue(cacadoreObj, Nothing))
'            Next

'            'Property TaskId As String
'            'Property Properties As TaskProperties
'            'ReadOnly Property ActionsCollection As ActionsCollection
'            'Property Triggers As TriggersCollection
'            'Property Errors As TaskErrors
'            'Property Variables As TaskVariables
'            'ReadOnly Property History As System.Collections.Generic.List(Of IHistory.HistoryEntry)
'            'Property State As TaskState
'            'Property ParentGroupId As String
'            'Property WaitBetweenActions As Boolean

'            Select Case cacadoreObj.GetType.Name
'                Case "Task"
'                    Dim insertCommands As New List(Of String)

'                    Dim tskState As TaskState = CType(propertiesDictionary("State"), TaskState)
'                    Dim tskPro As TaskProperties = CType(propertiesDictionary("Properties"), TaskProperties)

'                    insertParameters.Add("@id_tsk", propertiesDictionary("TaskId"))
'                    insertParameters.Add("@group_id_tsk", propertiesDictionary("ParentGroupId"))
'                    insertParameters.Add("@is_wait_between_actions_tsk", propertiesDictionary("WaitBetweenActions"))
'                    insertParameters.Add("@date_last_time_run_tsk", tskPro.LastTimeRun)
'                    insertParameters.Add("@last_time_run_result_tsk", tskState.TaskExecutionState)
'                    insertParameters.Add("@date_next_time_to_be_run_tsk", tskPro.NextTimeToBeRun)
'                    insertParameters.Add("@name_tsk", tskPro.TaskName)
'                    insertParameters.Add("@triggers_when_tsk", tskPro.TriggersWhen)
'                    insertParameters.Add("@is_enables_tsk", tskPro.IsEnabled)

'                    ' Dim ahmedvalue = cacadoreObj.GetType.GetProperty("TaskId").GetValue(cacadoreObj, Nothing)
'                    insertCommands.Add("INSERT INTO " +
'                                       "tasks_tsk" +
'                                       "(id_tsk,group_id_tsk,is_wait_between_actions_tsk,date_last_time_run_tsk,last_time_run_result_tsk,date_next_time_to_be_run_tsk,name_tsk,triggers_when_tsk,is_enables_tsk) VALUES (@id_tsk,@group_id_tsk,@is_wait_between_actions_tsk,@date_last_time_run_tsk,@last_time_run_result_tsk,@date_next_time_to_be_run_tsk,@name_tsk,@triggers_when_tsk,@is_enables_tsk)")

'                    'CacaduDatabase.Insert(insertCommands, insertParameters)
'                    CacaduDatabase.ExecuteUsingDataAdapter("INSERT INTO " +
'                                       "tasks_tsk" +
'                                       "(id_tsk,group_id_tsk,is_wait_between_actions_tsk,date_last_time_run_tsk,last_time_run_result_tsk,date_next_time_to_be_run_tsk,name_tsk,triggers_when_tsk,is_enables_tsk) VALUES (@id_tsk,@group_id_tsk,@is_wait_between_actions_tsk,@date_last_time_run_tsk,@last_time_run_result_tsk,@date_next_time_to_be_run_tsk,@name_tsk,@triggers_when_tsk,@is_enables_tsk)")
'            End Select

'            propertiesDictionary = Nothing

'            Return Nothing
'        End Function

'        Public Function Update(ByVal cacadoreObj As ICrudable, ByVal ParamArray args() As String) As Boolean Implements ICrud.Update
'            Return Nothing
'        End Function

'        Public Function Delete(ByVal cacadoreObj As ICrudable, ByVal ParamArray args() As String) As Boolean Implements ICrud.Delete
'            Return Nothing
'        End Function
'#End Region

'#Region "Events"
'        Private Sub CacaduDatabase_FillError(ByVal sender As Object, ByVal e As System.Data.FillErrorEventArgs) Handles CacaduDatabase.FillError

'        End Sub

'        Private Sub CacaduDatabase_RowUpdated(ByVal sender As Object, ByVal e As System.Data.Common.RowUpdatedEventArgs) Handles CacaduDatabase.RowUpdated

'        End Sub

'        Private Sub CacaduDatabase_RowUpdating(ByVal sender As Object, ByVal e As System.Data.Common.RowUpdatingEventArgs) Handles CacaduDatabase.RowUpdating

'        End Sub
'#End Region

'#Region "IDisposable Support"
'        Private disposedValue As Boolean ' To detect redundant calls

'        ' IDisposable
'        Protected Overridable Sub Dispose(disposing As Boolean)
'            If Not Me.disposedValue Then
'                If disposing Then
'                    ' TODO: dispose managed state (managed objects).
'                End If

'                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
'                ' TODO: set large fields to null.
'            End If
'            Me.disposedValue = True
'        End Sub

'        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
'        'Protected Overrides Sub Finalize()
'        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
'        '    Dispose(False)
'        '    MyBase.Finalize()
'        'End Sub

'        ' This code added by Visual Basic to correctly implement the disposable pattern.
'        Public Sub Dispose() Implements IDisposable.Dispose
'            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
'            Dispose(True)
'            GC.SuppressFinalize(Me)
'        End Sub
'#End Region

'    End Class

