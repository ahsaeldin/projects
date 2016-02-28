'So simple, a class derived from "Database" class, and you 
'can use for simple CRUD operations for a sqlite database 
'created or managed by dotconnect for sqlite provider http://www.devart.com/dotconnect/sqlite/

'We add a conditional compilation #if statement so I can handle the switch between
'Provider reference assemblies as easily as possible.

'DatabaseProvider is a global compilation constant for all configuration 
'used to define both Devart dot connect and sys.data.sqlite in the project, 
'so i can swap using them as i want.
#If DatabaseProvider = "Devart.Data.Sqlite" Then
Imports System.ComponentModel
Imports Devart.Data
Imports System.IO
Imports System.Data.Common
Imports Devart.Data.SQLite

Namespace Datler
    <LicenseProvider(GetType(Balora.Lixcer.RTALP))>
    Public Class DevartSQLiteDatabase
        Inherits Database
        Implements IDisposable

#Region "Properties"
        ''' <summary>
        ''' Holds the database connection used through out all the methods in the class.
        ''' </summary>
        ''' <remarks></remarks>
        Private dbConnectionValue As SQLiteConnection
        Public Overloads Property DatabaseConnection() As SQLiteConnection
            Get
                Return dbConnectionValue
            End Get
            Set(ByVal value As SQLiteConnection)
                dbConnectionValue = value
            End Set
        End Property

        Private WithEvents dataAdapterValue As New SQLiteDataAdapter
        Public Shadows Property DataAdapter() As SQLiteDataAdapter
            Get
                Return dataAdapterValue
            End Get
            Set(ByVal value As SQLiteDataAdapter)
                dataAdapterValue = value
            End Set
        End Property

        Private dataSetValue As New DataSet
        Public Property DataSet() As DataSet
            Get
                Return dataSetValue
            End Get
            Set(ByVal value As DataSet)
                dataSetValue = value
            End Set
        End Property
#End Region

#Region "Methods"

        ''' <summary>
        ''' Checks if a database table is already exists
        ''' </summary>
        ''' <param name="dbConn">Inherited class from DbConnection.</param>
        ''' <param name="tableName">Table name to be checked.</param>
        ''' <returns>true if exists</returns>
        ''' <remarks></remarks>
        Public Overrides Function IsTableExists(ByVal dbConn As DbConnection, ByVal tableName As String) As Boolean
            'Check here for more info about using GetSchema
            'http://www.codeproject.com/KB/database/checkdatabase.aspx?msg=2809530
            'http://www.devart.com/dotconnect/mysql/docs/MetaData.html
            Dim reConnectionResult As Boolean = ReConnect(dbConn)
            If reConnectionResult = False Then Return Nothing
            Try
                Dim restrictions(3) As String
                restrictions(2) = tableName
                Dim returnedDataTable As DataTable = dbConn.GetSchema("Tables", restrictions)
                If returnedDataTable.Rows.Count = 0 Then
                    'Table does not exist
                    Return False
                Else
                    'Table exists
                    Return True
                End If
            Catch ex As Exception
                Alerter.REP("", ex, True)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Create SQLite Table using a specified table name and "CREATE TABLE" command
        ''' </summary>
        ''' <param name="creationCommands">CREATE TABLE command</param>
        ''' <returns>true if succeed</returns>
        ''' <remarks>
        ''' more info about CREATE TABLE:- 
        ''' http://www.sqlite.org/lang_createtable.html
        ''' http://www.razorsql.com/docs/sqlite_create_table.html
        ''' </remarks>
        Public Overrides Function CreateTable(ByVal creationCommands As List(Of String)) As Boolean
            'For ex:
            ' ''''''''''''''''''''''''''''''''How To Create Table''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'CacaduDatabase.CreateTable("CREATE TABLE IF NOT EXISTS " & "TableNameHere" & "(id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT, testing1 REAL,testing2 INTEGER, testing3 NUMERIC, image BLOB);")
            ' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Return ExecuteCommand(creationCommands)
        End Function

        ''' <summary>
        ''' Manually executing SELECT statement (connection model)
        ''' </summary>
        ''' <param name="selectCommand">Select command text</param>
        ''' <param name="selectParameters">Parameters, if any</param>
        ''' <returns>Datatable of retrieved data</returns>
        Public Overrides Function Query(ByVal selectCommand As String, Optional ByVal selectParameters As Dictionary(Of String, Object) = Nothing) As DataTable
            ''''''''''''''''''''''''''''''''How To Query''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Dim selectCommand As String = Nothing
            'Dim selectParameters As New Dictionary(Of String, Object)
            'selectCommand = "SELECT * FROM TableNameHere WHERE testing2=@parameter1"

            'selectParameters.Add("@parameter1", "8")

            'Dim queryDatatable As DataTable = CacaduDatabase.Query(selectCommand, selectParameters)
            ''Retrieve Datatable
            'For Each row As DataRow In queryDatatable.Rows
            '    Dim values(row.ItemArray.Count - 1) As Object
            '    values = row.ItemArray
            '    For Each obj In values
            '        Balora.Util.SendToOutputWindow(obj.ToString)
            '    Next
            'Next
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Return ExecuteQuery(selectCommand, selectParameters)
        End Function

        ''' <summary>
        ''' Manually Execute Update command (connection model)
        ''' </summary>
        Public Overrides Function Update(ByVal updateCommands As List(Of String), Optional ByVal updateParameters As Dictionary(Of String, Object) = Nothing) As Boolean
            ''''''''''''''''''''''''''''''''How To Update'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Dim updateCommands As New List(Of String)
            'updateCommands.Add("UPDATE " + "TableNameHere" + " SET testing2='153284'")
            'CacaduDatabase.Update(updateCommands)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Return ExecuteCommand(updateCommands, updateParameters)
        End Function

        ''' <summary>
        ''' Manually Insert new rows into specified table. (connection model)
        '''  </summary>
        ''' <param name="inserationCommands">contains one or more insert statements to be added.</param>
        ''' <param name="insertParameters">contains one or more insert parameters</param>
        ''' <returns>True if succeed</returns>
        ''' <remarks></remarks>
        Public Overrides Function Insert(ByVal inserationCommands As List(Of String), Optional ByVal insertParameters As Dictionary(Of String, Object) = Nothing) As Boolean
            'Supported database by SQLite 3
            'http://www.sqlite.org/datatype3.html
            ''''''''''''''''''''''''''''''''How To Insert'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Dim insertCommands As New List(Of String)
            'Dim insertParameters As New Dictionary(Of String, Object)
            'insertCommands.Add("INSERT INTO " + "TableNameHere" + " (title, testing1,testing2,testing3,image) VALUES (@parameter1,@parameter2,@parameter3,@parameter4,@parameter5)")

            'insertParameters.Add("@parameter1", "parameter1 Value")
            'insertParameters.Add("@parameter2", 1)
            'insertParameters.Add("@parameter3", 2)
            'insertParameters.Add("@parameter4", 3)
            'insertParameters.Add("@parameter5", CacaduDatabase.ObjectToBlob(MainForm.PictureBox1.Image))

            'CacaduDatabase.Insert(insertCommands, insertParameters)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Return ExecuteCommand(inserationCommands, insertParameters)
        End Function

        ''' <summary>
        ''' Manually Delete rows, drop tables (connection model)
        ''' </summary>
        Public Overrides Function Delete(ByVal deleteCommands As List(Of String), Optional ByVal deleteParameters As Dictionary(Of String, Object) = Nothing) As Boolean
            ''''''''''''''''''''''''''''''''How to drop table?'''''''''
            'Dim deleteCommands As New List(Of String)
            'deleteCommands.Add("DROP TABLE " & "TableNameHere" & ";")

            'CacaduDatabase.Delete(deleteCommands)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ''''''''''''''''''''''''''''''''How to delete all rows?'''''
            'Dim deleteCommands As New List(Of String)
            'deleteCommands.Add("DELETE FROM " & "TableNameHere" & ";")

            'CacaduDatabase.Delete(deleteCommands)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ''''''''''''''''''''''''''''''''How to delete a specified row?'''''''''''''''''''''''''''''''
            'Dim deleteCommands As New List(Of String)
            'Dim deleteParameters As New Dictionary(Of String, Object)
            'deleteCommands.Add("DELETE FROM " & "TableNameHere WHERE ColumnNameHere=@parameter1" & ";")

            'deleteParameters.Add("@parameter1", 3)
            'CacaduDatabase.Delete(deleteCommands, deleteParameters)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Return ExecuteCommand(deleteCommands, deleteParameters)
        End Function

        ''' <summary>
        ''' Reconnect if connection closed.
        ''' </summary>
        ''' <param name="dbConn">Connection object to be reopened.</param>
        ''' <returns>True if succeed</returns>
        Private Function ReConnect(ByVal dbConn As DbConnection) As Boolean
            With dbConn
                If .State = ConnectionState.Closed Then
                    Try
                        .Open()
                    Catch ex As Exception
                        Alerter.REP("", ex, True)
                        Return False
                    End Try
                End If
            End With
            Return True
        End Function

        ''' <summary>
        ''' Read data in a SQLreader retrieved by a SELECT statement. (connection model)
        ''' </summary>
        Private Function ReadData(ByVal SQLreader As SQLiteDataReader) As DataTable
            Dim queryDataTable As New DataTable
            While SQLreader.Read()
                'Note you can get value by column title :SQLreader("title")
                Dim values(SQLreader.FieldCount - 1) As Object
                SQLreader.GetValues(values)
                If queryDataTable.Columns.Count = 0 Then
                    For i = 0 To SQLreader.FieldCount - 1
                        queryDataTable.Columns.Add(SQLreader.GetName(i), values(i).GetType)
                    Next
                End If
                queryDataTable.Rows.Add(values)
            End While
            Return queryDataTable
        End Function

        ''' <summary>
        ''' Execute commands more easily using SQLiteDataAdapter. (connectionless model)
        ''' </summary>
        ''' <param name="command">Command to be executed</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function ExecuteUsingDataAdapter(ByVal command As String) As DbDataAdapter
            'Õ Ï  ›Â„ ﬂÌ›  ”‰œ «·ﬁÌ„ «·„‰«”»… ··Œ’«∆’ «·À·«À «·«ŒÌ—…° „‰ «·ÃÌœ Ãœ« «” Ì⁄«»
            'Ê«· Ì  ﬂÊ‰ ﬂ«· «·Ì: ⁄‰œ„«  ‰ÊÌ  ÕœÌÀ «·»Ì«‰«  ≈·Ï DataAdapter ÿ—Ìﬁ… ⁄„· ﬂ«∆‰ «·„ÕÊ·
            '»«· Õﬁﬁ „‰ DataAdapter ”ÌﬁÊ„ ﬂ«∆‰ «·„ÕÊ· °Update() „’œ— »Ì«‰«  »«” œ⁄«¡ «·ÿ—Ìﬁ…
            '„‰ ”Ã·«  ﬂ«∆‰ «·»Ì«‰«  DataRow Ê«· «»⁄… ·ﬂ· ﬂ«∆‰ ”Ã· RowState «·Œ«’Ì…
            '”Ì „  ‰›Ì– Ã„·… «·«” ⁄·«„ ›Ì «·Œ«’Ì… Modified ›«‰ ﬂ«‰  °DataSource
            '”Ì „  ‰›Ì– «·«„— «· «»⁄ ··Œ«’Ì…
            '„‰ ﬂ «»  —ﬂÌ «·⁄”Ì—Ì

            'Check how to use below
            Dim reConnectionResult As Boolean = ReConnect(DatabaseConnection)
            If reConnectionResult = False Then Return Nothing
            AddHandler DataAdapter.FillError, AddressOf dataAdapterValue_FillError
            AddHandler Me.DataAdapter.RowUpdating, AddressOf dataAdapterValue_RowUpdating
            AddHandler Me.DataAdapter.RowUpdated, AddressOf dataAdapterValue_RowUpdated
            Dim sqlCommand As New SQLiteCommand(command)
            sqlCommand.Connection = DatabaseConnection
            Try
                If command.Contains("SELECT") Then
                    DataAdapter.SelectCommand = sqlCommand
                    'The Fill method has an overload that accepts one parameter for the DataSet only.  
                    'In that case, the table created has a default name of "table1" for the first table.
                    'The number will be incremented (table2, table3, ..., tableN) for each table added to 
                    'the(dataset) where the table name was not specified in the Fill method.
                    DataAdapter.Fill(dataSetValue)
                ElseIf command.Contains("INSERT") Then
                    DataAdapter.InsertCommand = sqlCommand
                    DataAdapter.Update(dataSetValue)
                ElseIf command.Contains("UPDATE") Then
                    DataAdapter.UpdateCommand = sqlCommand
                    DataAdapter.Update(dataSetValue)
                ElseIf command.Contains("DELETE") Then
                    DataAdapter.DeleteCommand = sqlCommand
                    DataAdapter.Update(dataSetValue)
                End If
                Return DataAdapter
            Catch ex As Exception
                Alerter.REP("", ex, True)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Manually Execute a SELECT statement without DataAdapter. (connection model)
        ''' </summary>
        ''' <param name="selectCommand"></param>
        ''' <param name="selectParameters"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function ExecuteQuery(ByVal selectCommand As String, Optional ByVal selectParameters As Dictionary(Of String, Object) = Nothing) As DataTable
            Dim SQLreader As SQLiteDataReader
            Dim reConnectionResult As Boolean = ReConnect(DatabaseConnection)
            If reConnectionResult = False Then Return Nothing
            Try
                Dim queryDataTable As New DataTable
                Using SQLcommand As New SQLiteCommand(dbConnectionValue.ConnectionString, dbConnectionValue)
                    SQLcommand.CommandText = selectCommand
                    If Not IsNothing(selectParameters) Then
                        For Each kvp As KeyValuePair(Of String, Object) In selectParameters
                            'This line performs 2 things
                            '   1.Add a parameter
                            '   2.Specify its datatype.
                            SQLcommand.Parameters.AddWithValue(kvp.Key, kvp.Value)
                        Next
                    End If
                    SQLreader = SQLcommand.ExecuteReader()
                End Using
                Dim dataTable As DataTable = ReadData(SQLreader)
                dbConnectionValue.Close()
                Return dataTable
            Catch ex As Exception
                Alerter.REP("", ex, True)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Manually Execute Update,Insert,Delete without DataAdapter. (connection model)
        ''' </summary>
        Private Function ExecuteCommand(ByVal commands As List(Of String), Optional ByVal commandParameters As Dictionary(Of String, Object) = Nothing) As Boolean
            Dim reConnectionResult As Boolean = ReConnect(DatabaseConnection)
            If reConnectionResult = False Then Return Nothing
            Try
                Using mytransaction As SQLiteTransaction = dbConnectionValue.BeginTransaction()
                    Using SQLcommand As New SQLiteCommand(dbConnectionValue.ConnectionString, dbConnectionValue)
                        For Each command As String In commands
                            SQLcommand.CommandText = command
                            If Not IsNothing(commandParameters) Then
                                For Each kvp As KeyValuePair(Of String, Object) In commandParameters
                                    'This line performs 2 things
                                    '   1.Add a parameter
                                    '   2.Specify its datatype.
                                    SQLcommand.Parameters.AddWithValue(kvp.Key, kvp.Value).DbType = CSystemTypeToDbType(kvp.Value.GetType())
                                Next
                            End If
                            SQLcommand.ExecuteNonQuery()
                        Next
                        SQLcommand.Dispose()
                    End Using
                    mytransaction.Commit()
                End Using
                dbConnectionValue.Close()
            Catch ex As Exception
                Alerter.REP("", ex, True)
                Return False
            End Try
            Return True
        End Function

        ''' <summary>
        ''' Checks if malformed by sending integrity_check pragma.
        ''' </summary>
        ''' <param name="connection">The connection.</param><returns></returns>
        Public Shared Function CheckIfMalformed(connection As DbConnection) As String
            Return DevartSQLiteDatabase.SendPragma(connection, "PRAGMA integrity_check")
        End Function

        ''' <summary>
        ''' Sends a pragma and receive the results.
        ''' </summary>
        ''' <param name="pragma">The pragma.</param><returns></returns>
        Public Function SendPragma(pragma As String) As String

            Return _privatePragma(dbConnectionValue, pragma)

        End Function

        ''' <summary>
        ''' Sends the pragma.
        ''' </summary>
        ''' <param name="connection">The connection to send the pragma.</param>
        ''' <param name="pragma">The pragma.</param><returns></returns>
        Public Shared Function SendPragma(connection As DbConnection, pragma As String) As String

            Return _privatePragma(CType(connection, SQLiteConnection), pragma)

        End Function

        ''' <summary>
        ''' contains the core lines of sending pragms.
        ''' </summary>
        ''' <param name="connection">The connection.</param>
        ''' <param name="pragma">The pragma.</param><returns></returns>
        Private Shared Function _privatePragma(connection As SQLiteConnection, pragma As String) As String

            Try
                If connection.State = ConnectionState.Closed Then
                    connection.Open()
                End If

            Catch ex As Exception

                'If failed to open database because it is malformed, no need to make malformation test :D
                If ex.Message.Contains("The database disk image is malformed") Then

                    Return "The database disk image is malformed"

                Else

                    'Alerter.REP("", ex, True)
                    Return ""

                End If

            End Try


            Dim pragmaCommand As SQLite.SQLiteCommand = connection.CreateCommand
            pragmaCommand.CommandText = pragma

            Try

                Return CStr(pragmaCommand.ExecuteScalar)

            Catch ex As Exception

                'Alerter.REP("", ex, True)
                Return ""

            End Try

        End Function

        ''' <summary>
        ''' Backup Database "Nooooooooot implemented Yet"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function BackupDatabase() As String
            'Using the SQLite Online Backup API
            'http://www.sqlite.org/backup.html
            'Online Backup API
            'http://www.sqlite.org/c3ref/backup_finish.html
            'SQLite Backup API
            'http://sqlite.phxsoftware.com/forums/p/2359/9382.aspx#9382
            'Using the Backup API
            'http://sqlite.phxsoftware.com/forums/p/2403/9511.aspx
            'You can give him encrypted database file or...
            'encrypted sql dump of the database.
            Return ""
        End Function

        ''' <summary>
        ''' Restore Database "Nooooooooot implemented Yet"
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function RestoreDatabase() As String
            'if SQL statements or the database file itself,
            'you must check the integrity and healthiness of the file
            'before dump or overwrite your database.
            Return ""
        End Function
#End Region

#Region "Events"
        'These events works only upon calling ExecuteUsingDataAdapter (connectionless model)
        Public Event FillError(ByVal sender As Object, ByVal e As System.Data.FillErrorEventArgs)
        Public Event RowUpdating(ByVal sender As Object, ByVal e As System.Data.Common.RowUpdatingEventArgs)
        Public Event RowUpdated(ByVal sender As Object, ByVal e As System.Data.Common.RowUpdatedEventArgs)

        Private Sub dataAdapterValue_FillError(ByVal sender As Object, ByVal e As System.Data.FillErrorEventArgs)
            RaiseEvent FillError(sender, e)
        End Sub

        Private Sub dataAdapterValue_RowUpdating(ByVal sender As Object, ByVal e As System.Data.Common.RowUpdatingEventArgs)
            RaiseEvent RowUpdating(sender, e)
        End Sub

        Private Sub dataAdapterValue_RowUpdated(ByVal sender As Object, ByVal e As System.Data.Common.RowUpdatedEventArgs)
            RaiseEvent RowUpdated(sender, e)
        End Sub
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
                dbConnectionValue.Dispose()
                dataAdapterValue.Dispose()
                dataSetValue.Dispose()
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


#Region "HowTOUse"
        ''' <summary>
        ''' This is not a method it is just a place for collecting all possible examples in one place.
        ''' </summary>
        ''' <remarks>'Do not uncomment, just use for a reference.</remarks>
        Private Sub HowToUse()
            'How to connect to another database or using another SQLite provider?
            '   1.Change the type of "CacaduDatabase" to the appropriate class or...
            '       leave it as is if you will just use another provider
            '   2.Change databaseType using DbType parameter of ConnectionFactory
            '       ===ConnectionFactory = New ConnectionFactory(CType(My.Settings.DatabaseType, DbType))
            '   3.If you will just change the provider type, you can use "DatabaseProviderName" name
            '   3.Get the new connection using CreateConnection method
            '   4.Convert this line to the new type
            '   5.CacaduDatabase.DatabaseConnection = CType(dbConn, SQLite.SQLiteConnection) 'Database Change: Only change this.(2/2)
            ''''''''''''''''''''''''''''''''How To Connect to a database''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Dim dbConn As DbConnection
            'Dim ConnectionFactory As ConnectionFactory
            'Dim databasePlace As String = Util.GetExecutablePath & "tectonic"
            'ConnectionFactory = New ConnectionFactory(DbType.SQLite, Dbprovider.DevartSQlite) 'Database Type Change: Change: Only one remainings.(3/4)
            'Try
            '    dbConn = ConnectionFactory.CreateConnection()
            '    dbConn.ConnectionString = "Data Source=" & databasePlace & ";"
            '    'Here i'm assigning the connection created by factory to CacaduDatabase 
            '    CacaduDatabase.DatabaseConnection = CType(dbConn, SQLite.SQLiteConnection) 'Convert to the appropriate connection class Change: Done.(4/4)
            'Catch ex As Exception
            '    Alerter.BoradCastExcpetion(ex)
            '    dbConn = Nothing
            'End Try
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ''''''''''''''''''''''''''''''''How To Create Table'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '    Dim creationCommands As New List(Of String)
            'creationCommands.Add("CREATE TABLE IF NOT EXISTS " & "TableNameHere" & "(id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT, testing1 REAL,testing2 INTEGER, testing3 NUMERIC, image BLOB);")
            'CacaduDatabase.CreateTable(creationCommands)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ''''''''''''''''''''''''''''''''How To Query''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Dim selectCommand As String = Nothing
            'Dim selectParameters As New Dictionary(Of String, Object)
            'selectCommand = "SELECT image FROM TableNameHere WHERE testing2=@parameter1"

            'selectParameters.Add("@parameter1", "153284")

            ''Retrieve Datatable
            'Dim queryDatatable As DataTable = CacaduDatabase.Query(selectCommand, selectParameters)
            'For Each row As DataRow In queryDatatable.Rows
            '    Dim values(row.ItemArray.Count - 1) As Object
            '    values = row.ItemArray
            '    For Each obj In values
            '        Dim img As Image = CType(CacaduDatabase.BlobToObject(CType(obj, Byte()), True), Image)
            '        img.Save("c:\image1.png")
            '        Balora.Util.SendToOutputWindow(obj.ToString)
            '    Next
            'Next
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ''''''''''''''''''''''''''''''''How To Insert'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '    Dim insertCommands As New List(Of String)
            '    Dim insertParameters As New Dictionary(Of String, Object)
            'insertCommands.Add("INSERT INTO " + "TableNameHere" + " (title, testing1,testing2,testing3,image) VALUES (@parameter1,@parameter2,@parameter3,@parameter4,@parameter5)")

            'insertParameters.Add("@parameter1", "parameter1 Value")
            'insertParameters.Add("@parameter2", 6)
            'insertParameters.Add("@parameter3", 8)
            'insertParameters.Add("@parameter4", 9)
            'insertParameters.Add("@parameter5", CacaduDatabase.ObjectToBlob(MainForm.PictureBox1.Image))

            'CacaduDatabase.Insert(insertCommands, insertParameters)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ''''''''''''''''''''''''''''''''How To Update'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Dim updateCommands As New List(Of String)
            'updateCommands.Add("UPDATE " + "TableNameHere" + " SET testing2='153284'")
            'CacaduDatabase.Update(updateCommands)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            ''''''''''''''''''''''''''''''''How To Drop table'''''''''
            'Dim deleteCommands As New List(Of String)
            'deleteCommands.Add("DROP TABLE " & "TableNameHere" & ";")

            'CacaduDatabase.Delete(deleteCommands)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ''''''''''''''''''''''''''''''''How To Delete All Rows'''''''''''''''''''''''''''''''''
            'Dim deleteCommands As New List(Of String)
            'Dim deleteParameters As New Dictionary(Of String, Object)
            'deleteCommands.Add("DELETE FROM " & "TableNameHere WHERE testing3=@parameter1" & ";")

            'deleteParameters.Add("@parameter1", 3)
            'CacaduDatabase.Delete(deleteCommands, deleteParameters)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            ''''''''''''''''''''''''''''''''How to perform all this using only ExecuteUsingDataAdapter method''''
            '   Dim command As String = "SELECT * FROM TableNameHere"
            '   Dim returnedDataAdapter As DbDataAdapter = CacaduDatabase.ExecuteUsingDataAdapter(command)
            '   'MainForm.DataGridView1.DataSource = CacaduDatabase.DataSet.Tables("Table") 'OR...
            '   MainForm.DataGridView1.DataSource = CacaduDatabase.DataSet.Tables(0)
            'And in on an CellEndEdit or any similar events in DataGridView update your database...
            '   Dim command As String = "UPDATE TableNameHere SET TITLE='yy'"
            '   CacaduDatabase.ExecuteUsingDataAdapter(command)
            ''''''''''''''''''''''''''''''''How to delete using ExecuteUsingDataAdapter
            'Note that nothing is done if nothing already done to the DataSet'''''''''''''''''''''''''''''''''''
            'CacaduDatabase.DataSet.Tables(0).Rows(0).Delete()
            'Dim command As String = "DELETE FROM TableNameHere WHERE TITLE='titlehere';"
            'CacaduDatabase.ExecuteUsingDataAdapter(command)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        End Sub
#End Region
    End Class
End Namespace
#End If