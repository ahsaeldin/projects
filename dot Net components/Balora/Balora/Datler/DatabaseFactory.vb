'Factory for creating  Classes of type Database, it doesn't create "databases",
'It created a classes that handle CRUD operations

'It has 2 constructors, one for creating the default database class, and
'the other for creating by specifying database type via DbType Enum declared
'in ConnectionFactory.vb

'The returned database class has its own connection string and can handle
'basic Crud operations

'How to support a new "Database" class:-
'1.Add the new database name to DbType Enum declared in ConnectionFactory.vb 
'2.Add a new case in "CreateDatabase" of the select statement.

'How to add a new provider for that used by a new "database" class for an existing database "like devart"
'1.Specify the provider name in the #If statement
'2.Add it just like below under the old provider in the specified case statement that handle this database.

Imports System.Data.Common
Imports System.ComponentModel
Imports System.Xml.Serialization

'DatabaseProvider is a global compilation constant for all configuration 
'used to define both Devart dot connect and sys.data.sqlite in the project, 
'so i can swap using them as i want.
#If DatabaseProvider = "Devart.Data.Sqlite" Then
Imports Devart.Data
Imports Devart.Data.SQLite
#ElseIf DatabaseProvider = "System.Data.Sqlite" Then
Imports System.Data.SQLite
#End If

Namespace Datler
    Friend Class DatabaseFactory
        Inherits BaloraBase

#Region "Constructors"
        Public Sub New(ByVal DatabaseType As DbType)
            'Set DatabaseType property 
            databaseTypeValue = DatabaseType
        End Sub
#End Region

#Region "Properties"
        ''' <summary>
        ''' Manually set the database type value, must be called if needed before "CreateDatabase"
        ''' </summary>
        ''' <remarks></remarks>
        Private databaseTypeValue As DbType
        Public Property DatabaseType() As DbType
            Get
                Return databaseTypeValue
            End Get
            Set(ByVal value As DbType)
                databaseTypeValue = value
            End Set
        End Property
#End Region

#Region "Methods"

        'Just a stub to hide the logic from obfuscators
        Public Function CreateDatabase(ByVal connectionString As String) As Database
            Return _createDatabase(connectionString)
        End Function

        ''' <summary>
        ''' You give me a connection string, I give you a class of type "database" 
        ''' or its derived in which you can use for simple CRUD operations.
        ''' </summary>
        Private Function _createDatabase(ByVal connectionString As String) As Database
            'how to create a database
            'Dim  dbFacroty As DatabaseFactory = New DatabaseFactory(0)
            'Dim connString As String = "Data Source=" & "c:\mydata.db3" & ";"
            'Dim cacaduSQLiteDatabase As Database = dbFacroty.CreateDatabase(connString)

            Dim db As Database = Nothing

            Select Case databaseTypeValue

                Case DbType.SQLite

#If DatabaseProvider = "Devart.Data.Sqlite" Then
                    db = New DevartSQLiteDatabase
                    Dim dbConn As DbConnection = CreateConnectionByDbConnection("System.Data.SQLite", connectionString)
                    If Not dbConn Is Nothing Then 'If connection created successfully.
                        Using dbConn
                            dbConn.Close()
                            Dim newSQLiteDatabase As New DevartSQLiteDatabase
                            newSQLiteDatabase.DatabaseConnection = CType(dbConn, SQLiteConnection)
                            RaiseEvent NewDatabaseCreated(CType(newSQLiteDatabase, Database))
                            Return newSQLiteDatabase
                        End Using
                    Else
                        Return Nothing
                    End If
#ElseIf DatabaseProvider = "System.Data.Sqlite" Then
                    db = New SQLiteDatabase
                    Dim dbConn As DbConnection = CreateConnectionByDbConnection("System.Data.SQLite", connectionString)
                    If Not dbConn Is Nothing Then 'If connection created successfully.
                        Using dbConn
                            dbConn.Close()
                            Dim newSQLiteDatabase As New SQLiteDatabase
                            newSQLiteDatabase.DatabaseConnection = CType(dbConn, SQLite.SQLiteConnection)
                            RaiseEvent NewDatabaseCreated(CType(newSQLiteDatabase, Database))
                            Return newSQLiteDatabase
                        End Using
                    Else
                        Return Nothing
                    End If
#End If
                Case DbType.MySQL '"Just as an example"

                    'db = New MySQLDatabase 'A new class that will be created by me and inherited from database.
                    'Dim dbConn As DbConnection = CreateConnectionByDbConnection("My SQL PROVIDER HERE", connectionString)
                    'If Not dbConn Is Nothing Then 'If connection created successfully.
                    '   Using dbConn
                    '       dbConn.Close()
                    '       Dim newSQLiteDatabase As New MySQLDatabase
                    '       MySQLDatabase.ConnectionString = connectionString
                    '       RaiseEvent NewDatabaseCreated(CType(MySQLDatabase, Database))
                    '       Return MySQLDatabase
                    '   End Using
                    'Else
                    '   Return Nothing
                    'End If
                    Return Nothing 'Added this return temporarily just for avoiding "doesn't return a value on all code paths." warning

                Case Else

                    Return Nothing 'Added this return temporarily just for avoiding "doesn't return a value on all code paths." warning

            End Select

        End Function

        ''' <summary>
        ''' Internal method for establishing the connection, used be CreateDatabase method.
        ''' </summary>
        ''' <param name="providerName"></param>
        ''' <param name="connectionString"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function CreateConnectionByDbConnection(ByVal providerName As String, ByVal connectionString As String) As DbConnection
            Try
                Dim dbFactory As DbProviderFactory = DbProviderFactories.GetFactory(providerName)
                Using dbConn As DbConnection = dbFactory.CreateConnection()
                    dbConn.ConnectionString = connectionString
                    dbConn.Open()
                    Return dbConn
                End Using
            Catch ex As Exception
                Balora.Alerter.REP("", ex, True)
                Return Nothing
            End Try
        End Function
#End Region

#Region "Events"
        Event NewDatabaseCreated(ByRef db As Database)
#End Region
    End Class
End Namespace






