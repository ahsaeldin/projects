'Factory for creating a connection to a database, you pass 
'this connection to database class in order to handle the CRUD operations.

Imports System.Data.Common
Imports System.ComponentModel
Imports System.Windows.Forms

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
    Friend Class ConnectionFactory
        Inherits BaloraBase

#Region "Properties"
        ''' <summary>
        ''' Decide the database connection type that will be returned by this factory
        ''' </summary>
        ''' <remarks>All returned connection must be inherited from DbConnection class</remarks>
        Private databaseTypeValue As DbType
        Public Property DatabaseType() As DbType
            Get
                Return databaseTypeValue
            End Get
            Set(ByVal value As DbType)
                databaseTypeValue = value
            End Set
        End Property

        ''' <summary>
        ''' Choose the database provider 
        ''' </summary>
        ''' <remarks>Default is Devart SQLite</remarks>
        Private databaseProviderNameValue As Dbprovider
        Public Property DatabaseProviderName() As Dbprovider
            Get
                Return databaseProviderNameValue
            End Get
            Set(ByVal value As Dbprovider)
                databaseProviderNameValue = value
            End Set
        End Property
#End Region

#Region "Methods"
        ''' <summary>
        ''' Create connection of the specified databaseTypeValue property. 
        ''' </summary>
        ''' <returns>returns a class inherited from DbConnection class</returns>
        ''' <remarks></remarks>
        Public Overridable Function CreateConnection() As DbConnection

            Dim dbConnection As DbConnection = Nothing

            Select Case databaseTypeValue

                Case DbType.SQLite

                    Select Case DatabaseProviderName

                        Case Dbprovider.DevartSQlite
#If DatabaseProvider = "Devart.Data.Sqlite" Then
                            Try
                                Dim connStringBuilder As New SQLiteConnectionStringBuilder
                                connStringBuilder.FailIfMissing = False
                                connStringBuilder.Locking = LockingMode.Normal
                                connStringBuilder.AutoVacuum = AutoVacuumMode.Full
                                connStringBuilder.ConnectionTimeout = 20
                                dbConnection = New SQLiteConnection(connStringBuilder.ConnectionString)
                            Catch ex As Exception
                                Alerter.REP("", ex, True)
                                Return Nothing
                            End Try
#End If

                            'Set the default database to SQLite using a default provider "System.Data.Sqlite"
                        Case Else 'Dbprovider.System_Data_SQLite
#If DatabaseProvider = "System.Data.Sqlite" Then
                            Try
                                Dim connStringBuilder As New SQLiteConnectionStringBuilder
                                'connStringBuilder.Password = "YourPasswordHere"
                                dbConnection = New SQLiteConnection()
                            Catch ex As Exception
                                Balora.Alerter.RaiseErrorUp("", ex, True)
                                Return Nothing
                            End Try
#End If

                    End Select

                Case DbType.MySQL

                Case DbType.SQLServer

                Case DbType.Oracle

                Case DbType.PostgreSQL

                Case DbType.Firebird

            End Select

            RaiseEvent NewConnectionCreated(databaseTypeValue, dbConnection)

            Return dbConnection

        End Function
#End Region

#Region "Events"
        ''' <summary>
        ''' Fired up whenever a new connection created using CreateConnection method.
        ''' </summary>
        ''' <param name="dbType">Database type Enum</param>
        ''' <param name="dbConnection">A class "like SQLiteConnection" inherited from DbConnection</param>
        ''' <remarks></remarks>
        Event NewConnectionCreated(ByVal dbType As DbType, ByRef dbConnection As DbConnection)
#End Region

    End Class
End Namespace

