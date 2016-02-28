'Base class for CRUD operations, most of method has to be implemented in the derived class.

'Abstract Database Class
Imports System.Data.Common
Imports System.ComponentModel

Namespace Datler
    Public MustInherit Class Database
        Inherits BaloraBase
        Implements IDatabase

#Region "Properties"

        ''' <summary>
        ''' Handle to the Connection created by ConnectionFactory.CreateConnection()
        ''' </summary>
        ''' <remarks>You can use it for set/get connection string.</remarks>
        Private dbConnectionValue As DbConnection
        Public Property DatabaseConnection() As DbConnection
            Get
                Return dbConnectionValue
            End Get
            Set(ByVal value As DbConnection)
                dbConnectionValue = value
            End Set
        End Property

        ''' <summary>
        ''' Set/Get database type using DbType Enum
        ''' </summary>
        ''' <remarks></remarks>
        Private databaseTypeValue As DbType
        Public Property DatabaseType() As DbType Implements IDatabase.DatabaseType
            Get
                Return databaseTypeValue
            End Get
            Set(ByVal value As DbType)
                databaseTypeValue = value
            End Set
        End Property

        Private WithEvents dataAdapterValue As DbDataAdapter
        Public Property DataAdapter As System.Data.Common.DbDataAdapter Implements IDatabase.DataAdapter
            Get
                Return dataAdapterValue
            End Get
            Set(ByVal value As System.Data.Common.DbDataAdapter)
                dataAdapterValue = value
            End Set
        End Property
#End Region

#Region "Methods"

        Public MustOverride Function CreateTable(ByVal creationCommands As List(Of String)) As Boolean Implements IDatabase.CreateTable
        Public MustOverride Function Query(ByVal selectCommand As String, Optional ByVal selectParameters As Dictionary(Of String, Object) = Nothing) As DataTable Implements IDatabase.Query
        Public MustOverride Function Update(ByVal updateCommands As List(Of String), Optional ByVal updateParameters As Dictionary(Of String, Object) = Nothing) As Boolean Implements IDatabase.Update
        Public MustOverride Function Insert(ByVal inserationCommands As List(Of String), Optional ByVal insertParameters As Dictionary(Of String, Object) = Nothing) As Boolean Implements IDatabase.Insert
        Public MustOverride Function Delete(ByVal deleteCommands As List(Of String), Optional ByVal deleteParameters As Dictionary(Of String, Object) = Nothing) As Boolean Implements IDatabase.Delete
        Public MustOverride Function BackupDatabase() As String Implements IDatabase.BackupDatabase
        Public MustOverride Function RestoreDatabase() As String Implements IDatabase.RestoreDatabase
        Public MustOverride Function IsTableExists(ByVal dbConn As System.Data.Common.DbConnection, ByVal tableName As String) As Boolean Implements IDatabase.IsTableExists
        Public MustOverride Function ExecuteUsingDataAdapter(ByVal command As String) As System.Data.Common.DbDataAdapter Implements IDatabase.ExecuteUsingDataAdapter
        ''' <summary>
        ''' Convert System.Type to System.Data.DbType using a hashtable
        ''' </summary>
        Public Overridable Function CSystemTypeToDbType(ByVal type As Type) As System.Data.DbType
            'Supported database by SQLite 3
            'http://www.sqlite.org/datatype3.html
            Dim dbTypeTable As New Hashtable
            dbTypeTable.Add(GetType(String), System.Data.DbType.String)
            dbTypeTable.Add(GetType(Int16), System.Data.DbType.Int16)
            dbTypeTable.Add(GetType(Int32), System.Data.DbType.Int32)
            dbTypeTable.Add(GetType(Int64), System.Data.DbType.Int64)
            dbTypeTable.Add(GetType(Byte()), System.Data.DbType.Binary)
            dbTypeTable.Add(GetType(Boolean), System.Data.DbType.Boolean)
            dbTypeTable.Add(GetType(Object), System.Data.DbType.Object)
            'Data.DbType.AnsiString
            'Data.DbType.AnsiStringFixedLength
            'Data.DbType.Currency
            'Data.DbType.Date
            'Data.DbType.DateTime
            'Data.DbType.DateTime2
            'Data.DbType.DateTimeOffset
            'Data.DbType.Decimal
            'Data.DbType.Double
            'Data.DbType.Guid
            'Data.DbType.SByte
            'Data.DbType.Single
            'Data.DbType.StringFixedLength
            'Data.DbType.Time
            'Data.DbType.UInt16
            'Data.DbType.UInt64
            'Data.DbType.VarNumeric
            'Data.DbType.Xml
            Dim myDbType As System.Data.DbType = CType(dbTypeTable(type), Data.DbType)
            Return myDbType
        End Function
#End Region

#Region "Events"

#End Region
    End Class
End Namespace

