'Implemented mainly by "Database" class declared in database.vb
Imports System.Data.Common
Namespace Datler
    Friend Interface IDatabase
        Property DatabaseType() As DbType
        Property DataAdapter() As DbDataAdapter
        Function BackupDatabase() As String 'You may use create file
        Function RestoreDatabase() As String 'You may use create file
        Function CreateTable(ByVal creationCommands As List(Of String)) As Boolean
        Function Query(ByVal selectCommand As String, Optional ByVal selectParameters As Dictionary(Of String, Object) = Nothing) As DataTable
        Function Update(ByVal updateCommands As List(Of String), Optional ByVal updateParameters As Dictionary(Of String, Object) = Nothing) As Boolean
        Function Insert(ByVal inserationCommands As List(Of String), Optional ByVal insertParameters As Dictionary(Of String, Object) = Nothing) As Boolean
        Function Delete(ByVal deleteCommands As List(Of String), Optional ByVal deleteParameters As Dictionary(Of String, Object) = Nothing) As Boolean
        Function IsTableExists(ByVal dbConn As DbConnection, ByVal tableName As String) As Boolean
        Function ExecuteUsingDataAdapter(ByVal command As String) As DbDataAdapter
    End Interface
End Namespace