Imports Balora
Imports Devart.Data
Imports System.Data
Imports Balora.Datler
Imports NUnit.Framework
Imports Balora.PathsHelper
Imports System.Data.Common
Imports System.Collections.Generic

Namespace BaloraUT
    <Ignore()>
    <TestFixture()> _
    Public Class SQLiteDatabaseTest
        Dim databasePlace As String = "D:\Rocknee\KV\Cacadu\Code\Trunk\Cacadu\bin\Debug\UnitTestingDatabase.db"
#If DatabaseProvider = "Devart.Data.Sqlite" Then
        Dim target As DevartSQLiteDatabase
#ElseIf DatabaseProvider = "System.Data.Sqlite" Then
    Dim target As SQLiteDatabase
#End If
        Dim newTableName As String = "TableToBeCreatedAndDeleted"
        <TestFixtureSetUp()> _
        Public Sub SetupMethods()

            Dim current As String = GetCurrentExecutingDirectory()
            If Not databasePlace.Contains("Rocknee") Then
                databasePlace = current & "UnitTestingDatabase.db"
            End If
            PrepareCacaduComponents()
#If DatabaseProvider = "Devart.Data.Sqlite" Then
            target = New DevartSQLiteDatabase() ' TODO: Initialize to an appropriate value
#ElseIf DatabaseProvider = "System.Data.Sqlite" Then
         target = New SQLiteDatabase()
#End If
        End Sub

        <TestFixtureTearDown()> _
        Public Sub TearDownMethods()
        End Sub

        <SetUp()> _
        Public Sub SetupTest()
        End Sub

        <TearDown()> _
        Public Sub TearDownTest()
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub IsDatabaseMalformed()
            Dim expected As Boolean = True
            Dim actual As Boolean = False
            Dim result As String = DevartSQLiteDatabase.CheckIfMalformed(CreateConnection(databasePlace))
            If result = "ok" Then
                actual = True
            End If
            Assert.AreEqual(expected, actual)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub IsTableExistsTest()
            'If IsNothing(x28549) Then
            'is_x28549(target.GetType, target)
            'End If
            Dim tableName As String = "UnitTestingTable" ' TODO: Initialize to an appropriate value
            Dim expected As Boolean = True ' TODO: Initialize to an appropriate value
            Dim actual As Boolean

            target.DatabaseConnection = CType(CreateConnection(databasePlace), SQLite.SQLiteConnection)
            actual = target.IsTableExists(target.DatabaseConnection, tableName) 'Test is here''''''''''''''''''''
            Assert.AreEqual(expected, actual)
        End Sub

        Private Function CreateConnection(databasePlace As String) As DbConnection
            Dim dbConn As DbConnection
            Dim ConnectionFactory As ConnectionFactory
            'Dim databasePlace As String = Util.GetCurrentExecutingDirectory & "UnitTestingDatabase.db"
#If DatabaseProvider = "Devart.Data.Sqlite" Then
            ConnectionFactory = New ConnectionFactory() 'Database Type Change: Change: Only one remainings.(3/4)
            ConnectionFactory.DatabaseProviderName = Dbprovider.DevartSQlite
#ElseIf DatabaseProvider = "System.Data.Sqlite" Then
        ConnectionFactory = New ConnectionFactory() 'Database Type Change: Change: Only one remainings.(3/4)
        ConnectionFactory.DatabaseProviderName = Dbprovider.System_Data_SQLite
#End If
            Try
                'If you didn't specify DatabaseProviderName  and/or DatabaseType Properties, defaults values used
                ConnectionFactory.DatabaseType = Balora.DbType.SQLite
                dbConn = ConnectionFactory.CreateConnection()
                dbConn.ConnectionString = "Data Source=" & databasePlace & ";"
                'Here i'm assigning the connection created by factory to CacaduDatabase 
                target.DatabaseConnection = CType(dbConn, SQLite.SQLiteConnection) 'Convert to the appropriate connection class Change: Done.(4/4)
            Catch ex As Exception
                MsgBox("Error Ya A7md")
                dbConn = Nothing
            End Try
            Return dbConn
        End Function


        '''<summary>
        '''A test for CreateTable
        '''</summary>
        <Test()> _
        <Category("ShortRunning")>
        Public Sub CreateTableTest()
            Console.WriteLine("Create table test started...")
            Dim creationCommands As New List(Of String)  ' TODO: Initialize to an appropriate value
            Dim expected As Boolean = True ' TODO: Initialize to an appropriate value
            Dim actual As Boolean
            Dim creationTableResult As Boolean = False
            creationCommands.Add("CREATE TABLE IF NOT EXISTS " & newTableName & "(id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT, testing1 REAL,testing2 INTEGER, testing3 NUMERIC, image BLOB);")
            target.DatabaseConnection = CType(CreateConnection(databasePlace), SQLite.SQLiteConnection)
            creationTableResult = target.CreateTable(creationCommands) 'Test is here''''''''''''''''''''
            If creationTableResult Then
                actual = target.IsTableExists(target.DatabaseConnection, "TestingCreateTableMethod")
            End If
            Console.WriteLine("Create table test ended.")
            Assert.AreEqual(expected, actual)
        End Sub

        '''<summary>
        '''A test for Query
        '''</summary>
        <Test()> _
        <Category("ShortRunning")>
        Public Sub QueryTest()
            Console.WriteLine("Query test started...")
            'I have a special database for testing located in Debug folder of unit test project...
            'The database has a table called "UnitTestingTable"...
            'This table has three columns (column1,column2,blobcolumn)
            'In this method we query for a row where column2=ahmed
            'that's it
            Dim selectCommand As String = Nothing
            Dim selectParameters As New Dictionary(Of String, Object)
            selectCommand = "SELECT * FROM UnitTestingTable WHERE column2=@parameter1"
            selectParameters.Add("@parameter1", "ahmed")
            target.DatabaseConnection = CType(CreateConnection(databasePlace), SQLite.SQLiteConnection)
            Dim queryDatatable As DataTable = target.Query(selectCommand, selectParameters)    'Test is here''''''''''''''''''''
            Dim expected As String = CStr(1)
            Dim actual As String = String.Empty
            'Check the successes of the query retrieving
            For Each row As DataRow In queryDatatable.Rows
                Dim values(row.ItemArray.Count - 1) As Object
                values = row.ItemArray
                For Each obj As Integer In values
                    If obj = 1 Then
                        actual = "1"
                        Exit For
                    End If
                Next
            Next
            Console.WriteLine("Query test ended.")
            Assert.AreEqual(expected, actual)
        End Sub

        '''<summary>
        '''A test for Update
        '''</summary>
        <Test()> _
        <Category("ShortRunning")>
        Public Sub UpdateTest()
            Console.WriteLine("Update test started...")
            Dim updateCommands As New List(Of String)  ' TODO: Initialize to an appropriate value
            updateCommands.Add("UPDATE " + "TestingUpdate" + " SET col1='153284'")
            target.DatabaseConnection = CType(CreateConnection(databasePlace), SQLite.SQLiteConnection)
            Dim expected As Boolean = True ' TODO: Initialize to an appropriate value
            Dim actual As Boolean
            actual = target.Update(updateCommands) 'Test is here''''''''''''''''''''
            'check the success of the update
            Dim selectCommand As String = "SELECT col1 FROM TestingUpdate"
            Dim queryDatatable As DataTable = target.Query(selectCommand)
            For Each row As DataRow In queryDatatable.Rows
                Dim values(row.ItemArray.Count - 1) As Object
                values = row.ItemArray
                For Each obj In values
                    If obj Is "153284" Then
                        actual = True
                        Exit For
                    End If
                Next
            Next
            Console.WriteLine("Update test ended.")
            Assert.AreEqual(expected, actual)
        End Sub

        '''<summary>
        '''A test for Delete
        '''</summary>
        <Test()> _
        <Category("ShortRunning")>
        Public Sub DeleteTest()
            Console.WriteLine("Delete test started...")
            'Delete all rows in newTableName variable
            Dim deleteCommands As New List(Of String)  ' TODO: Initialize to an appropriate value
            deleteCommands.Add("DELETE FROM " & newTableName & ";")
            Dim expected As Boolean = True  ' TODO: Initialize to an appropriate value
            Dim actual As Boolean
            target.DatabaseConnection = CType(CreateConnection(databasePlace), SQLite.SQLiteConnection)
            actual = target.Delete(deleteCommands) 'Test is here''''''''''''''''''''
            Console.WriteLine("Delete test ended.")
            Assert.AreEqual(expected, actual)
        End Sub

        '''<summary>
        '''A test for Insert
        '''</summary>
        <Test()> _
        <Category("ShortRunning")>
        Public Sub InsertTest()
            Console.WriteLine("Insert test started...")
            'Here we insert into "TableToTestInsert" then recheck with a query
            Dim insertCommands As New List(Of String)
            Dim insertParameters As New Dictionary(Of String, Object)
            insertCommands.Add("INSERT INTO " + "TableToTestInsert" + " (title, testing1,testing2,testing3,image) VALUES (@parameter1,@parameter2,@parameter3,@parameter4,@parameter5)")
            insertParameters.Add("@parameter1", Today.DayOfWeek.ToString)
            insertParameters.Add("@parameter2", 1)
            insertParameters.Add("@parameter3", 2)
            insertParameters.Add("@parameter4", 3)
            insertParameters.Add("@parameter5", "CacaduDatabase.ObjectToBlob(MainForm.PictureBox1.Image)")
            Dim expected As Boolean = True ' TODO: Initialize to an appropriate value
            Dim actual As Boolean
            target.DatabaseConnection = CType(CreateConnection(databasePlace), SQLite.SQLiteConnection)
            actual = target.Insert(insertCommands, insertParameters)  'Test is here''''''''''''''''''''

            If actual Then
                Dim selectCommand As String
                selectCommand = "SELECT * FROM TableToTestInsert WHERE title=@parameter1"
                Dim selectParameters As New Dictionary(Of String, Object)
                selectParameters.Add("@parameter1", Today.DayOfWeek.ToString)
                target.DatabaseConnection = CType(CreateConnection(databasePlace), SQLite.SQLiteConnection)
                Dim queryDatatable As DataTable = target.Query(selectCommand, selectParameters)
                If queryDatatable.Rows.Count <> 0 Then
                    actual = True
                End If
            End If
            Console.WriteLine("Insert test ended.")
            Assert.AreEqual(expected, actual)
        End Sub

        'EncryptTest & DecryptTest isn't for unit testing rather for...
        'Encrypt and decrypt Techtronic easily. 
        <Test()> _
        <Category("ShortRunning")>
        Public Sub EncryptTest()  ' TODO: Initialize to an appropriate value
            Dim expected As Boolean = True
            Dim actual As Boolean = False
            Console.WriteLine("Encrypt test started...")
            'target.DatabaseConnection = CType(CreateConnection(tectonicPlace), SQLite.SQLiteConnection)
            'target.Encrypt("SADFasdfq$%twe%tSAER#WQ$%RTYW$&DRSZDFGSER^W#$%^@$^#45")
            Assert.Inconclusive("Test commented for further fixes")
            Console.WriteLine("Encrypt table test ended.")
            Assert.AreEqual(expected, actual)
        End Sub

        'EncryptTest & DecryptTest isn't for unit testing rather for...
        'Encrypt and decrypt Techtronic easily. 
        <Test()> _
        <Category("ShortRunning")>
        Public Sub DecryptTest()
            Dim expected As Boolean = True
            Dim actual As Boolean = False
            Console.WriteLine("Decrypt test started...")
            'target.DatabaseConnection = CType(CreateConnection(tectonicPlace), SQLite.SQLiteConnection)
            'target.Decrypt("SADFasdfq$%twe%tSAER#WQ$%RTYW$&DRSZDFGSER^W#$%^@$^#45")
            Assert.Inconclusive("Test commented for further fixes")
            Console.WriteLine("Decrypt test ended.")
        End Sub

    End Class
End Namespace