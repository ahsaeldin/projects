Imports Balora
Imports NUnit.Framework

Namespace BaloraUT

    <TestFixture()> _
    Public Class PathsHelperTest

        <TestFixtureSetUp()> _
        Public Sub SetupMethods()
            PrepareCacaduComponents()
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
        Public Sub GetPathVariableListTest_CallAndReteriveList_CheckReturnedDirectoriesPaths()

            Dim pathVarList As List(Of String) = PathsHelper.GetPathVariableList()

            Console.WriteLine(pathVarList.Count)

            Assert.NotNull(pathVarList, "Cannot retrieve paths.")

            For Each _path In pathVarList
                If Not IO.Directory.Exists(_path) Then
                    Assert.Inconclusive("paths is invalid")
                End If
            Next

        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub GetExecutablesPathsListInPathEnVarFoldersTest_CallAndReteriveList_CheckReturnedFilesPaths()

            Dim _list As List(Of String) = PathsHelper.GetExecutablesPathsListInPathEnVarFolders

            Assert.NotNull(_list, "Cannot retrieve paths.")

            For Each _path In _list
                If Not IO.File.Exists(_path) Then
                    Assert.Inconclusive("paths is invalid")
                End If
            Next

        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub GetExecutablesNamesAndPathsInPathEnvVarFoldersTest_CallAndReteriveSortedList_CheckReturnedFilesPaths()

            Dim _dic As SortedList(Of String, String) = PathsHelper.GetExecutablesNamesAndPathsInPathEnvVarFolders()

            Assert.NotNull(_dic, "Cannot retrieve paths.")

            For Each kvp In _dic
                If Not IO.File.Exists(kvp.Value) Then
                    Assert.Inconclusive("paths is invalid")
                End If
            Next

        End Sub

    End Class

End Namespace
