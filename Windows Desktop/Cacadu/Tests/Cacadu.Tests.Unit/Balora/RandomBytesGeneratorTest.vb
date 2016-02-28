Imports Balora
Imports NUnit.Framework
Imports System.Collections
Imports System.Collections.Generic

Namespace BaloraUT
    '''<summary>
    '''This is a test class for RandomBytesGeneratorTest and is intended
    '''to contain all RandomBytesGeneratorTest Unit Tests
    '''</summary>
    <TestFixture()> _
    Public Class RandomBytesGeneratorTest

        <TestFixtureSetUp()> _
        Public Sub SetupMethods()
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
        Public Sub GenerateRNGCharacterMaskTest()
            Dim maxSize As Integer = 8
            Dim mask As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
            'Dim mask As String = "1234567890"
            Dim actual As Boolean
            Dim expected As Boolean = True
            Dim result As String = RandomBytesGenerator.GRCM(maxSize, mask)
            If result <> Nothing Then
                actual = True
                Assert.AreEqual(expected, actual)
            End If
        End Sub

        Public Sub CheckDuplicationTestHelper(Of T)()
            Dim someItemsDuplicated As Boolean = False ' TODO: Initialize to an appropriate value
            Dim someItemsDuplicatedExpected As Boolean = False ' TODO: Initialize to an appropriate value
            Dim expected As Boolean = True  ' TODO: Initialize to an appropriate value
            Dim len As Integer = 10000
            Dim keys As String() = New String(len - 1) {}
            Dim rndNumber As New Random
            For [loop] As Integer = 0 To len - 1
                'Three methods to generate random numbers
                keys([loop]) = RandomBytesGenerator.GRCM()
                'keys([loop]) = Guid.NewGuid.ToString
                'keys([loop]) = rndNumber.Next(1, 1000000).ToString
            Next
            Dim keysList As New List(Of String)
            'Holds all generated keys
            keysList = keys.ToList()
            'Holds all duplicated items in keysList with their index in keysList 
            Dim actual As Hashtable = RandomBytesGenerator.CheckDuplication(Of String)(keysList, someItemsDuplicated)
            Assert.AreEqual(someItemsDuplicatedExpected, someItemsDuplicated)
        End Sub

        <Test()> _
        <Category("ShortRunning")>
        Public Sub CheckDuplicationTest()
            CheckDuplicationTestHelper(Of String)()
        End Sub

    End Class
End Namespace