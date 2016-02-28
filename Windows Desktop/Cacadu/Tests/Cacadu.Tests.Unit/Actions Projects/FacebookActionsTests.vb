Imports NUnit.Framework
Imports System.Threading

Namespace CacadoreUT
    <TestFixture()> _
    Public Class FacebookActionsTests

        <TestFixtureSetUp()> _
        Public Sub SetupMethods()
            Unit.PrepareCacaduComponents()
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
        Public Sub TestLogin()
     
        End Sub

        <Test()> _
        Public Sub TestLogOff()

        End Sub

#Region "Threads"
        'Why we call using a thread.
        'http://stackoverflow.com/questions/1418466/single-threaded-apartment-cannot-instantiate-activex-control
        Private Sub login()
     
        End Sub
#End Region
    End Class
End Namespace