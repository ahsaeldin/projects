Imports NUnit.Framework

Namespace CacadoreIT
    <TestFixture()> _
    Public Class ExitProcessTest

        Private _cacadoreToTectonic As New CacadoreToTectonic

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
        <Category("ShortRunning")>
        Public Sub TestPresistanceOfExceptionsManager_CreateActionAndSetExceptionManager_ReteriveFromTectonicAndCheckResult()
            Dim tsk = CacadoreToTectonic.CreateNewTask(False)

            Dim newAction = CacadoreToTectonic.AttachActionToTask(tsk,
                                                                   Cacadore.IAction.ActionTypeEnum.ExitProcess,
                                                                   "seco seco man",
                                                                   False,
                                                                   Nothing,
                                                                   Nothing,
                                                                   Nothing)
            tsk.Save()

            Dim actionFromTectonic As Cacadore.Action = Cacadore.Commands.GetActionById(newAction.ActionId)

            Assert.AreEqual(newAction.ActionId, actionFromTectonic.ActionId)
            Assert.AreEqual(newAction.ExceptionsManager.PossibleExceptions.First.Key,
                            actionFromTectonic.ExceptionsManager.PossibleExceptions.First.Key)

        End Sub

    End Class
End Namespace
