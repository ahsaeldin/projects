Imports System.ComponentModel
Imports Balora
Imports NUnit.Framework
Imports System.Threading

Namespace BaloraUT
    <TestFixture()> _
    Public Class BackgroundWorkerHelperTest

        Dim manualResetEve As New ManualResetEvent(False)
        Dim sharedBoolean As Boolean = False


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

#Region "How to use"
        Delegate Sub MethodToBeInjectedDlg(sender As Object, e As System.ComponentModel.DoWorkEventArgs)
        Private Sub AHowToUseMethodImpppporrrrrrrrrrtaaaaaaaaaaaaaant()
            Dim bckGrndWrkHelper As New Balora.BackgroundWorkerHelper
            'You don't need to add handler of BackGroundWorkerObject.DoWork, to avoid duplicate call to it.
            AddHandler bckGrndWrkHelper.BackGroundWorkerObject.RunWorkerCompleted, AddressOf RunWorkerCompleted
            bckGrndWrkHelper.WorkToDo = New MethodToBeInjectedDlg(AddressOf WorkToDoMethodToBeInjected)
            bckGrndWrkHelper.RunWorkerCompletedDlg = New RunWorkerCompletedEventHandler(AddressOf RunWorkerCompleted)
            bckGrndWrkHelper.Start()
        End Sub
#End Region

        <Test()> _
        <NUnit.Framework.Category("ShortRunning")>
        Public Sub StartTest()
            Dim expected As Boolean = True
            Dim bckGrndWrkHelper As New Balora.BackgroundWorkerHelper
            AddHandler bckGrndWrkHelper.BackGroundWorkerObject.RunWorkerCompleted, AddressOf RunWorkerCompleted
            Dim workToDoDelegate As MethodToBeInjectedDlg = New MethodToBeInjectedDlg(AddressOf WorkToDoMethodToBeInjected)
            bckGrndWrkHelper.WorkToDo = workToDoDelegate
            bckGrndWrkHelper.RunWorkerCompletedDlg = New RunWorkerCompletedEventHandler(AddressOf RunWorkerCompleted)
            bckGrndWrkHelper.Start()
            manualResetEve.WaitOne()
            Assert.AreEqual(expected, sharedBoolean)
        End Sub

        Private Function WorkToDoMethodToBeInjected(sender As Object, e As System.ComponentModel.DoWorkEventArgs) As Boolean
            Return True
        End Function

        Private Function RunWorkerCompleted(sender As Object, e As ComponentModel.RunWorkerCompletedEventArgs) As Boolean
            sharedBoolean = True
            manualResetEve.Set()
            Return sharedBoolean
        End Function

    End Class
End Namespace
