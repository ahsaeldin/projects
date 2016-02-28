Imports Cacadore
Imports NUnit.Framework
Imports Cacadore.Actions.System.StartApplication

Namespace CacadoreIT
    <TestFixture()> _
    Public Class TaskVariablesTest

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
        Public Sub JustAnotherFuckenTest()
            Dim appNameVar As String = "calc"
            Dim newTask As New Task("testingId", False)
            Dim _global As New Globals()
            Dim newVarName = "CalculatorAppName1" & DateTime.Now.Millisecond
            Dim result As Boolean = newTask.Variables.AddVariable(newVarName, "calc")
            result = newTask.Variables.AddVariable("CalculatorAppName2" & DateTime.Now.Millisecond, "calc")
            result = _global.AddConstant("CalculatorAppName3" & DateTime.Now.Millisecond, "calc")
            result = _global.AddVariable("CalcResult" & DateTime.Now.Millisecond, "2")
            Dim errMessage As String = vbNullString
            TaskVariables.AddVariableNotation(newVarName, errMessage)
            Dim _value = newTask.Variables.GetVariableValue(newVarName)
            Dim startAction As New Actions.System.StartApplication
            Dim sss As New StartApplicationInputs(_value)
            startAction.Inputs = sss
            startAction.Execute()
            Assert.IsNotNull(_value)
            Assert.IsNull(errMessage)
        End Sub

        <Test()> _
        Sub Adding100Var()
            AddNVars(100)
        End Sub

        <Test()> _
        Sub Adding20Var()
            AddNVars(20)
        End Sub

        Sub AddNVars(n As Integer)
            Dim newTask As New Task(Balora.RandomBytesGenerator.GRCM(), False)
            newTask.Properties.TaskName = "A 100 vars task"
            newTask.Save()
            For i = 1 To n
                newTask.Variables.AddVariable(String.Format("var{0}{1}", i, DateTime.Now.Millisecond), "var value")
            Next
        End Sub
    End Class
End Namespace