'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   BIT Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/

Imports Balora
Imports Cacadore
Imports Cacadore.Actions.System
Imports Cacadore.Actions.System.StartApplication
Friend Class BIT
    Public Shared Function CreateBunchOfRunApplicationTasks() As Task
        Return _createBunchOfRunApplicationTasks()
    End Function

    Public Shared Function _createBunchOfRunApplicationTasks() As Task
        Dim myId = RandomBytesGenerator.GRCM()

        Dim miniTask As New Task(myId)

        miniTask.Properties.TaskName = "Tezk Ya A7md"

        For i As Integer = 1 To 5

            Dim myAction As New StartApplication

            Dim actionShape As New Shape
            Dim actionInput As New StartApplicationInputs("calc")
            Dim actionOutput As New StartApplicationOutputs

            actionInput.WaitForExitPeriod = 5000

            CacadoreToTectonic.AttachActionToTask(miniTask, IAction.ActionTypeEnum.StartApplication, "Action Desc", False, actionInput, actionOutput, actionShape)

        Next

        Return miniTask
    End Function
End Class