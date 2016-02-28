Imports Cacadore

'I have created this project in order to be able to reference all actions in one method (CreateActionByType) that
'will be used in TecDal
Public Class Util
    ''' <summary>
    ''' Creates new type of the action by IAction.ActionTypeEnum
    ''' </summary>
    ''' <param name="actionType">Type of the action.</param><returns></returns>
    Public Shared Function CreateActionByType(actionType As IAction.ActionTypeEnum) As Action
        Select Case actionType
            Case IAction.ActionTypeEnum.StartApplication
                Return New Cacadore.Actions.System.StartApplication
            Case IAction.ActionTypeEnum.ShowAlert
                Return New Cacadore.Actions.MessageBoxes.ShowAlert
            Case IAction.ActionTypeEnum.ShowMessageBox
                Return New Cacadore.Actions.MessageBoxes.ShowMessageBox
            Case IAction.ActionTypeEnum.TakeScreenShot
                Return New Cacadore.Actions.System.TakeScreenShot
            Case IAction.ActionTypeEnum.ExitProcess
                Return New Cacadore.Actions.System.ExitProcess
            Case IAction.ActionTypeEnum.ShutdownPC
                Return New Cacadore.Actions.System.ShutdownPC()
            Case IAction.ActionTypeEnum.LastFacebookNotification
                Return New Actions.FacebookActions.LastFacebookNotification
        End Select
        Return Nothing
    End Function
End Class
