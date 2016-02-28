'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   Action Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
#Region "How to inherit a new action class"
'Action creation loop:-

'   0.Get the name of the action you want to implement.

'   1.Know exactly what this action can do.

'   2.Know the .net class name responsible for performing this action

'   3.Know how you can make a remarkable action using this class.

'   4.Search for class hidden features

'   5.Get synonym of the action name using Word or thesaurus.com

'   6.If the action is a new project, name the namespace like 'Actions.FacebookActions', so you can group all external actions under one name space. Don't forget to add version and obfusction settings to the new project.

'   7.Create the new Class derived from Action Class.

'   8.Add Action name to ActionTypeEnum

'   9.Search for all places you need to add the new action type into it like "CreateActionByType" method
'	    Hint:You can search for "ActionTypeEnum"

'   10.Add the specified ToDo's from StartApplication.vb "Very Important"

'   11.Declare Inputs, Outputs class and make them implement IInputs, IOutputs, IErros
'      And make both of them  <Serializable()>

'   12.Assign the 3 props (Errors, inputs and outputs) to thier fields in the new action
'     and make this in New constructor that as you did in ShutDownPC action

'	13.Now design the mockup.	

'   14.Code GetFullDescription 

'   15.Add action icon to ActionsIconsSharedImageCollection and name it exactly like action.gettype.name, because GetActionIconIndex will get it by name.
'      Add 2 icons to "Actions Icons Backup" folder, one 16*16 for image list and one 32*32 for action picture box.

'   16.Add action to gallery and task flow context menu. And populate  tags for both with action name.

'   17.Code "Clone" method for the new action.

'   18.Use CheckOutputVariable & LoadOutputVariable in CommonActionConnector to load/save output.

'   19.Rename caption of WaitAppExitCheckEdit for the action form

'   20.Added action to CreateActionFormGalleryItem method & GetActionForm method.

'   21.Decide of, new added properties to action class based on the class you found.
'        *.Check ideas section in RunApplication.vb

'   22.Implement Execute man:-
'	    22.1.Get inputs and place them in Inputs Class
'	    22.2.Process.
'	    22.3.Get Output and place them in Outputs Class
'	    22.4.Return Outputs and Results.
'       22.5.Override WaitForExit and check for TaskWaitBetweenActions property. then use Thread.Sleep to implement the Delay tab.

'   23.Possible Exception:
'		23.1.Get all possible exception by checking the new action calling points. (calling points are places where you really execute or prepare to execute)
'		23.2.Overrides Sub SetPossibleExceptions() in the new action and fill ActionExceptionsManager.AddPossibleException with all possible exceptions.
'		23.3.In the new action, make a call to Me.ExceptionsManager.HandleException(ex) just like we did in StartApplication.GetErrors method and a call to 
'            Me.ParentTask.Errors.ListError(Me.ActionId, ex.Message).
'       23.4.Add a call to SetPossibleExceptions in the new action constructor.

'   24.Dispose all variables in the new Action, so to avoid any possible memory leak.
#End Region

'Mark as serializable so I can clone/duplicate it.
<Serializable()>
Public MustInherit Class Action
    Inherits CacadoreBase
    Implements IAction
    Implements ICrudable

#Region "Fields"
    Private _skipValue As Boolean
    Private parentTaskValue As Task
    Private _shapeValue As New Shape
    Private _actionIdValue As String
    Private _enabledValue As Boolean
    Private _parentTaskIdValue As String
    Private _commentsValue As String = "N/A"
    Private _resultValue As ITaskState.ExecutionState
    Private _outputValue As System.Collections.Generic.Dictionary(Of String, String)
#End Region

#Region "Properties"
#Region "Must Overriden"
    Public MustOverride ReadOnly Property ActionType As IAction.ActionTypeEnum Implements IAction.ActionType
#End Region

#Region "Overridable Properties"
    Public MustOverride Property Inputs As IAction.IInputs Implements IAction.Input
    Public MustOverride Property Errors As IAction.IErrors Implements IAction.Errors
    Public MustOverride Property Outputs As IAction.IOutputs Implements IAction.Outputs
#End Region

    Public Property ParentTask() As Task Implements IAction.ParentTask
        Get
            If IsNothing(parentTaskValue) Then
                If TaskId <> "" Then parentTaskValue = Commands.GetTaskById(TaskId)
            End If
            Return parentTaskValue
        End Get
        Set(ByVal value As Task)
            parentTaskValue = value
        End Set
    End Property

    Public ReadOnly Property TaskId As String Implements IAction.TaskId
        Get
            Return _parentTaskIdValue
        End Get
    End Property

    Public Property ExceptionsManager As New ActionExceptionsManager Implements IAction.ExceptionsManager

    'If it has no Id, then it is unattached action.
    'Actions Ids is given automatically upon attaching to a task
    Public ReadOnly Property ActionId As String Implements IAction.ActionId
        Get
            Return _actionIdValue
        End Get
    End Property

    ''' <summary>
    ''' Holds a comment that will be copied to Shape.Description.
    ''' </summary>
    Public Property Comments As String Implements IAction.Comments
        Get
            Return _commentsValue
        End Get
        Set(ByVal value As String)
            _commentsValue = value
        End Set
    End Property

    Public Property Result As ITaskState.ExecutionState Implements IAction.Result
        Get
            Return _resultValue
        End Get
        Set(ByVal value As ITaskState.ExecutionState)
            _resultValue = value
        End Set
    End Property

    Public Property Shape As Shape Implements IAction.Shape
        Get
            Return _shapeValue
        End Get
        Set(value As Shape)
            _shapeValue = value
        End Set
    End Property

    'Use to give user the ability to un active an action rather than deleting it.
    Public Property Enabled As Boolean Implements IAction.Enabled
        Get
            Return _enabledValue
        End Get
        Set(ByVal value As Boolean)
            _enabledValue = value
        End Set
    End Property

    Public Property Skip As Boolean Implements IAction.Skip
        Get
            Return _skipValue
        End Get
        Set(value As Boolean)
            _skipValue = value
        End Set
    End Property

    Public Property RetryTime As Integer Implements IAction.RetryTime
#End Region

#Region "Methods"
    Function GetFormattedComment() As String
        Dim comment As String = vbNullString
        If Me.Comments <> "" And Comments <> "N/A" And Comments <> "Comments..." Then
            If Me.Comments.Length > 10 Then
                comment = Strings.Left(Me.Comments, 20) + "..."
            Else
                comment = Me.Comments
            End If
        End If
        Return comment
    End Function

    Function GetActionNumber() As Integer
        'equivalent to:-
        'Dim str As String() = Split(ActionId, "-a")
        'Return CInt(str(1))
        Return CInt(Split(ActionId, "-a")(1))
    End Function

    Sub LogWait(waitPeriod As Decimal)
        Dim message As String = vbNullString
        Dim actionNo As Integer = Commands.ExtractActionNumber(ActionId)
        If waitPeriod > 0 Then
            message = "Wait " & waitPeriod / 1000 & " second(s) for action" & "#" & actionNo & " ..."
        Else
            message = "Wait for the exit of action #" & actionNo & " ..."
        End If
        Watcher.AddNewEntry(Me.TaskId,
                            Me.ParentTask.Properties.TaskName,
                            IHistoryEntry.HistoryEntryType.Action,
                            message,
                            DateTime.Now,
                            ITaskState.ExecutionState.WaitForAction)
    End Sub

#Region "Must Overriden"
    'Note that: i want to give the user the ability to execute only action manually.
    ''' <summary>
    ''' Executes the specified input.
    ''' </summary>
    Public MustOverride Function Execute(Optional retryTimes As Integer = 1) As ITaskState.ExecutionState Implements IAction.Execute
    Public MustOverride Function GetFullDescription() As String
    Public MustOverride Function Clone() As Cacadore.Action
#End Region

    Public Function Highlight() As Boolean Implements IAction.Highlight
        '    'وبعد ما تخلصها ارجع تاني لرحلة اكشن وابدا من جديد وفكر بئه'
        '    'Highlighting has 2 meanings: 
        '    '1.upon errors
        '    '2.while execution
        Return Nothing
    End Function

    Public Function Delete() As Boolean Implements IAction.Delete
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else

            Return CBool(InvokeTectonicDlg(Me, {}, DeleteOperationDlg, My.Resources.ActNotPresisted))

        End If
    End Function

    Public Overridable Sub SetPossibleExceptions() Implements IAction.SetPossibleExceptions

    End Sub

    'ToDo check in the overridden method for TaskWaitBetweenActions property.
    Protected Overridable Sub WaitForExit(Optional ByRef state As Object = Nothing) Implements IAction.WaitForExit

    End Sub

    Public Overridable Sub SetOutput(memberName As String, newValue As String)
        If memberName <> vbNullString Then
            If Commands.IsRegisteredGlobalVariable(memberName) Then
                Dim globs As New Globals
                globs.EditVariable(memberName, newValue)
            End If

            If Commands.IsRegisteredTaskVariable(TaskId, memberName) Then
                Me.ParentTask.Variables.UpdateVariable(memberName, newValue)
            End If
        End If
    End Sub

    Protected Sub RaiseCompletedEvent([action] As Action, e As EventArgs)
        RaiseEvent Completed([action], e)
    End Sub

#Region "Setter Methods"
    Public Sub SetActionId(actionId As String)
        _actionIdValue = actionId
    End Sub

    Public Sub SetParentTaskId(taskId As String)
        _parentTaskIdValue = taskId
    End Sub
#End Region

#End Region

#Region "Events"
    Public Event Completed(sender As Object, e As System.EventArgs) Implements IAction.Completed
#End Region
End Class