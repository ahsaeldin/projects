'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   Trigger Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
'Useful Resources:-
'   http://msdn.microsoft.com/en-us/magazine/cc164015.aspx
'   http://www.albahari.com/threading/part3.aspx#_Timers
'   http://www.codeproject.com/KB/dotnet/ABTransClockArticle.aspx
'   http://www.codeproject.com/KB/cs/tsnewlib.aspx 
'   http://quartznet.sourceforge.net/faq.html
'   http://quartznet.sourceforge.net/tutorial/index.html
'   http://quartznet.sourceforge.net/apidoc/
'   http://www.quartz-scheduler.org/documentation/quartz-1.x/tutorials/crontrigger
'   http://www.cronmaker.com
'   http://cron.nmonitoring.com/cron-generator.html
'   http://www.abunchofutils.com/utils/developer/cron-expression-helper

<Serializable()> Public MustInherit Class Trigger
    Inherits CacadoreBase
    Implements ITrigger
    Implements ICrudable

#Region "Fields"
    Private _triggerIdValue As String
    Private _parentTaskIdValue As String
    'AhSaElDin 20111228: Make it writable property.
    Private _descriptionValue As String = "Abstract Trigger"
#End Region

#Region "Properties"
    'AhSaElDin 20111231: Added to check it trigger is still scheduled.
    Public MustOverride ReadOnly Property MayFireAgain As Boolean
    'Not all of these properties are storable in Tectonic.
    Public MustOverride ReadOnly Property StartTimeUTC As Nullable(Of DateTime)
    Public MustOverride ReadOnly Property EndTimeUTC As Nullable(Of DateTime)
    Public MustOverride ReadOnly Property PreviousFireTimeUTC As Nullable(Of DateTime)
    Public MustOverride ReadOnly Property NextFireTimeUTC As Nullable(Of DateTime)
    Public MustOverride Property Task As Task Implements ITrigger.Task
    Public MustOverride ReadOnly Property TriggerType As ITrigger.TriggerTypeEnum Implements ITrigger.TriggerType
    'Reference to the actual object responsible for handling the triggers work.
    Public MustOverride Property InnerTrigger As Object Implements ITrigger.InnerTrigger
    Public MustOverride Property Calendar As Object Implements ITrigger.Calendar


    Private _misFireCommand As ITrigger.MisFireCommandsEnum = ITrigger.MisFireCommandsEnum.RunNow
    Public Property MisFireCommand As ITrigger.MisFireCommandsEnum Implements ITrigger.MisFireCommand
        Get
            Return _misFireCommand
        End Get
        Set(value As ITrigger.MisFireCommandsEnum)
            _misFireCommand = value
        End Set
    End Property

    Public ReadOnly Property TriggerId As String Implements ITrigger.TriggerId
        Get
            Return _triggerIdValue
        End Get
    End Property
    Public ReadOnly Property ParentTaskId As String Implements ITrigger.ParentTaskId
        Get
            Return _parentTaskIdValue
        End Get
    End Property

    Public Overridable Property Description As String Implements ITrigger.Description
        Get
            Return _descriptionValue
        End Get
        Set(value As String)
            _descriptionValue = value
        End Set
    End Property

    Public Property IsExecuting As Boolean = False Implements ITrigger.IsExecuting
#End Region

#Region "Methods"
    Public MustOverride Function StartListening() As Object Implements ITrigger.StartListening

    Public MustOverride Function StopListening() As Boolean Implements ITrigger.StopListening

    ''' <summary>
    ''' Deletes this instance.
    ''' </summary><returns></returns>
    ''' <remarks>
    ''' Delete is just like UnAttachTrigger, however, Task.UnAttachTrigger 
    ''' changes the execution state of the task to NotScheduled
    ''' if there is no trigger is attached to it after deletion.
    ''' </remarks>
    Public Overridable Function Delete() As Boolean Implements ITrigger.Delete
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return CBool(InvokeTectonicDlg(Me, {}, DeleteOperationDlg, My.Resources.TrgNotPresisted))
        End If
    End Function

    Public Function Save(tsk As Task) As String Implements ITrigger.Save
        Return _save(tsk)
    End Function

    Public Function Update(tsk As Task) As String Implements ITrigger.Update
        Return _save(tsk, True)
    End Function

    Public Sub SetParentTaskId(taskId As String) Implements ITrigger.SetParentTaskId
        _parentTaskIdValue = taskId
    End Sub

    Public Sub SetTriggerId(triggerId As String) Implements ITrigger.SetTriggerId
        _triggerIdValue = triggerId
    End Sub

    Private Sub getNewTriggerId(ByVal tsk As Task, ByVal update As Boolean)
        If Not update Then
            Dim lastTriggerId As String
            Dim newTriggerNo As Integer
            If tsk.MyTriggers.TriggersCount > 0 Then
                lastTriggerId = tsk.MyTriggers.Triggers.Last.Key
                newTriggerNo = Commands.ExtractTriggerNumber(lastTriggerId) + 1
            Else
                newTriggerNo = 1
            End If
            Dim newTriggerId As String = String.Format("{0}-t{1}", tsk.TaskId, newTriggerNo)
            _triggerIdValue = newTriggerId
            'Balora.Util.SetPrivateField(Me, "_triggerIdValue", newTriggerId)'AhSaElDin 20120202: والنبي أنت عسل عشان تضيف سطر زي دة!
        End If
    End Sub

    Public Overridable Sub PauseTrigger()

    End Sub

    Public Overridable Sub ResumeTrigger()

    End Sub

    Private Sub setParentTask(ByVal tsk As Task)
        Me.SetParentTaskId(tsk.TaskId)
    End Sub

    Private Sub sendToTectonic(ByVal update As Boolean)
        If update Then
            InvokeTectonicDlg(Me, {}, UpdateOperationDlg, My.Resources.TrgNotPresisted)
        Else
            InvokeTectonicDlg(Me, {}, InsertOperationDlg, My.Resources.TrgNotPresisted)
        End If
    End Sub

    Private Function _save(tsk As Task, Optional update As Boolean = False) As String
        Try
            getNewTriggerId(tsk, update)
            'Set the action id (attach action to the task)
            setParentTask(tsk)
        Catch ex As Exception
            If Debugger.IsAttached Then Debugger.Break()
            _privateRaiseErrorUp("", ex, True)
            Return ""
        End Try
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            sendToTectonic(update)
            Return Me.TriggerId
        End If
    End Function
#End Region
End Class