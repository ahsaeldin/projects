Imports Cacadore

'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   TaskMeta Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
'This class will be used to exchange a task and its 
'triggers standard data between Cacadu and CacaSer.
<Serializable()> Public Class TaskMeta
    Implements IDisposable
    ''' <summary>
    ''' Used to hold all commands and string message between Cacadu
    ''' and CacaSer, and will be used by both of them.
    ''' </summary>
    ''' <remarks >
    ''' The only member here that is public to access from Cacadu.
    ''' </remarks>
    Public Enum StringMessages
        'Common Commands...

        'For unit testing only...
        swear 'For unit testing only (send back by CacaSer to Cacadu while unit testing.)
        CC 'For unit testing only.
        'Use this enum to tell the receiver 
        'of the message (usually CacaSer) that
        'you want to create a attach a new 
        'trigger to a task with the id specified
        'in TaskId property and with scheduling option
        'as the one in CronExpression.
        CreateNewTask
        TaskCreated
        CreateAndAttachTrigger
        TriggerCreatedAndAttached
        ScheduleTrigger
        TriggerScheduled
        TriggerFired
        RunTask
        TaskDone
    End Enum

#Region "Properties"
    'AhSaElDin 20120102: After disabling CacaSer, there is 
    'no need for a lot of property is, however, I will left 
    'them, who knows?!
    Property TaskId As String
    Property Task As Cacadore.Task
    'AhSaElDin 20120102:
    'Property Command As StringMessages
    Property CronExpression As String
    Property CronExpressionDescription As String
    Property CalendarCronExpression As String
    Property CalendarCronExpressionDescription As String
    Property StartDateTimeUTC As DateTime?
    Property EndDateTimeUTC As DateTime?
    Property LastAttachedTriggerId As String
    Property ParentGroupId As String
    Property BreakOnFailed As Boolean
    Property WaitBetweenActions As Boolean
    Property IsEnabled As Boolean
    Property Name As String
    Property TriggerType As ITrigger.TriggerTypeEnum
    Property RepeatCount As Integer
    Property RunOnce As Boolean
    Property MisFire As ITrigger.MisFireCommandsEnum
    'AhSaElDin 20120102: 
    'Property Message As String 'an optional message to be sent between receiver and sender.
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If Not IsNothing(Task) Then Task.Dispose()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
            Task = Nothing
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
