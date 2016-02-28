'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   ITrigger Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
''' <summary>
''' Base interface for Trigger class.
''' </summary>
Public Interface ITrigger
#Region "Enums"
    Enum MisFireCommandsEnum
        RunNow
        Skip
        AskMe
    End Enum

    Enum TriggerTypeEnum
        'Every time you inherit a new trigger, add its type here and...
        'Set this type in the Overridable ReadOnly Property TriggerType
        'in the new derived trigger class, just as you do in Trigger.vb
        DumpTrigger
        'Todo Note: Very important Note:-
        '*.For each new Trigger class, add it to 
        '      *.Commands.CreateTriggerByType()
        '      *.Set Description property in the new trigger class with TriggersDescriptions structure.
        ScheduleTrigger
        CronTrigger
        SimpleTrigger
    End Enum

    Structure TriggersDescriptions
        Public Shared DumpTrigger As String = "A dump trigger."
        Public Shared ScheduleTrigger As String = "Schedule a trigger to be fired at a prespecified periods."
    End Structure
#End Region

#Region "Properties"
    Property Task As Task
    Property Calendar As Object
    Property Description As String
    Property IsExecuting As Boolean
    ReadOnly Property TriggerId As String
    ReadOnly Property ParentTaskId As String
    Property MisFireCommand As MisFireCommandsEnum
    ReadOnly Property TriggerType As TriggerTypeEnum
    Property InnerTrigger As Object 'Reference to the actual object responsible for handling the triggers work.
#End Region

#Region "Methods"
    Function StartListening() As Object
    Function StopListening() As Boolean
    Function Save(tsk As Task) As String
    Function Update(tsk As Task) As String
    Function Delete() As Boolean
    Sub SetParentTaskId(taskId As String)
    Sub SetTriggerId(triggerId As String)
#End Region
End Interface