'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   IHistoryEntry Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
'Represents a row in History in class.
Public Interface IHistoryEntry
    ''' <summary>
    ''' Used to tell history class about the type of entry 
    ''' </summary>
    ''' <remarks></remarks>
    Enum HistoryEntryType
        Task
        Action
        Trigger
    End Enum

    Property TaskId As String
    Property EntryType As HistoryEntryType
    Property Description As String
    Property EntryDateTime As DateTime
    Property Status As ITaskState.ExecutionState
End Interface