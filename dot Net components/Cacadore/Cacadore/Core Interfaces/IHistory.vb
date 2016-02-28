'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   IHistory Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
''' <summary>
''' Base interface for History class.
''' </summary> 
Public Interface IHistory
#Region "Structures & Enums"
    ''' <summary>
    ''' Specify a time range used to clear the history.
    ''' </summary>
    ''' <remarks></remarks>
    Enum TimeRange
        LastHour
        LastTwoHours
        LastFourHours
        Today
        Everything
        CustomDate
    End Enum
#End Region
End Interface

