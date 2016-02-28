'Till 20111211, History doesn't save any data neither about triggers nor action. 
'Only about tasks @ChangeExecutionState (historyObject.AddNewEntry(historyEntry). 
'In the future you can support this.
'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   History Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
''' <summary>
''' Full mixed log (Tasks,Actions and Triggers) for the entire Cacadore.
''' </summary>
Public Class History
    Inherits CacadoreBase
    Implements IHistory
    Implements ICrudable
    Implements IDisposable

    Private Enum EventName
        Cleared
        EnabledChanged
        NewEntryAdded
    End Enum

    Enum ClearType
        All
        ByEntryType
        ByTimeRange
    End Enum

#Region "Methods"
#Region "Add Entries Methods"
    ''' <summary>
    ''' Add a new history entry to database 
    ''' using HistoryEntry structure. 
    ''' </summary>
    ''' <returns>True, if added successfully</returns>
    Public Shared Function AddNewEntry(ByVal entry As HistoryEntry) As Boolean
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            If Not Settings.HistoryEnabled Then Return False
            Dim result As Boolean = CBool(InvokeTectonicDlg(entry, {}, InsertOperationDlg, My.Resources.HistNotPresisted))
            raiseMyEvent(EventName.NewEntryAdded)
            Return result
        End If
    End Function

    Public Shared Function AddNewEntry(TaskId As String,
                                       EntryType As IHistoryEntry.HistoryEntryType,
                                       Description As String,
                                       EntryDateTime As DateTime,
                                       Status As ITaskState.ExecutionState) As Boolean
        Dim histEntry As New HistoryEntry
        With histEntry
            .TaskId = TaskId
            .EntryType = EntryType
            .Description = Description
            .EntryDateTime = EntryDateTime
            .Status = Status
        End With
        Return AddNewEntry(histEntry)
    End Function
#End Region

#Region "Get Methods"
    Private Shared Function GetHistory(args() As String) As List(Of HistoryEntry)
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return CType(InvokeTectonicDlg(New History, args, QueryOperationDlg, My.Resources.HistNotPresisted), List(Of HistoryEntry))
        End If
    End Function

    ''' <summary>
    ''' Get all Tasks history as a generic list of HistoryEntry
    ''' </summary>
    Public Shared Function GetFullTasksHistory() As List(Of HistoryEntry)
        Return GetHistory({CStr(IHistoryEntry.HistoryEntryType.Task)})
    End Function

    ''' <summary>
    ''' Get Task history using its id.
    ''' </summary>
    Public Shared Function GetTaskHistory(ByVal taskId As String) As System.Collections.Generic.List(Of HistoryEntry)
        Return GetHistory({taskId})
    End Function

    ''' <summary>
    ''' Get all actions history as a generic list of HistoryEntry
    ''' </summary>
    Public Shared Function GetFullActionsHistory() As System.Collections.Generic.List(Of HistoryEntry)
        Return GetHistory({CStr(IHistoryEntry.HistoryEntryType.Action)})
    End Function

    ''' <summary>
    ''' Get all Triggers history as a generic list of HistoryEntry
    ''' </summary>
    Public Shared Function GetFullTriggersHistory() As System.Collections.Generic.List(Of HistoryEntry)
        Return GetHistory({CStr(IHistoryEntry.HistoryEntryType.Trigger)})
    End Function
#End Region

#Region "Clear Methods"
    ''' <summary>
    ''' Clear all Cacadore history.
    ''' </summary>
    ''' <returns>True, if cleared successfully</returns>
    Public Shared Function ClearHistory() As Boolean
        Dim result As Boolean = _clearHistory(ClearType.All.ToString)
        If result Then
            raiseMyEvent(EventName.Cleared)
        End If
        Return result
    End Function

    ''' <summary>
    ''' Clear all entries specified by EntryType Enum (Task - Action - Trigger)
    ''' </summary>
    ''' <param name="byEntryType">(Task - Action - Trigger)</param>
    ''' <returns>True, if cleared successfully</returns>
    ''' <remarks>Overload of ClearHistory</remarks>
    Public Shared Function ClearHistory(ByVal byEntryType As IHistoryEntry.HistoryEntryType) As Boolean
        Dim result As Boolean = _clearHistory(ClearType.ByEntryType.ToString, CStr(byEntryType))
        If result Then
            raiseMyEvent(EventName.Cleared)
        End If
        Return result
    End Function

    ''' <summary>
    ''' Clear by choosing a time range to clear like what FireFox do.
    ''' </summary>
    ''' <param name="clearTimeRange"></param>
    ''' <param name="customDate">
    ''' 'If clearTimeRange = CustomDate then you must 
    '''  set customDate to a date of your choice.
    ''' </param>
    ''' <returns>True, if cleared successfully</returns>
    ''' <remarks>Overload of ClearHistory</remarks>
    Public Function ClearHistory(ByVal clearTimeRange As IHistory.TimeRange, Optional ByVal customDate As Date = #12:00:00 AM#) As Boolean
        Dim result As Boolean = _clearHistory(ClearType.ByTimeRange.ToString, CStr(clearTimeRange), customDate.ToString)
        If result Then
            raiseMyEvent(EventName.Cleared)
        End If
        Return result
    End Function

    Private Shared Function _clearHistory(ParamArray args() As String) As Boolean
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return CBool(InvokeTectonicDlg(New History, args, DeleteOperationDlg, My.Resources.HistNotPresisted))
        End If
    End Function

    Private Shared Sub raiseMyEvent(eventName As EventName, ParamArray args As Object())
        Select Case eventName
            Case History.EventName.Cleared
                RaiseEvent Cleared()
            Case History.EventName.EnabledChanged
                RaiseEvent EnabledChanged(CBool(args(0)))
            Case History.EventName.NewEntryAdded
                RaiseEvent NewEntryAdded(CType(args(0), HistoryEntry))
        End Select
    End Sub

    Friend Shared Sub OnEnabledChanged(change As Boolean)
        raiseMyEvent(EventName.EnabledChanged, change)
    End Sub
#End Region
#End Region

#Region "Events"
    Public Shared Event Cleared()
    Public Shared Event EnabledChanged(ByVal state As Boolean)
    Public Shared Event NewEntryAdded(ByVal entry As HistoryEntry)
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class