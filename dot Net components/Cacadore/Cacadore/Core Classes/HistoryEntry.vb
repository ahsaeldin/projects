'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   HistoryEntry Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
''' <summary>
''' Represents a row in History in class.(HistoryEntry has no methods), 
''' created just to be able to pass it to Crud.Insert(ICrudable)
''' </summary>
Public Class HistoryEntry
    Inherits CacadoreBase
    Implements ICrudable
    Implements IHistoryEntry
    Implements IDisposable

#Region "Fields"
    Private taskIdValue As String
    Private messageValue As String
    Private entryDateTimeValue As Date
    Private statusValue As ITaskState.ExecutionState
    Private entryTypeValue As IHistoryEntry.HistoryEntryType
#End Region

#Region "Properties"
    Public Property EntryDateTime As Date Implements IHistoryEntry.EntryDateTime
        Get
            Return entryDateTimeValue
        End Get
        Set(value As Date)
            entryDateTimeValue = value
        End Set
    End Property

    Public Property EntryType As IHistoryEntry.HistoryEntryType Implements IHistoryEntry.EntryType
        Get
            Return entryTypeValue
        End Get
        Set(value As IHistoryEntry.HistoryEntryType)
            entryTypeValue = value
        End Set
    End Property

    Public Property Description As String Implements IHistoryEntry.Description
        Get
            Return messageValue
        End Get
        Set(value As String)
            messageValue = value
        End Set
    End Property

    Public Property Status As ITaskState.ExecutionState Implements IHistoryEntry.Status
        Get
            Return statusValue
        End Get
        Set(value As ITaskState.ExecutionState)
            statusValue = value
        End Set
    End Property

    Public Property TaskId As String Implements IHistoryEntry.TaskId
        Get
            Return taskIdValue
        End Get
        Set(value As String)
            taskIdValue = value
        End Set
    End Property
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
