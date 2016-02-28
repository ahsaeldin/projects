'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   TriggersCollection Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
<Serializable()> Public Class TriggersCollection
    Inherits CacadoreBase
    Implements ITriggersCollection
    Implements ICrudable
    Implements IDisposable

#Region "Fields"
    Friend TaskId As String
#End Region

#Region "Properties"
    Public ReadOnly Property TriggersCount As Integer Implements ITriggersCollection.TriggersCount
        Get
            Dim count As Integer = Commands.GetTaskTriggersCount(TaskId)
            Return count
        End Get
    End Property

    Public ReadOnly Property Triggers As System.Collections.Generic.Dictionary(Of String, Trigger) Implements ITriggersCollection.Triggers
        Get
            Return getTriggers()
        End Get
    End Property
#End Region

#Region "Methods"
    Private Function getTriggers(ParamArray args() As String) As Dictionary(Of String, Trigger)
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return CType(InvokeTectonicDlg(Me, {Me.TaskId}, QueryOperationDlg, My.Resources.TrgNotPresisted), Dictionary(Of String, Trigger))
        End If
    End Function
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