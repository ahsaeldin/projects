'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   TaskErrors Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
<Serializable()> Public Class TaskErrors
    Inherits CacadoreBase
    Implements ITaskErrors
    Implements IDisposable

#Region "Fields"
    Private _errorsValue As New System.Collections.Generic.Dictionary(Of String, List(Of String))
#End Region

#Region "Properties"
    Public Property Errors As System.Collections.Generic.Dictionary(Of String, List(Of String)) Implements ITaskErrors.Errors
        Get
            Return _errorsValue
        End Get
        Set(ByVal value As System.Collections.Generic.Dictionary(Of String, List(Of String)))
            _errorsValue = value
        End Set
    End Property
#End Region

#Region "Methods"
    ''' <summary>
    ''' You add an error by add the error no (extracted from ActionId) with the error message
    ''' </summary>
    Public Sub ListError(actionId As String, errorMessage As String) Implements ITaskErrors.ListError
        If Not _errorsValue.ContainsKey(actionId) Then
            Dim actionErrorsList As New List(Of String)
            actionErrorsList.Add(errorMessage)
            _errorsValue.Add(actionId, actionErrorsList)
        Else
            _errorsValue(actionId).Add(errorMessage)
        End If
    End Sub
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