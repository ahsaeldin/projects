'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   TaskFlowPresenter Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
''' <summary>
''' Handle the entire representation of a selected task actions flow
''' </summary>
''' <remarks></remarks>
Public Class TaskFlowPresenter
    Inherits CacadoreBase
    Implements ITaskFlowPresenter
    Implements IDisposable

    'How to start
    'http://www.angelfire.com/tx4/cus/shapes/vbnet.html  COOOOOOOOOOOOOL
    'System.Drawing	 'تركي العسيري

#Region "Fields"
    Private _arrowTypeValue As ITaskFlowPresenter.ArrowTypeEnum
    Private _drawingEngineValue As ITaskFlowPresenter.DrawingEngineEnum
#End Region

#Region "Properties"
    Public Property ArrowType As ITaskFlowPresenter.ArrowTypeEnum Implements ITaskFlowPresenter.ArrowType
        Get
            Return _arrowTypeValue
        End Get
        Set(ByVal value As ITaskFlowPresenter.ArrowTypeEnum)
            _arrowTypeValue = value
        End Set
    End Property

    ''' <summary>
    ''' Choose the library/Engine you want to use to
    ''' build your Task Flow Chart.
    ''' 'Note: You can use this property in the future,
    ''' 'to change the library you use to build the task
    ''' 'flow chart.
    ''' </summary>
    Public Property DrawingEngine As ITaskFlowPresenter.DrawingEngineEnum Implements ITaskFlowPresenter.DrawingEngine
        Get
            Return _drawingEngineValue
        End Get
        Set(ByVal value As ITaskFlowPresenter.DrawingEngineEnum)
            _drawingEngineValue = value
        End Set
    End Property
#End Region

#Region "Methods"
    ''' <summary>
    ''' Draw actions with the arrows in between
    ''' </summary>
    Public Function Draw(ByVal task As Task, ByVal control As System.Windows.Forms.Control) As Boolean Implements ITaskFlowPresenter.Draw
        RaiseEvent StartDraw(task)
        'based on DrawingEngine type, it will draw the chart on the specified control 

        'If Drawn Then
        RaiseEvent Drawn(task)
        'Else
        RaiseEvent DrawnFail(task)
        'End If

        Return Nothing
    End Function
#End Region

#Region "Events"
    Event StartDraw(ByRef task As Task) Implements ITaskFlowPresenter.StartDraw
    Event Drawn(ByRef task As Task) Implements ITaskFlowPresenter.Drawn
    Event DrawnFail(ByRef task As Task) Implements ITaskFlowPresenter.DrawnFail
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

