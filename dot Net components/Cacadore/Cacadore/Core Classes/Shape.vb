'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   Shape Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
Imports System.Drawing
''' <summary>
''' Will be used to graphically represent a single action
''' </summary>
''' <remarks></remarks>
<Serializable()>
Public Class Shape
    Inherits CacadoreBase
    Implements IShape
    Implements IDisposable

    'How to start
    'http://www.angelfire.com/tx4/cus/shapes/vbnet.html  COOOOOOOOOOOOOL
    'System.Drawing	 'تركي العسيري

#Region "Fields"
    Private _contentValue As String
    Private _descriptionValue As String
    Private _shapeTypeValue As IShape.ShapeType
#End Region

#Region "Properties"
    ''' <summary>
    ''' Holds a comment that may written in side the shape it self.
    ''' </summary>
    Public Property Description As String Implements IShape.Description
        Get
            Return _descriptionValue
        End Get
        Set(ByVal value As String)
            _descriptionValue = value
        End Set
    End Property

    ''' <summary>
    ''' Tells you what action will be performed
    ''' </summary>
    Public Property Content As String Implements IShape.Content
        Get
            Return _contentValue
        End Get
        Set(ByVal value As String)
            _contentValue = value
        End Set
    End Property

    ''' <summary>
    ''' Specify the shape type.
    ''' </summary>
    ''' <remarks>Will be used later in TaskFlowPresenter to draw the shape.</remarks>
    Public Property Type As IShape.ShapeType Implements IShape.Type
        Get
            Return _shapeTypeValue
        End Get
        Set(ByVal value As IShape.ShapeType)
            _shapeTypeValue = value
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

