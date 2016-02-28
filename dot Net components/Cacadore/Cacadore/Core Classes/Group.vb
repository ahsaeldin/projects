'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   Group Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
''' <summary>
''' Group class stores data about a specified 
''' tasks group, and "has a" TasksCollection class.
''' </summary>

<Serializable()>
Public Class Group
    Inherits CacadoreBase
    Implements IGroup
    Implements ICrudable
    Implements IDisposable

#Region "Fields"
    Private _nameValue As String
    Private _groupIdValue As String
    Private _taskCollectionValue As New TasksCollection
#End Region

#Region "Constructors"
    Public Sub New(groupId As String)
        _groupIdValue = groupId
        _taskCollectionValue.GroupId = groupId
    End Sub
#End Region

#Region "Properties"
    ''' <summary>
    ''' Every group has a name (name isn't unique and may duplicate)
    ''' </summary>
    Public Property Name As String Implements IGroup.Name
        Get
            Return _nameValue
        End Get
        Set(ByVal value As String)
            _nameValue = value
        End Set
    End Property

    ''' <summary>
    ''' Every group has a unique id.
    ''' </summary>
    Public ReadOnly Property GroupId As String Implements IGroup.GroupId
        Get
            Return _groupIdValue
        End Get
    End Property

    ''' <summary>
    ''' Gets the tasks collection.
    ''' </summary>
    ''' <value>
    ''' The tasks collection.
    ''' </value>
    Public ReadOnly Property TasksCollection As TasksCollection Implements IGroup.TasksCollection
        Get
            Return _taskCollectionValue
        End Get
    End Property
#End Region

#Region "Methods"
    ''' <summary>
    ''' Saves the group.
    ''' </summary><returns></returns>
    Public Function Save() As Boolean Implements IGroup.Save
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return CBool(InvokeTectonicDlg(Me, {}, InsertOperationDlg, My.Resources.GrpNotPresisted))
        End If
    End Function

    ''' <summary>
    ''' Deletes the entire group.
    ''' </summary><returns></returns>
    Public Function Delete() As Boolean Implements IGroup.Delete
        If Not _isCrudObjectDefined() Then
            'In-memory Cacadore scenario.
            _privateShowMessageBox("In-memory Cacadore scenario isn't developed yet!")
            Return Nothing
        Else
            Return CBool(InvokeTectonicDlg(Me, {}, DeleteOperationDlg, My.Resources.GrpNotPresisted))
        End If
    End Function

    Public Sub Pause() Implements IGroup.Pause
        For Each tskKVP In TasksCollection.Tasks
            tskKVP.Value.Pause()
        Next
    End Sub

    Public Sub [Resume]() Implements IGroup.Resume
        For Each tskKVP In TasksCollection.Tasks
            tskKVP.Value.[Resume]()
        Next
    End Sub
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                If Not IsNothing(_taskCollectionValue) Then _taskCollectionValue.Dispose()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
            _taskCollectionValue = Nothing
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