'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   Crud Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
'1.Crud & ICrud is defined and implemented in Cacadore.
'2.Cacadore classes invoke CRUD operations using Delegates.
'3.Delegates execuate corresponding methods in an instance of Crud
'4.Now the instance of CRUD execute the corresponding CRUD method in
'the shared CrudObject property of settings class.
'5.So that's it, all you need to do from Cacadu is to set CrudObject to an ICrud
'class the implements the CRUD functions.
Imports Cacadore
''' <summary>
''' A proxy for CRUD operations to avoid dealing
''' directly with Datler or any other database technology 
''' directly from within the code, so I can change the CRUD
''' in the future.
''' </summary>
Public Class Crud
    'Never ever inherits this class from CacadoreBase... 
    'coz you need to set it in "PrepareCacadore" so how could you set it 
    'Inherits CacadoreBase
    Implements ICrud
    Implements IDisposable

#Region "Properties"
    'Link to the external Crud class that will handle all the Crud stuff.
    Private _crudObjectValue As ICrud
    Public Property CrudObject() As ICrud
        Get
            Return _crudObjectValue
        End Get
        Set(ByVal value As ICrud)
            _crudObjectValue = value
        End Set
    End Property
#End Region

#Region "Methods"
    Friend Function Query(ByVal cacadoreObj As ICrudable, ByVal ParamArray args() As String) As Object Implements ICrud.Query
        If _crudObjectValue Is Nothing Then
            _privateRaiseErrorUp(My.Resources.CrudObjNotSet, Nothing, False)
            Return False
        Else
            Return (_crudObjectValue.Query(cacadoreObj, args))
        End If
    End Function

    'args used to send "allrows" to delete all rows
    'or used to send global variable name to delete using "DeleteGlobalVariable"
    Friend Function Delete(ByVal cacadoreObj As ICrudable, ByVal ParamArray args() As String) As Boolean Implements ICrud.Delete
        If _crudObjectValue Is Nothing Then
            _privateRaiseErrorUp(My.Resources.CrudObjNotSet, Nothing, False)
            Return False
        Else
            Return _crudObjectValue.Delete(cacadoreObj, args)
        End If
    End Function

    Friend Function Insert(ByVal cacadoreObj As ICrudable, ByVal ParamArray args() As String) As Boolean Implements ICrud.Insert
        If _crudObjectValue Is Nothing Then
            _privateRaiseErrorUp(My.Resources.CrudObjNotSet, Nothing, False)
            Return False
        Else
            Return _crudObjectValue.Insert(cacadoreObj, args)
        End If
    End Function

    Friend Function Update(ByVal cacadoreObj As ICrudable, ByVal ParamArray args() As String) As Boolean Implements ICrud.Update
        If _crudObjectValue Is Nothing Then
            _privateRaiseErrorUp(My.Resources.CrudObjNotSet, Nothing, False)
            Return False
        Else
            Return _crudObjectValue.Update(cacadoreObj, args)
        End If
    End Function
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If Not _crudObjectValue Is Nothing Then _crudObjectValue.Dispose()
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