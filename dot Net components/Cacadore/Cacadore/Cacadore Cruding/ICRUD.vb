'1.Crud & ICrud is defined and implemented in Cacadore.
'2.Cacadore classes invoke CRUD operations using Delegates.
'3.Delegates execuate corresponding methods in an instance of Crud
'4.Now the instance of CRUD execute the corresponding CRUD method in
'the shared CrudObject property of settings class.
'5.So that's it, all you need to do from Cacadu is to set CrudObject to an ICrud
'class the implements the CRUD functions.
Imports Cacadore

Public Interface ICrud
    Inherits IDisposable
#Region "Methods"
    Function Query(ByVal cacadoreObj As ICrudable, ByVal ParamArray args() As String) As Object
    Function Insert(ByVal cacadoreObj As ICrudable, ByVal ParamArray args() As String) As Boolean
    Function Update(ByVal cacadoreObj As ICrudable, ByVal ParamArray args() As String) As Boolean
    Function Delete(ByVal cacadoreObj As ICrudable, ByVal ParamArray args() As String) As Boolean
#End Region
End Interface