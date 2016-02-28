'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   ITaskVariables Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
''' <summary>
''' Base interface for TaskVariables class.
''' </summary>
Friend Interface ITaskVariables
#Region "Properties"
    ReadOnly Property taskId As String
    ReadOnly Property Variables As VariablesCollection
#End Region

#Region "Methods"
    Function DeleteVariables() As Boolean
    Function GetVariableValue(name As String) As String
    Function RemoveVariable(ByVal variableName As String) As Boolean
    Function AddVariable(name As String, value As String) As Boolean
    Function UpdateVariables(varColl As VariablesCollection) As Boolean
    Function UpdateVariable(name As String, newValue As String) As Boolean
#End Region
End Interface