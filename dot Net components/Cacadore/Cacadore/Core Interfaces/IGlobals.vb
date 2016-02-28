'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   IGlobals Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
''' <summary>
''' Base interface for Globals class.
''' </summary>
Public Interface IGlobals
    Enum GlobalType
        Constant
        Variable
        UnKnown
    End Enum

#Region "Properties"
    ReadOnly Property Variables As Dictionary(Of String, String)
    ReadOnly Property Constants As Dictionary(Of String, String)
#End Region

#Region "Methods"
    Function AddVariable(name As String,
                         value As String,
                         Optional errorMessage As String = vbNullString) As Boolean
    Function AddConstant(name As String, value As String, Optional errorMessage As String = vbNullString) As Boolean
    Function EditVariable(name As String,
                          newValue As String,
                          Optional errorMessage As String = vbNullString) As Boolean
    Function RemoveMemberByName(ByVal memberName As String) As Boolean
    Function IsRegisteredMember(memberName As String) As Boolean
#End Region
End Interface

Friend Interface IGlobalEntry
    Property Name As String
    Property Value As String
    Property type As IGlobals.GlobalType
End Interface