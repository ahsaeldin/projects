'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   IGroup Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
''' <summary>
''' Base interface for Group class.
''' </summary>
Friend Interface IGroup
#Region "Properties"
    Property Name As String
    ReadOnly Property GroupId As String
    ReadOnly Property TasksCollection As TasksCollection
#End Region

#Region "Methods"
    Sub Pause()
    Sub [Resume]()
    Function Save() As Boolean
    Function Delete() As Boolean
#End Region
End Interface