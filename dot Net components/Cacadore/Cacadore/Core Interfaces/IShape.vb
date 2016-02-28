'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   IShape Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
''' <summary>
''' Base interface for Shape class.
''' </summary>
Public Interface IShape
#Region "Structures & Enums"
    Enum ShapeType
        Rectangular
        Cirular
        Square
        polygon
    End Enum
#End Region

#Region "Properties"
    Property Type As ShapeType
    Property Content As String
    Property Description As String
#End Region
End Interface

