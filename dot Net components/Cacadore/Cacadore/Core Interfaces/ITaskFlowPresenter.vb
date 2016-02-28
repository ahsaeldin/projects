'/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\
'                                                   ITaskFlowPresenter Class                                                    
'\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/
Imports System.Windows.Forms
''' <summary>
''' Base interface for TaskFlowPresenter class.
''' </summary>
Public Interface ITaskFlowPresenter
#Region "Structures & Enums"
    Enum ArrowTypeEnum
        nameATypeHere
    End Enum

    Enum DrawingEngineEnum
        _dotNetDrawings
        nevron
    End Enum
#End Region

#Region "Properties"
    Property ArrowType As ArrowTypeEnum
    Property DrawingEngine As DrawingEngineEnum
#End Region

#Region "Methods"
    Function Draw(ByVal task As Task, ByVal control As Control) As Boolean
#End Region

#Region "Events"
    Event StartDraw(ByRef task As Task)
    Event Drawn(ByRef task As Task)
    Event DrawnFail(ByRef task As Task)
#End Region
End Interface

