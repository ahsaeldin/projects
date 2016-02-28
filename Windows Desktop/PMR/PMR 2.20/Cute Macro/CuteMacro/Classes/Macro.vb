Option Explicit On

<Serializable()> Friend Enum MacroEventType
    MouseMove
    MouseDown
    MouseUp
    MouseWheel
    KeyDown
    KeyUp
    ForeGroundWindow
End Enum

<Serializable()> Friend Structure ForeGroundWindowAttr
    Dim WinTitle As String
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    '28/09/2010
    'اضافة اسم برنامج النافذة ومساره, لكي ابحث عنه
    'اذا لم يكن يعمل بالفعل لفتحه في وظيفة
    'PlaybackTheList
    'Dim ProcessName As String 'FoxSmr 
    'Dim ProcessFilePath As String 'FoxSmr 
End Structure

<Serializable()> Friend Structure ClonedMouseEventArgs '//because MouseEventArgs ain't serilizable, i will define my own type that clone it

    Dim Button As System.Windows.Forms.MouseButtons
    Dim Clicks As Integer
    Dim Delta As Integer
    Dim Location As System.Drawing.Point
    Dim X As Integer
    Dim Y As Integer

End Structure

<Serializable()> Friend Structure ClonedKeyEventArgs
    Dim Alt As Boolean
    Dim Control As Boolean
    Dim Handled As Boolean
    Dim KeyCode As System.Windows.Forms.Keys
    Dim KeyData As System.Windows.Forms.Keys
    Dim KeyValue As Integer
    Dim Modifiers As System.Windows.Forms.Keys
    Dim Shift As Boolean
    Dim SuppressKeyPress As Boolean
End Structure

<Serializable()> Friend Class MacroEvent

    <NonSerialized()> Public Event_Args As EventArgs
    Friend ForeWinStruct As New ForeGroundWindowAttr
    Friend ClonedKeyEventArgs As New ClonedKeyEventArgs
    Friend ClonedMouseEventArgs As New ClonedMouseEventArgs
    Friend Time_SinceLastEvent As Integer
    Friend Macro_Event_Type As MacroEventType

    '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Friend Sub New(ByVal macroEventType As MacroEventType, ByVal eventArgs As EventArgs, ByVal timeSinceLastEvent As Integer, Optional ByVal OtherItems As Object = "")
        '//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

        Macro_Event_Type = macroEventType

        Select Case macroEventType

            Case CuteMacro.MacroEventType.MouseMove, CuteMacro.MacroEventType.MouseDown, CuteMacro.MacroEventType.MouseUp, CuteMacro.MacroEventType.MouseWheel

                Dim MouseEvents As MouseEventArgs = eventArgs
                ClonedMouseEventArgs.Button = MouseEvents.Button
                ClonedMouseEventArgs.Clicks = MouseEvents.Clicks
                ClonedMouseEventArgs.Delta = MouseEvents.Delta
                ClonedMouseEventArgs.X = MouseEvents.X
                ClonedMouseEventArgs.Y = MouseEvents.Y

            Case CuteMacro.MacroEventType.KeyDown, CuteMacro.MacroEventType.KeyUp

                Dim KeyEvents As KeyEventArgs = eventArgs
                ClonedKeyEventArgs.Alt = KeyEvents.Alt
                ClonedKeyEventArgs.Control = KeyEvents.Control
                ClonedKeyEventArgs.Handled = KeyEvents.Handled
                ClonedKeyEventArgs.KeyCode = KeyEvents.KeyCode
                ClonedKeyEventArgs.KeyData = KeyEvents.KeyData
                ClonedKeyEventArgs.KeyValue = KeyEvents.KeyValue
                ClonedKeyEventArgs.Modifiers = KeyEvents.Modifiers
                ClonedKeyEventArgs.Shift = KeyEvents.Shift
                ClonedKeyEventArgs.SuppressKeyPress = KeyEvents.SuppressKeyPress

            Case CuteMacro.MacroEventType.ForeGroundWindow

                ForeWinStruct = OtherItems

        End Select

        Event_Args = eventArgs

        Time_SinceLastEvent = timeSinceLastEvent

    End Sub

End Class
