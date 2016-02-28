Attribute VB_Name = "ModEditMacros"
Option Explicit

'"0x0ffee11" == Form_Load in frmedit.
'"0x0ffee12" == OpenMacroForEditing.
'"0x0ffee13" == GetListedIndexMacroInTheCombo.
'"0x0ffee14" == AddMouseEvent.
'"0x0ffee15 == AddKeyboardEvent.
'"0x0ffee16" == TranslateKey.
'"0x0ffee17" == AddDelayEvent.

Private Declare Function loword Lib "TLBINF32" (ByVal DWord As Long) As Integer
Private Declare Function hiword Lib "TLBINF32" (ByVal DWord As Long) As Integer
'//cprin 4Sep2006: added next 2 declares to get the key and write it into the list view
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Private Declare Function MapVirtualKeyEx Lib "user32" Alias "MapVirtualKeyExA" (ByVal uCode As Long, ByVal uMapType As Long, ByVal dwhkl As Long) As Long

'//cprin 28Aug2006: added next function to add a mouse event to the list view of frm edit
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub AddMouseEvent(EventType As String, EventArrForEditing As EVENTMSG, index As Long, Optional PrevTime As Long)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: AddMouseEvent
' Description: AddMouseEvent Description
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

'Parameters of this function:
'Event Type : Contains the Type of the mouse event [Move, Left Button Down, Left Button Up, Wheel,Right Button Down, Right Button Up]
'EventArrForEditing:Contains the EVENTMSG element that contains the data
'index:the index of the EventArrForEditing and will be used in this function as
    '1.In the key of the item to be added in listview line labeled 10
    '2.To Check if this is the first element in line labeled 20
    '3.To be saved in CurItem.Tag because may be we will use later in modifing the element
'PrevTime: Contains the time of EVENTMSG for the previous element and i make it optional because
'//in OpenMacroForEditing function if we process the element nubmer 0, no need for this option
'//hence rather than pass 0 to this parametet, i make it an optional parameter.

    On Error GoTo AddMouseEventErr
        
    Dim CurItem As ListItem
    
10  Set CurItem = frmEdit.CommandsListView.ListItems.Add(, "Key" & Str(index), "Mouse", , MouseIcon)  '//add the event
    
    CurItem.ForeColor = MouseEventColor '//Set the color the row.
    
    '//cprin 9Sep2006: added next if else statement to get the delay between events
20  If index = 0 Then '//if the first event
        '//then make the delay equal to the event time
        '//cprin 17Sep2006: change next line because sometimes the first event daley time is less than0
        '//Like this -188 hence i will make it 0
        'CurItem.SubItems(1) = EventArrForEditing.time '//Enter the timeout for this event.
        'CurItem.SubItems(4) = 0 'EventArrForEditing.time '//Enter the timeout for this event.
    
    Else
        '//else substract this event time from the prev event time to get the delay for this event
        'CurItem.SubItems(4) = EventArrForEditing.time - PrevTime   '// PrevTime is the time of the previous element in the event array
    End If
    
    CurItem.SubItems(1) = EventType '//Enter the Event Type
    CurItem.SubItems(2) = "X=" & EventArrForEditing.paramL '//Enter X horizontal position of cursor
    CurItem.SubItems(3) = "Y=" & EventArrForEditing.paramH '//vertical position of cursor
    CurItem.ListSubItems.Item(1).ForeColor = MouseEventColor '//Set the color the Sub Item.
    CurItem.ListSubItems.Item(2).ForeColor = MouseEventColor '//Set the color the Sub Item.
    CurItem.ListSubItems.Item(3).ForeColor = MouseEventColor '//Set the color the Sub Item.
'    CurItem.ListSubItems.Item(4).ForeColor = MouseEventColor   '//Set the color the Sub Item.
30  CurItem.Tag = index  '//Tag will contains the array index to use later.

AddMouseEventExit:
    On Error Resume Next

    Exit Sub
    
AddMouseEventErr:

    '//0x0ffee14 number is a unique number for AddMouseEvent because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    SysError Err, "0x0ffee14"

GoTo AddMouseEventExit

End Sub
'//cprin 28Aug2006: added next function to add a mouse event to the list view of frm edit
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub AddKeyboardEvent(EventType As String, EventArrForEditing As EVENTMSG, index As Long, Optional PrevTime As Long)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: AddMouseEvent
' Description: AddMouseEvent Description
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

'Parameters of this function:
'Event Type : Contains the Type of the keyboard event [Key Up, Key Down, System Key Down, System Key Up]
'EventArrForEditing:Contains the EVENTMSG element that contains the data
'index:the index of the EventArrForEditing and will be used in this function as
    '1.In the key of the item to be added in listview line labeled 10
    '2.To Check if this is the first element in line labeled 20
    '3.To be saved in CurItem.Tag because may be we will use later in modifing the element
'PrevTime: Contains the time of EVENTMSG for the previous element and i make it optional because
'//in OpenMacroForEditing function if we process the element nubmer 0, no need for this option
'//hence rather than pass 0 to this parametet, i make it an optional parameter.

    On Error GoTo AddKeyboardEventErr
        
    Dim CurItem As ListItem
10  Set CurItem = frmEdit.CommandsListView.ListItems.Add(, "Key" & Str(index), "Keyboard", , KeyboardIcon)  '//add the event
    CurItem.ForeColor = KeyboradEventColor '//Set the color the row.
    '//cprin 9Sep2006: added next if else statement to get the delay between events
20   If index = 0 Then '//if the first event
        '//then make the delay equal to the event time
          '//cprin 17Sep2006: change next line because sometimes the first event daley time is less than0
        '//Like this -188 hence i will make it 0
        'CurItem.SubItems(1) = EventArrForEditing.time '//Enter the timeout for this event.
        CurItem.SubItems(5) = 0 'EventArrForEditing.time '//Enter the timeout for this event.
   Else
        '//else substract this event time from the prev event time to get the delay for this event
        'CurItem.SubItems(5) = EventArrForEditing.time - PrevTime '// PrevTime is the time of the previous element in the event array
    End If
    CurItem.SubItems(1) = EventType '//Enter the Event Type
    '//cprin 4Sep2006: added next Select Case statement to get the key and write it into the list view
    'Select Case (EventArrForEditing.paramL And &HFF)
    '       Case 0 To 31, 90 To 159
    '             CurItem.SubItems(3) = TranslateKey(EventArrForEditing.paramL And &HFF)
    '       Case Else
    '             CurItem.SubItems(3) = Chr(MapVirtualKeyEx(EventArrForEditing.paramL And &HFF, 2, GetKeyboardLayout(GetWindowThreadProcessId(GetForegroundWindow, 0))))
    'End Select
    
    '//cprin 9Sep2006: added next line to write the key into the list view.
    CurItem.SubItems(2) = TranslateKey(EventArrForEditing.paramL And &HFF)
    
    CurItem.SubItems(3) = ""
    CurItem.ListSubItems.Item(1).ForeColor = KeyboradEventColor '//Set the color the Sub Item.
    CurItem.ListSubItems.Item(2).ForeColor = KeyboradEventColor '//Set the color the Sub Item.
    CurItem.ListSubItems.Item(3).ForeColor = KeyboradEventColor '//Set the color the Sub Item.
'    CurItem.ListSubItems.Item(4).ForeColor = KeyboradEventColor '//Set the color the Sub Item.
30  CurItem.Tag = index '//Tag will contains the array index to use later.
    


AddKeyboardEventExit:
    On Error Resume Next

    Exit Sub
    
AddKeyboardEventErr:

    '//0x0ffee15 number is a unique number for AddKeyboardEvent because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    SysError Err, "0x0ffee15"

GoTo AddKeyboardEventExit

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub AddDelayEvent(EventType As String, EventArrForEditing As EVENTMSG, index As Long, Optional PrevTime As Long)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: AddDelayEvent
' Description: AddDelayEvent Description
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

'Parameters of this function:
'Event Type : Contains the Type of the mouse event [Move, Left Button Down, Left Button Up, Wheel,Right Button Down, Right Button Up]
'EventArrForEditing:Contains the EVENTMSG element that contains the data
'index:the index of the EventArrForEditing and will be used in this function as
    '1.In the key of the item to be added in listview line labeled 10
    '2.To Check if this is the first element in line labeled 20
    '3.To be saved in CurItem.Tag because may be we will use later in modifing the element
'PrevTime: Contains the time of EVENTMSG for the previous element and i make it optional because
'//in OpenMacroForEditing function if we process the element nubmer 0, no need for this option
'//hence rather than pass 0 to this parametet, i make it an optional parameter.

    On Error GoTo AddDelayEvent
        
    Dim CurItem As ListItem
    
10  Set CurItem = frmEdit.CommandsListView.ListItems.Add(, "Key" & Str(index), "Delay", , DelayIcon) '//add the event
    
    CurItem.ForeColor = DelayEventColor '//Set the color the row.
    
    '//cprin 9Sep2006: added next if else statement to get the delay between events
20  If index = 0 Then '//if the first event
        '//then make the delay equal to the event time
        '//cprin 17Sep2006: change next line because sometimes the first event daley time is less than0
        '//Like this -188 hence i will make it 0
        'CurItem.SubItems(1) = EventArrForEditing.time '//Enter the timeout for this event.
        CurItem.SubItems(4) = 0 'EventArrForEditing.time '//Enter the timeout for this event.
    
    Else
        '//else substract this event time from the prev event time to get the delay for this event
        CurItem.SubItems(4) = EventArrForEditing.time - PrevTime  '// PrevTime is the time of the previous element in the event array
           
    End If
    
    CurItem.SubItems(1) = CurItem.SubItems(4) '//Enter the Event Type
    'CurItem.SubItems(2) = "X=" & EventArrForEditing.paramL '//Enter X horizontal position of cursor
    'CurItem.SubItems(3) = "Y=" & EventArrForEditing.paramH '//vertical position of cursor
    CurItem.ListSubItems.Item(1).ForeColor = DelayEventColor '//Set the color the Sub Item.
    CurItem.ListSubItems.Item(2).ForeColor = DelayEventColor '//Set the color the Sub Item.
    CurItem.ListSubItems.Item(3).ForeColor = DelayEventColor '//Set the color the Sub Item.
    CurItem.ListSubItems.Item(4).ForeColor = vbWhite   '//Set the color the Sub Item.
30  CurItem.Tag = index  '//Tag will contains the array index to use later.

AddDelayEventExit:
    On Error Resume Next

    Exit Sub
    
AddDelayEvent:

    '//0x0ffee14 number is a unique number for AddDelayEvent because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    SysError Err, "0x0ffee17"

GoTo AddDelayEventExit

End Sub

'//cprin 4Sep2006: added next function to Translate Keys from 0 To 31, 90 To 159
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function TranslateKey(Key As Integer) As String
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: TranslateKey
' Description:  Translate Keys from 0 To 31, 90 To 159
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo TranslateKeyErr
      
    Select Case Key
           
        Case vbKeyControl
                TranslateKey = "Control"
        Case vbKeyMenu
                TranslateKey = "Alt"
        Case vbKeyBack
                TranslateKey = "Back"
        Case vbKeyTab
                TranslateKey = "Tab"
        Case vbKeyClear
                TranslateKey = "Clear"
        Case vbKeyReturn
                TranslateKey = "Enter"
        Case vbKeyShift
                TranslateKey = "Shift"
        Case vbKeyMenu
                TranslateKey = "Alt"
        Case vbKeyPause
                TranslateKey = "Pause"
        Case vbKeyCapital
                TranslateKey = "Capital"
        Case vbKeyEscape
                TranslateKey = "Escape"
        Case vbKeySpace
                TranslateKey = "Space"
        Case vbKeyPrior
                TranslateKey = "Page up"
        Case vbKeyNext
                TranslateKey = "Page down"
        Case vbKeyEnd
                TranslateKey = "End"
        Case vbKeyHome
                TranslateKey = "Home"
        Case vbKeyLeft
                TranslateKey = "Left"
        Case vbKeyRight
                TranslateKey = "Right"
        Case vbKeyUp
                TranslateKey = "Up"
        Case vbKeyDown
                TranslateKey = "Down"
        Case vbKeySnapshot
               TranslateKey = "Snapshot"
        Case vbKeyInsert
               TranslateKey = "Insert"
        Case vbKeyDelete
               TranslateKey = "Delete"
        Case vbKey0
               TranslateKey = "0"
        Case vbKey1
               TranslateKey = "1"
        Case vbKey2
               TranslateKey = "2"
        Case vbKey3
               TranslateKey = "3"
        Case vbKey4
               TranslateKey = "4"
        Case vbKey5
               TranslateKey = "5"
        Case vbKey6
               TranslateKey = "6"
        Case vbKey7
               TranslateKey = "7"
        Case vbKey8
               TranslateKey = "8"
        Case vbKey9
               TranslateKey = "9"
        Case vbKey0
               TranslateKey = "0"
        Case vbKeyA
               TranslateKey = "A"
        Case vbKeyB
               TranslateKey = "B"
        Case vbKeyC
               TranslateKey = "C"
        Case vbKeyD
               TranslateKey = "A"
        Case vbKeyE
               TranslateKey = "E"
        Case vbKeyF
               TranslateKey = "F"
        Case vbKeyG
               TranslateKey = "G"
        Case vbKeyH
               TranslateKey = "H"
        Case vbKeyI
               TranslateKey = "I"
        Case vbKeyJ
               TranslateKey = "J"
        Case vbKeyK
               TranslateKey = "K"
        Case vbKeyL
               TranslateKey = "L"
        Case vbKeyM
               TranslateKey = "M"
        Case vbKeyN
               TranslateKey = "N"
        Case vbKeyO
               TranslateKey = "O"
        Case vbKeyP
               TranslateKey = "P"
        Case vbKeyQ
               TranslateKey = "Q"
        Case vbKeyR
               TranslateKey = "R"
        Case vbKeyS
               TranslateKey = "S"
        Case vbKeyT
               TranslateKey = "T"
        Case vbKeyU
               TranslateKey = "U"
        Case vbKeyV
               TranslateKey = "V"
        Case vbKeyW
               TranslateKey = "W"
        Case vbKeyX
               TranslateKey = "X"
        Case vbKeyY
               TranslateKey = "Y"
        Case vbKeyZ
               TranslateKey = "Z"
        Case &H5B 'VK_LWIN
               TranslateKey = "Left Win"
        Case &H5C 'VK_RWIN
               TranslateKey = "Right Win"
        Case vbKeyNumlock
               TranslateKey = "Numlock"
        Case vbKeyNumpad0
               TranslateKey = "Numpad0"
        Case vbKeyNumpad1
               TranslateKey = "Numpad1"
        Case vbKeyNumpad2
               TranslateKey = "Numpad2"
        Case vbKeyNumpad3
               TranslateKey = "Numpad3"
        Case vbKeyNumpad4
               TranslateKey = "Numpad4"
        Case vbKeyNumpad5
               TranslateKey = "Numpad5"
        Case vbKeyNumpad6
               TranslateKey = "Numpad6"
        Case vbKeyNumpad7
               TranslateKey = "Numpad7"
        Case vbKeyNumpad8
               TranslateKey = "Numpad8"
        Case vbKeyNumpad9
               TranslateKey = "Numpad9"
        Case vbKeyMultiply
               TranslateKey = "Multiply"
        Case vbKeyAdd
               TranslateKey = "Add"
        Case vbKeySeparator
               TranslateKey = "Separator"
        Case vbKeySubtract
               TranslateKey = "Subtract"
        Case vbKeyDecimal
               TranslateKey = "Decimal"
        Case vbKeyDivide
               TranslateKey = "Divide"
        Case vbKeyF1
               TranslateKey = "F1"
        Case vbKeyF2
               TranslateKey = "F2"
        Case vbKeyF3
               TranslateKey = "F3"
        Case vbKeyF4
               TranslateKey = "F4"
        Case vbKeyF5
               TranslateKey = "F5"
        Case vbKeyF6
               TranslateKey = "F6"
        Case vbKeyF7
               TranslateKey = "F7"
        Case vbKeyF8
               TranslateKey = "F8"
        Case vbKeyF9
               TranslateKey = "F9"
        Case vbKeyF10
               TranslateKey = "F10"
        Case vbKeyF11
               TranslateKey = "F11"
        Case vbKeyF12
               TranslateKey = "F12"
        Case vbKeyF13
               TranslateKey = "F13"
        Case vbKeyF14
               TranslateKey = "F14"
        Case vbKeyF15
               TranslateKey = "F15"
        Case vbKeyF16
               TranslateKey = "F16"
        Case vbKeyScrollLock
               TranslateKey = "ScrollLock"
        Case Else
               TranslateKey = Chr(Key)
   End Select
    

   ' @@@@@@@@@@@@@@@@@@@@@@@@@
   ' Select Case Chr(Key)
   '        Case "!"
   '             TranslateKey = "!"
   ' End Select
    
TranslateKeyExit:
    On Error Resume Next

    Exit Function
    
TranslateKeyErr:

    '//0x0ffee16 number is a unique number for TranslateKey because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    SysError Err, "0x0ffee16"

GoTo TranslateKeyExit

End Function
