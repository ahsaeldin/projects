Option Explicit On

Module Win32API


    Public Const VK_LWIN = &H5B
    Public Const KEYEVENTF_KEYUP = &H2
    Public Const VK_F2 = &H71
    Public Const MOD_ALT = &H1
    Public Const SW_RESTORE = 9
    Public Const MOD_SHIFT = &H4
    Public Const VK_SHIFT = &H10
    Public Const MOD_CONTROL = &H2
    Public Const WM_HOTKEY = &H312
    Public Const VK_CONTROL = &H11
    Public Const HWND_TOPMOST = -1
    Public Const HWND_NOTOPMOST = -2
    Public Const HWND_TOP As Long = 0
    Public Const SWP_SHOWWINDOW = &H40
    Public Const VK_MENU As Long = &H12
    Public Const SMTO_ABORTIFHUNG = &H2
    Public Const VK_LBUTTON As Long = &H1
    Public Const VK_DELETE As Long = &H2E
    Public Const WM_COPYDATA As Integer = &H4A
    Public Const SWP_HIDEWINDOW As Long = &H80

    Public Structure RECT
        Dim Left As Integer
        Dim Top As Integer
        Dim Right As Integer
        Dim Bottom As Integer
    End Structure

    <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)> _
    Public Structure COPYDATASTRUCT
        Public dwData As IntPtr
        Public cbData As Integer
        Public lpData As IntPtr
    End Structure

    Public Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Integer, ByVal wBar As Integer, ByVal bShow As Integer) As Integer

    Public Declare Function GetParent Lib "user32.dll" (ByVal hwnd As IntPtr) As Integer

    Public Declare Sub SwitchToThisWindow Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal fAltTab As Boolean)

    Public Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Int32, _
                                                      ByVal id As Int32, _
                                                      ByVal fsModifiers As Int32, _
                                                      ByVal vk As Int32) As Boolean

    Public Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Int32, ByVal id As Int32) As Boolean


    Public Declare Function GetForegroundWindow Lib "user32.dll" () As IntPtr
    Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
    Public Declare Function BlockInput Lib "user32" (ByVal fBlockIt As Boolean) As Boolean
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Integer
    Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As IntPtr) As Integer
    Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Long

    Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" _
    (ByVal hWnd As Int32) As Integer

    Public Declare Function GetWindowRect Lib "User32" Alias "GetWindowRect" _
    (ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Int32

    Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" _
    (ByVal hWnd As IntPtr, ByVal lpString As String, ByVal cch As Integer) As Integer

    Public Declare Sub SetWindowPos Lib "User32" _
    (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, _
     ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)

    Public Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Integer, _
                                                              ByVal yPoint As Integer) As Integer

    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                                                                         ByVal lpWindowName As String) As IntPtr

    '//2 overloads of SendMessage
    Public Declare Auto Function SendMessage Lib "user32.dll" (ByVal hWnd As IntPtr, _
                                                    ByVal msg As Integer, ByVal wParam As IntPtr, _
                                                    ByVal lParam As IntPtr) As IntPtr

    Public Declare Auto Function SendMessage Lib "user32" ( _
    ByVal hwnd As IntPtr, _
    ByVal wMsg As Integer, _
    ByVal wParam As Integer, _
    ByRef lParam As COPYDATASTRUCT _
    ) As Integer

    Public Declare Function SendMessageTimeout _
          Lib "user32" Alias "SendMessageTimeoutA" _
          (ByVal hwnd As IntPtr, _
          ByVal msg As Integer, _
          ByVal wParam As Integer, _
          ByVal lParam As Integer, _
          ByVal fuFlags As Integer, _
          ByVal uTimeout As Integer, _
          ByVal lpdwResult As Integer) As Integer


    Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
    ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

    '28/09/2010 FoxSmr "First use of implicit continuation"
    Public Declare Auto Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As IntPtr,
                                                                    ByRef lpdwProcessId As IntPtr
                                                                   ) As IntPtr
End Module
