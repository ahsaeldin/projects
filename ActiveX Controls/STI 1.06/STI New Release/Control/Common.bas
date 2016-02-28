Attribute VB_Name = "Common"
Option Explicit

Public Enum EErrorWindowProc
  eeExternalError = 13080
End Enum

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetVersionEx Lib "KERNEL32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Public Const VER_PLATFORM_WIN32_NT As Long = 2

Public Const TBSTATE_CHECKED As Long = &H1
Public Const TBSTATE_ELLIPSES As Long = &H40
Public Const TBSTATE_ENABLED As Long = &H4
Public Const TBSTATE_HIDDEN As Long = &H8
Public Const TBSTATE_INDETERMINATE As Long = &H10
Public Const TBSTATE_MARKED As Long = &H80
Public Const TBSTATE_PRESSED As Long = &H2
Public Const TBSTATE_WRAP As Long = &H20

Public Const TBSTYLE_ALTDRAG As Long = &H400
Public Const TBSTYLE_AUTOSIZE As Long = &H10
Public Const TBSTYLE_BUTTON As Long = &H0
Public Const TBSTYLE_CHECK As Long = &H2
Public Const TBSTYLE_GROUP As Long = &H4
Public Const TBSTYLE_CHECKGROUP As Long = (TBSTYLE_GROUP Or TBSTYLE_CHECK)
Public Const TBSTYLE_CUSTOMERASE As Long = &H2000
Public Const TBSTYLE_DROPDOWN As Long = &H8
Public Const TBSTYLE_EX_DOUBLEBUFFER As Long = &H80
Public Const TBSTYLE_EX_DRAWDDARROWS As Long = &H1
Public Const TBSTYLE_EX_HIDECLIPPEDBUTTONS As Long = &H10
Public Const TBSTYLE_EX_MIXEDBUTTONS As Long = &H8
Public Const TBSTYLE_FLAT As Long = &H800
Public Const TBSTYLE_LIST As Long = &H1000
Public Const TBSTYLE_NOPREFIX As Long = &H20
Public Const TBSTYLE_REGISTERDROP As Long = &H4000
Public Const TBSTYLE_SEP As Long = &H1
Public Const TBSTYLE_TOOLTIPS As Long = &H100
Public Const TBSTYLE_TRANSPARENT As Long = &H8000
Public Const TBSTYLE_WRAPABLE As Long = &H200

Public Const WM_USER As Long = &H400
Public Const TB_GETBUTTON As Long = (WM_USER + 23)
Public Const TB_BUTTONCOUNT As Long = (WM_USER + 24)

Public Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Public Const PROCESS_VM_OPERATION As Long = (&H8)
Public Const PROCESS_VM_READ As Long = (&H10)
Public Const PROCESS_VM_WRITE As Long = (&H20)
Public Const PROCESS_VM As Long = PROCESS_QUERY_INFORMATION Or PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE 'Or PROCESS_DUP_HANDLE
Public Const MEM_COMMIT As Long = &H1000
Public Const PAGE_READWRITE As Long = &H4
Public Const MEM_RELEASE As Long = &H8000

Public Type TBBUTTON
    iBitmap As Long
    idCommand As Long
    fsState As Byte
    fsStyle As Byte
    bReserved(1 To 2) As Byte
    dwData As Long
    iString As Long
End Type

Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hwnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Public Declare Function OpenProcess Lib "KERNEL32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CloseHandle Lib "KERNEL32.dll" (ByVal hObject As Long) As Long

Public Declare Function VirtualAllocEx Lib "KERNEL32.dll" (ByVal hProcess As Long, lpAddress As Any, ByRef dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function VirtualFreeEx Lib "KERNEL32.dll" (ByVal hProcess As Long, lpAddress As Any, ByRef dwSize As Long, ByVal dwFreeType As Long) As Long

Public Declare Function ReadProcessMemory Lib "KERNEL32.dll" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function WriteProcessMemory Lib "KERNEL32.dll" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Public Declare Function DrawIcon Lib "user32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long


Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Type NotifyEntry
    dwImageIndex As Long
    NotifyDataA As NOTIFYICONDATA
    
    Path As String
End Type

Public Type NotifyListHeader
    dwCount As Long
    lpNotifyDataList() As NotifyEntry
End Type

Public Type Systray
    hTrayNotifyWnd As Long
    hTrayClockWClass As Long
    hToolTips As Long
    dxNotify As Long
    dyNotify As Long
    dwNotifyWidth As Long
    dwNotifyHeight As Long
    lpsysListHeader As NotifyListHeader
    hicons As Long
    bkColor As Long
    dwImageCount As Long
End Type


Public Type TRAY_ME
    hwnd As Long
    uId As Long
    ucallbackMessage As Long
    Unknown1(1 To 2) As Long
    hIcon As Long
    Unknown2(1 To 3) As Long
    Path As String * 255
    sTip As String * 255
    iBitmap As Long
    idCommand As Long
    fsState As Byte
    fsStyle As Byte
    dwData As Long
    iString As Long
End Type

Dim bybass_Error As Boolean

Public Function IsWindowsNT() As Boolean
    Dim os As OSVERSIONINFO
    
    os.dwOSVersionInfoSize = Len(os)
    GetVersionEx os

    IsWindowsNT = (os.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Public Function IsWindowsME() As Boolean
    Dim os As OSVERSIONINFO

    os.dwOSVersionInfoSize = Len(os)
    GetVersionEx os

    IsWindowsME = ((os.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS) And (os.dwMajorVersion = 4) And (os.dwMinorVersion = 90))
End Function


Public Function GethWndTray() As Long
Dim hwnd As Long, hwnd2 As Long
hwnd = FindWindow("Shell_TrayWnd", vbNullString)
hwnd = FindWindowEx(hwnd, ByVal 0&, "TrayNotifyWnd", vbNullString)
hwnd2 = FindWindowEx(hwnd, ByVal 0&, "SysPager", vbNullString)
If (hwnd2 = 0) Then hwnd2 = hwnd ' ME,2000
hwnd2 = FindWindowEx(hwnd2, ByVal 0&, "ToolbarWindow32", vbNullString) ' ME, 2000, XP...
If (hwnd2 = 0) Then
    GethWndTray = hwnd ' 95,98
Else
    GethWndTray = hwnd2  ' ME, 2000, XP...
End If
End Function
Public Sub ErrRaise(e As Long, Optional objectName As String)
Dim sText As String
  
   If e > 1000 Then
      Select Case e
      Case eeExternalError
           sText = objectName & " Fail::" & "External Error,Can't Get System Tray Handle"
      End Select
      MsgBox sText, vbCritical
   Else
      ' Raise standard Visual Basic error
      MsgBox sText, vbCritical
   End If
End Sub

Public Property Get bybassError() As Boolean
 bybassError = bybass_Error
End Property

Public Property Let bybassError(ByVal vNewValue As Boolean)
 bybass_Error = vNewValue
End Property

