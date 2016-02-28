Attribute VB_Name = "ModAPICalls"
Option Explicit

Public Const MEM_RESERVE = &H2000
Public Const MEM_COMMIT = &H1000
Public Const MEM_RELEASE = &H8000

Public Const TTF_TRACK = &H20
Public Const TTF_IDISHWND = &H1
Public Const TTF_CENTERTIP = &H2

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SECTION_QUERY = &H1
Public Const SECTION_MAP_WRITE = &H2
Public Const SECTION_MAP_READ = &H4
Public Const SECTION_MAP_EXECUTE = &H8
Public Const SECTION_EXTEND_SIZE = &H10
Public Const SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE
Public Const FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS

'============NT Shared memory staff======================
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Const PROCESS_VM_OPERATION = &H8
Public Const PROCESS_VM_READ = &H10
Public Const PROCESS_VM_WRITE = &H20
Public Const PROCESS_ALL_ACCESS = 0
Public Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Public Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'==========Memory access constants===========
Public Const PAGE_NOACCESS = &H1&
Public Const PAGE_READONLY = &H2&
Public Const PAGE_READWRITE = &H4&
Public Const PAGE_WRITECOPY = &H8&
Public Const PAGE_EXECUTE = &H10&
Public Const PAGE_EXECUTE_READ = &H20&
Public Const PAGE_EXECUTE_READWRITE = &H40&
Public Const PAGE_EXECUTE_WRITECOPY = &H80&
Public Const PAGE_GUARD = &H100&
Public Const PAGE_NOCACHE = &H200&


Public Const TTS_NOPREFIX = 2
Public Const GWL_STYLE = (-16)
Public Const TTS_BALLOON = &H40
Public Const GWL_HINSTANCE = (-6)
Public Const TTS_ALWAYSTIP = &H1

'  ----- Symbols -----
Public Const VS_VERSION_INFO = 1
Public Const VS_USER_DEFINED = 100
 
'  ----- VS_VERSION.dwFileFlags -----
Public Const VS_FFI_SIGNATURE = &HFEEF04BD
Public Const VS_FFI_STRUCVERSION = &H10000
Public Const VS_FFI_FILEFLAGSMASK = &H3F&

'  ----- VS_VERSION.dwFileFlags -----
Public Const VS_FF_DEBUG = &H1&
Public Const VS_FF_PRERELEASE = &H2&
Public Const VS_FF_PATCHED = &H4&
Public Const VS_FF_PRIVATEBUILD = &H8&
Public Const VS_FF_INFOINFERRED = &H10&
Public Const VS_FF_SPECIALBUILD = &H20&

'  ----- VS_VERSION.dwFileOS -----
Public Const VOS_UNKNOWN = &H0&
Public Const VOS_DOS = &H10000
Public Const VOS_OS216 = &H20000
Public Const VOS_OS232 = &H30000
Public Const VOS_NT = &H40000

Public Const VOS__BASE = &H0&
Public Const VOS__WINDOWS16 = &H1&
Public Const VOS__PM16 = &H2&
Public Const VOS__PM32 = &H3&
Public Const VOS__WINDOWS32 = &H4&

Public Const VOS_DOS_WINDOWS16 = &H10001
Public Const VOS_DOS_WINDOWS32 = &H10004
Public Const VOS_OS216_PM16 = &H20002
Public Const VOS_OS232_PM32 = &H30003
Public Const VOS_NT_WINDOWS32 = &H40004

'  ----- VS_VERSION.dwFileType -----
Public Const VFT_UNKNOWN = &H0&
Public Const VFT_APP = &H1&
Public Const VFT_DLL = &H2&
Public Const VFT_DRV = &H3&
Public Const VFT_FONT = &H4&
Public Const VFT_VXD = &H5&
Public Const VFT_STATIC_LIB = &H7&

'  ----- VS_VERSION.dwFileSubtype for VFT_WINDOWS_DRV -----
Public Const VFT2_UNKNOWN = &H0&
Public Const VFT2_DRV_PRINTER = &H1&
Public Const VFT2_DRV_KEYBOARD = &H2&
Public Const VFT2_DRV_LANGUAGE = &H3&
Public Const VFT2_DRV_DISPLAY = &H4&
Public Const VFT2_DRV_MOUSE = &H5&
Public Const VFT2_DRV_NETWORK = &H6&
Public Const VFT2_DRV_SYSTEM = &H7&
Public Const VFT2_DRV_INSTALLABLE = &H8&
Public Const VFT2_DRV_SOUND = &H9&
Public Const VFT2_DRV_COMM = &HA&
Public Const VFT2_DRV_INPUTMETHOD = &HB&

'  ----- VS_VERSION.dwFileSubtype for VFT_WINDOWS_FONT -----
Public Const VFT2_FONT_RASTER = &H1&
Public Const VFT2_FONT_VECTOR = &H2&
Public Const VFT2_FONT_TRUETYPE = &H3&

'  ----- VerFindFile() flags -----
Public Const VFFF_ISSHAREDFILE = &H1

Public Const VFF_CURNEDEST = &H1
Public Const VFF_FILEINUSE = &H2
Public Const VFF_BUFFTOOSMALL = &H4

'  ----- VerInstallFile() flags -----
Public Const VIFF_FORCEINSTALL = &H1
Public Const VIFF_DONTDELETEOLD = &H2

Public Const VIF_TEMPFILE = &H1&
Public Const VIF_MISMATCH = &H2&
Public Const VIF_SRCOLD = &H4&

Public Const VIF_DIFFLANG = &H8&
Public Const VIF_DIFFCODEPG = &H10&
Public Const VIF_DIFFTYPE = &H20&

Public Const VIF_WRITEPROT = &H40&
Public Const VIF_FILEINUSE = &H80&
Public Const VIF_OUTOFSPACE = &H100&
Public Const VIF_ACCESSVIOLATION = &H200&
Public Const VIF_SHARINGVIOLATION = &H400&
Public Const VIF_CANNOTCREATE = &H800&
Public Const VIF_CANNOTDELETE = &H1000&
Public Const VIF_CANNOTRENAME = &H2000&
Public Const VIF_CANNOTDELETECUR = &H4000&
Public Const VIF_OUTOFMEMORY = &H8000&

Public Const VIF_CANNOTREADSRC = &H10000
Public Const VIF_CANNOTREADDST = &H20000
Public Const VIF_BUFFTOOSMALL = &H40000

Public Const ICC_BAR_CLASSES = &H4

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const GWL_WNDPROC = (-4)
Public Const WM_USER As Long = &H400
Public Const TB_GETBUTTON As Long = (WM_USER + 23)
Public Const TB_BUTTONCOUNT As Long = (WM_USER + 24)
Public Const WM_MYHOOK As Long = WM_USER + 2
Public Const TTM_GETTOOLCOUNT = (WM_USER + 13)
Public Const TTM_ENUMTOOLSA = (WM_USER + 14)
Public Const TTM_ENUMTOOLSW = (WM_USER + 58)
Public Const TTM_GETTEXTA = (WM_USER + 11)
Public Const TTM_GETTEXTW = (WM_USER + 56)
Public Const TTM_ADDTOOLA = (WM_USER + 4)
Public Const TTM_TRACKACTIVATE = WM_USER + 17
Public Const TTM_TRACKPOSITION = WM_USER + 18

Public Const WM_LBUTTONUP = &H202
Public Const WM_MOUSEMOVE = &H200
Public Const WM_MBUTTONUP = &H208
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_RBUTTONDBLCLK = &H206
Public Const TTM_SETTITLE = (WM_USER + 32)
Public Const WM_BALLOONSHOW = (WM_USER + 2)
Public Const WM_BALLOONHIDE = (WM_USER + 3)
Public Const WM_BALLOONRCLK = (WM_USER + 4)
Public Const WM_BALLOONLCLK = (WM_USER + 5)
Public Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Public Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Public Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Public Const TOOLTIPS_CLASSA = "tooltips_class32"

Public Const NOTIFYICON_VERSION = 3

Public Const TH32CS_SNAPPROCESS As Long = &H2

Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
'  ----- Types and structures -----

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type TOOLINFO
    cbSize      As Long
    uFlags      As Long
    hwnd        As Long
    uId         As Long
    cRect       As RECT
    hInst       As Long
    lpszText    As Long 'LPCSTR
    lParam      As Long
End Type

Public Type TOOLTEXT
    sTipText As String * 80
End Type

Public Type udtTrayList
  trTipText As String
  trToolInfo As TOOLINFO
End Type

Public Type VS_FIXEDFILEINFO
        dwSignature As Long
        dwStrucVersion As Long         '  e.g. 0x00000042 = "0.42"
        dwFileVersionMS As Long        '  e.g. 0x00030075 = "3.75"
        dwFileVersionLS As Long        '  e.g. 0x00000031 = "0.31"
        dwProductVersionMS As Long     '  e.g. 0x00030010 = "3.10"
        dwProductVersionLS As Long     '  e.g. 0x00000031 = "0.31"
        dwFileFlagsMask As Long        '  = 0x3F for version "0.42"
        dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
        dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
        dwFileType As Long             '  e.g. VFT_DRIVER
        dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
        dwFileDateMS As Long           '  e.g. 0
        dwFileDateLS As Long           '  e.g. 0
End Type

Public Enum EErrorWindowProc
  eeAlreadyAttached = 13080 ' Message already handled by another class
  eeInvalidWindow           ' Invalid window
  eeNoExternalWindow        ' Can't modify external window
  eeNotInSystray
  eeCantSubClass
  eeCantShowIcon
  eeAlreayhandle
  eeWrongInterval
  eeTimerOverFlow
  eeAlreadyAnimated
  eeNoImageFound
  eeInvalidImgLst
  eeNoAnimationFound
  eeInvalidIcon
  eeCantShowBallon
  eeExternalError
End Enum

Public Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Public Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uId As Long
  uFlags As Long
  ucallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type

Public Enum enProductType
  VER_NT_WORKSTATION = &H1
  VER_NT_DOMAIN_CONTROLLER = &H2
  VER_NT_SERVER = &H3
End Enum

Public Enum enSuiteMask
  VER_SUITE_SMALLBUSINESS = &H1
  VER_SUITE_ENTERPRISE = &H2
  VER_SUITE_BACKOFFICE = &H4
  VER_SUITE_COMMUNICATIONS = &H8
  VER_SUITE_TERMINAL = &H10
  VER_SUITE_SMALLBUSINESS_RESTRICTED = &H20
  VER_SUITE_EMBEDDEDNT = &H40
  VER_SUITE_DATACENTER = &H80
  VER_SUITE_SINGLEUSERTS = &H100
  VER_SUITE_PERSONAL = &H200
  VER_SUITE_BLADE = &H400
End Enum

Public Enum enPlatformID
  VER_PLATFORM_WIN32s = 0          'WINDOWS 3.1
  VER_PLATFORM_WIN32_WINDOWS = 1   'WINDOWS 9x
  VER_PLATFORM_WIN32_NT = 2        'WINDOWS NT, 2000, XP
End Enum


Public Type udtOSVERSIONINFO
    lMajorVersion As Long
    lMinorVersion As Long
    lBuildNumber As Long
    lPlatformId As enPlatformID
    sCSDVersion As String
    iServicePackMajor As Integer
    iServicePackMinor As Integer
    iSuiteMask As enSuiteMask
    bytProductType As enProductType
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type OSVERSIONINFOEX
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
  wServicePackMajor As Integer
  wServicePackMinor As Integer
  wSuiteMask As Long
  wProductType As Byte
  wReserved As Byte
End Type

Public Type tagINITCOMMONCONTROLSEX
   dwSize As Long   ' size of this structure
   dwICC As Long    ' flags indicating which classes to be
                    ' initialized.
End Type

Public Type POINTAPI
    X As Long
    y As Long
End Type


Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

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


Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function EndMenu Lib "user32.dll" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Public Declare Function GetSubMenu Lib "user32" (ByVal hmenu As Long, ByVal nPos As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Public Declare Function GetVersion Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function Process32Next Lib "KERNEL32.dll" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Public Declare Function Process32First Lib "KERNEL32.dll" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
Public Declare Function GetVersionEx1 Lib "KERNEL32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "KERNEL32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef lpcbNeeded As Long) As Long
Public Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lpBuffer As Any, nVerSize As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hwnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hmenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Any) As Long
Public Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpFileMappigAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Public Declare Function VerFindFile Lib "version.dll" Alias "VerFindFileA" (ByVal uFlags As Long, ByVal szFileName As String, ByVal szWinDir As String, ByVal szAppDir As String, ByVal szCurDir As String, lpuCurDirLen As Long, ByVal szDestDir As String, lpuDestDirLen As Long) As Long
Public Declare Function VerInstallFile Lib "version.dll" Alias " VerInstallFileA" (ByVal uFlags As Long, ByVal szSrcFileName As String, ByVal szDestFileName As String, ByVal szSrcDir As String, ByVal szDestDir As String, ByVal szCurDir As String, ByVal szTmpFile As String, lpuTmpFileLen As Long) As Long
Public Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hmenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long






Public m_lSysTrayHWnd As Long
Public tTrayList() As udtTrayList
Public NWThreadID As Long, NWPid As Long



