Attribute VB_Name = "ModUtils"
Option Explicit

Public Type TBBUTTON
    iBitmap As Long
    idCommand As Long
    fsState As Byte
    fsStyle As Byte
    bReserved(1 To 2) As Byte
    dwData As Long
    iString As Long
End Type

Public Type RawNotifyListHeader
    dwCount As Long
    lpNotifyDataList As Long
End Type

Public Type RawSystray
    hTrayNotifyWnd As Long
    hTrayClockWClass As Long
    hToolTips As Long
    dxNotify As Long
    dyNotify As Long
    dwNotifyWidth As Long
    dwNotifyHeight As Long
    lpsysListHeader As Long
    hicons As Long
    bkColor As Long
    dwImageCount As Long
End Type

Private Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Const NOTIFYICONDATA_V1_SIZE As Long = 88  'pre-5.0 structure size
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488 'pre-6.0 structure size
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure size

Public IconNumber As Long
Dim bybass_Error As Boolean

Public NOTIFYICONDATA_SIZE As Long
Public TrayCollection As New Collection
Public BalloonCollection As New Collection
Public SubClassCollection As New Collection
 
Public tOSVERSIONINFO As udtOSVERSIONINFO

Private m_pUdtOSVersion As OSVERSIONINFO
Private m_pUdtOSVersionEX As OSVERSIONINFOEX

Public Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Public Const PROCESS_VM_OPERATION As Long = (&H8)
Public Const PROCESS_VM_READ As Long = (&H10)
Public Const PROCESS_VM_WRITE As Long = (&H20)


Public Const PROCESS_VM As Long = PROCESS_QUERY_INFORMATION Or PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE 'Or PROCESS_DUP_HANDLE

Private Last_Error As String
Public NoIconFound As Boolean
Private Last_ErrorForMe As String
 
 '//cg 18April2008: Do you rembmember the error of ECW and "daniel.roesler@convergit.de"=order number for ASTC is U27236360"
'//cg 18Apr2008: added next 2 lines to use in "GetTrayIcons" method to solve the problem of the overflow caused by frequntly using of "VirtualAllocEx"
Public hProc As Long
Public ptrVM As Long

Public Sub ErrRaise(e As Long, Optional objectName As String, Optional objectName2 As String)
Dim AlreadySubClassed As Boolean
Dim sText As String
   If e > 1000 Then
      Select Case e
      Case eeExternalError
         sText = objectName & " Fail::" & "External Error,Can't Get System Tray Handle"
      Case eeAlreadyAttached
         sText = "[" & objectName & " error]" & " Window already Subclassed by " & objectName2 & " Control."
         AlreadySubClassed = True
      Case eeInvalidWindow
         sText = "'" & objectName & "'" & " Invalid window"
      Case eeNoExternalWindow
         sText = "'" & objectName & "'" & "Can't handle external window"
      Case eeNotInSystray
         sText = objectName & " Fail::" & "Icon isn't in the System Tray"
      Case eeCantSubClass
         sText = "Subclassing Error by" & "'" & objectName & "'" & " control"
      Case eeCantShowIcon
         sText = objectName & " Fail::" & "Can't Show the icon"
      Case eeAlreayhandle
          sText = objectName & " already handle a window"
      Case eeTimerOverFlow
          sText = objectName & " Fail::" & "Interval can't exceed 65,535 milliseconds"
      Case eeWrongInterval
          sText = objectName & " Fail::" & "Wrong Interval"
      Case eeAlreadyAnimated
          sText = objectName & " Fail::" & "Icon is already animated in System Tray"
      Case eeNoImageFound
          sText = objectName & " Fail::" & "No images Found in the ImageList"
       Case eeInvalidImgLst
          sText = objectName & " Fail::" & "Invalid ImageList"
       Case eeNoAnimationFound
          sText = objectName & " Fail::" & "Icon isn't animated"
       Case eeInvalidIcon
          sText = objectName & " Error::" & "Invalid Icon"
       Case eeCantShowBallon
          sText = objectName & " Fail::" & "Can't show Balloon while the icon is animated"
     End Select
      LastError = sText
      If AlreadySubClassed Then MsgBox sText
   Else
      ' Raise standard Visual Basic error
      LastError = sText
   End If
End Sub

Public Property Get LastError() As Variant
LastError = Last_Error
End Property

Public Property Let LastError(ByVal vNewValue As Variant)
Last_Error = vNewValue
End Property

Public Property Get LastErrorForMe() As Variant
LastErrorForMe = Last_ErrorForMe
End Property

Public Property Let LastErrorForMe(ByVal vNewValue As Variant)
Last_ErrorForMe = vNewValue
End Property

Public Function IsWindowLocal(ByVal hwnd As Long) As Boolean
    Dim idWnd As Long
    Call GetWindowThreadProcessId(hwnd, idWnd)
    IsWindowLocal = (idWnd = GetCurrentProcessId())
End Function
Public Sub SetShellVersion()

  
      If IsShellVersion(6) Then
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE '6.0 structure size
      
      ElseIf IsShellVersion(5) Then
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE 'pre-6.0 structure size
      
      Else
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE 'pre-5.0 structure size
      
      End If

End Sub
Private Function IsShellVersion(ByVal version As Long) As Boolean

  'returns True if the Shell version
  '(shell32.dll) is equal or later than
  'the value passed as 'version'
   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim nVerMajor As Integer
   Dim bBuffer() As Byte
   
   Const sDLLFile As String = "shell32.dll"
   
   nBufferSize = GetFileVersionInfoSize(sDLLFile, nUnused)
   
   If nBufferSize > 0 Then
    
      ReDim bBuffer(nBufferSize - 1) As Byte
    
      Call GetFileVersionInfo(sDLLFile, 0&, nBufferSize, bBuffer(0))
    
      If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
         
         CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
        
         IsShellVersion = nVerMajor >= version
      
      End If  'VerQueryValue
    
   End If  'nBufferSize
  
End Function

Public Property Get IsAttached(ByVal hwnd As Long) As Boolean
Dim sName As String
   sName = hwnd & "#" & "AttachState"
   IsAttached = GetProp(hwnd, sName)
End Property

Public Property Let IsAttached(ByVal hwnd As Long, ByVal State As Boolean)
Dim m_f As Long
Dim sName As String
   sName = hwnd & "#" & "AttachState"
   m_f = SetProp(hwnd, sName, State)
   If m_f = 0 Then Last_ErrorForMe = "Error  Let IsAttached"
   If (State = False) Then
      RemoveProp hwnd, sName
   End If
   
End Property

Public Sub InitComctl32(dwFlags As Long)
    
   Dim CommControlsVersion As String
    
   CommControlsVersion = CommonControlsVersion
    
   On Error Resume Next
   
   Dim icc As tagINITCOMMONCONTROLSEX
   Dim lRetVal As Long
   
   icc.dwSize = Len(icc)
   icc.dwICC = dwFlags
   lRetVal = InitCommonControlsEx(icc)

   If lRetVal = 0 Then InitCommonControls

End Sub
Public Property Get CommonControlsVersion() As String

  Dim sVersion() As String
  
  CommonControlsVersion = FileVersion("COMCTL32.DLL")
   
  sVersion = Split(CommonControlsVersion, ".")
  
  If Val(sVersion(0)) < 5 Then
    MsgBox "Download MSCommon Controls DLL plz"
  ElseIf Val(sVersion(0)) = 5 And Val(sVersion(1)) < 80 Then
    MsgBox "Download MSCommon Controls DLL plz"
  End If
  
End Property
Private Function FileVersion(ByVal sFile As String) As String

   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim nVerMajor As Integer
   Dim bBuffer() As Byte
   Dim sBuffer As String
   Dim tFixedFileInfo As VS_FIXEDFILEINFO
   Dim i As Integer, sVer As String
   Dim lRet As Long
   
   
   nBufferSize = GetFileVersionInfoSize(sFile, nUnused)
   
   If nBufferSize > 0 Then
    
      ReDim bBuffer(nBufferSize - 1) As Byte
    
      Call GetFileVersionInfo(sFile, 0&, nBufferSize, bBuffer(0))
      
      If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
         
         CopyMemory tFixedFileInfo, ByVal lpBuffer, Len(tFixedFileInfo)
         
         If tFixedFileInfo.dwFileVersionMS > 0 Then
         
           sVer = HiWord(tFixedFileInfo.dwFileVersionMS) & "." & LoWord(tFixedFileInfo.dwFileVersionMS) & "." _
               & HiWord(tFixedFileInfo.dwFileVersionLS) & "." & LoWord(tFixedFileInfo.dwFileVersionLS)
         
         End If
         
         FileVersion = sVer
      
      End If  'VerQueryValue
    
   End If  'nBufferSize
  
End Function

Private Function HiWord(ByVal wParam As Long) As Integer
   HiWord = (wParam And &HFFFF0000) \ (&H10000)

End Function

Private Function LoWord(ByVal wParam As Long) As Integer
   LoWord = wParam And &HFFFF&

End Function

Private Function HiByte(ByVal wParam As Integer) As Byte
   
   HiByte = (wParam And &HFF00&) \ (&H100)

End Function

Private Function LoByte(ByVal wParam As Integer) As Byte

   LoByte = wParam And &HFF&

End Function



Public Sub GetOSVersion()

  On Error GoTo ErrorHandler
  Dim pUdtOSVersion As OSVERSIONINFO
  Dim pUdtOSVersionEX As OSVERSIONINFOEX
  Dim pOSVERSIONINFO As udtOSVERSIONINFO
  Dim plMajorVersion  As Long
  Dim plMinorVersion As Long
  Dim plPlatformID As Long
  Dim sCSDVer As String
  Dim lRetVal As Long
    
    m_pUdtOSVersion = pUdtOSVersion
    m_pUdtOSVersionEX = pUdtOSVersionEX
    
    pUdtOSVersion.dwOSVersionInfoSize = Len(pUdtOSVersion)
    lRetVal = GetVersion(pUdtOSVersion)
    If lRetVal = 0 Then GoTo lblEnd
    m_pUdtOSVersion = pUdtOSVersion
    
    plMajorVersion = pUdtOSVersion.dwMajorVersion
    plMinorVersion = pUdtOSVersion.dwMinorVersion
    plPlatformID = pUdtOSVersion.dwPlatformId
    sCSDVer = UCase(pUdtOSVersion.szCSDVersion)
    
    If plMajorVersion = 5 Then
      pUdtOSVersionEX.dwOSVersionInfoSize = Len(pUdtOSVersionEX)
      lRetVal = GetVersionEx(pUdtOSVersionEX)
      If lRetVal <> 0 Then m_pUdtOSVersionEX = pUdtOSVersionEX
    End If

With pOSVERSIONINFO
  .lBuildNumber = pUdtOSVersion.dwBuildNumber
  .lMajorVersion = pUdtOSVersion.dwMajorVersion
  .lMinorVersion = pUdtOSVersion.dwMinorVersion
  .lPlatformId = pUdtOSVersion.dwPlatformId
  .sCSDVersion = fNullCut(pUdtOSVersion.szCSDVersion)
  .iServicePackMajor = pUdtOSVersionEX.wServicePackMajor
  .iServicePackMinor = pUdtOSVersionEX.wServicePackMinor
  .iSuiteMask = pUdtOSVersionEX.wSuiteMask
  .bytProductType = pUdtOSVersionEX.wProductType
End With


lblEnd:
tOSVERSIONINFO = pOSVERSIONINFO

lblExit:
  Exit Sub

ErrorHandler:
  tOSVERSIONINFO = pOSVERSIONINFO
  Resume lblExit
  
End Sub

Private Function fNullCut(ByVal myString As String) As String
  Dim i As Long
  i = InStr(myString, vbNullChar)
  If i > 0& Then
    fNullCut = Trim$(Left$(myString, i - 1&))
  Else
    fNullCut = Trim$(myString)
  End If
End Function

Public Function OSSupportsBalloons() As Boolean
  If tOSVERSIONINFO.lPlatformId = VER_PLATFORM_WIN32_NT Then
    If tOSVERSIONINFO.lMajorVersion >= 5 Then
      OSSupportsBalloons = True
    End If
  ElseIf tOSVERSIONINFO.lPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    If tOSVERSIONINFO.lMajorVersion >= 4 And tOSVERSIONINFO.lMinorVersion = 90 Then
      OSSupportsBalloons = True
    End If
  End If
End Function

Private Sub GetTrayList()
   ReDim tTrayList(0)
   NWThreadID = 0
   NWPid = 0
   NWThreadID = GetWindowThreadProcessId(GetTrayNotifyWnd, NWPid)
   EnumWindows AddressOf EnumWinProc, 0
   
End Sub

Private Function GetTrayNotifyWnd() As Long
   GetTrayNotifyWnd = FindWindowEx(FindWindow("Shell_TrayWnd", vbNullString), 0, "TrayNotifyWnd", vbNullString)
End Function

Private Function EnumWinProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
   Dim pid As Long, tid As Long, lStyle As Long
   Dim hProcess As Long, nCount As Long, lWritten As Long, i As Long
   Dim lpSysShared As Long, hFileMapping As Long, dwSize As Long
   Dim lpSysShared2 As Long, hFileMapping2 As Long
   Dim h As Long
   Dim ti As TOOLINFO
   Dim tt As TOOLTEXT
   Dim rc As RECT
   Static sTip As String
   
   tid = GetWindowThreadProcessId(hwnd, pid)
   lStyle = GetWindowLong(hwnd, GWL_STYLE)
   
   If pid = NWPid And GetWndClass(hwnd) = "tooltips_class32" And (lStyle And TTS_NOPREFIX) <> TTS_NOPREFIX Then
      nCount = SendMessage(hwnd, TTM_GETTOOLCOUNT, 0&, ByVal 0&)
      If nCount <= 0 Then
        EnumWinProc = 0
        Exit Function
      End If
      
      ReDim tTrayList(nCount - 1)
       ti.cbSize = Len(ti)
      dwSize = ti.cbSize
      
    '//cprin 23May2007: added the part of "If IsWindowsNT Then" in support of solving the problem of Roy Kaplan roy.kaplan@gmail.com
    '//concerning trying to put a message in the balloon more than 256 char
    If IsWindowsNT Then
    
             lpSysShared = GetMemSharedNT(pid, dwSize, hProcess)
             If lpSysShared = 0 Then GoTo lblEnd
             lpSysShared2 = GetMemSharedNT(pid, LenB(tt), hProcess)
             If lpSysShared2 = 0 Then GoTo lblEnd
             WriteProcessMemory hProcess, ByVal lpSysShared, ti, dwSize, lWritten
             
             For i = 0 To nCount - 1
                 tt.sTipText = String(80, Chr(0))
                 WriteProcessMemory hProcess, ByVal lpSysShared2, tt, LenB(tt), lWritten
                 Call SendMessage(hwnd, TTM_ENUMTOOLSW, i, ByVal lpSysShared)
                 ReadProcessMemory hProcess, ByVal lpSysShared, ti, dwSize, lWritten
                 ti.lpszText = lpSysShared2
                 WriteProcessMemory hProcess, ByVal lpSysShared, ti, dwSize, lWritten
                 Call SendMessage(hwnd, TTM_GETTEXTW, 0&, ByVal lpSysShared)
                 ReadProcessMemory hProcess, ByVal lpSysShared2, tt, LenB(tt), lWritten
                 
                 tTrayList(i).trToolInfo = ti
                 
                 sTip = StrConv(tt.sTipText, vbFromUnicode)
                 sTip = StripTerminator(sTip)
                 'Debug.Print "sTip=" & sTip & ", hwnd=" & hwnd
                 tTrayList(i).trTipText = sTip
                 If i = nCount - 1 Then Exit For
             Next i
             
             FreeMemSharedNT hProcess, lpSysShared, dwSize
             FreeMemSharedNT hProcess, lpSysShared2, LenB(tt)
    
    Else
            lpSysShared = GetMemShared95(dwSize, hFileMapping)
             If lpSysShared = 0 Then GoTo lblEnd
             lpSysShared2 = GetMemShared95(Len(tt), hFileMapping2)
             If lpSysShared2 = 0 Then GoTo lblEnd
      
             CopyMemory ByVal lpSysShared, ti, dwSize
             For i = 0 To nCount - 1
                 tt.sTipText = String(80, Chr(0))
                 CopyMemory ByVal lpSysShared2, tt, Len(tt)
                 Call SendMessage(hwnd, TTM_ENUMTOOLSA, i, ByVal lpSysShared)
                 CopyMemory ti, ByVal lpSysShared, dwSize
                 ti.lpszText = lpSysShared2
                 CopyMemory ByVal lpSysShared, ti, dwSize
                 Call SendMessage(hwnd, TTM_GETTEXTA, 0&, ByVal lpSysShared)
                 CopyMemory tt, ByVal lpSysShared2, Len(tt)
      
                 tTrayList(i).trToolInfo = ti
                 sTip = StripTerminator(tt.sTipText)
                 tTrayList(i).trTipText = sTip
                 If i = nCount - 1 Then Exit For
             Next i
             FreeMemShared95 hFileMapping, lpSysShared
             FreeMemShared95 hFileMapping2, lpSysShared2
      
    End If
lblEnd:
      EnumWinProc = 0
      Exit Function
   
   End If
   EnumWinProc = 1
End Function
Private Function GetWndClass(hwnd As Long) As String
   Dim k As Long, sName As String
   sName = Space$(128)
   k = GetClassName(hwnd, sName, 128)
   If k > 0 Then sName = Left$(sName, k) Else sName = ""
   GetWndClass = sName
End Function
Public Function GetTrayIconRect() As RECT
    
    Dim sTrayListArray() As String
    Dim rctSysTray As RECT
    Dim rctTrayIcon As RECT
    Dim i As Integer, bIconFound As Boolean
    
  
    m_lSysTrayHWnd = GetSysTrayHWnd()
    If m_lSysTrayHWnd = 0 Then GoTo lblExit
    Call GetWindowRect(m_lSysTrayHWnd, rctSysTray)
    
    Call GetTrayList
    If UBound(tTrayList) = 0 And tTrayList(0).trToolInfo.hwnd = 0 Then GoTo lblExit
    
    For i = 0 To UBound(tTrayList)
      'Debug.Print i & "-'" & tTrayList(i).trTipText & "'"
      If tTrayList(i).trTipText = "cPrIn" & App.hInstance Then
        bIconFound = True
        Exit For
      End If
    Next i
    
    If bIconFound Then
      With tTrayList(i).trToolInfo.cRect
        rctTrayIcon.Left = rctSysTray.Left + .Left
        rctTrayIcon.Right = rctSysTray.Left + .Right
        rctTrayIcon.Top = rctSysTray.Top + .Top
        rctTrayIcon.Bottom = rctSysTray.Top + .Bottom
      End With
    
      GetTrayIconRect = rctTrayIcon
    Else
       GetTrayIconRect = rctSysTray
      ' MsgBox "Error Icon Not Found"
    End If
    
lblExit:

End Function

Public Function GetSysTrayHWnd() As Long
    Const GW_CHILD = 5
    Const GW_HWNDNEXT = 2
    '
    Dim lngTaskbarHwnd As Long
    Dim lHWnd As Long
    Dim lTrayHwnd As Long
    Dim strClassName As String * 250
    Dim sClassName As String
    Dim hInstance As Long
    Dim rc As RECT
    Dim n As Long, hIcon As Long
    
    '
    'Get taskbar handle
    lngTaskbarHwnd = FindWindow("Shell_TrayWnd", vbNullString)
    If lngTaskbarHwnd = 0 Then GoTo lblExit
    '
    'Get system tray handle
    'Call PirntText("-------TrayNotifyWnd-------")
    lHWnd = GetWindow(lngTaskbarHwnd, GW_CHILD)
    If lHWnd = 0 Then GoTo lblExit
    
    Do Until lHWnd = 0
        strClassName = String(250, Chr(0))
        GetClassName lHWnd, strClassName, 250
        sClassName = StripTerminator(strClassName)
        'Call PirntText("sClassName=" & sClassName & ", Hwnd=" & lHwnd)
        If sClassName = "TrayNotifyWnd" Then
          lTrayHwnd = lHWnd
          Exit Do
        End If
        lHWnd = GetWindow(lHWnd, GW_HWNDNEXT)
    Loop
    '
    'Call PirntText("-------SysPager-------")
    lHWnd = GetWindow(lTrayHwnd, GW_CHILD)
    If lHWnd = 0 Then GoTo lblExit
    
    Do Until lHWnd = 0
        strClassName = String(250, Chr(0))
        GetClassName lHWnd, strClassName, 250
        sClassName = StripTerminator(strClassName)
        'Call PirntText("sClassName=" & sClassName & ", Hwnd=" & lHwnd)
        If sClassName = "SysPager" Then
          lTrayHwnd = lHWnd
          Exit Do
        End If
        lHWnd = GetWindow(lHWnd, GW_HWNDNEXT)
    Loop
    
    'Call PirntText("-------ToolbarWindow32-------")
    lHWnd = GetWindow(lTrayHwnd, GW_CHILD)
    If lHWnd = 0 Then GoTo lblExit
    
    Do Until lHWnd = 0
        strClassName = String(250, Chr(0))
        GetClassName lHWnd, strClassName, 250
        sClassName = StripTerminator(strClassName)
        'Call PirntText("sClassName=" & sClassName & ", Hwnd=" & lHwnd)
        If sClassName = "ToolbarWindow32" Then
          lTrayHwnd = lHWnd
          Exit Do
        End If
        lHWnd = GetWindow(lHWnd, GW_HWNDNEXT)
    Loop
    

lblExit:
    GetSysTrayHWnd = lTrayHwnd

End Function

Private Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos = 1 Then
        StripTerminator = ""
    ElseIf intZeroPos > 1 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Public Function GetMemShared95(ByVal memSize As Long, hFile As Long) As Long
    hFile = CreateFileMapping(&HFFFFFFFF, 0, PAGE_READWRITE, 0, memSize, vbNullString)
    If hFile <> 0 Then
      GetMemShared95 = MapViewOfFile(hFile, FILE_MAP_ALL_ACCESS, 0, 0, 0)
    End If
End Function

Public Sub FreeMemShared95(ByVal hFile As Long, ByVal lpMem As Long)
    UnmapViewOfFile lpMem
    CloseHandle hFile
End Sub

Public Function GetMemSharedNT(ByVal pid As Long, ByVal memSize As Long, hProcess As Long) As Long
    hProcess = OpenProcess(PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE, False, pid)
    If hProcess <> 0 Then
      GetMemSharedNT = VirtualAllocEx(ByVal hProcess, ByVal 0&, ByVal memSize, MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
    End If
End Function

Public Sub FreeMemSharedNT(ByVal hProcess As Long, ByVal MemAddress As Long, ByVal memSize As Long)
   Call VirtualFreeEx(hProcess, ByVal MemAddress, memSize, MEM_RELEASE)
   CloseHandle hProcess
End Sub

Public Property Get bybassError() As Boolean
 bybassError = bybass_Error
End Property

Public Property Let bybassError(ByVal vNewValue As Boolean)
 bybass_Error = vNewValue
End Property


Public Function IsWindowsME() As Boolean
    Dim os As OSVERSIONINFO

    os.dwOSVersionInfoSize = Len(os)
    GetVersionEx1 os

    IsWindowsME = ((os.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS) And (os.dwMajorVersion = 4) And (os.dwMinorVersion = 90))
End Function
Public Function IsWindowsNT() As Boolean
    Dim os As OSVERSIONINFO
    
    os.dwOSVersionInfoSize = Len(os)
    GetVersionEx1 os

    IsWindowsNT = (os.dwPlatformId = VER_PLATFORM_WIN32_NT)
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

Public Sub GetProcessNameFromPID_NT(pid As Long, lpName As String)
Dim hP As Long
    
    hP = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, False, pid)

    If (hP) Then
        Dim hMod As Long
        Dim ret As Long
        
        EnumProcessModules hP, hMod, 4&, ret
        
        If (hMod) Then
            
            lpName = Space(255)
            
            GetModuleFileNameExA hP, hMod, lpName, 255
        
            lpName = Left$(lpName, InStr(lpName, vbNullChar) - 1)
        End If
        
        CloseHandle hP
    End If
End Sub

Public Function GetTrayIcons() As TrayIconInfo()
Dim hTray As Long
Dim pid As Long
Dim ret As Long
Dim ptrTip As Long
Dim tbut As TBBUTTON
Dim TChar As Integer
Dim X As Long
Dim nIcons As Long
Dim Icons() As TrayIconInfo

hTray = GethWndTray

If hTray = 0 Then
   ErrRaise eeExternalError
   GoTo Leave:
End If

Static IsVirtualMemCreated As Boolean
If Not IsVirtualMemCreated Then
    '//cg 18April2008 : commented next 4 lines coz i will use "CreateVirtualMem" instead to solve the problem of the overflow error
    '//GetWindowThreadProcessId hTray, pid
    '//hProc = OpenProcess(PROCESS_VM, 0, pid)
    '//If hProc = -1 Then Err.Raise &H80070000 & Err.LastDllError 'cprin check one
    '//ptrVM = VirtualAllocEx(hProc, ByVal 0&, Len(tbut), MEM_COMMIT, PAGE_READWRITE)
    CreateVirutalMem
    IsVirtualMemCreated = True
End If

'//pt 12July2007: the next line causes "Over Flow Error" in Easy Window & System Tray Icons hider
'//because of frequently calling to Iconcount Prop in a timer
If ptrVM = 0 Then Err.Raise &H80070000 & Err.LastDllError

nIcons = SendMessage(hTray, TB_BUTTONCOUNT, ByVal 0&, ByVal 0&)
 
If nIcons = 0 Then LastErrorForMe = "No Icon Found in systray": NoIconFound = True

If nIcons = 0 Then GoTo Leave:

ReDim Icons(nIcons - 1)

IconNumber = nIcons - 1

NoIconFound = False

For X = 0 To nIcons - 1
    
    
    SendMessage hTray, TB_GETBUTTON, X, ByVal ptrVM

    ReadProcessMemory hProc, ByVal ptrVM, tbut, Len(tbut), ret
    
    
    Icons(X).Data = tbut.dwData
    Icons(X).State = tbut.fsState
    Icons(X).Style = tbut.fsStyle
    Icons(X).Bitmap = tbut.iBitmap
    Icons(X).Command = tbut.idCommand
    Icons(X).str = tbut.iString

    ptrTip = tbut.iString
    Do
        
        ReadProcessMemory hProc, ByVal ptrTip, TChar, 2&, ret
        Icons(X).ToolTip = Icons(X).ToolTip & ChrW$(TChar)
        ptrTip = ptrTip + 2
        
    Loop While TChar
    
    ReadProcessMemory hProc, ByVal tbut.dwData, Icons(X).hwnd, 36&, ret
        
    
    GetWindowThreadProcessId Icons(X).hwnd, pid
    
    GetProcessNameFromPID_NT pid, Icons(X).APath
Next



GetTrayIcons = Icons

'//cg 18April2008 : commented next 2 lines coz i will use "CreateVirtualMem" instead to solve the problem of the overflow error
'VirtualFreeEx hProc, ByVal ptrVM, 0&, MEM_RELEASE
'CloseHandle hProc

Exit Function
Leave:
ReDim Icons(0)
GetTrayIcons = Icons

End Function
Public Sub GetProcessNameFromPID_9x(pid As Long, lpName As String)
    Dim hSnap As Long
    Dim ret As Long
    
    
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, ByVal 0&)
    
    Dim Process As PROCESSENTRY32

    Process.dwSize = Len(Process)

    
    ret = Process32First(hSnap, Process)

    While ((ret) And (Process.th32ProcessID <> pid))
        
        Process.dwSize = Len(Process)
        
        ret = Process32Next(hSnap, Process)
    Wend
    

    CloseHandle hSnap

    
    If (ret) Then
    
        lpName = Left(Process.szExeFile, InStr(Process.szExeFile, vbNullChar) - 1)
    End If
End Sub

Public Function GetTrayIcons_9x() As TrayIconInfo()
 
    Dim hTray As Long
    Dim sys As Systray
    Dim lpsystray As RawSystray
    Dim lpsysListHeader As RawNotifyListHeader
    Dim temp() As Long
    Dim lpData As Long
    Dim pid  As Long
    Dim hP As Long
    Dim X As Long
    Dim hicons() As TrayIconInfo
    
    hTray = GethWndTray
    
    
    lpData = GetWindowLong(hTray, 0)
    
    If (lpData) Then
        
        GetWindowThreadProcessId hTray, pid
        
        hP = OpenProcess(PROCESS_VM_READ Or PROCESS_VM_OPERATION, 0&, pid)
        
        If (hP) Then
            
            If (ReadProcessMemory(hP, ByVal lpData, ByVal VarPtr(lpsystray), 44&, ByVal 0&)) Then
            
                lpData = lpsystray.lpsysListHeader
                
                sys.bkColor = lpsystray.bkColor
                sys.dwImageCount = lpsystray.dwImageCount
                sys.dwNotifyHeight = lpsystray.dwNotifyHeight
                sys.dwNotifyWidth = lpsystray.dwNotifyWidth
                sys.dxNotify = lpsystray.dxNotify
                sys.dyNotify = lpsystray.dyNotify
                sys.hicons = lpsystray.hicons
                sys.hToolTips = lpsystray.hToolTips
                sys.hTrayClockWClass = lpsystray.hTrayClockWClass
                sys.hTrayNotifyWnd = lpsystray.hTrayNotifyWnd
            
                If (ReadProcessMemory(hP, ByVal lpData, ByVal VarPtr(lpsysListHeader), 8&, ByVal 0&)) Then
                        Dim dwCount As Long
                        
                        dwCount = lpsysListHeader.dwCount
                        
                        lpData = lpsysListHeader.lpNotifyDataList
                        
                        sys.lpsysListHeader.dwCount = dwCount
                        
      If dwCount = 0 Then NoIconFound = True: GoTo Leave:
       
      If dwCount = 0 Then GoTo Leave:
                        
                                         
                        
                        ReDim temp(dwCount - 1)
                        
                        ReDim sys.lpsysListHeader.lpNotifyDataList(dwCount - 1)
                        ReDim hicons(dwCount - 1)
                        
                        If (ReadProcessMemory(hP, ByVal lpData, ByVal VarPtr(temp(0)), dwCount * 4, ByVal 0&)) Then
                            IconNumber = dwCount - 1
                            NoIconFound = False
                            For X = 0 To dwCount - 1
                                   ReadProcessMemory hP, ByVal temp(X), ByVal VarPtr(sys.lpsysListHeader.lpNotifyDataList(X)), 4 + 24 + 64, ByVal 0&

                                     sys.lpsysListHeader.lpNotifyDataList(X).NotifyDataA.szTip = StrConv(sys.lpsysListHeader.lpNotifyDataList(X).NotifyDataA.szTip, vbUnicode)
                                    GetWindowThreadProcessId sys.lpsysListHeader.lpNotifyDataList(X).NotifyDataA.hwnd, pid
                                    GetProcessNameFromPID_9x pid, sys.lpsysListHeader.lpNotifyDataList(X).Path
                            
     hicons(X).APath = sys.lpsysListHeader.lpNotifyDataList(X).Path
     hicons(X).ToolTip = sys.lpsysListHeader.lpNotifyDataList(X).NotifyDataA.szTip
     hicons(X).hwnd = sys.lpsysListHeader.lpNotifyDataList(X).NotifyDataA.hwnd
     hicons(X).uId = sys.lpsysListHeader.lpNotifyDataList(X).NotifyDataA.uId
     hicons(X).Bitmap = sys.lpsysListHeader.lpNotifyDataList(X).dwImageIndex
     hicons(X).ucallbackMessage = sys.lpsysListHeader.lpNotifyDataList(X).NotifyDataA.ucallbackMessage
     hicons(X).hIcon = sys.hicons
     hicons(X).Command = sys.lpsysListHeader.lpNotifyDataList(X).NotifyDataA.uId
                           Next
                        End If
                End If
            End If
            
            CloseHandle hP
        End If
    End If

GetTrayIcons_9x = hicons

Exit Function
Leave:
ReDim hicons(0)
GetTrayIcons_9x = hicons
 
End Function

'//cg 18April2008: Do you rembmember the error of ECW and "daniel.roesler@convergit.de"=order number for ASTC is U27236360"
'//the problem is: frequently calls to "GetSysTrayIcons" function causes overflow error due to frequntly use of "VirtualAllocEx"
'//hence i added CreateVirutalMem function to use VirtualAllocEx only one time in the program life time
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub CreateVirutalMem()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
   
    Dim pid As Long
    
    Dim tbut As TBBUTTON
    
    GetWindowThreadProcessId GethWndTray, pid
    
    hProc = OpenProcess(PROCESS_VM, 0, pid)
    
    ptrVM = VirtualAllocEx(hProc, ByVal 0&, Len(tbut), MEM_COMMIT, PAGE_READWRITE)

End Sub
