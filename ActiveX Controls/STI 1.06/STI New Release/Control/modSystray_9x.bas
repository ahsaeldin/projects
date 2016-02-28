Attribute VB_Name = "modSystray_9x"
Option Explicit


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

Private Declare Function GlobalAlloc Lib "KERNEL32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "KERNEL32.dll" (ByVal hMem As Long) As Long
Private Const GMEM_FIXED As Long = &H0
Private Const GMEM_ZEROINIT As Long = &H40

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

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
Private Const TH32CS_SNAPPROCESS As Long = &H2
Private Declare Function CreateToolhelp32Snapshot Lib "KERNEL32.dll" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "KERNEL32.dll" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "KERNEL32.dll" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long

Public Function GetTrayIcons_9x() As TrayIcon()
 
    Dim hTray As Long
    Dim sys As Systray
    Dim lpsystray As RawSystray
    Dim lpsysListHeader As RawNotifyListHeader
    Dim temp() As Long
    Dim lpData As Long
    Dim pid  As Long
    Dim hP As Long
    Dim x As Long
    Dim hicons() As TrayIcon
    
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
                        
      If Not bybassError Then If dwCount = 0 Then MsgBox "No Icon Found in Systray", vbExclamation: GoTo Leave:
       
      If dwCount = 0 Then GoTo Leave:
                        
                                         
                        
                        ReDim temp(dwCount - 1)
                        
                        ReDim sys.lpsysListHeader.lpNotifyDataList(dwCount - 1)
                        ReDim hicons(dwCount - 1)
                        
                        If (ReadProcessMemory(hP, ByVal lpData, ByVal VarPtr(temp(0)), dwCount * 4, ByVal 0&)) Then
                            IconNumber = dwCount - 1
                            For x = 0 To dwCount - 1
                                   ReadProcessMemory hP, ByVal temp(x), ByVal VarPtr(sys.lpsysListHeader.lpNotifyDataList(x)), 4 + 24 + 64, ByVal 0&

                                     sys.lpsysListHeader.lpNotifyDataList(x).NotifyDataA.szTip = StrConv(sys.lpsysListHeader.lpNotifyDataList(x).NotifyDataA.szTip, vbUnicode)
                                    GetWindowThreadProcessId sys.lpsysListHeader.lpNotifyDataList(x).NotifyDataA.hwnd, pid
                                    GetProcessNameFromPID_9x pid, sys.lpsysListHeader.lpNotifyDataList(x).Path
                            
     hicons(x).APath = sys.lpsysListHeader.lpNotifyDataList(x).Path
     hicons(x).ToolTip = sys.lpsysListHeader.lpNotifyDataList(x).NotifyDataA.szTip
     hicons(x).hwnd = sys.lpsysListHeader.lpNotifyDataList(x).NotifyDataA.hwnd
     hicons(x).uId = sys.lpsysListHeader.lpNotifyDataList(x).NotifyDataA.uId
     hicons(x).Bitmap = sys.lpsysListHeader.lpNotifyDataList(x).dwImageIndex
     hicons(x).ucallbackMessage = sys.lpsysListHeader.lpNotifyDataList(x).NotifyDataA.ucallbackMessage
     hicons(x).hIcon = sys.hicons
     hicons(x).Command = sys.lpsysListHeader.lpNotifyDataList(x).NotifyDataA.uId
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

Public Function GetTrayIcons_ME() As TRAY_ME()
Dim hProc As Long
Dim hTray As Long
Dim pid As Long
Dim ptrVM As Long
Dim ret As Long
Dim ptrTip As Long
Dim tbut As TBBUTTON
Dim TChar As Integer
Dim x As Long
Dim nIcons As Long
Dim Icons() As TrayIcon


hTray = GethWndTray
If hTray = 0 Then Exit Function
GetWindowThreadProcessId hTray, pid
hProc = OpenProcess(PROCESS_VM, 0, pid)
If hProc = -1 Then Err.Raise &H80070000 & Err.LastDllError
ptrVM = GlobalAlloc(GMEM_ZEROINIT Or GMEM_FIXED, Len(tbut))
If ptrVM = 0 Then Err.Raise &H80070000 & Err.LastDllError
nIcons = SendMessage(hTray, TB_BUTTONCOUNT, ByVal 0&, ByVal 0&)
ReDim Icons(nIcons - 1)
For x = 0 To nIcons - 1
    SendMessage hTray, TB_GETBUTTON, x, ByVal ptrVM
    ReadProcessMemory hProc, ByVal ptrVM, tbut, Len(tbut), ret
    
    Icons(x).Data = tbut.dwData
    Icons(x).State = tbut.fsState
    Icons(x).Style = tbut.fsStyle
    Icons(x).Bitmap = tbut.iBitmap
    Icons(x).Command = tbut.idCommand
    Icons(x).str = tbut.iString
    
    ptrTip = tbut.iString
    Do
        ReadProcessMemory hProc, ByVal ptrTip, TChar, 2&, ret
        Icons(x).ToolTip = Icons(x).ToolTip & ChrW$(TChar)
        ptrTip = ptrTip + 2
    Loop While TChar
    ReadProcessMemory hProc, ByVal tbut.dwData, Icons(x).hwnd, 36&, ret
    
    ptrTip = tbut.dwData + 36
    Do
        ReadProcessMemory hProc, ByVal ptrTip, TChar, 2&, ret
        Icons(x).APath = Icons(x).APath & ChrW$(TChar)
        ptrTip = ptrTip + 2
    Loop While TChar
Next

GetTrayIcons_ME = Icons

Cleanup:
GlobalFree ptrVM
CloseHandle hProc
End Function

