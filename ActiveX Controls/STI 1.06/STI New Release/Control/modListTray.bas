Attribute VB_Name = "modListTray_NT"
Option Explicit
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef lpcbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Public IconNumber As Long
Public Function GetTrayIcons() As TrayIcon()
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

If hTray = 0 Then
   ErrRaise eeExternalError
   GoTo Leave:
End If

GetWindowThreadProcessId hTray, pid

hProc = OpenProcess(PROCESS_VM, 0, pid)

If hProc = -1 Then Err.Raise &H80070000 & Err.LastDllError

ptrVM = VirtualAllocEx(hProc, ByVal 0&, Len(tbut), MEM_COMMIT, PAGE_READWRITE)

If ptrVM = 0 Then Err.Raise &H80070000 & Err.LastDllError

nIcons = SendMessage(hTray, TB_BUTTONCOUNT, ByVal 0&, ByVal 0&)

If Not bybassError Then If nIcons = 0 Then MsgBox "No Icon Found in Systray", vbExclamation:

If nIcons = 0 Then GoTo Leave:

ReDim Icons(nIcons - 1)

IconNumber = nIcons - 1
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
        
    
    GetWindowThreadProcessId Icons(x).hwnd, pid
    
    GetProcessNameFromPID_NT pid, Icons(x).APath
Next



GetTrayIcons = Icons

Cleanup:

VirtualFreeEx hProc, ByVal ptrVM, 0&, MEM_RELEASE

CloseHandle hProc

Exit Function
Leave:
ReDim Icons(0)
GetTrayIcons = Icons

End Function


Public Function GetState(fsState As Long) As String
If (fsState And TBSTATE_CHECKED) = TBSTATE_CHECKED Then GetState = GetState & "Coché - "
If (fsState And TBSTATE_ELLIPSES) = TBSTATE_ELLIPSES Then GetState = GetState & "Ellipsé - "
If (fsState And TBSTATE_ENABLED) = TBSTATE_ENABLED Then GetState = GetState & "Activé - "
If (fsState And TBSTATE_HIDDEN) = TBSTATE_HIDDEN Then GetState = GetState & "Caché - "
If (fsState And TBSTATE_INDETERMINATE) = TBSTATE_INDETERMINATE Then GetState = GetState & "Grisé - "
If (fsState And TBSTATE_MARKED) = TBSTATE_MARKED Then GetState = GetState & "Marqué - "
If (fsState And TBSTATE_PRESSED) = TBSTATE_PRESSED Then GetState = GetState & "Pressé - "
If (fsState And TBSTATE_WRAP) = TBSTATE_WRAP Then GetState = GetState & "Suivi d'une barre - "
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
