Attribute VB_Name = "Utils"

                    '//"·«  —ﬂ“ ⁄·Ï «·„«÷Ï , ›ﬁÿ «” ⁄„·Â ··Œ—ÊÃ »«·Õﬂ„… Ê«·›«∆œ… , À„ « —ﬂÂ Œ·›ﬂ

Option Explicit
   
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_QUERY_INFORMATION = &H400
      
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function EnumProcesses Lib "PSAPI.DLL" (lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Private Declare Function GetModuleBaseName Lib "PSAPI.DLL" Alias "GetModuleBaseNameA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long


Private Type PROCESSENTRY32
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

Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Private Const PROCESS_ALL_ACCESS = 0
Private Const TH32CS_SNAPPROCESS As Long = 2&

Private Const WINNT As Integer = 2
Private Const WIN98 As Integer = 1

Private KillAppReturn As Boolean

'//Check if an exe is running
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function IsProcessRunning(ByVal sProcess As String) As Boolean
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Const MAX_PATH As Long = 260

    Dim sName As String
    Dim lProcesses() As Long, lModules() As Long, N As Long, lRet As Long, hProcess As Long
      
    sProcess = UCase$(sProcess)
    ReDim lProcesses(1023) As Long
    If EnumProcesses(lProcesses(0), 1024 * 4, lRet) Then
       For N = 0 To (lRet \ 4) - 1
           hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(N))
           If hProcess Then
              ReDim lModules(1023)
              If EnumProcessModules(hProcess, lModules(0), 1024 * 4, lRet) Then
                 sName = String$(MAX_PATH, vbNullChar)
                 GetModuleBaseName hProcess, lModules(0), sName, MAX_PATH
                 sName = Left$(sName, InStr(sName, vbNullChar) - 1)
                 If Len(sName) = Len(sProcess) Then
                    If sProcess = UCase$(sName) Then IsProcessRunning = True: Exit Function
                 End If
               End If
            End If
            CloseHandle hProcess
        Next N
      End If
      
End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function ExtractFileNameFromPath(path As String) As String
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
'Print ExtractFileNameFromPath("C:\WINDOWS\system32\ctfmon.exe")
'Returns ctfmon.exe

'Note : Must call "ExtractPath" function before calling this if the path may contains quotaions

   Dim WhereBackSlash As Integer
   WhereBackSlash = InStrRev(path, "\")
   ExtractFileNameFromPath = Mid(path, WhereBackSlash + 1, Len(path) - WhereBackSlash)
   
End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function ExtractFileLocationFromPath(path As String) As String
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'
'Print ExtractFileLocationFromPath("C:\WINDOWS\system32\ctfmon.exe")
'Returns C:\WINDOWS\system32\

'Note : Must call "ExtractPath" function before calling this if the path may contains quotaions
   Dim WhereBackSlash As Integer
   WhereBackSlash = InStrRev(path, "\")
   'ExtractFileNameFromPath = Mid(path, WhereBackSlash + 1, Len(path) - WhereBackSlash)
   ExtractFileLocationFromPath = Left(path, WhereBackSlash) '(path, WhereBackSlash + 1, Len(path) - WhereBackSlash)
      
End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function ExtractFileExtensionFromPath(path As String) As String
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'Print ExtractFileLocationFromPath("C:\WINDOWS\system32\ctfmon.exe")
'Returns exe
Dim ExtensionCharsNum As Integer

'Note : Must call "ExtractPath" function before calling this if the path may contains quotaions
   Dim WhereBackSlash As Integer
   WhereBackSlash = InStrRev(path, ".")
   'ExtractFileNameFromPath = Mid(path, WhereBackSlash + 1, Len(path) - WhereBackSlash)
   ExtensionCharsNum = Len(path) - WhereBackSlash
   ExtractFileExtensionFromPath = Mid(path, WhereBackSlash + 1, ExtensionCharsNum) '(path, WhereBackSlash + 1, Len(path) - WhereBackSlash)
      
End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetVersion() As Integer
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

   Dim udtOSInfo As OSVERSIONINFO
   Dim intRetVal As Integer
       
   With udtOSInfo
       .dwOSVersionInfoSize = 148
       .szCSDVersion = Space$(128)
   End With
   
   intRetVal = GetVersionExA(udtOSInfo)
 
   GetVersion = udtOSInfo.dwPlatformId

End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function Killapp(myName As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Select Case GetVersion()
           Case WIN98
                Killapp9X (myName)
           Case WINNT
                KillappNT (myName)
    End Select

End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function KillappNT(myName As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
       
       Dim uProcess As PROCESSENTRY32
       Dim rProcessFound As Long
       Dim hSnapshot As Long
       Dim szExename As String
       Dim exitCode As Long
       Dim myProcess As Long
       Dim AppKill As Boolean
       Dim appCount As Integer
       Dim I As Integer
       On Local Error GoTo Finish
       appCount = 0
       uProcess.dwSize = Len(uProcess)
       hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
       rProcessFound = ProcessFirst(hSnapshot, uProcess)
       Do While rProcessFound
           I = InStr(1, uProcess.szExeFile, Chr(0))
           szExename = LCase$(Left$(uProcess.szExeFile, I - 1))
           If Right$(szExename, Len(myName)) = LCase$(myName) Then
               KillAppReturn = True
               appCount = appCount + 1
               myProcess = OpenProcess(1&, -1&, uProcess.th32ProcessID)
               AppKill = TerminateProcess(myProcess, 0&)
               Call CloseHandle(myProcess)
           End If
           rProcessFound = ProcessNext(hSnapshot, uProcess)
       Loop
       Call CloseHandle(hSnapshot)
Finish:
    KillAppReturn = False

End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function Killapp9X(myName As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
       
       Dim uProcess As PROCESSENTRY32
       Dim rProcessFound As Long
       Dim hSnapshot As Long
       Dim szExename As String
       Dim exitCode As Long
       Dim myProcess As Long
       Dim AppKill As Boolean
       Dim appCount As Integer
       Dim I As Integer
       On Local Error GoTo Finish
       appCount = 0
       uProcess.dwSize = Len(uProcess)
       hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
       rProcessFound = ProcessFirst(hSnapshot, uProcess)
       Do While rProcessFound
           I = InStr(1, uProcess.szExeFile, Chr(0))
           szExename = LCase$(Left$(uProcess.szExeFile, I - 1))
           If Right$(szExename, Len(myName)) = LCase$(myName) Then
               KillAppReturn = True
               appCount = appCount + 1
               myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
               AppKill = TerminateProcess(myProcess, exitCode)
               Call CloseHandle(myProcess)
           End If
           rProcessFound = ProcessNext(hSnapshot, uProcess)
       Loop
       Call CloseHandle(hSnapshot)
Finish:
    KillAppReturn = False

End Function
