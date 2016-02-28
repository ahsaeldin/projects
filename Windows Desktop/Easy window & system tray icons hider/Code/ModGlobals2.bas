Attribute VB_Name = "ModGlobals2"

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '“Bill Gates is a very rich man today .. and do you want to know why? The answer is one word: versions.”'
        'Dave Barry quotes                                                                                      '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'//pt 1July2007: Why I added ModGlobals2
'//because i will add to it all the global variables that i will use only in this project
'//so ModGlobal will be portable and can be use in the further software isA

Option Explicit

'//pt 29May2007: Public Task Item Structure
Public Type TASK_STRUCT
    TaskName As String 'contains window title
    TaskID As Long     'contains window hwnd
End Type

'//pt 29May2007: Structure filled by FillTaskList Sub call
Public TaskList(200) As TASK_STRUCT

'//pt 31May2007: added next Var as an global object for CWindows Class
Public CVisWindows As CWindows

'//pt 9July2007: added next line as a global variable to use in checking Hot Track rather than open ini file
Public IsHotTrack As Boolean

Public IsNeverTrayIcon As Boolean

Public IsAlwaysTrayIcon As Boolean

Public IsTrayIconOnlyWhenMini As Boolean

'//added next var to set to true in mnuExit in order to check
'//in form_queryunload to avoid processing form_queryunload that's cozs
'//the software to not to be closed
Public ExitMeNow As Boolean

Public IsShowTopMost As Boolean

'//pt 2Feb2008 6:25PM : add "IsRunAtStartUp" to enable Run the software at startup
Public IsRunAtStartUp As Boolean

Public Type SysIcoToolTip
        hwnd As Long
        ToolTip As String * 64
End Type

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'//OurTraget: Update the listview if the tooltip or icon changed
'//to achieve this target we need to compare listview items with the current systray info
'//so we need to get systray info 3ly tool
'//if we get systray info 3ly tool, we will get an "over flow" error after a period because of using "VirtualAllocEx" in ASTC Control
'//hence i added CreateVirutalMem function to use VirtualAllocEx only once and to get the icons info manually
Public Type TBBUTTON
    iBitmap As Long
    idCommand As Long
    fsState As Byte
    fsStyle As Byte
    bReserved(1 To 2) As Byte
    dwData As Long
    iString As Long
End Type

Public Const PROCESS_VM_WRITE As Long = (&H20)
Public Const PROCESS_VM_READ As Long = (&H10)
Public Const PROCESS_VM_OPERATION As Long = (&H8)
Public Const PROCESS_QUERY_INFORMATION As Long = (&H400)

Public Const PROCESS_VM As Long = PROCESS_QUERY_INFORMATION Or PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE 'Or PROCESS_DUP_HANDLE

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Public Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long

Public Enum DesktopTopIconsState
       HideIcons = 0
       ShowIcons
End Enum

Public Sub ShowDesktopIcons(State As DesktopTopIconsState)
    
    Dim FolderViewHwnd As Long
    Dim ProgramManagerHwnd As Long
    Dim SHELLDLL_DefViewHwnd As Long
         
    ProgramManagerHwnd = FindWindow(vbNullString, "Program Manager")
    SHELLDLL_DefViewHwnd = FindWindowEx(ProgramManagerHwnd, 0&, "SHELLDLL_DefView", vbNullString)
    If IsWindowsNT Then
       FolderViewHwnd = FindWindowEx(SHELLDLL_DefViewHwnd, 0&, vbNullString, "FolderView")
    Else
       FolderViewHwnd = FindWindowEx(SHELLDLL_DefViewHwnd, 0&, "SysListView32", vbNullString)
    End If
    ShowWindow FolderViewHwnd, State
       
End Sub

Public Function IsWindowsNT() As Boolean
    Dim os As OSVERSIONINFO
    
    os.dwOSVersionInfoSize = Len(os)
    GetVersionEx os

    IsWindowsNT = (os.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Public Sub CreateShortcut(ByVal TargetPath As String, ByVal ShortcutPath As String, ByVal ShortcutName As String)

    Dim VbsObj As Object
    
    Dim MyShortcut As Object
    
    Dim FullShortCutPath As String
    
    '//We must check here is Failed
    Set VbsObj = CreateObject("WScript.Shell")
     
    If ShortcutPath = "Startup" Then
       '//We must check here is Failed
       FullShortCutPath = VbsObj.SpecialFolders("Startup")
    Else
       FullShortCutPath = My_PathAddBackslash(ShortcutPath)
    End If
   
    '//We must check here is Failed
    Set MyShortcut = VbsObj.CreateShortcut(FullShortCutPath & "\" & ShortcutName & ".lnk")
        
    MyShortcut.TargetPath = TargetPath
    
    '//We must check here is Failed
    MyShortcut.Save

    '//Returen a value
    
    Set VbsObj = Nothing
    
    Set MyShortcut = Nothing
    
End Sub

Public Sub DeleteShortcut(ByVal TargetPath As String, ByVal ShortcutPath As String, ByVal ShortcutName As String)

    Dim VbsObj As Object
    
    Dim MyShortcut As String
    
    Dim FullShortCutPath As String
    
    '//We must check here is Failed
    Set VbsObj = CreateObject("WScript.Shell")
     
    If ShortcutPath = "Startup" Then
       '//We must check here is Failed
       FullShortCutPath = VbsObj.SpecialFolders("Startup")
    Else
       FullShortCutPath = My_PathAddBackslash(ShortcutPath)
    End If
   
    '//We must check here
     MyShortcut = FullShortCutPath & "\" & ShortcutName & ".lnk"
     
     On Error Resume Next
     Kill MyShortcut
    
     Set VbsObj = Nothing
     
End Sub

'//////////nooooooooooote Don't use this function, Use another one called "GetSpecialFolderPathbyAPI" in HSM
' Get a special folder's path.
'//pt 7Feb2008 5:22PM : added to get the Path of the current user's startup folder
'//pt 7Feb2008 6:04PM : I discoverd that i can use ".SpecialFolders" in CreateShortcut function
'//to get startup folder but anyway i won't delete this function .. may be i use it in the future
Public Function GetSpecialFolderPath(ByVal folder_number As Long) As String

    Dim path As String
    Const MAX_PATH = 260
    
    path = Space$(MAX_PATH)

    If SHGetSpecialFolderPath(frmOptions.hwnd, path, folder_number, False) Then
        '//note that we added "- 1" in the next line to delete the null termintating
        '//hence when we use "My_PathAddBackslash", we can add the slash
        GetSpecialFolderPath = Left$(path, InStr(path, Chr$(0)) - 1)
    Else
        GetSpecialFolderPath = "Failed"
    End If

End Function

Public Function My_PathAddBackslash(path As String)
    
    If Right(path, 1) <> "\" Then My_PathAddBackslash = path + "\" Else My_PathAddBackslash = path

End Function
