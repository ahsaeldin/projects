Attribute VB_Name = "ModGlobals2"
Option Explicit

'//////////////
'//because i will add to it all the global variables that i will use only in this project
'//so ModGlobal will be portable and can be use in the further software isA

'//added next var to set to true in mnuExit in order to check
'//in form_queryunload to avoid processing form_queryunload that's cozs
'//the software to not to be closed
Public ExitMeNow As Boolean

Public IsHotTrack As Boolean
Public IsShowTopMost As Boolean
Public IsRunAtStartUp As Boolean
Public rCheckatStartup As Boolean
Public rMonitorAllowed As Boolean
Public rchkNewlyAdded As Boolean
Public rchkRemoved As Boolean
Public IsNeverTrayIcon As Boolean
Public IsAlwaysTrayIcon As Boolean
Public IsTrayIconOnlyWhenMini As Boolean

Public CUStartupFolderPath As String
Public AllUStartupFolderPath As String

Public Const CSIDL_STARTUP = &H7
Public Const CSIDL_COMMON_STARTUP = &H18

Public Declare Function InstallKey Lib "ArmAccess.DLL" (ByVal Name$, ByVal Code$) As Byte
Public Declare Function ShowReminderMessage Lib "ArmAccess.DLL" (ByVal ParentWindow As Long)
Public Declare Function ArmEnviron Lib "ArmAccess.DLL" Alias "Environ" (ByVal Name$) As String
Public Declare Function ShowEnterKeyDialog Lib "ArmAccess.DLL" (ByVal ParentWindow As Long) As Byte
      
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub CreateShortcut(ByVal TargetPath As String, ByVal ShortcutPath As String, ByVal ShortcutName As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

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

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub DeleteShortcut(ByVal TargetPath As String, ByVal ShortcutPath As String, ByVal ShortcutName As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Dim VbsObj As Object
    
    Dim MyShortcut As String
    
    Dim FullShortCutPath As String
     
    '//We must check here is Failed
    Set VbsObj = CreateObject("WScript.Shell")
     
    If ShortcutPath = "Startup" Then
       '//We must check here is Failed
       FullShortCutPath = VbsObj.SpecialFolders("Startup")
    ElseIf ShortcutPath = "AllUsersStartup" Then
       FullShortCutPath = VbsObj.SpecialFolders("AllUsersStartup")
    Else
       FullShortCutPath = My_PathAddBackslash(ShortcutPath)
    End If
    
    '//We must check here
     MyShortcut = FullShortCutPath & "\" & ShortcutName
     
     Dim ext As String
     ext = ExtractFileExtensionFromPath(ShortcutName)
     If ext <> "lnk" Or ext <> "url" Then
       On Error Resume Next
       Kill MyShortcut
     Else
        Kill MyShortcut
     End If
            
    
     Set VbsObj = Nothing
     
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function My_PathAddBackslash(path As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    If Right(path, 1) <> "\" Then My_PathAddBackslash = path + "\" Else My_PathAddBackslash = path

End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetSpecialFolderPathbyShell(FolderName As String) As String
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: GetSpecialFolderPathbyShell
' Author: Ahmed Saad
' Date: 14/02/08 6:51:29 ã
' Copyright © 2006 - 2008,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo GetSpecialFolderPathbyShellErr
        
'//For Example:
'//GetSpecialFolderPathbyShell("StartUp")
'//GetSpecialFolderPathbyShell ("AllUsersStartup")

    Dim VbsObj As Object
    
    Dim FullShortCutPath As String
    
    Set VbsObj = CreateObject("WScript.Shell")
    
    Select Case FolderName
           Case "StartUp"
           FullShortCutPath = VbsObj.SpecialFolders("StartUp")
           Case "AllUsersStartup"
           FullShortCutPath = VbsObj.SpecialFolders("AllUsersStartup")
    End Select
    
    Set VbsObj = Nothing
    
    GetSpecialFolderPathbyShell = My_PathAddBackslash(FullShortCutPath)
    
GetSpecialFolderPathbyShellExit:
    On Error Resume Next

    Exit Function
    
GetSpecialFolderPathbyShellErr:

    '//0x0f4hsm112 number is a unique number for MyGetSpecialFolderPathbyShell because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f4hsm112"
    
GoTo GetSpecialFolderPathbyShellExit

End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetSpecialFolderPathbyAPI(WinHwnd As Long, ByVal folder_number As Long) As String
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    '//For Example:
    '//GetSpecialFolderPathbyAPI (hwnd, CSIDL_STARTUP)
    '//GetSpecialFolderPathbyAPI (hwnd, CSIDL_COMMON_STARTUP)
    
    Dim path As String
    
    path = Space$(MAX_PATH)
    
    If SHGetSpecialFolderPath(WinHwnd, path, folder_number, False) Then
        
        '// "-1" to delete the null terminateing string
        GetSpecialFolderPathbyAPI = My_PathAddBackslash(Left$(path, InStr(path, Chr$(0)) - 1))
    
    Else
        
        GetSpecialFolderPathbyAPI = "null"
    
    End If

End Function
