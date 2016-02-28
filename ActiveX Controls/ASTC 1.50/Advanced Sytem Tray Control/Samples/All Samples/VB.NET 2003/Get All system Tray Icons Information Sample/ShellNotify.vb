Imports System.Runtime.InteropServices

Module ShellNotify
    Public OldTrayIcon As ASTC.TrayIconInfo

    Public Const NIM_ADD As Int32 = &H0
    Public Const NIF_TIP As Int32 = &H4
    Public Const NIF_ICON As Int32 = &H2
    Public Const NIF_MESSAGE As Int32 = &H1

    Public Structure NOTIFYICONDATA        Dim cbSize As Int32
        Dim hwnd As Int32
        Dim uID As Int32
        Dim uFlags As Int32
        Dim uCallbackMessage As Int32
        Dim hIcon As Int32
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=64)> Dim szTip As String    End Structure

    Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Int32, ByRef lpData As NOTIFYICONDATA) As Boolean
End Module
