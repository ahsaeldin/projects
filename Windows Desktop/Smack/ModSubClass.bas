Attribute VB_Name = "ModSubClass"
Option Explicit

Public Const SW_SHOW = 5
Public Const SW_RESTORE = 9
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWMAXIMIZED = 3


Public Const WM_USER As Long = &H400
Public Const WM_MYHOOK As Long = WM_USER + 2

Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const GWL_WNDPROC = (-4)

'//cprin 10Aug2006 implment the Evulation [search the other]
Public Ending As Boolean
