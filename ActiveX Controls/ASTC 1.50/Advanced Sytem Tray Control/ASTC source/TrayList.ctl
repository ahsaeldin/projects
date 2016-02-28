VERSION 5.00
Begin VB.UserControl TrayList 
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   540
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   540
   ScaleWidth      =   540
   ToolboxBitmap   =   "TrayList.ctx":0000
   Begin VB.Image Image1 
      Height          =   540
      Left            =   0
      Picture         =   "TrayList.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   540
   End
End
Attribute VB_Name = "TrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type TrayIconInfo
    hwnd As Long
    uId As Long
    ucallbackMessage As Long
    Param1(1) As Long
    hIcon As Long
    Param2(2) As Long
    APath As String
    ToolTip As String
    Bitmap As Long
    Command As Long
    State As Byte
    Style As Byte
    Data As Long
    str As Long
End Type

Private Sub UserControl_Resize()
UserControl.Width = 540
UserControl.Height = 540
End Sub

Public Function GetSysTrayIcons() As TrayIconInfo()
Attribute GetSysTrayIcons.VB_Description = "Get all icons displayed in system tray area.\r\n\r\n"
Attribute GetSysTrayIcons.VB_HelpID = 1

Dim ticons() As TrayIconInfo

    If IsWindowsNT Then
       
       ticons = GetTrayIcons
       
       GetSysTrayIcons = ticons
       
    Else
       
       ticons = GetTrayIcons_9x
      
       GetSysTrayIcons = ticons
       
    End If


End Function

Friend Function ExtenderName() As String
  ExtenderName = Extender.Name
End Function

Public Property Get IconCount() As Variant
Attribute IconCount.VB_Description = "Returns the number of the icons located in system tray area."
Dim TrayList() As TrayIconInfo
TrayList = GetSysTrayIcons
If NoIconFound Then IconCount = 0: Exit Property
IconCount = IconNumber + 1
End Property

Public Sub HideIcon(index As Integer)
Attribute HideIcon.VB_Description = "Hide a specified icon from system tray area."
On Error Resume Next
Dim TrayList() As TrayIconInfo
Dim TrayI As NOTIFYICONDATA
Const NIF_TIP = &H4
Const NIF_ICON = &H2
Const NIF_MESSAGE = &H1
        
    bybassError = True
    TrayList = GetSysTrayIcons
    TrayI.cbSize = Len(TrayI)
    TrayI.hwnd = TrayList(index - 1).hwnd
    TrayI.uId = TrayList(index - 1).uId
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.hIcon = TrayList(index - 1).hIcon
    Shell_NotifyIcon NIM_DELETE, TrayI
 
End Sub

Private Sub UserControl_Terminate()

'//cg 18April2008: Do you rembmember the error of ECW and "daniel.roesler@convergit.de"=order number for ASTC is U27236360"
Cleanup:
    
    '//clear the virutal memory we used
    VirtualFreeEx hProc, ByVal ptrVM, 0&, MEM_RELEASE
    
    CloseHandle hProc

End Sub
