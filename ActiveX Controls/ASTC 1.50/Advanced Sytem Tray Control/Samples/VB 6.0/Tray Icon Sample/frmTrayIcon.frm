VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{30E5838E-8E26-11DA-A1C2-807A50C10000}#1.0#0"; "astc.ocx"
Begin VB.Form Form1 
   Caption         =   "TrayIcon Sample Code"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   3000
   StartUpPosition =   2  'CenterScreen
   Begin ASTC.TrayIcon TrayIcon1 
      Left            =   1200
      Top             =   720
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdchg 
         Caption         =   "&ChangeIcon"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdHide 
         Caption         =   "&Hide Icon"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdani 
         Caption         =   "Animate"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton stopani 
         Caption         =   "Stop Animation"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdShowBalloon 
         Caption         =   "ShowBallon"
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCloseBalloon 
         Caption         =   "CloseBallon"
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdTrack 
         Caption         =   "TrackPopUp"
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton stoptrack 
         Caption         =   "Stop Track"
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdshow 
         Caption         =   "&Show Icon"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1200
         Top             =   1320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTrayIcon.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTrayIcon.frx":031A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTrayIcon.frx":0634
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTrayIcon.frx":094E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTrayIcon.frx":0C68
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTrayIcon.frx":0F82
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTrayIcon.frx":129C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTrayIcon.frx":15B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTrayIcon.frx":18D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTrayIcon.frx":1BEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTrayIcon.frx":1F04
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTrayIcon.frx":221E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu menu1 
      Caption         =   "File"
      Begin VB.Menu menu2 
         Caption         =   "Show"
      End
      Begin VB.Menu mexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
' Description:   A sample code demonstrate how to add your software icon to system tray area
'                As well as how to call the advanced features related to system tray area from
'                within your application.
'==============================================================================================
Option Explicit
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hmenu As Long, ByVal nPos As Long) As Long

'ASTC provide you with a very simple object, TrayIcon object which is the responsible to add the most of system
'Tray functions to your application.
'
'With a few lines we can demonstrate how TrayIcon object can help you
'
'Use TrayIcon.Show method to add an icon to system tray area
'Use TrayIcon.Hide method to hide the icon you had add by TrayIcon.Show method from system tray
'Use TrayIcon.ChangeIcon method to change the icon you had added by another icon
'Use TrayIcon.Animate method to animate the icon in the system tray
'Use TrayIcon.StopAnimateing method to stop animating the icon
'Use TrayIcon.ShowBalloon to display a tooltip balloon
'Use TrayIcon.HideBalloon to hide the balloon
'Use TrayIcon.Popup to display a popup menu
'Use TrayIcon.TrackPopupmenu to track the popup menu whenever it lost focus in order to close it.
'Use TrayIcon.IsDisplayed to check if the icon is displayed in the system tray area
'
'Note that you have a custom types of events associated with TrayIcon Object like
'BalloonShow, BalloonHide, MedMouseUp, LeftMouseUp, RightMouseUp, MedMouseDown, LeftMouseDown
'RightMouseDown, MedMouseDBLCLK, LeftMouseDBLCLK, BalloonLeftClick, BalloonRightClick, RightMouseDBLCLK ()
'
'Note that TrayIcon Object doesn't subclass the window in which it contains TrayIcon control
'because TrayIcon Object dynamically creates an internal window related to it's instance,
'and destroy this window whenever you call Hide method or close your application,

'Don't hesitate to contact us At support@cpringold.com if you have any further questions or need more sample

'==============================================================================================
' Method:        cmdshow_Click
'
' Description:   Add the icon to System Tray Area.
'==============================================================================================
Private Sub cmdshow_Click()
    
   'Unlike the previous version of ASTC, you don't need to pass a window handle to Show function because
   'ASTC dynamically creates an internal window related to ASTC instance, and destroy this window
   'whenever you call Hide method or close your application.
    TrayIcon1.Show Me.Icon, "Your Tooltip goes here." 'Add this Form icon to System Tray Area.

End Sub

'==============================================================================================
' Method:        cmdchg_Click
'
' Description:   Changes the icon in System Tray Area.
'==============================================================================================
Private Sub cmdchg_Click()

   If TrayIcon1.IsDisplayed Then 'If the icon is displayed in System Tray Area
      TrayIcon1.ChangeIcon ImageList1.ListImages.Item(1).ExtractIcon 'then change the icon in system tray area by another one in ImageList
   End If

End Sub

'==============================================================================================
' Method:        cmdani_Click
'
' Description:   Animate the icon in System Tray Area using the icons set in imagelist control.
'==============================================================================================
Private Sub cmdani_Click()

   If TrayIcon1.IsDisplayed Then 'If the icon is displayed in System Tray Area
       If TrayIcon1.AnimateState = False Then 'Then if the icon is already not Animated in System Tray Area
          TrayIcon1.Animate ImageList1, 1 'Then Start animating the icons in System Tray Area by using the icons stored in the imagelist
         '''Note: You can replace the icons with your own'''
       End If
   End If

End Sub

'==============================================================================================
' Method:       stopani_Click
'
' Description:  Stop animating the icons in system tray.
'==============================================================================================
Private Sub stopani_Click()

   If TrayIcon1.IsDisplayed Then 'If the icon is displayed in System Tray Area
      If TrayIcon1.AnimateState = True Then 'Then if the icon is already Animated in System Tray Area
          TrayIcon1.StopAnimateing ImageList1.ListImages.Item(2).ExtractIcon 'Then stop animateing the icon and use icon no 2 in the ImageList as a new icon.
      End If
   End If

End Sub

'==============================================================================================
' Method:       cmdShowBalloon_Click
'
' Description:  Displays a balloon tooltip points to the icon.
'==============================================================================================
Private Sub cmdShowBalloon_Click()

   If TrayIcon1.AnimateState = True Then MsgBox ("You can't Display a balloon while animating the icon.") 'Can't Display the Balloon Tooltip if the icon is already Animated in System Tray Area.
   If TrayIcon1.IsDisplayed Then 'If the icon is displayed in System Tray Area
      TrayIcon1.ShowBalloon "Click me", "My Balloon", Info, 5000 'Then Display a balloon ToolTip in System Tray Area that points to the icon.
   End If

End Sub

'==============================================================================================
' Method:       cmdCloseBalloon_Click
'
' Description:  Close the balloon tooltip.
'==============================================================================================
Private Sub cmdCloseBalloon_Click()
   
   If TrayIcon1.IsDisplayed Then 'If the icon is displayed in System Tray Area
      TrayIcon1.CloseBalloon 'Then Close and Hide the Balloon ToolTip
   End If

End Sub

'==============================================================================================
' Method:       cmdTrack_Click
'
' Description:  Start Tracking the popup menu in system tray area.
'==============================================================================================
Private Sub cmdTrack_Click()
   
  
   TrayIcon1.TrackPopMenu = True 'Enable the Track of the popup menu so ASTC can destroy it whenever it lost focus
   MsgBox "when you right click the icon a popupmenu will appear,ASTC will Track the popupmenu and close it lost focus"
   stoptrack.Enabled = True

End Sub

'==============================================================================================
' Method:       stoptrack_Click
'
' Description:  Stop track the popup menu in system tray area.
'==============================================================================================
Private Sub stoptrack_Click()
   
   TrayIcon1.TrackPopMenu = False 'Disable the Track of the popup menu hence ASTC will not destroy it whenever it lost focus
   MsgBox "ASTC stop tracking the popupmenu,Now right click the icon again and you will see that popupmenu willn't close unless you click an item in it."
   stoptrack.Enabled = False

End Sub

Private Sub menu2_Click()
   
   MsgBox "Menu item 'Show' had been clicked."

End Sub

'==============================================================================================
' Method:       TrayIcon1_BalloonLeftClick
'
' Description:  Display a message box whenever you left click the balloon tooltip.
'==============================================================================================
Private Sub TrayIcon1_BalloonLeftClick()
   
   MsgBox "The Balloon Tooltip had been left clicked." 'The Balloon ToolTip has been Left clicked.

End Sub

'==============================================================================================
' Method:       TrayIcon1_RightMouseUp
'
' Description:  Display a popup menu whenever you right click the icon in system tray.
'==============================================================================================
Private Sub TrayIcon1_RightMouseUp() 'The icon has been right clicked in the System Tray Area.
 
    Dim Menuhandle As Long
    Dim SubMenuHandle As Long
    Menuhandle = GetMenu(hwnd) 'Get a handle to the menu of this form [menu1]
    SubMenuHandle = GetSubMenu(Menuhandle, 0) 'Get a handel to the submenu of menu1
    TrayIcon1.PopUp SubMenuHandle, hwnd 'Pass the handle of the submenu and handle of this window to PopUp function in order to display a popup menu whenever the user right click this form

End Sub

'==============================================================================================
' Method:       cmdHide_Click
'
' Description:  Removes the icon from system tray.
'==============================================================================================
Private Sub cmdHide_Click()
   
    If TrayIcon1.IsDisplayed Then 'If the icon is displayed in System Tray Area
        TrayIcon1.Hide 'Then Destroy and Hide the icon from the System Tray Area.
    End If

End Sub

'==============================================================================================
' Method:        mexit_Click
'
' Description:   Removes the icon from system tray.
'==============================================================================================
Private Sub mexit_Click()
  
  If TrayIcon1.IsDisplayed Then 'If the icon is displayed in System Tray Area
     TrayIcon1.Hide 'Then Destroy and Hide the icon from the System Tray Area.
     End
  End If

End Sub

'==============================================================================================
' Method:       Form_Unload
'
' Description:  Removes the icon from system tray.
'==============================================================================================
Private Sub Form_Unload(Cancel As Integer)
   
   If TrayIcon1.IsDisplayed Then 'If the icon is displayed in System Tray Area
      TrayIcon1.Hide  'Then Destroy and Hide the icon from the System Tray Area.
   End If

End Sub

