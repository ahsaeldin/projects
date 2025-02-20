VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{30E5838E-8E26-11DA-A1C2-807A50C10000}#1.0#0"; "astc.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TrayIcon Sample Code"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
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
Option Explicit
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hmenu As Long, ByVal nPos As Long) As Long
'=========================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
' Description:   A sample code demonstrate how to add your software icon to system tray area
'                As well as how to call the advanced features related to system tray area from
'                within your application.
'==========================================================================================
'ASTC provide you with a very simple object, TrayIcon object which is the responsible to add the most of system
'Tray functions to your application.
'
'With a few lines i can demonstrate how TrayIcon object can help you
'
'Use TrayIcon.Show method to add an icon to system tray area
'Use TrayIcon.Hide method to hide the icon you had add by TrayIcon.Show method from system tray
'Use TrayIcon. ChangeIcon method to change the icon you had added by another icon
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
'so don 't worry about subclassing
'If you have any further questions or need more sample code don't hesitate to contact us
'At support@cpringold.com

'==============================================================================
' Method:        cmdshow_Click
'
' Description:  Shows the icon in system tray.
'==============================================================================
Private Sub cmdshow_Click()

   'Unlike the previous version of ASTC, you don't need to pass a window handle to Show function because
   'ASTC dynamically creates an internal window related to ASTC instance, and destroy this window whenever
   'you call Hide method or close your application.
    TrayIcon1.Show Me.Icon, "Your Tooltip goes here."

End Sub

'==============================================================================
' Method:        cmdchg_Click
'
' Description:  Changes the icon in system tray area.
'==============================================================================
Private Sub cmdchg_Click()
   'only if the icon is displayed
   If TrayIcon1.IsDisplayed Then
      'change the icon in system tray area
      'by another one in ImageList
      TrayIcon1.ChangeIcon ImageList1.ListImages.Item(1).ExtractIcon
   End If
End Sub

'==================================================================================
' Method:        cmdani_Click
'
' Description:  Animate the icon in system tray using in icons in imagelist control.
'==================================================================================
Private Sub cmdani_Click()
   'only if the icon is displayed
   If TrayIcon1.IsDisplayed Then
 
    If TrayIcon1.AnimateState = False Then 'only if the icon is not Animated
       'Start animating the icons in system tray
       'we use the icons in the imagelist control
       'Then
       'We pass the ImageList1 control to Animate method
        '''you can replace the icons with your own'''
       TrayIcon1.Animate ImageList1, 1
   End If

   End If
End Sub

'==============================================================================
' Method:       stopani_Click
'
' Description:  Stop animating the icons in system tray.
'==============================================================================
Private Sub stopani_Click()

   If TrayIcon1.IsDisplayed Then 'only if the icon is displayed
      'only if the icon is animated
      If TrayIcon1.AnimateState = True Then
          'stop animateing the icon
          TrayIcon1.StopAnimateing ImageList1.ListImages.Item(2).ExtractIcon
      End If
   End If

End Sub

'==============================================================================
' Method:       cmdShowBalloon_Click
'
' Description:  Displays a balloon tooltip points to the icon.
'==============================================================================
Private Sub cmdShowBalloon_Click()
   If TrayIcon1.AnimateState = True Then MsgBox ("You can't Display a balloon while animating the icon.")
   If TrayIcon1.IsDisplayed Then 'only if the icon is displayed
      'display a balloon tooltip pointing to the icon
      TrayIcon1.ShowBalloon "Click me", "My Balloon", Info, 5000
   End If
End Sub

'==============================================================================
' Method:       cmdCloseBalloon_Click
'
' Description:  Close the balloon tooltip.
'==============================================================================
Private Sub cmdCloseBalloon_Click()
   If TrayIcon1.IsDisplayed Then 'only if the icon is displayed
      'Close the balloon tooltip
      TrayIcon1.CloseBalloon
      Me.Show
   End If
End Sub
'==============================================================================
' Method:       cmdTrack_Click
'
' Description:  Start Tracking the popup menu in system tray area.
'==============================================================================
Private Sub cmdTrack_Click()
   'Track the popup menu so ASTC can destroy it whenever it lost focus
   TrayIcon1.TrackPopMenu = True
   MsgBox "when you right click the icon a popupmenu will appear,ASTC will Track the popupmenu and close it lost focus"
   stoptrack.Enabled = True
End Sub

'==============================================================================
' Method:       stoptrack_Click
'
' Description:  Stop track the popup menu in system tray area.
'==============================================================================
Private Sub stoptrack_Click()
   'stop tracking the popup menu
   TrayIcon1.TrackPopMenu = False
   MsgBox "ASTC stop tracking the popupmenu,Now right click the icon again and you will see that popupmenu willn't close unless you click an item in it."
   stoptrack.Enabled = False
End Sub
'==============================================================================
' Method:        mexit_Click
'
' Description:  Removes the icon from system tray.
'==============================================================================
Private Sub mexit_Click()
  If TrayIcon1.IsDisplayed Then
     TrayIcon1.Hide
     End
  End If
End Sub

Private Sub menu2_Click()
   MsgBox "Menu item 'Show' had been clicked."
End Sub

'==============================================================================
' Method:       TrayIcon1_BalloonLeftClick
'
' Description:  Display a message box whenever you left click the balloon tooltip.
'==============================================================================
Private Sub TrayIcon1_BalloonLeftClick()
   MsgBox "The Balloon Tooltip had been left clicked."
End Sub

'====================================================================================
' Method:       TrayIcon1_RightMouseUp
'
' Description:  Display a popup menu whenever you right click the icon in system tray.
'====================================================================================
Private Sub TrayIcon1_RightMouseUp()
 Dim Menuhandle As Long
 Dim SubMenuHandle As Long
   Menuhandle = GetMenu(hwnd)
   SubMenuHandle = GetSubMenu(Menuhandle, 0)
   TrayIcon1.PopUp SubMenuHandle, hwnd
End Sub

'==============================================================================
' Method:       cmdHide_Click
'
' Description:  Removes the icon from system tray.
'==============================================================================
Private Sub cmdHide_Click()
   'only if the icon is displayed
   If TrayIcon1.IsDisplayed Then
      TrayIcon1.Hide
   End If
End Sub

'==============================================================================
' Method:       Form_Unload
'
' Description:  Removes the icon from system tray whenever the
'               terminates the program.
'==============================================================================

Private Sub Form_Unload(Cancel As Integer)
   'only if the icon is displayed
   If TrayIcon1.IsDisplayed Then
      TrayIcon1.Hide
   End If
End Sub
