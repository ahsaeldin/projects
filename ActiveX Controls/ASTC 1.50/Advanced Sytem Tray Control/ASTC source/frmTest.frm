VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "*\AASTC.vbp"
Begin VB.Form frmTrayIcon 
   Caption         =   "Form1"
   ClientHeight    =   2730
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   3105
   DrawMode        =   15  'Merge Pen Not
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   3105
   StartUpPosition =   2  'CenterScreen
   Begin ASTC.TrayList TrayList1 
      Left            =   1320
      Top             =   1080
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin ASTC.TrayIcon TrayIcon1 
      Left            =   1440
      Top             =   2040
      _extentx        =   953
      _extenty        =   953
   End
   Begin VB.Timer tmrCheck 
      Interval        =   100
      Left            =   2160
      Top             =   2160
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton cmdshow 
         Caption         =   "&Show Icon"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton stoptrack 
         Caption         =   "Stop Track"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdTrack 
         Caption         =   "TrackPopUp"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdCloseBalloon 
         Caption         =   "CloseBallon"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdShowBalloon 
         Caption         =   "ShowBallon"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton stopani 
         Caption         =   "Stop Animation"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdani 
         Caption         =   "Animate"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdHide 
         Caption         =   "&Hide Icon"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdchg 
         Caption         =   "&ChangeIcon"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1215
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1080
         Top             =   0
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
               Picture         =   "frmTest.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTest.frx":031A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTest.frx":0634
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTest.frx":094E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTest.frx":0C68
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTest.frx":0F82
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTest.frx":129C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTest.frx":15B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTest.frx":18D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTest.frx":1BEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTest.frx":1F04
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTest.frx":221E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu menu1 
      Caption         =   "File"
      Begin VB.Menu menu2 
         Caption         =   "msg"
      End
      Begin VB.Menu mexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hmenu As Long, ByVal nPos As Long) As Long

'1.Errors and msgbox and LastError Property


'2.ˆ«·”ƒ«· «·Â«„ Â· «·»—‰«„Ã Ê «·Â·» ”Â· «·«” œ«„ «·Ï Õœ «·Â»·


'=========================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
' Description:  A sample code demonstrate how to add your software icon to system tray area
'                as well as how to use the advanced features related to system tray area
'==========================================================================================
'ASTC provide you with a very simple object ,TrayIcon object which is the responsible to add the most of system
'tray functions to your applications.
'
'with a  few lines i can demonstrate how TrayIcon object can help you
'
'use TrayIcon.Show method to add an icon to system tray area
'use TrayIcon.Hide method to hide the icon you had add by TrayIcon.Show method from system tray
'use TrayIcon.ChangeIcon method to change the icon you had add by another icon
'use TrayIcon.Animate method to animate the icon in the system tray
'use TrayIcon.StopAnimateing method to stop animateding the icon
'use TrayIcon.ShowBalloon to display a tooltip balloon
'use TrayIcon.HideBalloon to hide the balloon
'use TrayIcon.Popup to display a popup menu
'use TrayIcon.TrackPopupmenu to track the popup menu whenever it lostfocus in order to close it.
'use TrayIcon.IsDisplayed to check if the icon is displayed in the system tray area
'
'note that you have a custom types of events associated with TrayIcon Object like
'BalloonShow,BalloonHide,MedMouseUp,LeftMouseUp,RightMouseUp,MedMouseDown,LeftMouseDown
'RightMouseDown ,MedMouseDBLCLK ,LeftMouseDBLCLK ,BalloonLeftClick ,BalloonRightClick ,RightMouseDBLCLK()
'
'if you have any further questions or need more sample code don't hesitate to contact us
'at support@cpringold.com


'==============================================================================
' Method:        cmdshow_Click
'
' Description:  show the icon in system tray
'==============================================================================
Private Sub cmdshow_Click()
 'add the icon to system tray
TrayIcon1.Show Me.Icon, "welcome1"
  
 cmdchg.Enabled = True
 cmdani.Enabled = True
 cmdTrack.Enabled = True
 cmdCloseBalloon.Enabled = True
 cmdHide.Enabled = True
 cmdShowBalloon.Enabled = True
 stopani.Enabled = True

End Sub

'==============================================================================
' Method:        cmdchg_Click
'
' Description:  change the icon in system tray area
'==============================================================================
Private Sub cmdchg_Click()
If TrayIcon1.IsDisplayed Then 'only if the icon is displayed
 'change the icon in system tray area by another one in ImageList
 TrayIcon1.ChangeIcon ImageList1.ListImages.Item(1).ExtractIcon
End If
End Sub

'==============================================================================
' Method:       cmdHide_Click
'
' Description:  remove the icon from system tray
'==============================================================================
Private Sub cmdHide_Click()

If TrayIcon1.IsDisplayed Then 'only if the icon is displayed
   'remove the icon from system tray
   TrayIcon1.Hide
End If
 
 cmdchg.Enabled = False
 cmdani.Enabled = False
 cmdTrack.Enabled = False
 cmdCloseBalloon.Enabled = False
 cmdHide.Enabled = False
 cmdShowBalloon.Enabled = False
 stopani.Enabled = False
End Sub

'==================================================================================
' Method:        cmdani_Click
'
' Description:  animate the icon in system tray using in icons in imagelist control
'==================================================================================
Private Sub cmdani_Click()
If TrayIcon1.IsDisplayed Then 'only if the icon is displayed
 
 If TrayIcon1.AnimateState = False Then 'only if the icon is not Animated
    'start animateing the icons in system tray
    'we use the icons in the imagelist control
    ''you replace the icons with your icons''
    'then
    'we pass the ImageList1 control to Animate method
    TrayIcon1.Animate ImageList1, 1
 End If

End If
End Sub



Private Sub Form_Load()
Dim TrayList() As TrayIconInfo

10 TrayList = TrayList1.GetSysTrayIcons
End Sub

'==============================================================================
' Method:       stopani_Click
'
' Description:  stop animateing the icons in system tray
'==============================================================================
Private Sub stopani_Click()

If TrayIcon1.IsDisplayed Then 'only if the icon is displayed
   
   If TrayIcon1.AnimateState = True Then 'only if the icon is animated
       'stop animateing the icon
       TrayIcon1.StopAnimateing ImageList1.ListImages.Item(2).ExtractIcon
   End If
End If

End Sub

'==============================================================================
' Method:       cmdShowBalloon_Click
'
' Description:  Displays a balloon tooltip points to the icon
'==============================================================================
Private Sub cmdShowBalloon_Click()
If TrayIcon1.IsDisplayed Then 'only if the icon is displayed
   'display a balloon tooltip pointing to the icon
'   TrayIcon1.ShowBalloon "Clickjsld;kfjgl;sdjflglksdjfglksd" & (13) & Chr(10) & "jflglsk;djfgkjsldkfjglkjsdflkgjlksdjfglkjdflgjlkdf" & Chr(13) & Chr(10) & "jglkdjflgkjdflgjdlfjglkdfjgljdflg;jdflgjlkdsjgljdflkgjldfsjglk;dfjglkjdflgjdflkgjdlfkjglkdfjgldjfglj" & Chr(13) & Chr(10) & "sflgjldsfjgldsjfgljsdlfgjls;djfg;ljk me gogogogogogogogogogogo", "My Balloon", Info, 8000
   TrayIcon1.ShowBalloon "asdf", "asdf", Warning
End If
End Sub

'==============================================================================
' Method:       cmdCloseBalloon_Click
'
' Description:  close the balloon tooltip
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
' Description:  setup the flexgrid control to receive the data
'==============================================================================
Private Sub cmdTrack_Click()
 'Track the popup menu so ASTC can destroy it whenever it lostfocus
 TrayIcon1.TrackPopMenu = True
 MsgBox "when you right click the icon a popupmenu will appear,ASTC will Track the popupmenu and close it lost focus"
 stoptrack.Enabled = True
End Sub

'==============================================================================
' Method:       stoptrack_Click
'
' Description:  stop track the popup menu in system tray area.
'==============================================================================
Private Sub stoptrack_Click()
 'stop tracking the popup menu
 TrayIcon1.TrackPopMenu = False
 MsgBox "ASTC stop tracking the popupmenu,Now right click the icon again and you will see that popupmenu willn't close unless you click an item in it."
 stoptrack.Enabled = False
End Sub

Private Sub menu2_Click()
MsgBox "you click show."
''''frmTrayList.Show
End Sub

'==============================================================================
' Method:        mexit_Click
'
' Description:  remove the icon from system tray if it was displayed
'==============================================================================
Private Sub mexit_Click()
If TrayIcon1.IsDisplayed Then
 TrayIcon1.Hide
 End
End If
End Sub

'==============================================================================
' Method:       TrayIcon1_BalloonLeftClick
'
' Description:  display a message box whenever you left click the balloon tooltip
'==============================================================================
Private Sub TrayIcon1_BalloonLeftClick()
 MsgBox "you clicked the balloon"
End Sub

'====================================================================================
' Method:       TrayIcon1_RightMouseUp
'
' Description:  display a popup menu whenever you right click the icon in system tray
'====================================================================================
Private Sub TrayIcon1_RightMouseUp()
Dim Menuhandle As Long
Dim SubMenuHandle As Long
   
   Menuhandle = GetMenu(hWnd)
   SubMenuHandle = GetSubMenu(Menuhandle, 0)
   TrayIcon1.PopUp SubMenuHandle, hWnd
   
End Sub

'==============================================================================
' Method:        Form_Unload
'
' Description:  remove the icon from system tray if it was displayed
'==============================================================================
Private Sub Form_Unload(cancel As Integer)


If TrayIcon1.IsDisplayed Then
'''TrayIcon1.Hide
End If
End Sub


