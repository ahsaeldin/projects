VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmBalloon 
   Caption         =   "Form3"
   ClientHeight    =   9120
   ClientLeft      =   10335
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form3"
   ScaleHeight     =   9120
   ScaleWidth      =   7800
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3360
      Top             =   4320
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   4680
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1931
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   1935
   End
   Begin VB.PictureBox Balloon1 
      Height          =   480
      Left            =   6720
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   4
      Top             =   240
      Width           =   1200
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmTest3.frx":0000
      Height          =   2895
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   $"frmTest3.frx":0190
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   6480
      Width           =   7575
   End
End
Attribute VB_Name = "frmBalloon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'1.Errors and msgbox and LastError Property


'2.ˆ«·”ƒ«· «·Â«„ Â· «·»—‰«„Ã Ê «·Â·» ”Â· «·«” œ«„ «·Ï Õœ «·Â»·


'=========================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
' Description:  A sample code demonstrate how to use Balloon Object .
'==========================================================================================

'Use the Optional parameter PHwnd to insure that the balloon tooltip will not appear outside a specified
'window or a child window region like a command button or a text box.
'note that:
'
'1.you can't set this parameter to any external window i.e. you can't set to an external Apps window
'
'2.you can set this parameter to the window which contains the current instance of the balloon control
'
'i.e.  for the following code
'
'balloon1.ShowBalloon From1.hwnd
'
'The hwnd is a handle to the window in which contains the Balloon control and that's means that the balloon
'will only appear if the mouse is over the Form itself not any of its child controls [child windows] in which
'it contains.
'
'3.you can set this parameter to any control of the window that contains the current instance of the balloon
'control i.e if you have a Form called Form1 which has a command button called Command1 and a text box called
'text1,you can set the parameter as the following
'
'balloon1.ShowBalloon command1.hwnd
'
'balloon1.ShowBalloon text1.hwnd
'
'and that's means the balloon will only appear if the mouse is over the control itself
'
'4. if you didn't pass this optional  parameter,  Balloon1 object will use its parent window automatically
'i.e. if you have a form called Form1 ,Balloon object will Form1.Hwnd for PHwnd parameter silently but in
'this case the balloon tooltip may appear at any part of the Form or it's Controls
'
'use DelayTime parameter to set the Delay Time before the balloon appear, the default value for this
'parameter is 2000 milliseconds and The maximum, value for DelayTime Parameter is 65,535 milliseconds,
'is equivalent to just over 1 minute.
'The maximum, value for Timeout Parameter is 65,535 milliseconds, is equivalent to just over 1 minute.
'

Private Type POINTAPI
    X As Long
    y As Long
End Type

Private Type HDHITTESTINFO
    pt As POINTAPI
    flags As Long
    iItem As Long
End Type

Private Sub Balloon1_BalloonLeftClick()
MsgBox "left"
End Sub

Private Sub Balloon1_BalloonRightClick()
MsgBox "Right"
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  Balloon1.BackColor = RGB(255, 255, 255)
  Balloon1.ForeColor = RGB(255, 50, 150)
  
'Style Property Returns or sets the balloon tooltip shape [balloon or Rectangle].
Balloon1.Style = BalloonType
  
'ShowBalloon Method
'Parameters:
'PHwnd      Long Optional. Handle to a window in which you may assign the balloon for.
'Prompt     String  Optional. the balloon title.
'Title      String Optional. the balloon title.
'IconType   BIcoType Optional. the Icon Type in which will be displayed in the balloon.
'DelayTime  Long Optional. The  Delay Time before the balloon appear.
'Timeout    Long Optional. The Timeout for the balloon to be closed automatically.
 Balloon1.ShowBalloon Command1.hWnd, "Command1 balloon tooltip", "", Info, 1000, 4000

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
  Balloon1.Destroy
End Sub

