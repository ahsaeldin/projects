VERSION 5.00
Object = "{30E5838E-8E26-11DA-A1C2-807A50C10000}#1.0#0"; "ASTC.OCX"
Begin VB.Form cmdBalloon 
   Caption         =   "Balloon ToolTip Sample"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin ASTC.Balloon Balloon2 
      Left            =   240
      Top             =   1080
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin ASTC.Balloon Balloon1 
      Left            =   4200
      Top             =   1080
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton cmdBalloon3 
         Caption         =   "Balloon3"
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdBalloon1 
         Caption         =   "Balloon1"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdBalloon2 
         Caption         =   "Balloon2"
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Move the Cursor to every command button to see the balloon tooltip."
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   3975
      End
   End
End
Attribute VB_Name = "cmdBalloon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=========================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
' Description:  A sample code demonstrates how to use Balloon Object.
'==========================================================================================

'Use the Optional parameter PHwnd to insure that the balloon tooltip will not appear outside a specified
'Window or a child window region like a command button or a text box.
'Note that:
'
'1.you can't set this parameter to any external window i.e. you can't set to an external Apps window
'
'2.you can set this parameter to the window which contains the current instance of the balloon control
'
'i.e.  For the following code
'
'balloon1.ShowBalloon From1.hwnd
'
'The hwnd is a handle to the window in which contains the Balloon control and that's means that the balloon
'Will only appear if the mouse is over the Form itself not any of its child controls [child windows] in which
'It contains.
'
'3.you can set this parameter to any control of the window that contains the current instance of the balloon
'Control i.e if you have a Form called Form1 which has a command button called Command1 and a text box called
'text1, you can set the parameter as the following
'
'balloon1.ShowBalloon command1.hwnd
'
'balloon1.ShowBalloon text1.hwnd
'
'And that's means the balloon will only appear if the mouse is over the control itself
'
'4. If you didn't pass this optional parameter, Balloon1 object will use its parent window automatically
'I.e. if you have a form called Form1, Balloon object will Form1.Hwnd for PHwnd parameter silently but in
'This case the balloon tooltip may appear at any part of the Form or its Controls
'
'Use DelayTime parameter to set the Delay Time before the balloon appear, the default value for this
'Parameter is 2000 milliseconds and the maximum, value for DelayTime Parameter is 65,535 milliseconds,
'Is equivalent to just over 1 minute?
'The maximum, value for Timeout Parameter is 65,535 milliseconds, is equivalent to just over 1 minute.
'
'==============================================================================
' Method:        Form_MouseMove
'
' Description:   Destroy the Balloon Tooltip.
'==============================================================================
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Balloon1.Destroy
   Balloon2.Destroy
End Sub

'==============================================================================
' Method:        Frame1_MouseMove
'
' Description:   Destroy the Balloon Tooltip.
'==============================================================================
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Balloon1.Destroy
   Balloon2.Destroy
End Sub

'==============================================================================
' Method:        Label1_MouseMove
'
' Description:   Destroy the Balloon Tooltip.
'==============================================================================

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Balloon1.Destroy
   Balloon2.Destroy
End Sub

'==============================================================================
' Method:        cmdBalloon1_MouseMove
'
' Description:  Display a balloon tooltip for cmdBalloon1.
'==============================================================================
Private Sub cmdBalloon1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Balloon1.Style = BalloonType
   Balloon1.ShowBalloon cmdBalloon1.hWnd, "Click Me", "Balloon1", Info, 500
End Sub

'==============================================================================
' Method:        Balloon1_BalloonLeftClick
'
' Description:   Display a message Box when the left click the balloon.
'==============================================================================
Private Sub Balloon1_BalloonLeftClick()
  MsgBox "The Balloon Tooltip had been Clicked"
End Sub

'==============================================================================
' Method:        cmdBalloon2_MouseMove
'
' Description:   Display a balloon tooltip for cmdBalloon2.
'==============================================================================
Private Sub cmdBalloon2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Balloon1.Style = RectangleType
   Balloon1.ShowBalloon cmdBalloon2.hWnd, "Rectangle Type", "Balloon2", Info, 500
End Sub

'==============================================================================
' Method:        cmdBalloon3_MouseMove
'
' Description:   Display a balloon tooltip for cmdBalloon3.
'==============================================================================
Private Sub cmdBalloon3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Balloon2.BackColor = RGB(255, 255, 255)
   Balloon2.ForeColor = RGB(255, 50, 150)
   Balloon2.ShowBalloon cmdBalloon3.hWnd, "Customizable balloon tooltip", "Balloon3", Info, 500
End Sub

