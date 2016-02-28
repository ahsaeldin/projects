VERSION 5.00
Object = "{30E5838E-8E26-11DA-A1C2-807A50C10000}#1.0#0"; "astc.ocx"
Begin VB.Form frmSubClass 
   Caption         =   "SubClass Sample Code"
   ClientHeight    =   4830
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   StartUpPosition =   2  'CenterScreen
   Begin ASTC.Balloon Balloon2 
      Left            =   5520
      Top             =   2520
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin ASTC.Balloon Balloon1 
      Left            =   5520
      Top             =   3240
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Frame Frame2 
         ForeColor       =   &H00FF0000&
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5775
         Begin ASTC.SubClass SubClass1 
            Left            =   4680
            Top             =   2880
            _ExtentX        =   953
            _ExtentY        =   953
         End
         Begin VB.CommandButton cmdSubclass 
            Caption         =   "Subclass"
            Height          =   375
            Left            =   1440
            TabIndex        =   7
            Top             =   3000
            Width           =   1335
         End
         Begin VB.CommandButton cmdStopSubclass 
            Caption         =   "Stop Subclass"
            Height          =   375
            Left            =   2880
            TabIndex        =   6
            Top             =   3000
            Width           =   1335
         End
         Begin VB.ListBox List1 
            Height          =   1035
            ItemData        =   "frmSubClass.frx":0000
            Left            =   600
            List            =   "frmSubClass.frx":0010
            TabIndex        =   2
            Top             =   1440
            Width           =   4455
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmSubClass.frx":0030
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   480
            TabIndex        =   8
            Top             =   3600
            Width           =   4815
         End
         Begin VB.Label Label4 
            Caption         =   $"frmSubClass.frx":00AA
            ForeColor       =   &H00FF0000&
            Height          =   705
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   5535
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmSubClass.frx":019F
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   600
            TabIndex        =   4
            Top             =   960
            Width           =   4815
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   600
            TabIndex        =   3
            Top             =   2520
            Width           =   3855
         End
      End
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu menu2 
         Caption         =   "Click Me"
      End
   End
End
Attribute VB_Name = "frmSubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
'Description:  A sample code demonstrates how to use ASTC control to subclass any Form.
'==========================================================================================
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

'What is the Subclass?
'Subclassing is the processing of intercepting Windows messages that your program normally wouldn't
'Receive, so it extends your ability to process more windows messages and add new features to your program.
''
'With a few lines i can demonstrate how Subclass object can help you
'
'1.use SubClass.BeginSubClass to start subclassing a specified window.
'2.use SubClass.EndSubClass to end subclassing.
'3.use SubClass_MessageReceived event to process the message you want.
'
'About this sample code
'This sample code show you the benefits of subclassing by demonstrate
'How to extends a listbox functions?

'Don't hesitate to contact us At support@cpringold.com if you have any further questions or need more sample code
'

'==============================================================================
' Method:        cmdSubclass_Click
'
' Description:  Start subclassing List1.
'==============================================================================
Private Sub cmdSubclass_Click()
  SubClass1.BeginSubClass List1.hWnd
End Sub

'=========================================================================================
' Method:        SubClass1_MessageReceived
'
' Description:   Extends the listbox functionality by using a subclass and balloon tooltip.
'=========================================================================================
Private Sub Subclass1_MessageReceived(hWnd As Long, Msg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
    
    'Hwnd    Long    Handle of the subclassed window or control.
    'Msg     Long    The ID of the intercepted message.
    'wParam  Long    The wParam value of the intercepted message.
    'lParam  Long    The lParam value of the intercepted message.
    'Cancel  Boolean Further processing state.
    Dim ret As Long
    Dim i As Integer
    
    Static OldItem As Integer
    
    Const WM_MOUSEMOVE = &H200
    Const LB_ITEMFROMPOINT = 425
    
    
    If Msg = WM_MOUSEMOVE Then
       
       'retrieve the zero-based index of the item nearest a specified point in the list box.
        ret = SendMessage(hWnd, LB_ITEMFROMPOINT, ByVal 0&, ByVal lParam)
        For i = 0 To List1.ListCount - 1
        
            If ret = i Then
                If OldItem <> i Then Balloon1.Destroy: Balloon2.Destroy
                List1.ListIndex = i
                Label2.Caption = List1.List(i) & " had been hovered."
               'Extends Listbox functionality
                If ret = 3 Then
                   Balloon2.ShowBalloon List1.hWnd, "Click Me.", "SubClass", Info, 100, 5000
                Else
                   Balloon1.ShowBalloon List1.hWnd, List1.List(i) & " had been hovered.", "SubClass", Info, 100, 5000
                End If
                OldItem = i
            End If
            
        Next i
    
    End If

End Sub

Private Sub Balloon2_BalloonLeftClick()
  MsgBox List1.List(List1.ListIndex) & " had been clicked"
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Balloon1.Destroy
  Balloon2.Destroy
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Balloon1.Destroy
   Balloon2.Destroy
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Balloon1.Destroy
   Balloon2.Destroy
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Balloon1.Destroy
   Balloon2.Destroy
End Sub

'==============================================================================
' Method:        cmdStopSubclass_Click
'
' Description:  Stop Subclassing List1.
'==============================================================================
Private Sub cmdStopSubclass_Click()
   Label2.Caption = ""
   SubClass1.EndSubClass
  'if you didn't end the subclassing
  'that ASTC will terminate the subclassing automatically
  'when your program closed.
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Balloon1.Destroy
  Balloon2.Destroy
  SubClass1.EndSubClass
End Sub
