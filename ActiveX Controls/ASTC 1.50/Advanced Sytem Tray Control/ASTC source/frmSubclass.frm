VERSION 5.00
Object = "*\AI:\Documents and Settings\Ahmed\Desktop\ASTC 1.50 Pro\ASTC.vbp"
Begin VB.Form frmsubclass 
   Caption         =   "Subclass Sample Code"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   6060
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   11805
   Begin ASTC.SubClass SubClass3 
      Left            =   5760
      Top             =   1560
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin ASTC.SubClass SubClass2 
      Left            =   4440
      Top             =   2040
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin ASTC.SubClass SubClass1 
      Left            =   4800
      Top             =   840
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin ASTC.TrayIcon TrayIcon2 
      Left            =   4680
      Top             =   3720
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1080
      TabIndex        =   16
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      Caption         =   "Sample3"
      Height          =   1335
      Left            =   480
      TabIndex        =   12
      Top             =   5160
      Width           =   4575
      Begin VB.Label Label7 
         Caption         =   "Try to resize this Form before and after you click the subclass command buttton and watch the difference."
         Height          =   615
         Left            =   360
         TabIndex        =   13
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sample1"
      Height          =   2175
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Width           =   3375
      Begin VB.TextBox Text1 
         Height          =   975
         Left            =   480
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "frmSubclass.frx":0000
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   1560
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sample2"
      Height          =   2895
      Left            =   6960
      TabIndex        =   2
      Top             =   1680
      Width           =   4455
      Begin VB.ListBox List1 
         Height          =   1035
         ItemData        =   "frmSubclass.frx":0063
         Left            =   120
         List            =   "frmSubclass.frx":0073
         TabIndex        =   3
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   3855
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Right Click any item in listbox before you click the Subclass command button."
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdStopSubclass 
      Caption         =   "Stop Subclass"
      Height          =   375
      Left            =   9360
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdSubclass 
      Caption         =   "Subclass"
      Height          =   375
      Left            =   7800
      TabIndex        =   0
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "check that every Method paramter is byval for ,net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   975
      Left            =   5400
      TabIndex        =   17
      Top             =   5280
      Width           =   4695
   End
   Begin VB.Label Label10 
      Caption         =   "write in Help that in ASTC has a help senstive"
      Height          =   615
      Left            =   5280
      TabIndex        =   15
      Top             =   7920
      Width           =   4095
   End
   Begin VB.Label Label9 
      Caption         =   "check that every Method paramter is byval for ,net"
      Height          =   615
      Left            =   5400
      TabIndex        =   14
      Top             =   6480
      Width           =   3495
   End
   Begin VB.Label Label6 
      Caption         =   "You have three samples of how to use subclass"
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "you can use Subclassing to disable the Textbox context menu in which it appears when the user right click the Text box"
      Height          =   735
      Left            =   360
      TabIndex        =   10
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   $"frmSubclass.frx":0093
      Height          =   975
      Left            =   6960
      TabIndex        =   9
      Top             =   600
      Width           =   4335
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu"
      Begin VB.Menu menu2 
         Caption         =   "Click me"
      End
   End
End
Attribute VB_Name = "frmsubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'1.Errors and msgbox and LastError Property


'2.öÇáÓÄÇá ÇáåÇã åá ÇáÈÑäÇãÌ æ ÇáåáÈ Óåá ÇáÇÓÊÏÇã Çáì ÍÏ ÇáåÈá


'=========================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
' Description:  A sample code demonstrate how to use ASTC control to subclass any Form.
'==========================================================================================
'First of all what is the Subclass?
'Subclassing is the processing of intercepting Windows messages that your program normally wouldn't
'receive,so it extends your ability to process more windows messages and add new features to your program.
''
'with a  few lines i can demonstrate how Subclass object can help you
'
'1.use SubClass.BeginSubClass to start subclassing a specified window.
'2.use SubClass.EndSubClass to end subclassing.
'3.use SubClass_MessageReceived event to process the message you want.
'
'about this sample code
'this sample code show you the benefits of subclassing by demonstrate
'how to disable the context menu in a text box?
'how to extend a listbox functions?
'how to stop the form from being resized below or above a user-defined amount?

'that's it
'if you have any fruther questions or need more sample code don't hesitate to contact us
'at support@cpringold.com

'==============================================================================
' Method:        cmdSubclass_Click
'
' Description:  start subclassing the form,Text1 and List1
'==============================================================================
Dim ahmed As Boolean

Private Sub cmdSubclass_Click()
  SubClass1.BeginSubClass hWnd
  SubClass2.BeginSubClass Text1.hWnd
  SubClass3.BeginSubClass List1.hWnd
End Sub

'==============================================================================
' Method:        cmdStopSubclass_Click
'
' Description:  End Subclassing
'==============================================================================
Private Sub cmdStopSubclass_Click()
  Label1.Caption = ""
  Label2.Caption = ""
  SubClass1.EndSubclass
  SubClass2.EndSubclass
  SubClass3.EndSubclass
  'if you didn't end the subclassing
  'that ASTC will terminate the subclassing automatically
  'when your program closed.
End Sub

Private Sub Command1_Click()
SubClass2.BeginSubClass cmdSubclass.hWnd
''SubClass3.BeginSubClass Label1
End Sub

Private Sub menu2_Click()
  MsgBox List1.List(List1.ListIndex) & " had been clicked"
End Sub

'==============================================================================
' Method:        SubClass1_MessageReceived
'
' Description:  stop the form from being resized below or above a user-defined amount
'==============================================================================

Private Sub Subclass1_MessageReceived(hWnd As Long, Msg As Long, wParam As Long, lParam As Long, cancel As Boolean)

'hwnd Long Handle of the Subclassed window or control.
'Msg Long  The ID of the intercepted message.
'wParam Long  The wParam value of the intercepted message.
'lParam Long The lParam value of the intercepted message.
'cancel Boolean  Further processing state.

Dim mmiMsg As MINMAXINFO

If Msg = WM_GETMINMAXINFO Then
            
            CopyMemory mmiMsg, ByVal lParam, LenB(mmiMsg)
            If mmiMsg.ptMinTrackSize.X > 0 Then mmiMsg.ptMinTrackSize.X = 200
            If mmiMsg.ptMinTrackSize.y > 0 Then mmiMsg.ptMinTrackSize.y = 300
            'If mmiMsg.ptMinTrackSize.x > 0 Then mmiMsg.ptMinTrackSize.x = 800
            'If mmiMsg.ptMinTrackSize.y > 0 Then mmiMsg.ptMinTrackSize.x = 1000
            CopyMemory ByVal lParam, mmiMsg, LenB(mmiMsg)
            
           'Cancel any further processing.
            cancel = True

End If
            
End Sub

'==============================================================================
' Method:        SubClass2_MessageReceived
'
' Description:   Disable the context menu in a text box
'==============================================================================
Private Sub SubClass2_MessageReceived(hWnd As Long, Msg As Long, wParam As Long, lParam As Long, cancel As Boolean)

'hwnd Long Handle of the Subclassed window or control.
'Msg Long  The ID of the intercepted message.
'wParam Long  The wParam value of the intercepted message.
'lParam Long The lParam value of the intercepted message.
'cancel Boolean  Further processing state.

If Msg = WM_CONTEXTMENU Then
    Label1.Caption = "Context Menu had been disabled"
    'Cancel any further processing.
    cancel = True
End If
End Sub

'==============================================================================
' Method:        SubClass3_MessageReceived
'
' Description:   Extend a listbox functionality
'==============================================================================

Private Sub SubClass3_MessageReceived(hWnd As Long, Msg As Long, wParam As Long, lParam As Long, cancel As Boolean)

'hwnd Long Handle of the Subclassed window or control.
'Msg Long  The ID of the intercepted message.
'wParam Long  The wParam value of the intercepted message.
'lParam Long The lParam value of the intercepted message.
'cancel Boolean  Further processing state.

Dim i As Integer
Dim ret As Long

If Msg = WM_RBUTTONUP Then

    'retrieve the zero-based index of the item nearest a specified point in the list box.
    ret = SendMessage(hWnd, LB_ITEMFROMPOINT, ByVal 0&, ByVal lParam)
    For i = 0 To List1.ListCount - 1
        If ret = i Then
           List1.ListIndex = i
           Label2.Caption = List1.List(i) & " had been clicked"
           'Extend Listbox functionality
           PopupMenu menu1
        End If
        
    Next i

End If

End Sub

