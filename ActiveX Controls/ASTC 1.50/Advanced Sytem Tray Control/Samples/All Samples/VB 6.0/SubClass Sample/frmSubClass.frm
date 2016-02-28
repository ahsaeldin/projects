VERSION 5.00
Object = "{30E5838E-8E26-11DA-A1C2-807A50C10000}#1.0#0"; "astc.ocx"
Begin VB.Form frmSubClass 
   Caption         =   "SubClass Sample Code"
   ClientHeight    =   8985
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.Frame Frame4 
         Caption         =   "SubClass Sample 3"
         ForeColor       =   &H00FF0000&
         Height          =   3015
         Left            =   120
         TabIndex        =   8
         Top             =   5760
         Width           =   4215
         Begin ASTC.SubClass SubClass3 
            Left            =   3600
            Top             =   2400
            _ExtentX        =   953
            _ExtentY        =   953
         End
         Begin VB.CommandButton cmdStopSubText 
            Caption         =   "Stop SubClass"
            Height          =   375
            Left            =   2160
            TabIndex        =   15
            Top             =   2520
            Width           =   1335
         End
         Begin VB.CommandButton cmdSubClassText 
            Caption         =   "SubClass"
            Height          =   375
            Left            =   720
            TabIndex        =   14
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Height          =   975
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   9
            Text            =   "frmSubClass.frx":0000
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label8 
            Caption         =   "You can use Subclassing to disable the Textbox context menu in which it appears when the user right click the Text box."
            Height          =   735
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   840
            TabIndex        =   10
            Top             =   2040
            Width           =   2415
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "SubClass Sample 2"
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   120
         TabIndex        =   6
         Top             =   4320
         Width           =   4215
         Begin ASTC.SubClass SubClass2 
            Left            =   3600
            Top             =   720
            _ExtentX        =   953
            _ExtentY        =   953
         End
         Begin VB.CommandButton cmdStopSubClassForm 
            Caption         =   "Stop SubClass"
            Height          =   375
            Left            =   2160
            TabIndex        =   13
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton cmdSubForm 
            Caption         =   "SubClass Form"
            Height          =   375
            Left            =   720
            TabIndex        =   12
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Try to resize this Form before and after you click the subclass Form command buttton and watch the difference."
            Height          =   615
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Subclass Sample 1"
         ForeColor       =   &H00FF0000&
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4215
         Begin ASTC.SubClass SubClass1 
            Left            =   3600
            Top             =   3360
            _ExtentX        =   953
            _ExtentY        =   953
         End
         Begin VB.CommandButton cmdSubclass 
            Caption         =   "Subclass"
            Height          =   375
            Left            =   720
            TabIndex        =   17
            Top             =   3480
            Width           =   1335
         End
         Begin VB.CommandButton cmdStopSubclass 
            Caption         =   "Stop Subclass"
            Height          =   375
            Left            =   2160
            TabIndex        =   16
            Top             =   3480
            Width           =   1335
         End
         Begin VB.ListBox List1 
            Height          =   1035
            ItemData        =   "frmSubClass.frx":0063
            Left            =   120
            List            =   "frmSubClass.frx":0073
            TabIndex        =   2
            Top             =   1920
            Width           =   3855
         End
         Begin VB.Label Label4 
            Caption         =   $"frmSubClass.frx":0093
            Height          =   975
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Right Click any item in listbox before you click the Subclass command button."
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   120
            TabIndex        =   4
            Top             =   1440
            Width           =   3735
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Then click SubClass command and Right Click any item and you will see the difference."
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   120
            TabIndex        =   3
            Top             =   3000
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
Option Explicit

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

'=========================================================================================
'Publisher      CprinGold Software.
'               http://www.cpringold.com
'               support@cpringold.com
'
'
'Description:  A sample code demonstrates how to use ASTC control to subclass any Form.
'==========================================================================================
'First of all what is the Subclass?
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
'How to disable the context menu in a text box?
'How to extend a listbox functions?
'How to stop the form from being resized below or above a user-defined amount?

'That's it
'If you have any further questions or need more sample code don't hesitate to contact us
'At support@cpringold.com

'==============================================================================
' Method:        cmdSubclass_Click
'
' Description:  Start subclassing List1.
'==============================================================================
Private Sub cmdSubclass_Click()
  SubClass1.BeginSubClass List1.hWnd
End Sub

'==============================================================================
' Method:        SubClass1_MessageReceived
'
' Description:   Extends the listbox functionality.
'==============================================================================

Private Sub Subclass1_MessageReceived(hWnd As Long, Msg As Long, wParam As Long, lParam As Long, cancel As Boolean)
'Hwnd    Long    Handle of the subclassed window or control.
'Msg     Long    The ID of the intercepted message.
'wParam  Long    The wParam value of the intercepted message.
'lParam  Long    The lParam value of the intercepted message.
'Cancel  Boolean Further processing state.
Dim i As Integer
Dim ret As Long

Const WM_RBUTTONUP = 517
Const LB_ITEMFROMPOINT = 425

If Msg = WM_RBUTTONUP Then

    'retrieve the zero-based index of the item nearest a specified point in the list box.
    ret = SendMessage(hWnd, LB_ITEMFROMPOINT, ByVal 0&, ByVal lParam)
    For i = 0 To List1.ListCount - 1
        If ret = i Then
           List1.ListIndex = i
           Label2.Caption = List1.List(i) & " had been clicked"
           'Extends Listbox functionality
           PopupMenu menu1
        End If
        
    Next i

End If

End Sub

Private Sub menu2_Click()
  MsgBox List1.List(List1.ListIndex) & " had been clicked"
End Sub

'==============================================================================
' Method:        cmdStopSubclass_Click
'
' Description:  Stop Subclassing List1.
'==============================================================================
Private Sub cmdStopSubclass_Click()
   SubClass1.EndSubClass
  'if you didn't end the subclassing
  'that ASTC will terminate the subclassing automatically
  'when your program closed.
End Sub

'==============================================================================
' Method:        cmdSubForm_Click
'
' Description:  Start subclassing the Form.
'==============================================================================
Private Sub cmdSubForm_Click()
  SubClass2.BeginSubClass Me.hWnd
End Sub

'==============================================================================
' Method:        SubClass2_MessageReceived
'
' Description:  Stop the form from being resized below or above a user-defined amount.
'==============================================================================

Private Sub SubClass2_MessageReceived(hWnd As Long, Msg As Long, wParam As Long, lParam As Long, cancel As Boolean)

'Hwnd    Long    Handle of the subclassed window or control.
'Msg     Long    The ID of the intercepted message.
'wParam  Long    The wParam value of the intercepted message.
'lParam  Long    The lParam value of the intercepted message.
'Cancel  Boolean Further processing state.

Const WM_GETMINMAXINFO = 36

Dim mmiMsg As MINMAXINFO

   If Msg = WM_GETMINMAXINFO Then
            
            CopyMemory mmiMsg, ByVal lParam, LenB(mmiMsg)
            If mmiMsg.ptMinTrackSize.X > 0 Then mmiMsg.ptMinTrackSize.X = 4770 / Screen.TwipsPerPixelX
            If mmiMsg.ptMinTrackSize.Y > 0 Then mmiMsg.ptMinTrackSize.Y = 9495 / Screen.TwipsPerPixelY
            CopyMemory ByVal lParam, mmiMsg, LenB(mmiMsg)
            
           'Cancel any further processing.
            cancel = True

   End If
            
End Sub

'==============================================================================
' Method:        cmdStopSubClassForm_Click
'
' Description:  Stop Subclassing Form1.
'==============================================================================
Private Sub cmdStopSubClassForm_Click()
   SubClass2.EndSubClass
End Sub

'==============================================================================
' Method:        cmdSubClassText_Click
'
' Description:  Starts subclassing Text1.
'==============================================================================

Private Sub cmdSubClassText_Click()
    SubClass3.BeginSubClass Text1.hWnd
End Sub

'==============================================================================
' Method:        SubClass3_MessageReceived
'
' Description:   Disable a context menu in a text box.
'==============================================================================
Private Sub SubClass3_MessageReceived(hWnd As Long, Msg As Long, wParam As Long, lParam As Long, cancel As Boolean)
'Hwnd    Long    Handle of the subclassed window or control.
'Msg     Long    The ID of the intercepted message.
'wParam  Long    The wParam value of the intercepted message.
'lParam  Long    The lParam value of the intercepted message.
'Cancel  Boolean Further processing state.
        
Const WM_CONTEXTMENU = 123

If Msg = WM_CONTEXTMENU Then
    Label6.Caption = "Context Menu had been disabled"
    'Cancel any further processing.
    cancel = True
End If
End Sub
'==============================================================================
' Method:        cmdStopSubText_Click
'
' Description:  Stop Subclassing Text1.
'==============================================================================
Private Sub cmdStopSubText_Click()
   SubClass3.EndSubClass
   Label6.Caption = ""
End Sub

