VERSION 5.00
Object = "*\A..\..\..\..\Documents and Settings\Ahmed\Desktop\ASTC 1.50 Pro\ASTC.vbp"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTrayList 
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   Icon            =   "frmTest2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9000
      Top             =   3720
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   100
      Cols            =   8
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send Double Click Message"
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   3720
      Width           =   1935
   End
   Begin ASTC.TrayList TrayList1 
      Left            =   9720
      Top             =   3720
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   7080
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   3255
      Begin VB.CommandButton cmdFeedback 
         Caption         =   "Ask Questions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdHide 
         Caption         =   "&Hide Icon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmTrayList"
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
' Description:  A sample code demonstrate how to use ASTC control to get all system tray icons
'                informations as well as how to hide a specified icon from system tray area.
'==========================================================================================

'first of all don't afraid from all these lines of codes ,the main line which return the system
'tray information is only one line which located in FillFlexGrid method [CoreLine] at line 10
'all the remaining lines is to fill the returnd data from GetSysTrayIcons method into FlexGrid
'to display the data in a good manner.
'how i return these informations?
'well
'GetSysTrayIcons method return an array of TrayIconInfo type
'wait what does it means?
'we have a Type Called TrayIconInfo , you can use Object Browser to see this type in ASTC
'when you call GetSysTrayIcons method ,it returns an array of this type contains the informations
'i don't understand
'well take this example
'if the system tray area contains three icons say [sound control icon ,
'mcafee icon and network connection icon].
'then the returned data will be organized in an Array and the array will have three elements
'and every element will be a type called TrayIconInfo and TrayIconInfo will contains the information.
'
'so what is these informations?
'
'Public Type TrayIconInfo
'    hwnd As Long                 Handle to the window that will receive notification messages associated with an icon in the taskbar status area.
'    uId As Long                  Application-defined identifier of the taskbar icon.
'    ucallbackMessage As Long     Application-defined message identifier. The system uses this identifier for notification messages that it sends to the window identified in hWnd. These notifications are sent when a mouse event occurs in the bounding rectangle of the icon.
'    Param1(1) As Long            unknown interpretation
'    hIcon As Long                Handle to the icon
'    Param2(2) As Long            unknown interpretation
'    APath As String              The path of the appication which add the icon to system tray area
'    ToolTip As String            The tooltip associated with the icon
'    Bitmap As Long               index of the posted icon
'    Command As Long              command id
'    State As Byte                icon state
'    Style As Byte                icon style
'    Data As Long                 pointer towards the data of the tray
'    str As Long                  pointor to tooltip
'End Type

'i think that the only problem comes from you must know that the variable which will receive the
'returned data type must be an array of TrayIconInfo
'see this example
'
'Dim TrayList() As TrayIconInfo
'TrayList = TrayIcons1.GetSysTrayIcons
'
'as you see the variable TrayList is an array of TrayIconInfo type

'how to hide icon from system tray area?
'because every icon has an element in the returned array,all you need to do is to pass
'the element index to HideIcon method like this
'
'TrayList1.HideIcon 1
'which 1 is the index of the icon element(TrayIconInfo) in the array

'that's it
'if you have any fruther questions or need more sample code don't hesitate to contact us
'at support@cpringold.com

Dim TrayI As NOTIFYICONDATA



'==============================================================================
' Method:       initFlexGrid
'
' Description:  setup the flexgrid control to receive the data
'==============================================================================

Private Sub initFlexGrid()


  MSFlexGrid1.ColWidth(0) = 300
  MSFlexGrid1.TextMatrix(0, 0) = "I"
  MSFlexGrid1.ColWidth(1) = 3000
  MSFlexGrid1.TextMatrix(0, 1) = "Application Path"
  MSFlexGrid1.ColWidth(2) = 1300
  MSFlexGrid1.TextMatrix(0, 2) = "uID"
  MSFlexGrid1.ColWidth(3) = 1300
  MSFlexGrid1.TextMatrix(0, 3) = "hWnd"
  MSFlexGrid1.ColWidth(4) = 1300
  MSFlexGrid1.TextMatrix(0, 4) = "hIcon"
  MSFlexGrid1.ColWidth(5) = 1700
  MSFlexGrid1.TextMatrix(0, 5) = "ToolTip"
  MSFlexGrid1.ColWidth(6) = 1500
  MSFlexGrid1.TextMatrix(0, 6) = "uCallbackMessage"
  MSFlexGrid1.ForeColorFixed = vbBlue
  
   
End Sub

'==============================================================================
' Method:       FillFlexGrid
'
' Description:  the core function here we get the informations from GetSysTrayIcons
'               then fill the FlexGrid control with these info.
'==============================================================================

Public Sub FillFlexGrid()

Dim i As Integer
Dim TrayList() As TrayIconInfo

initFlexGrid

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CoreLine:
'Get the System Tray Data
10 TrayList = TrayList1.GetSysTrayIcons 'this is the core line you can consider
'that this is the sample code and all the remaining is nothing.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

MSFlexGrid1.Clear
'here we Fill the FlexGrid with Data
For i = 0 To TrayList1.IconCount - 1
    
    MSFlexGrid1.TextMatrix(i + 1, 0) = i + 1
    If TrayList(i).APath = "" Then
       MSFlexGrid1.TextMatrix(i + 1, 1) = "N/A"
    Else
       MSFlexGrid1.TextMatrix(i + 1, 1) = TrayList(i).APath
    End If
    MSFlexGrid1.TextMatrix(i + 1, 2) = TrayList(i).Uid
    MSFlexGrid1.TextMatrix(i + 1, 3) = TrayList(i).hWnd
    MSFlexGrid1.TextMatrix(i + 1, 4) = TrayList(i).hicon
    MSFlexGrid1.TextMatrix(i + 1, 5) = TrayList(i).ToolTip
    MSFlexGrid1.TextMatrix(i + 1, 6) = "&H" & Hex$(TrayList(i).uCallbackMessage)
   
    
Next i
 
If TrayList1.IconCount = 0 Then
    MSFlexGrid1.TextMatrix(1, 1) = "No Icon Found in systray"
End If
 
End Sub

Private Sub Command1_Click()
MsgBox TrayList1.IconCount
End Sub

Private Sub Command3_Click()
    SendMessage MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 3), MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 6), ByVal 0&, WM_LBUTTONDBLCLK
End Sub

'==============================================================================
' Method:       Form_Load
'
' Description:  call FillFlexGrid to display the data
'==============================================================================

Private Sub Form_Load()
 
 'Fill the FlexGrid control with System Tray Data
 FillFlexGrid

End Sub

'==============================================================================
' Method:       cmdHide_Click
'
' Description:  cool function used to hide any icon from system tray icon
'                then repopulate the FlexGrid contrrol.
'==============================================================================

Private Sub cmdHide_Click(Index As Integer)
 Dim OldUid As Long
 Dim OldHwnd As Long
 Dim OldHicon As Long
 Dim TrayList() As TrayIconInfo
 
 If MSFlexGrid1.Visible = True Then
    'Hide the selected Item from the System Tray area
    'TrayList = TrayList1.GetSysTrayIcons
    'OldHwnd = TrayList(MSFlexGrid1.RowSel - 1).hWnd
    'OldUid = TrayList(MSFlexGrid1.RowSel - 1).Uid
    'OldHicon = TrayList(MSFlexGrid1.RowSel - 1).hicon
    '
    TrayList1.HideIcon MSFlexGrid1.RowSel
    
   ' sdfasdf OldHwnd, OldUid, OldHicon
    
    MSFlexGrid1.Clear
    FillFlexGrid
 End If

End Sub

Private Sub sdfasdf(hWnd As Long, Uid As Long, hicon As Long)
  TrayI.cbSize = Len(TrayI)
    'Set the window's handle (this will be used to hook the specified window)
    TrayI.hWnd = hWnd
    'Application-defined identifier of the taskbar icon
    TrayI.Uid = Uid
    'Set the flags
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    'Set the callback message
    TrayI.uCallbackMessage = WM_LBUTTONDOWN
    'Set the picture (must be an icon!)
    TrayI.hicon = hicon
    
    'Create the icon
    Shell_NotifyIcon NIM_ADD, TrayI
    
End Sub

Private Sub Timer1_Timer()
FillFlexGrid
End Sub
