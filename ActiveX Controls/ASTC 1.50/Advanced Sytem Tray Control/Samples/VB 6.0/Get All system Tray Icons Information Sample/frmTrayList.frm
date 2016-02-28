VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{30E5838E-8E26-11DA-A1C2-807A50C10000}#1.0#0"; "astc.ocx"
Begin VB.Form Form1 
   Caption         =   "TrayList Sample Code"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   10650
   StartUpPosition =   2  'CenterScreen
   Begin ASTC.TrayList TrayList1 
      Left            =   10080
      Top             =   3720
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "R&efresh"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "&Restore Last"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   3720
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
      Top             =   3720
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   100
      Cols            =   7
      ForeColorSel    =   16711680
      GridColorFixed  =   255
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "Form1"
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
'Description:  A sample code demonstrate how to use ASTC control to get all system tray icons
'                Information as well as how to hide a specified icon from system tray area.
'==========================================================================================
Option Explicit

Dim OldTrayIcon As TrayIconInfo

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'The main line that returns the system Tray information is only one line which located
'in FillFlexGrid method [CoreLine] at line 10. All the remaining lines is to fill the
'FlexGrid with the returned data from GetSysTrayIcons method to display the data in a good manner.

'How can you return System Tray Icons information?
'GetSysTrayIcons method returns an array of TrayIconInfo type
'in other words
'We have a Type Called TrayIconInfo [you can use Object Browser to see this type in ASTC Control Library]
'When you call GetSysTrayIcons method, it returns an array of this type contains System Tray Icons information

'For Example:
'If the system tray area contains three icons say [sound control icon,'MacAfee icon and network connection icon],
'then the returned data will be organized in an Array and the array will have three elements
'And every element type will be TrayIconInfo and TrayIconInfo will contains this icon information.
'
'What is these information?
'
'Public Type TrayIconInfo
'    Hwnd As Long                 Handle to the window that will receive notification messages associated with an icon in the taskbar status area.
'    uId As Long                  Application-defined identifier of the taskbar icon.
'    ucallbackMessage As Long     Application-defined message identifier. The system uses this identifier for notification messages that it sends to the window identified in hWnd. These notifications are sent when a mouse event occurs in the bounding rectangle of the icon.
'    Param1(1) As Long            unknown interpretation
'    hIcon As Long                Handle to the icon
'    Param2(2) As Long            unknown interpretation
'    APath As String              The path of the application which adds the icon to system tray area
'    ToolTip As String            The tooltip associated with the icon
'    Bitmap as Long               index of the posted icon
'    Command as Long              command id
'    State as Byte                icon state
'    Style as Byte                icon style
'    Data as Long                 pointer towards the data of the tray
'    str As Long                  pointer to tooltip
'End Type

'Note that: You must know that the variable which will receive the returned data type must be an array of TrayIconInfo
'See this example
'
'Dim TrayList() As TrayIconInfo'//array of TrayIconInfo type
'TrayList = TrayIcons1.GetSysTrayIcons'//GetSysTrayIcons function returns System Tray Icons information to TrayList array
'
'The variable TrayList is an array of TrayIconInfo type

'How can i hide an icon from system tray area?
'Because every icon has an element in the returned array, all you need to do is to pass
'The element index to HideIcon method like this
'
'TrayList1.HideIcon 1
'Which 1 is the index of the icon element (TrayIconInfo) in the array?

'Don't hesitate to contact us At support@cpringold.com if you have any further questions or need more sample code
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
' Description:  [The core function] here we get the information from GetSysTrayIcons
'               then fill the FlexGrid control with these information.
'==============================================================================
Public Sub FillFlexGrid()

    Dim i As Integer
    Dim TrayList() As TrayIconInfo
    
    initFlexGrid
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CoreLine:
    'Get the System Tray Data
10     TrayList = TrayList1.GetSysTrayIcons 'this is the core line you can consider
    'that this is the sample code and all the remaining is nothing.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    'here we Fill the FlexGrid with Data
    For i = 0 To TrayList1.IconCount - 1
        
        MSFlexGrid1.TextMatrix(i + 1, 0) = i + 1
        If TrayList(i).APath = "" Then
           MSFlexGrid1.TextMatrix(i + 1, 1) = "N/A"
        Else
           MSFlexGrid1.TextMatrix(i + 1, 1) = TrayList(i).APath
        End If
        MSFlexGrid1.TextMatrix(i + 1, 2) = TrayList(i).uId
        MSFlexGrid1.TextMatrix(i + 1, 3) = TrayList(i).hWnd
        MSFlexGrid1.TextMatrix(i + 1, 4) = TrayList(i).hIcon
        MSFlexGrid1.TextMatrix(i + 1, 5) = TrayList(i).ToolTip
        MSFlexGrid1.TextMatrix(i + 1, 6) = "&H" & Hex$(TrayList(i).ucallbackMessage)
       
        
    Next i
     
    If TrayList1.IconCount = 0 Then
        MSFlexGrid1.TextMatrix(1, 1) = "No Icon Found in systray"
    End If
 
End Sub

'==============================================================================
' Method:       cmdRefresh_Click
'
' Description:  Refresh FillFlexGrid to display any new data
'==============================================================================
Private Sub cmdRefresh_Click()
 'Fill the FlexGrid control with System Tray Data
 FillFlexGrid
End Sub

'==============================================================================
' Method:       Form_Load
'
' Description:  Call FillFlexGrid to display the data
'==============================================================================
Private Sub Form_Load()
 
 'Fill the FlexGrid control with System Tray Data
 FillFlexGrid

End Sub

'==============================================================================
' Method:       cmdHide_Click
'
' Description:  Used to hide any icon from system tray icon
'               then refresh the FlexGrid control.
'==============================================================================
Private Sub cmdHide_Click(Index As Integer)
 Dim TrayList() As TrayIconInfo
 
    TrayList = TrayList1.GetSysTrayIcons
    'Save the Icon Info before Remove it in order to use
    'when we restore the icon
    OldTrayIcon = TrayList(MSFlexGrid1.RowSel - 1)
    
    TrayList1.HideIcon MSFlexGrid1.RowSel
    
    MSFlexGrid1.Clear
    
    FillFlexGrid


End Sub

'==============================================================================
' Method:       RestoreIcon
'
' Description:  Restore any icon you removed with Hide Method.
'==============================================================================
Private Sub RestoreIcon(TrayIconInfo As TrayIconInfo)

    Const NIF_TIP = &H4
    Const NIF_ICON = &H2
    Const NIM_ADD = &H0
    Const NIF_MESSAGE = &H1
    
    Dim TrayI As NOTIFYICONDATA
    
    On Error Resume Next
    
         If TrayIconInfo.hWnd = 0 Then Exit Sub
         
         TrayI.cbSize = Len(TrayI)
         TrayI.hWnd = TrayIconInfo.hWnd
         TrayI.uId = TrayIconInfo.uId
         TrayI.hIcon = TrayIconInfo.hIcon
         TrayI.szTip = TrayIconInfo.ToolTip
         TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         TrayI.ucallbackMessage = TrayIconInfo.ucallbackMessage
         
         'Restore the icon
          Shell_NotifyIcon NIM_ADD, TrayI
        
          MSFlexGrid1.Clear
         
          FillFlexGrid
     
End Sub

Private Sub cmdRestore_Click()
    RestoreIcon OldTrayIcon
End Sub
