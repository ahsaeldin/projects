VERSION 5.00
Object = "{30E5838E-8E26-11DA-A1C2-807A50C10000}#1.0#0"; "acomte445.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   7545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHelp 
      Caption         =   "He&lp"
      Height          =   495
      Left            =   6000
      TabIndex        =   19
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Timer tmrDisableSysListHorzBar 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7680
      Top             =   1200
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Options"
      Height          =   495
      Left            =   4560
      TabIndex        =   12
      Top             =   8280
      Width           =   1335
   End
   Begin VB.CommandButton cmdSysHide 
      Caption         =   "Hi&de"
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton cmdSysShowAll 
      Caption         =   "Show &All"
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton cmdSysHideAll 
      Caption         =   "Hid&e All"
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton cmdSysRevert 
      Caption         =   "Re&vert"
      Height          =   495
      Left            =   6000
      TabIndex        =   11
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton cmdSysShow 
      Caption         =   "Sho&w"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton CmdRevert 
      Caption         =   "&Revert"
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdShowAll 
      Caption         =   "Sh&ow All"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "&Hide"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton CmdHideAll 
      Caption         =   "H&ide all"
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdshowwindow 
      Caption         =   "&Show"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   1335
   End
   Begin VB.PictureBox PicSysIco 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7680
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox PIco 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7680
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer tmrNoTopMost 
      Interval        =   1
      Left            =   7680
      Top             =   720
   End
   Begin VB.Timer tmrUpdateListView 
      Interval        =   4000
      Left            =   7680
      Top             =   240
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   8895
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   7335
      Begin MSComctlLib.ListView lstSystrayIcons 
         Height          =   2895
         Left            =   120
         TabIndex        =   6
         Top             =   4560
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5106
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Systray Icon"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView Lstwindows 
         Height          =   2895
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Window Caption"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   2469
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Taskbar Windows"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "System Tray Icons"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   -240
         X2              =   7320
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         X1              =   0
         X2              =   7320
         Y1              =   8160
         Y2              =   8160
      End
   End
   Begin ASTC.TrayIcon TrayIcon 
      Left            =   7680
      Top             =   2880
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin MSComctlLib.ImageList ILWindowsIcons 
      Left            =   7680
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin ASTC.TrayList TrayList 
      Left            =   7680
      Top             =   3480
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin MSComctlLib.ImageList ILSysIcons 
      Left            =   7680
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstHiddenItems 
      Height          =   2895
      Left            =   360
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Systray Icon"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IconHwnd"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IconuId"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Hicon"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Tooltip"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ucallbackmessage"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuSystray 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Software Name"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuWebSite 
         Caption         =   "Software WebSite"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu menuHelp2 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuHelp3 
         Caption         =   "Help"
      End
      Begin VB.Menu menuWebSite2 
         Caption         =   "Web Site"
      End
      Begin VB.Menu menuAbout2 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

    ''''''''''''''''''''''''''''''''''''''´´´´¶¶¶¶¶¶´´´´´´¶¶¶¶¶¶
    '''''''''''''''''''''''''''''''''''''´´¶¶¶¶¶¶¶¶¶¶´´¶¶¶¶¶¶¶¶¶¶
    '''''''''''''''''''''''''''''''''''''´¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶´´´´¶¶¶¶
    '''''''''''''''''''''''''''''''''''''¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶´´´´¶¶¶¶
    '''''''''''''''''''''''''''''''''''''¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶´´¶¶¶¶¶
    '''''''''''''''''''''''''''''''''''''¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶ ´¶¶¶¶¶
    '''''''''''''''''''''''''''''''''''''´¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶
    '''''''''''''''''''''''''''''''''''''´´´¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶
    '''''''''''''''''''''''''''''''''''''´´´´´¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶
    '''''''''''''''''''''''''''''''''''''´´´´´´´¶¶¶¶¶¶¶¶¶¶¶¶¶
    '''''''''''''''''''''''''''''''''''''´´´´´´´´´¶¶¶¶¶¶¶¶

    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *'
    '“Love always brings difficulties, that is true, but the good side of it is that it gives energy.”'
    '                                                                                  Vincent van Gog'
    '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *'

'//for unknown reason NOTIFYICONDATA and Shell_NotifyIcon inclued in win32.tlb isn't like the one provided by API-Viewer, this problem apperes in
'//szTip As String * 64 and i can't copy the tooltip to szTip hence i declared it here
Private Type NOTIFYICONDATA
     cbSize As Long
     hwnd As Long
     uId As Long
     uFlags As Long
     ucallbackMessage As Long
     hIcon As Long
     szTip As String * 64
End Type

'//for unknown reason NOTIFYICONDATA and Shell_NotifyIcon inclued in win32.tlb isn't like the one provided by API-Viewer, this problem apperes in
'//szTip As String * 64 and i can't copy the tooltip to szTip hence i declared it here
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Load()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    CreateVirutalMem

    Dim FadeSettings As Integer '//Buffer that will store a value from the ini that tell us if this form Faded before or not
    
    InnerSysError = True   '//Enable Inner SysError message boxs
    
    OnErrorSysError = True   '//Enable OnError SysError message boxs
    
    IsLogError = True '//Enable Error Log
    
    Me.Caption = SoftwareName '//assign the title of this form to be equal to this app title
    
    FadeSettings = GetINIProfileInt("Fade", 0)  '//Check if this form Faded before or not
    
    If FadeSettings = 0 Then FadeForm Me '//Fade this S if it doesn't faded before

    Set CVisWindows = New CWindows

    HiddenWindowForeColor = RGB(192, 192, 192) '//pt 8June2007: SET the color of the hidden windows in the listview

    If App.PrevInstance Then End

    '//Check if Hot Track is enabled to enable or disable IsHotTrack
    Dim res As Integer
    res = GetINIProfileInt("chkHotTrack", 1)
    If res Then IsHotTrack = True
    
    '//Check if always display a TrayIcon
    res = GetINIProfileInt("optalways", 1)
    If res Then IsAlwaysTrayIcon = True
    
    res = GetINIProfileInt("optnever", 0)
    If res Then IsNeverTrayIcon = True
   
    res = GetINIProfileInt("optmini", 0)
    If res Then IsTrayIconOnlyWhenMini = True
   
    '//commented next line because it cozs the form to be appear first at poz 0,0
'   res = GetINIProfileInt("chkShowTopMost", 0)
'
'   If res Then
'       IsShowTopMost = True
'       'SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'   End If
    
    TrayIcon.TrackPopMenu = True
     
    mnuOpen.Caption = "Open " & SoftwareName
      
    mnuWebSite.Caption = SoftwareName & " WebSite"
                      
    '//added next if block from tmrUpdateListView_Timer to display my trayIcon before calling CheckSysTrayList below
    '//in order to appear my icon in the list
    If IsAlwaysTrayIcon Then
       If Not TrayIcon.IsDisplayed Then
          'If IsHiddenWindow <> 7878 Then '//Only show me if (my state isn't hide in the listview and this option "Hide my system tray icon when hiding me" is unchecked
              TrayIcon.Show Icon, SoftwareName
          'End If
       End If
    ElseIf IsNeverTrayIcon Then
        If TrayIcon.IsDisplayed Then TrayIcon.Hide
    ElseIf IsTrayIconOnlyWhenMini Then
        If TrayIcon.IsDisplayed Then
           If IsWindowVisible(hwnd) <> 0 Then TrayIcon.Hide
        End If
    End If
     
    '//For unknown reason, this line in CheckSysTrayList function
    '//i = frmmain.TrayList.IconCount
    '//causes a delay in the program startup, you can notice this if you click
    '//frmopions just after you run the exe
    '//hence, i will make tmrUpdateListView.Interval = 4000 in the property pages
    '//then and i will update the listviews in form_load
    '//also inside tmrUpdateListView_Timer, i will return the timer interval again to 1, just after the startup period get passed
    CheckSysTrayList
            
    tmrDisableSysListHorzBar.Enabled = True

End Sub

'//PT 4June2007: added next function to process the click event for cmdshowwindow
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub cmdshowwindow_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: cmdshowwindow_Click
' Description:  added next function to process the click event for cmdshowwindow
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

'/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/
'“The computer programmer is a creator of universes for which he alone is responsible. Universes of virtually unlimited complexity can be created in the form of computer programs.” '
'                                                                                                                                                                                    '
'                                                                                                                                                            Joseph Weizenbaum quotes'
'/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/*\*/
 
    On Error GoTo cmdshowwindow_ClickErr
        
    Dim IsHidden As Long '//Check the window visiblity
    Dim WindowHwnd As Long '//Get window hwnd stored in Item.Tag
    Dim ListViewItemsIndex As Integer '//Used To loop through ListView
    
    '//if Show as TopMost then no need to proceed
    '//pt 21July2007: changed next line and added "And IsWindowVisible(hwnd)"
    '//If Not IsShowTopMost Then
    If Not IsShowTopMost And IsWindowVisible(hwnd) Then
        '//Set Me Top most while the click then set Me "No top most" using tmrNoTopMost_Timer
        '//Okay but why=>>well the next line to make Me TopMost while the user uses Me
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        '//and No top most in tmrNoTopMost_Timer to delete Top Most state to avoid nagging the user
    End If
    
    For ListViewItemsIndex = 1 To Lstwindows.ListItems.count '//Loop through the listview
        '//Show Selected windows only
        If Lstwindows.ListItems.Item(ListViewItemsIndex).Checked Then
            'Debug.Print Lstwindows.ListItems.Item(ListViewItemsIndex).Text
            WindowHwnd = Lstwindows.ListItems.Item(ListViewItemsIndex).Tag '//Get Window's Hwnd stored in the Tag
            If WindowHwnd = 0 Then
               LogError "Can't get window's Hwnd."
                Err.Description = "Can't get window's Hwnd."
                If InnerSysError Then SysError Err, "0x0f3ee110"
            End If
            IsHidden = GetProp(WindowHwnd, "VisibiltyState")  '//if the window is previously hidden
            If IsHidden = 7878 Then '//then
                RemoveProp WindowHwnd, "VisibiltyState" '//Show it
                Lstwindows.ListItems.Item(ListViewItemsIndex).SubItems(1) = "Show"
                Lstwindows.ListItems.Item(ListViewItemsIndex).ForeColor = RGB(0, 0, 0)
                ShowWindow WindowHwnd, SW_SHOW
            End If
        End If
    Next ListViewItemsIndex

    
cmdshowwindow_ClickExit:
    On Error Resume Next

    Exit Sub
    
cmdshowwindow_ClickErr:

    '//0x0f3ee110 number is a unique number for MyTemplate because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee110"
    
GoTo cmdshowwindow_ClickExit

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub cmdShowAll_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: cmdShowAll_Click
' Description: cmdShowAll_Click Description
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo cmdShowAll_ClickErr
        
    Dim IsHidden As Long '//Check the window visiblity
    Dim WindowHwnd As Long '//Get window hwnd stored in Item.Tag
    Dim ListViewItemsIndex As Integer '//Used To loop through ListView
    
    '//if Show as TopMost then no need to proceed
    '//pt 21July2007: changed next line and added "And IsWindowVisible(hwnd)"
    '//If Not IsShowTopMost Then
    If Not IsShowTopMost And IsWindowVisible(hwnd) Then
        '//Set Me Top most while the click then set Me "No top most" using tmrNoTopMost_Timer
        '//Okay but why=>>well the next line to make Me TopMost while the user uses Me
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        '//and No top most in tmrNoTopMost_Timer to delete Top Most state to avoid nagging the user
    End If
    
    For ListViewItemsIndex = 1 To Lstwindows.ListItems.count '//Loop through the listview
            'Debug.Print Lstwindows.ListItems.Item(ListViewItemsIndex).Text
            WindowHwnd = Lstwindows.ListItems.Item(ListViewItemsIndex).Tag '//Get Window's Hwnd stored in the Tag
            If WindowHwnd = 0 Then
               LogError "Can't get window's Hwnd."
                Err.Description = "Can't get window's Hwnd."
                If InnerSysError Then SysError Err, "0x0f3ee110"
            End If
            IsHidden = GetProp(WindowHwnd, "VisibiltyState") '//if the window is previously hidden
            If IsHidden = 7878 Then '//then
                RemoveProp WindowHwnd, "VisibiltyState" '//Show it
                Lstwindows.ListItems.Item(ListViewItemsIndex).SubItems(1) = "Show"
                Lstwindows.ListItems.Item(ListViewItemsIndex).ForeColor = RGB(0, 0, 0)
                ShowWindow WindowHwnd, SW_SHOW
            End If
    Next ListViewItemsIndex
    
    Unload frmtip
    
cmdShowAll_ClickExit:
    On Error Resume Next

    Exit Sub
    
cmdShowAll_ClickErr:

    '//0x0f3ee113 number is a unique number for MycmdShowAll_Click because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee113"
    
GoTo cmdShowAll_ClickExit

End Sub

'//PT 8June2007:
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub cmdHide_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: cmdHide_Click
' Description: cmdHide_Click Description
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo cmdHide_ClickErr
        
    Dim IsHidden As Long '//Check the window visiblity
    Dim WindowHwnd As Long '//Get window hwnd stored in Item.Tag
    Dim ListViewItemsIndex As Integer '//Used To loop through ListView
    
    '//if Show as TopMost then no need to proceed
    '//pt 21July2007: changed next line and added "And IsWindowVisible(hwnd)"
    '//Why added And IsWindowVisible(hwnd) ??
    '//Well,uncomment the next line and do the following:
    '//1.Send me to systray
    '//Hide the foregroundwindow using the shortcut key
    '//notice that me will showed up again and my state in the list is Hide
    '//that's why i added
    '//If Not IsShowTopMost "And IsWindowVisible(hwnd)" Then
    If Not IsShowTopMost And IsWindowVisible(hwnd) Then
        '//Set Me Top most while the click then set Me "No top most" using tmrNoTopMost_Timer
        '//Okay but why=>>well the next line to make Me TopMost while the user uses Me
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        '//and No top most in tmrNoTopMost_Timer to delete Top Most state to avoid nagging the user
    End If
    
    For ListViewItemsIndex = 1 To Lstwindows.ListItems.count '//Loop through the listview
        '//Hide Selected windows only
        If Lstwindows.ListItems.Item(ListViewItemsIndex).Checked Then
            'Debug.Print Lstwindows.ListItems.Item(ListViewItemsIndex).Text
            WindowHwnd = Lstwindows.ListItems.Item(ListViewItemsIndex).Tag '//Get Window's Hwnd stored in the Tag
            If WindowHwnd = 0 Then
               LogError "Can't get window's Hwnd."
                Err.Description = "Can't get window's Hwnd."
                If InnerSysError Then SysError Err, "0x0f3ee112"
            End If
            IsHidden = GetProp(WindowHwnd, "VisibiltyState") '//if the window is previously hidden
            If IsHidden = 0 Or IsHidden = Null Then '//then
                SetProp WindowHwnd, "VisibiltyState", 7878
                Lstwindows.ListItems.Item(ListViewItemsIndex).SubItems(1) = "Hide"
                Dim res As Integer
                res = GetINIProfileInt("chkGray", 1)
                If res Then
                   Lstwindows.ListItems.Item(ListViewItemsIndex).ForeColor = HiddenWindowForeColor
                End If
                
                '//Coz google talk doesn't want to be hidden
                If InStr(1, Lstwindows.ListItems.Item(ListViewItemsIndex).Text, "Google Talk") <> 0 Then
                   ShowWindow WindowHwnd, SW_MINIMIZE
                End If
                
                ShowWindow WindowHwnd, SW_HIDE
                If WindowHwnd = Me.hwnd Then
                   Dim TipSettings As Integer
                   TipSettings = GetINIProfileInt("Tip1", 0)
                   If TipSettings = 0 Then frmtip.Show
                End If
            End If
        End If
    Next ListViewItemsIndex
        
cmdHide_ClickExit:
    On Error Resume Next

    Exit Sub
    
cmdHide_ClickErr:

    '//0x0f3ee112 number is a unique number for MycmdHide_Click because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee112"
    
GoTo cmdHide_ClickExit

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub CmdHideAll_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: cmdShowAll_Click
' Description: cmdShowAll_Click Description
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo cmdShowAll_ClickErr
        
    Dim IsHidden As Long '//Check the window visiblity
    Dim WindowHwnd As Long '//Get window hwnd stored in Item.Tag
    Dim ListViewItemsIndex As Integer '//Used To loop through ListView
    
    '//if Show as TopMost then no need to proceed
    If Not IsShowTopMost Then
        '//Set Me Top most while the click then set Me "No top most" using tmrNoTopMost_Timer
        '//Okay but why=>>well the next line to make Me TopMost while the user uses Me
        '//we add the "If IsWindowVisible(hwnd) Then" to avoid displaying this window if the user hide it using the shortcut key or even with the normal way
        If IsWindowVisible(hwnd) Then
           SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        End If
        '//and No top most in tmrNoTopMost_Timer to delete Top Most state to avoid nagging the user
    End If
     
    '//Why add "If IsWindowVisible(hwnd) Then" ??
    '//If we don't do this then multipal calls to a hiding key will shows frmtip
    '//for example if "Hide All Window" key is Q then every press to Q will shows up frmtip
    If IsWindowVisible(hwnd) Then
        Dim TipSettings As Integer
        TipSettings = GetINIProfileInt("Tip1", 0)
        If TipSettings = 0 Then frmtip.Show
    End If
 
    For ListViewItemsIndex = 1 To Lstwindows.ListItems.count '//Loop through the listview
            'Debug.Print Lstwindows.ListItems.Item(ListViewItemsIndex).Text
            WindowHwnd = Lstwindows.ListItems.Item(ListViewItemsIndex).Tag '//Get Window's Hwnd stored in the Tag
            If WindowHwnd = 0 Then
               LogError "Can't get window's Hwnd."
                Err.Description = "Can't get window's Hwnd."
                If InnerSysError Then SysError Err, "0x0f3ee112"
            End If
            IsHidden = GetProp(WindowHwnd, "VisibiltyState") '//if the window is previously hidden
            If IsHidden = 0 Or IsHidden = Null Then '//then
                SetProp WindowHwnd, "VisibiltyState", 7878
                Lstwindows.ListItems.Item(ListViewItemsIndex).SubItems(1) = "Hide"
                Dim res As Integer
                res = GetINIProfileInt("chkGray", 1)
                If res Then
                   Lstwindows.ListItems.Item(ListViewItemsIndex).ForeColor = HiddenWindowForeColor
                End If
                
               '//Coz google talk doesn't want to be hidden
                If InStr(1, Lstwindows.ListItems.Item(ListViewItemsIndex).Text, "Google Talk") <> 0 Then
                    ShowWindow WindowHwnd, SW_MINIMIZE
                End If
                
                ShowWindow WindowHwnd, SW_HIDE
            End If
    Next ListViewItemsIndex
            
cmdShowAll_ClickExit:
    On Error Resume Next

    Exit Sub
    
cmdShowAll_ClickErr:

    '//cmdShowAll_Clickunique number is a unique number for MycmdShowAll_Click because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "cmdShowAll_Clickunique number"
    
GoTo cmdShowAll_ClickExit

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub CmdRevert_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: CmdRevert_Click
' Description: CmdRevert_Click Description
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo CmdRevert_ClickErr
        
    Dim IsHidden As Long '//Check the window visiblity
    Dim WindowHwnd As Long '//Get window hwnd stored in Item.Tag
    Dim ListViewItemsIndex As Integer '//Used To loop through ListView
    
    
    '//if Show as TopMost then no need to proceed
    '//pt 21July2007: changed next line and added "And IsWindowVisible(hwnd)"
    '//If Not IsShowTopMost Then
    If Not IsShowTopMost And IsWindowVisible(hwnd) Then
        '//Set Me Top most while the click then set Me "No top most" using tmrNoTopMost_Timer
        '//Okay but why=>>well the next line to make Me TopMost while the user uses Me
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        '//and No top most in tmrNoTopMost_Timer to delete Top Most state to avoid nagging the user
    End If
    
    '//Why add "If IsWindowVisible(hwnd) Then" ??
    '//If we don't do this then multipal calls to a hiding key will shows frmtip
    '//for example if "Hide All Window" key is Q then every press to Q will shows up frmtip
    If IsWindowVisible(hwnd) Then
        Dim TipSettings As Integer
        TipSettings = GetINIProfileInt("Tip1", 0)
        If TipSettings = 0 Then frmtip.Show
    End If
    
    For ListViewItemsIndex = 1 To Lstwindows.ListItems.count '//Loop through the listview
            'Debug.Print Lstwindows.ListItems.Item(ListViewItemsIndex).Text
            WindowHwnd = Lstwindows.ListItems.Item(ListViewItemsIndex).Tag '//Get Window's Hwnd stored in the Tag
            If WindowHwnd = 0 Then
               LogError "Can't get window's Hwnd."
                Err.Description = "Can't get window's Hwnd."
                If InnerSysError Then SysError Err, "0x0f3ee110"
            End If
            IsHidden = GetProp(WindowHwnd, "VisibiltyState") '//if the window is previously hidden
            If IsHidden = 7878 Then '//then
                RemoveProp WindowHwnd, "VisibiltyState" '//Show it
                Lstwindows.ListItems.Item(ListViewItemsIndex).SubItems(1) = "Show"
                Lstwindows.ListItems.Item(ListViewItemsIndex).ForeColor = RGB(0, 0, 0)
                ShowWindow WindowHwnd, SW_SHOW
            ElseIf IsHidden = 0 Or IsHidden = Null Then
                SetProp WindowHwnd, "VisibiltyState", 7878
                Lstwindows.ListItems.Item(ListViewItemsIndex).SubItems(1) = "Hide"
                Dim res As Integer
                res = GetINIProfileInt("chkGray", 1)
                If res Then
                   Lstwindows.ListItems.Item(ListViewItemsIndex).ForeColor = HiddenWindowForeColor
                End If
                ShowWindow WindowHwnd, SW_HIDE
            End If
    Next ListViewItemsIndex
    
CmdRevert_ClickExit:
    On Error Resume Next

    Exit Sub
    
CmdRevert_ClickErr:

    '//CmdRevert_Clickunique number is a unique number for MyCmdRevert_Click because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee115"
    
GoTo CmdRevert_ClickExit

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub lstSystrayIcons_ItemClick(ByVal Item As MSComctlLib.ListItem)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     
    Exit Sub 'Until we found a solution for the order problem
    
    Dim FFile As Integer
    Dim TotalRecordsNumber As Long
    Dim ArrSelectedIndex As Integer
    Dim SysListItemIndex As Integer
    Dim SysToolTipTypeRecordLen As Long
    Dim SysToolTipType As SysIcoToolTip
    Dim TrayListArrType() As TrayIconInfo
    Dim SeletedItemInArrType As TrayIconInfo
    
    TrayListArrType = TrayList.GetSysTrayIcons

    If Item.SubItems(1) = "Show" Then
       For ArrSelectedIndex = 0 To UBound(TrayListArrType)
           
           If TrayListArrType(ArrSelectedIndex).hIcon = Item.Tag Then
                '//why we added the next line??
                '//because of two reasons:
                '//1.we loop through TrayListArrType(ArrSelectedIndex) array and this array contains a non tooltip
                '//icon and we don't want to process non tooltip
                '//2. to fix the problem of connection status  "Local Area Connection  Speed: 100.0 Mbps   Status: Connected"
                '//because this icon has a non tooltip icons "about 3 icons" and this coz this sub to be processed more than one time
                '//and overwrite the info of the icon and cozs to not display the tootip and the right click menu after re show it
                If TrayListArrType(ArrSelectedIndex).ToolTip <> "" Then
                   Item.SubItems(1) = "Hide"
                   Item.SubItems(2) = TrayListArrType(ArrSelectedIndex).hwnd
                   Item.SubItems(3) = TrayListArrType(ArrSelectedIndex).uId
                   Item.SubItems(4) = TrayListArrType(ArrSelectedIndex).hIcon
                   Item.SubItems(5) = TrayListArrType(ArrSelectedIndex).ToolTip
                   Item.SubItems(6) = TrayListArrType(ArrSelectedIndex).ucallbackMessage
                   '//Save these as probs in the window in case me closed before showing the icon again
                   SetProp TrayListArrType(ArrSelectedIndex).hwnd, "SysHwnd", TrayListArrType(ArrSelectedIndex).hwnd
                   SetProp TrayListArrType(ArrSelectedIndex).hwnd, "SysuId", TrayListArrType(ArrSelectedIndex).uId
                   SetProp TrayListArrType(ArrSelectedIndex).hwnd, "SyshIcon", TrayListArrType(ArrSelectedIndex).hIcon
                   SetProp TrayListArrType(ArrSelectedIndex).hwnd, "SysucallbackMessage", TrayListArrType(ArrSelectedIndex).ucallbackMessage
                               
                   '//Save Icon tooltip to a file
                   SysToolTipType.hwnd = TrayListArrType(ArrSelectedIndex).hwnd
                   '//we 'll use "#&" as a delimiter to get the end of tooltip text in order to
                   '//avoid displaying the tooltip like this after re-show it "snerfe ..."
                   SysToolTipType.ToolTip = TrayListArrType(ArrSelectedIndex).ToolTip & "#&"
                   FFile = FreeFile()
                   SysToolTipTypeRecordLen = Len(SysToolTipType)
                              
                   Open AppPath + "ico.dat" For Random As FFile Len = SysToolTipTypeRecordLen
                        TotalRecordsNumber = FileLen(AppPath + "ico.dat") / SysToolTipTypeRecordLen
                        Put #FFile, TotalRecordsNumber + 1, SysToolTipType
                   Close FFile
                                 
                   ' Dim IconHandle As Long
                   '
                   ' frmmain.PicSysIco.Cls
                   '
                   ' IconHandle = CopyIcon(TrayListArrType(ArrSelectedIndex).hIcon)
                   '
                   ' DrawIcon frmmain.PicSysIco.hDC, 0, 0, TrayListArrType(ArrSelectedIndex).hIcon
                   '
                   ' frmmain.ILLostIcons.ListImages.Add 1, , frmmain.PicSysIco.Image
                                 
                   '//Hide Icon//////////////////////////////////////////
                   Dim TrayI As NOTIFYICONDATA
                   Const NIF_TIP = &H4
                   Const NIF_ICON = &H2
                   Const NIF_MESSAGE = &H1
                   TrayI.cbSize = Len(TrayI)
                   TrayI.hwnd = TrayListArrType(ArrSelectedIndex).hwnd
                   TrayI.uId = TrayListArrType(ArrSelectedIndex).uId
                   TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
                   TrayI.hIcon = frmmain.Icon '  TrayListArrType(ArrSelectedIndex).hIcon
                   Shell_NotifyIcon NIM_DELETE, TrayI
                   'Dim i As Long
                   'Dim hTray As Long
                   'hTray = GetHwndTray
                   'i = SendMessage(hTray, TB_HIDEBUTTON, TrayListArrType(ArrSelectedIndex).BITMAP + 1, ByVal 1&)
                   '//////////////////////////////////////////////////////
           End If
          End If
        Next ArrSelectedIndex
    ElseIf Item.SubItems(1) = "Hide" Then
           If Item.SubItems(1) = "Hide" Then
              ShowSysTrayIcon Item.Index
              '//Remove the probs from the window
              RemoveProp Item.SubItems(2), "SysuId"
              RemoveProp Item.SubItems(2), "SysHwnd"
              RemoveProp Item.SubItems(2), "SyshIcon"
              RemoveProp Item.SubItems(2), "SysucallbackMessage"
           End If
           '//Rewrite the ico.dat again to remove the icon tooltip and hwnd that already show now
           ReWriteIcoDatFile
    End If
 
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Lstwindows_KeyDown(KeyCode As Integer, Shift As Integer)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    '//hide the hidden columns in the sys list view
    If KeyCode = vbKeyRight Then KeyCode = 0

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub lstSystrayIcons_KeyDown(KeyCode As Integer, Shift As Integer)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    '//hide the hidden columns in the sys list view
    If KeyCode = vbKeyRight Then KeyCode = 0

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Lstwindows_ItemClick(ByVal Item As MSComctlLib.ListItem)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: Lstwindows_ItemClick
' Description: Lstwindows_ItemClick Description
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    
   ''''''''''''\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/'
   ''''''''''''“Girls bored me, they still do. I love Mickey Mouse more than any woman I've ever known.” '
   ''''''''''''                                                                               Walt Disney'
   ''''''''''''/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\'

    On Error GoTo Lstwindows_ItemClickErr
    Debug.Print "2"
    Dim ItemHwnd As Long
    Dim IsHidden As Long
    Dim TaskListIndex As Integer
    
    Dim chkReverseStateRes As Integer
    chkReverseStateRes = GetINIProfileInt("chkReverseState", 1)
    If chkReverseStateRes <> 1 Then Exit Sub
        
    '//if Show as TopMost then no need to proceed
    If Not IsShowTopMost Then
        '//Set Me Top most while the click then set Me "No top most" using tmrNoTopMost_Timer
        '//Okay but why=>>well the next line to make Me TopMost while the user uses Me
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        '//and No top most in tmrNoTopMost_Timer to delete Top Most state to avoid nagging the user
    End If
    
    If GetAsyncKeyState(VK_RBUTTON) Then Exit Sub '//exit function if right mouse button
    
    '//pt 8June2007: Find the window that has the title in which the user clicks in the listview
    'ItemHwnd = FindWindow(vbNullString, Item.Text)
    'here i could get the window hwnd by using the index of the selected item and use it as an index for
    'the tasklist array however may be this cause unexpected results or get an invalid window hwnd or another
    'window hwnd ==>>> """"""HENCE i will use Findwindow and Item.Text""""
    
    '//HOWEVER for any reason, may be FindWindow returns zero
    '//SO i put an a7tamal THAT "Item.Text" will larger or smaller
    '//in actual window title, hence i will make a check if ItemHwnd == 0
    If ItemHwnd = 0 Then
            '//Chr(34) is the Quotaion "
            'LogError "Lstwindows_ItemClick::FindWindow returns Zero, can't found " & Chr(34) & Item.Text & Chr(34) & " Window"
            '//note that: i added the hwnd of the window to the item tag in loadlistview function
            '//in the next line, i 'll check if the hwnd stored in the Tag is a handle to a valid window
            If IsWindow(Item.Tag) Then '//if already there is a window with this handle
                 
                    For TaskListIndex = 0 To 200 'then search TaskList array for this Hwnd
                        
                        '//if i found this hwnd in TaskList array
                        If TaskList(TaskListIndex).TaskID = Item.Tag Then
                           '//then i 'll check if the TaskName "Window title" assosisted with this Hwnd in
                           '//the TaskList array is the same as "Item.Text"
                           '//so first i 'll check if "Len(TaskList(TaskListIndex).TaskName)" larger than
                           '//"Len(Item.Text)" by one char "null terminating"
                           If Len(TaskList(TaskListIndex).TaskName) - Len(Item.Text) = 1 Then
                              '//then i 'll check if both are same title
                              If InStr(1, TaskList(TaskListIndex).TaskName, Item.Text) Then
                                  'okay now i make sure that the title stored in TaskList is the same title in "Item.Text"
                                   ItemHwnd = Item.Tag '//then we 'll use this handle
                                   'LogError "Found window's Hwnd in the Item tag"
                              End If
                           End If
                        End If
                        
                    Next TaskListIndex
                    
             End If
    End If
    
    '//if ItemHwnd is still 0 then try for FindWindow by the window's title in the listview
    If ItemHwnd = 0 Then ItemHwnd = FindWindow(vbNullString, Item.Text)
    
    '//if ItemHwnd is still 0, so i'll arise an error
    If ItemHwnd = 0 Then
       Err.Description = "Can't hide " & Chr(34) & Item.Text & Chr(34) & " window."
       If InnerSysError Then SysError Err, "0x0f3ee111"
       LogError "Lstwindows_ItemClick::Can't hide " & Chr(34) & Item.Text & Chr(34) & " window."
    End If
    
    'the advangtage of using getprob by getting priveows hidden windows
    IsHidden = GetProp(ItemHwnd, "VisibiltyState")
    If IsHidden = 7878 Then
        RemoveProp ItemHwnd, "VisibiltyState"
        Item.SubItems(1) = "Show"
        Item.ForeColor = RGB(0, 0, 0)
        ShowWindow ItemHwnd, SW_SHOW
        Unload frmtip
    ElseIf IsHidden = 0 Or IsHidden = Null Then
        SetProp ItemHwnd, "VisibiltyState", 7878
        Item.SubItems(1) = "Hide"
        Dim res As Integer
        res = GetINIProfileInt("chkGray", 1)
        If res Then
           Item.ForeColor = HiddenWindowForeColor
        End If
         
        '//Coz google talk doesn't want to be hidden
        If InStr(1, Item.Text, "Google Talk") <> 0 Then
           ShowWindow Item.Tag, SW_MINIMIZE
        End If
                
        ShowWindow ItemHwnd, SW_HIDE
        If ItemHwnd = Me.hwnd Then
           Dim TipSettings As Integer
           TipSettings = GetINIProfileInt("Tip1", 0)
           If TipSettings = 0 Then frmtip.Show
        End If
    End If
        
    'added to avoid a lost focus and get focus again and this causes a r3sha
    PutFocus Me.hwnd
    
Lstwindows_ItemClickExit:
    On Error Resume Next

    Exit Sub
    
Lstwindows_ItemClickErr:

    '//Lstwindows_ItemClickunique number is a unique number for MyLstwindows_ItemClick because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee111"
    
GoTo Lstwindows_ItemClickExit
End Sub

'//PT 8June2007: Used to hide or show windows in the listview
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Lstwindows_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: Lstwindows_MouseUp
' Description: Lstwindows_MouseUp Description
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    
''''@   ''''''''''''\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/\*/'
''''@   ''''''''''''“Girls bored me, they still do. I love Mickey Mouse more than any woman I've ever known.” '
''''@   ''''''''''''                                                                               Walt Disney'
''''@   ''''''''''''/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\/*\'

''''@    On Error GoTo Lstwindows_MouseUpErr
''''@    Debug.Print "2"
''''@    Dim ItemHwnd As Long
''''@    Dim IsHidden As Long
''''@    Dim TaskListIndex As Integer
    
''''@    '//Set Me Top most while the click then set Me "No top most" using tmrNoTopMost_Timer
''''@    '//Okay but why=>>well the next line to make Me TopMost while the user uses Me
''''@    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
''''@    '//and No top most in tmrNoTopMost_Timer to delete Top Most state to avoid nagging the user
    
''''@    If Button = 2 Then Exit Sub '//exit function if right mouse button
                     
''''@    '//pt 8June2007: Find the window that has the title in which the user clicks in the listview
''''@    'ItemHwnd = FindWindow(vbNullString, Lstwindows.SelectedItem.Text)
''''@    'here i could get the window hwnd by using the index of the selected item and use it as an index for
''''@    'the tasklist array however may be this cause unexpected results or get an invalid window hwnd or another
''''@    'window hwnd ==>>> """"""HENCE i will use Findwindow and Lstwindows.SelectedItem.Text""""
    
''''@    '//HOWEVER for any reason, may be FindWindow returns zero
''''@    '//SO i put an a7tamal THAT "Lstwindows.SelectedItem.Text" will larger or smaller
''''@    '//in actual window title, hence i will make a check if ItemHwnd == 0
''''@    If ItemHwnd = 0 Then
''''@            '//Chr(34) is the Quotaion "
''''@            'LogError "Lstwindows_MouseUp::FindWindow returns Zero, can't found " & Chr(34) & Lstwindows.SelectedItem.Text & Chr(34) & " Window"
''''@            '//note that: i added the hwnd of the window to the item tag in loadlistview function
''''@            '//in the next line, i 'll check if the hwnd stored in the Tag is a handle to a valid window
''''@            If IsWindow(Lstwindows.SelectedItem.Tag) Then '//if already there is a window with this handle
                 
''''@                    For TaskListIndex = 0 To 200 'then search TaskList array for this Hwnd
                        
''''@                        '//if i found this hwnd in TaskList array
''''@                        If TaskList(TaskListIndex).TaskID = Lstwindows.SelectedItem.Tag Then
''''@                           '//then i 'll check if the TaskName "Window title" assosisted with this Hwnd in
''''@                           '//the TaskList array is the same as "Lstwindows.SelectedItem.Text"
''''@                           '//so first i 'll check if "Len(TaskList(TaskListIndex).TaskName)" larger than
''''@                           '//"Len(Lstwindows.SelectedItem.Text)" by one char "null terminating"
''''@                           If Len(TaskList(TaskListIndex).TaskName) - Len(Lstwindows.SelectedItem.Text) = 1 Then
''''@                              '//then i 'll check if both are same title
''''@                              If InStr(1, TaskList(TaskListIndex).TaskName, Lstwindows.SelectedItem.Text) Then
''''@                                  'okay now i make sure that the title stored in TaskList is the same title in "Lstwindows.SelectedItem.Text"
''''@                                   ItemHwnd = Lstwindows.SelectedItem.Tag '//then we 'll use this handle
''''@                                   'LogError "Found window's Hwnd in the Item tag"
''''@                              End If
''''@                           End If
''''@                        End If
                        
''''@                    Next TaskListIndex
                    
''''@             End If
''''@    End If
    
''''@    '//if ItemHwnd is still 0 then try for FindWindow by the window's title in the listview
''''@    If ItemHwnd = 0 Then ItemHwnd = FindWindow(vbNullString, Lstwindows.SelectedItem.Text)
    
''''@    '//if ItemHwnd is still 0, so i'll arise an error
''''@    If ItemHwnd = 0 Then
''''@       Err.Description = "Can't hide " & Chr(34) & Lstwindows.SelectedItem.Text & Chr(34) & " window."
''''@       If InnerSysError Then SysError Err, "0x0f3ee111"
''''@       LogError "Lstwindows_MouseUp::Can't hide " & Chr(34) & Lstwindows.SelectedItem.Text & Chr(34) & " window."
''''@    End If
    
''''@    'the advangtage of using getprob by getting priveows hidden windows
''''@    IsHidden = GetProp(ItemHwnd, "VisibiltyState")
''''@    If IsHidden = 7878 Then
''''@        RemoveProp ItemHwnd, "VisibiltyState"
''''@        Lstwindows.SelectedItem.SubItems(1) = "Show"
''''@        Lstwindows.SelectedItem.ForeColor = RGB(0, 0, 0)
''''@        ShowWindow ItemHwnd, SW_SHOW
''''@    ElseIf IsHidden = 0 Or IsHidden = Null Then
''''@        SetProp ItemHwnd, "VisibiltyState", 7878
''''@        Lstwindows.SelectedItem.SubItems(1) = "Hide"
''''@        Lstwindows.SelectedItem.ForeColor = HiddenWindowForeColor
''''@        ShowWindow ItemHwnd, 0
''''@    End If
        
''''@    'added to avoid a lost focus and get focus again and this causes a r3sha
''''@    PutFocus Me.hwnd
    
''''@Lstwindows_MouseUpExit:
''''@    On Error Resume Next

''''@    Exit Sub
    
''''@Lstwindows_MouseUpErr:

''''@    '//Lstwindows_MouseUpunique number is a unique number for MyLstwindows_MouseUp because i don't want to write
''''@    '//the function name in the error message display to the User for 2 reasons
''''@    '//1.It is my internal structure
''''@    '//2.May be Function name will be useful in Reverse engineering "may be"
''''@    If OnErrorSysError Then SysError Err, "0x0f3ee111"
    
''''@GoTo Lstwindows_MouseUpExit

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub cmdSysHide_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
       
    Dim FFile As Integer
    Dim TotalRecordsNumber As Long
    Dim ArrSelectedIndex As Integer
    Dim SysListItemIndex As Integer
    Dim SysToolTipTypeRecordLen As Long
    Dim SysToolTipType As SysIcoToolTip
    Dim TrayListArrType() As TrayIconInfo
    Dim SeletedItemInArrType As TrayIconInfo
    
    TrayListArrType = TrayList.GetSysTrayIcons
     
    For SysListItemIndex = 1 To lstSystrayIcons.ListItems.count
        If lstSystrayIcons.ListItems.Item(SysListItemIndex).Checked Then
           If lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(1) = "Show" Then
              For ArrSelectedIndex = 0 To UBound(TrayListArrType)
                  
                  If TrayListArrType(ArrSelectedIndex).hIcon = lstSystrayIcons.ListItems.Item(SysListItemIndex).Tag Then
                     '//why we added the next line??
                     '//because of two reasons:
                     '//1.we loop through TrayListArrType(ArrSelectedIndex) array and this array contains a non tooltip
                     '//icon and we don't want to process non tooltip
                     '//2. to fix the problem of connection status  "Local Area Connection  Speed: 100.0 Mbps   Status: Connected"
                     '//because this icon has a non tooltip icons "about 3 icons" and this coz this sub to be processed more than one time
                     '//and overwrite the info of the icon and cozs to not display the tootip and the right click menu after re show it
                     If TrayListArrType(ArrSelectedIndex).ToolTip <> "" Then
                         lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(1) = "Hide"
                         lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(2) = TrayListArrType(ArrSelectedIndex).hwnd
                         lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(3) = TrayListArrType(ArrSelectedIndex).uId
                         lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(4) = TrayListArrType(ArrSelectedIndex).hIcon
                         lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(5) = TrayListArrType(ArrSelectedIndex).ToolTip
                         lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(6) = TrayListArrType(ArrSelectedIndex).ucallbackMessage
                         '//Save these as probs in the window in case me closed before showing the icon again
                         SetProp TrayListArrType(ArrSelectedIndex).hwnd, "SysHwnd", TrayListArrType(ArrSelectedIndex).hwnd
                         SetProp TrayListArrType(ArrSelectedIndex).hwnd, "SysuId", TrayListArrType(ArrSelectedIndex).uId
                         SetProp TrayListArrType(ArrSelectedIndex).hwnd, "SyshIcon", TrayListArrType(ArrSelectedIndex).hIcon
                         SetProp TrayListArrType(ArrSelectedIndex).hwnd, "SysucallbackMessage", TrayListArrType(ArrSelectedIndex).ucallbackMessage
                         
                         '//Save Icon tooltip to a file
                         SysToolTipType.hwnd = TrayListArrType(ArrSelectedIndex).hwnd
                         '//we 'll use "#&" as a delimiter to get the end of tooltip text in order to
                         '//avoid displaying the tooltip like this after re-show it "snerfe ..."
                         SysToolTipType.ToolTip = TrayListArrType(ArrSelectedIndex).ToolTip & "#&"
                         FFile = FreeFile()
                         SysToolTipTypeRecordLen = Len(SysToolTipType)
                         
                         Open AppPath + "ico.dat" For Random As FFile Len = SysToolTipTypeRecordLen
                              TotalRecordsNumber = FileLen(AppPath + "ico.dat") / SysToolTipTypeRecordLen
                              Put #FFile, TotalRecordsNumber + 1, SysToolTipType
                         Close FFile
                            
                        ' Dim IconHandle As Long
                        '
                        ' frmmain.PicSysIco.Cls
                        '
                        ' IconHandle = CopyIcon(TrayListArrType(ArrSelectedIndex).hIcon)
                        '
                        ' DrawIcon frmmain.PicSysIco.hDC, 0, 0, TrayListArrType(ArrSelectedIndex).hIcon
                        '
                        ' frmmain.ILLostIcons.ListImages.Add 1, , frmmain.PicSysIco.Image
                            
                         '//Hide Icon//////////////////////////////////////////
                         Dim TrayI As NOTIFYICONDATA
                         Const NIF_TIP = &H4
                         Const NIF_ICON = &H2
                         Const NIF_MESSAGE = &H1
                         TrayI.cbSize = Len(TrayI)
                         TrayI.hwnd = TrayListArrType(ArrSelectedIndex).hwnd
                         TrayI.uId = TrayListArrType(ArrSelectedIndex).uId
                         TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
                         TrayI.hIcon = frmmain.Icon '  TrayListArrType(ArrSelectedIndex).hIcon
                         Shell_NotifyIcon NIM_DELETE, TrayI
                         'Dim i As Long
                         'Dim hTray As Long
                         'hTray = GetHwndTray
                         'i = SendMessage(hTray, TB_HIDEBUTTON, TrayListArrType(ArrSelectedIndex).BITMAP + 1, ByVal 1&)
                         '//////////////////////////////////////////////////////
                  End If
                  End If
              Next ArrSelectedIndex
           End If
         End If
    Next SysListItemIndex
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub cmdSysShow_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Dim SysListItemIndex As Integer
      
    For SysListItemIndex = 1 To lstSystrayIcons.ListItems.count
        If lstSystrayIcons.ListItems.Item(SysListItemIndex).Checked Then
           If lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(1) = "Hide" Then
              ShowSysTrayIcon SysListItemIndex
              '//Remove the probs from the window
              RemoveProp lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(2), "SysuId"
              RemoveProp lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(2), "SysHwnd"
              RemoveProp lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(2), "SyshIcon"
              RemoveProp lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(2), "SysucallbackMessage"
           End If
        End If
    Next SysListItemIndex
     
    '//Rewrite the ico.dat again to remove the icon tooltip and hwnd that already show now
    ReWriteIcoDatFile
    
End Sub

Public Sub ReWriteIcoDatFile()
    
    'the idea here, badl ma a3ml search in the file, i will delete it and write it again
    Dim FFile As Integer
    Dim iIndex As Integer
    Dim TotalRecordsNumber As Long
    Dim SysToolTipType As SysIcoToolTip
    Dim SysToolTipTypeRecordLen As Long
        
    On Error Resume Next
    Kill AppPath + "ico.dat"
    On Error GoTo 0
                                
    For iIndex = 1 To lstSystrayIcons.ListItems.count
        If lstSystrayIcons.ListItems.Item(iIndex).SubItems(1) = "Hide" Then
           SysToolTipType.hwnd = lstSystrayIcons.ListItems.Item(iIndex).SubItems(2)
           SysToolTipType.ToolTip = lstSystrayIcons.ListItems.Item(iIndex).SubItems(5)
           FFile = FreeFile()
           SysToolTipTypeRecordLen = Len(SysToolTipType)
           Open AppPath + "ico.dat" For Random As FFile Len = SysToolTipTypeRecordLen
                TotalRecordsNumber = FileLen(AppPath + "ico.dat") / SysToolTipTypeRecordLen
                Put #FFile, TotalRecordsNumber + 1, SysToolTipType
           Close FFile
         End If
    Next iIndex
                           
End Sub
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub cmdSysHideAll_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
        
    Dim FFile As Integer
    Dim TotalRecordsNumber As Long
    Dim ArrSelectedIndex As Integer
    Dim SysListItemIndex As Integer
    Dim SysToolTipTypeRecordLen As Long
    Dim SysToolTipType As SysIcoToolTip
    Dim TrayListArrType() As TrayIconInfo
    Dim SeletedItemInArrType As TrayIconInfo
    
    TrayListArrType = TrayList.GetSysTrayIcons
     
    For SysListItemIndex = 1 To lstSystrayIcons.ListItems.count
           If lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(1) = "Show" Then
              For ArrSelectedIndex = 0 To UBound(TrayListArrType)
                  
                  If TrayListArrType(ArrSelectedIndex).hIcon = lstSystrayIcons.ListItems.Item(SysListItemIndex).Tag Then
                     '//why we added the next line??
                     '//because of two reasons:
                     '//1.we loop through TrayListArrType(ArrSelectedIndex) array and this array contains a non tooltip
                     '//icon and we don't want to process non tooltip
                     '//2. to fix the problem of connection status  "Local Area Connection  Speed: 100.0 Mbps   Status: Connected"
                     '//because this icon has a non tooltip icons "about 3 icons" and this coz this sub to be processed more than one time
                     '//and overwrite the info of the icon and cozs to not display the tootip and the right click menu after re show it
                     If TrayListArrType(ArrSelectedIndex).ToolTip <> "" Then
                         lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(1) = "Hide"
                         lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(2) = TrayListArrType(ArrSelectedIndex).hwnd
                         lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(3) = TrayListArrType(ArrSelectedIndex).uId
                         lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(4) = TrayListArrType(ArrSelectedIndex).hIcon
                         lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(5) = TrayListArrType(ArrSelectedIndex).ToolTip
                         lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(6) = TrayListArrType(ArrSelectedIndex).ucallbackMessage
                         '//Save these as probs in the window in case me closed before showing the icon again
                         SetProp TrayListArrType(ArrSelectedIndex).hwnd, "SysHwnd", TrayListArrType(ArrSelectedIndex).hwnd
                         SetProp TrayListArrType(ArrSelectedIndex).hwnd, "SysuId", TrayListArrType(ArrSelectedIndex).uId
                         SetProp TrayListArrType(ArrSelectedIndex).hwnd, "SyshIcon", TrayListArrType(ArrSelectedIndex).hIcon
                         SetProp TrayListArrType(ArrSelectedIndex).hwnd, "SysucallbackMessage", TrayListArrType(ArrSelectedIndex).ucallbackMessage
                         
                         '//Save Icon tooltip to a file
                         SysToolTipType.hwnd = TrayListArrType(ArrSelectedIndex).hwnd
                         '//we 'll use "#&" as a delimiter to get the end of tooltip text in order to
                         '//avoid displaying the tooltip like this after re-show it "snerfe ..."
                         SysToolTipType.ToolTip = TrayListArrType(ArrSelectedIndex).ToolTip & "#&"
                         FFile = FreeFile()
                         SysToolTipTypeRecordLen = Len(SysToolTipType)
                         
                         Open AppPath + "ico.dat" For Random As FFile Len = SysToolTipTypeRecordLen
                              TotalRecordsNumber = FileLen(AppPath + "ico.dat") / SysToolTipTypeRecordLen
                              Put #FFile, TotalRecordsNumber + 1, SysToolTipType
                         Close FFile
                            
                        ' Dim IconHandle As Long
                        '
                        ' frmmain.PicSysIco.Cls
                        '
                        ' IconHandle = CopyIcon(TrayListArrType(ArrSelectedIndex).hIcon)
                        '
                        ' DrawIcon frmmain.PicSysIco.hDC, 0, 0, TrayListArrType(ArrSelectedIndex).hIcon
                        '
                        ' frmmain.ILLostIcons.ListImages.Add 1, , frmmain.PicSysIco.Image
                            
                         '//Hide Icon//////////////////////////////////////////
                         Dim TrayI As NOTIFYICONDATA
                         Const NIF_TIP = &H4
                         Const NIF_ICON = &H2
                         Const NIF_MESSAGE = &H1
                         TrayI.cbSize = Len(TrayI)
                         TrayI.hwnd = TrayListArrType(ArrSelectedIndex).hwnd
                         TrayI.uId = TrayListArrType(ArrSelectedIndex).uId
                         TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
                         TrayI.hIcon = frmmain.Icon '  TrayListArrType(ArrSelectedIndex).hIcon
                         Shell_NotifyIcon NIM_DELETE, TrayI
                         'Dim i As Long
                         'Dim hTray As Long
                         'hTray = GetHwndTray
                         'i = SendMessage(hTray, TB_BUTTONCOUNT, ByVal 0&, ByVal 0&)
                         'Dim k As Long
                         'For k = 0 To i - 1
                         '   Call SendMessage(hTray, TB_HIDEBUTTON, k, ByVal 1&)
                         'Next k
                         'Exit Sub
                          '//////////////////////////////////////////////////////
                  End If
                  End If
              Next ArrSelectedIndex
           End If
    Next SysListItemIndex

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub cmdSysRevert_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Dim FFile As Integer
    Dim TotalRecordsNumber As Long
    Dim ArrSelectedIndex As Integer
    Dim SysListItemIndex As Integer
    Dim SysToolTipTypeRecordLen As Long
    Dim SysToolTipType As SysIcoToolTip
    Dim TrayListArrType() As TrayIconInfo
    Dim SeletedItemInArrType As TrayIconInfo
    
    TrayListArrType = TrayList.GetSysTrayIcons


    For SysListItemIndex = 1 To lstSystrayIcons.ListItems.count
           If lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(1) = "Hide" Then
              
              'Dim i As Long
              'Dim hTray As Long
              'hTray = GetHwndTray
              'i = SendMessage(hTray, TB_BUTTONCOUNT, ByVal 0&, ByVal 0&)
              'Dim k  As Long
              'For k = 0 To i - 1
              '    Call SendMessage(hTray, TB_HIDEBUTTON, k, ByVal 0&)
              'Next k
              ShowSysTrayIcon SysListItemIndex
              '//Remove the probs from the window
              RemoveProp lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(2), "SysuId"
              RemoveProp lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(2), "SysHwnd"
              RemoveProp lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(2), "SyshIcon"
              RemoveProp lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(2), "SysucallbackMessage"
           
           ElseIf lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(1) = "Show" Then
           
                  For ArrSelectedIndex = 0 To UBound(TrayListArrType)
                  
                     If TrayListArrType(ArrSelectedIndex).hIcon = lstSystrayIcons.ListItems.Item(SysListItemIndex).Tag Then
                        '//why we added the next line??
                        '//because of two reasons:
                        '//1.we loop through TrayListArrType(ArrSelectedIndex) array and this array contains a non tooltip
                        '//icon and we don't want to process non tooltip
                        '//2. to fix the problem of connection status  "Local Area Connection  Speed: 100.0 Mbps   Status: Connected"
                        '//because this icon has a non tooltip icons "about 3 icons" and this coz this sub to be processed more than one time
                        '//and overwrite the info of the icon and cozs to not display the tootip and the right click menu after re show it
                        If TrayListArrType(ArrSelectedIndex).ToolTip <> "" Then
                           lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(1) = "Hide"
                           lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(2) = TrayListArrType(ArrSelectedIndex).hwnd
                           lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(3) = TrayListArrType(ArrSelectedIndex).uId
                           lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(4) = TrayListArrType(ArrSelectedIndex).hIcon
                           lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(5) = TrayListArrType(ArrSelectedIndex).ToolTip
                           lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(6) = TrayListArrType(ArrSelectedIndex).ucallbackMessage
                           '//Save these as probs in the window in case me closed before showing the icon again
                           SetProp TrayListArrType(ArrSelectedIndex).hwnd, "SysHwnd", TrayListArrType(ArrSelectedIndex).hwnd
                           SetProp TrayListArrType(ArrSelectedIndex).hwnd, "SysuId", TrayListArrType(ArrSelectedIndex).uId
                           SetProp TrayListArrType(ArrSelectedIndex).hwnd, "SyshIcon", TrayListArrType(ArrSelectedIndex).hIcon
                           SetProp TrayListArrType(ArrSelectedIndex).hwnd, "SysucallbackMessage", TrayListArrType(ArrSelectedIndex).ucallbackMessage
                          
                           '//Save Icon tooltip to a file
                           SysToolTipType.hwnd = TrayListArrType(ArrSelectedIndex).hwnd
                           '//we 'll use "#&" as a delimiter to get the end of tooltip text in order to
                           '//avoid displaying the tooltip like this after re-show it "snerfe ..."
                           SysToolTipType.ToolTip = TrayListArrType(ArrSelectedIndex).ToolTip & "#&"
                           FFile = FreeFile()
                           SysToolTipTypeRecordLen = Len(SysToolTipType)
                         
                           Open AppPath + "ico.dat" For Random As FFile Len = SysToolTipTypeRecordLen
                                TotalRecordsNumber = FileLen(AppPath + "ico.dat") / SysToolTipTypeRecordLen
                                Put #FFile, TotalRecordsNumber + 1, SysToolTipType
                           Close FFile
                            
                           ' Dim IconHandle As Long
                           '
                           ' frmmain.PicSysIco.Cls
                           '
                           ' IconHandle = CopyIcon(TrayListArrType(ArrSelectedIndex).hIcon)
                           '
                           ' DrawIcon frmmain.PicSysIco.hDC, 0, 0, TrayListArrType(ArrSelectedIndex).hIcon
                           '
                           ' frmmain.ILLostIcons.ListImages.Add 1, , frmmain.PicSysIco.Image
                            
                           '//Hide Icon//////////////////////////////////////////
                           Dim TrayI As NOTIFYICONDATA
                           Const NIF_TIP = &H4
                           Const NIF_ICON = &H2
                           Const NIF_MESSAGE = &H1
                           TrayI.cbSize = Len(TrayI)
                           TrayI.hwnd = TrayListArrType(ArrSelectedIndex).hwnd
                           TrayI.uId = TrayListArrType(ArrSelectedIndex).uId
                           TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
                           TrayI.hIcon = frmmain.Icon '  TrayListArrType(ArrSelectedIndex).hIcon
                           Shell_NotifyIcon NIM_DELETE, TrayI
                           'Dim i As Long
                           'Dim hTray As Long
                           'hTray = GetHwndTray
                           'i = SendMessage(hTray, TB_BUTTONCOUNT, ByVal 0&, ByVal 0&)
                           'Dim k As Long
                           'For k = 0 To i - 1
                           ' Call SendMessage(hTray, TB_HIDEBUTTON, k, ByVal 1&)
                           'Next k
                           'Exit Sub
                            '//////////////////////////////////////////////////////
                       End If
                    End If
                Next ArrSelectedIndex
           End If
    Next SysListItemIndex

    '//Rewrite the ico.dat again to remove the icon tooltip and hwnd that already show now
    ReWriteIcoDatFile

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub cmdSysShowAll_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
   
   Dim SysListItemIndex As Integer
      
    For SysListItemIndex = 1 To lstSystrayIcons.ListItems.count
           If lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(1) = "Hide" Then
              'Dim i As Long
              'Dim hTray As Long
              'hTray = GetHwndTray
              'i = SendMessage(hTray, TB_BUTTONCOUNT, ByVal 0&, ByVal 0&)
              'Dim k  As Long
              'For k = 0 To i - 1
              '    Call SendMessage(hTray, TB_HIDEBUTTON, k, ByVal 0&)
              'Next k
              ShowSysTrayIcon SysListItemIndex
              '//Remove the probs from the window
              RemoveProp lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(2), "SysuId"
              RemoveProp lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(2), "SysHwnd"
              RemoveProp lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(2), "SyshIcon"
              RemoveProp lstSystrayIcons.ListItems.Item(SysListItemIndex).SubItems(2), "SysucallbackMessage"
           End If
    Next SysListItemIndex
     
    '//Rewrite the ico.dat again to remove the icon tooltip and hwnd that already show now
    ReWriteIcoDatFile

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuOpen_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If IsWindowVisible(hwnd) = 0 Then
       ShowMe
       Me.WindowState = 0
       Me.SetFocus
    End If
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub menuAbout2_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    frmabout.Show 1, frmmain

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub menuWebSite2_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    mnuWebSite_Click
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuHelp_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Shell "hh " & AppPath & "help.chm", vbMaximizedFocus
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuHelp3_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    mnuHelp_Click
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuWebSite_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    ShellExecute Me.hwnd, vbNullString, SiteName, vbNullString, "", SW_SHOWNORMAL

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuExit_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    tmrUpdateListView.Enabled = False
    
    ExitMeNow = True
    
    Unload Me

End Sub

'//I know that this timer to disable the Scroll bars, however because it occurs only one time
'//then i will use it in some process that needs to run only once like GetList
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub tmrDisableSysListHorzBar_Timer()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     
    '//This Timer run only once to disable the systray list horz bar just after loading frmmain
    ShowScrollBar Lstwindows.hwnd, SB_HORZ, False
    ShowScrollBar lstSystrayIcons.hwnd, SB_HORZ, False
    
    '//because the following lines in form_load cozs the form first to be appear at 0,0
    '//3ashan kada ana 7ateeet el-lines de hana
    Dim res As Long
    res = GetINIProfileInt("chkShowTopMost", 0)
   
    If res Then
       IsShowTopMost = True
       SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If
    
    '//moved "GetList: from Form_Load
    '//Why??
    '//Because frmmain itself doesn't appears in listview immediatly, we need to wait for tmrupdateListview
    '//because in Form_Load is already doesn't appeard yet hence i moved GetList to here to add Frmmain to listview
    '//just after you run the proram
    GetList
    
    '//No need to this timer again
    tmrDisableSysListHorzBar.Enabled = False
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub tmrUpdateListView_Timer()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
        
     Lstwindows.HotTracking = IsHotTrack
     Lstwindows.HoverSelection = IsHotTrack
     
     '//handle the option of "Hide my system tray icon when hiding me"
     'Dim res As Integer
     'res = GetINIProfileInt("chkHideTrayIcon", 1)
     'If res Then
     '   If IsWindowVisible(hwnd) = 0 Then '//if the me is already hidden
     '      Dim IsHiddenWindow As Long
     '      IsHiddenWindow = GetProp(hwnd, "VisibiltyState") '//check if THIS window is hidden
     '      If IsHiddenWindow = 7878 Then '//and if also my state is Hide in the listview
     '         If TrayIcon.IsDisplayed Then TrayIcon.Hide '//then this is the case of "Hide my system tray icon when hiding me"
     '      End If
     '   End If
     'End If
    
     If IsAlwaysTrayIcon Then
        If Not TrayIcon.IsDisplayed Then
           'If IsHiddenWindow <> 7878 Then '//Only show me if (my state isn't hide in the listview and this option "Hide my system tray icon when hiding me" is unchecked
              TrayIcon.Show Icon, SoftwareName
           'End If
        End If
     ElseIf IsNeverTrayIcon Then
        If TrayIcon.IsDisplayed Then TrayIcon.Hide
     ElseIf IsTrayIconOnlyWhenMini Then
        If TrayIcon.IsDisplayed Then
           If IsWindowVisible(hwnd) <> 0 Then TrayIcon.Hide
        End If
     End If
      
     GetList 'Update the listview whenever a change occure
     CheckSysTrayList
     
     '//For unknown reason, this line in CheckSysTrayList function
     '//i = frmmain.TrayList.IconCount
     '//causes a delay in the program startup, you can notice this if you click
     '//frmopions just after you run the exe
     '//hence, i will make tmrUpdateListView.Interval = 4000 in the property pages
     '//then and i will update the listviews in form_load
     '//also inside tmrUpdateListView_Timer, i will return the timer interval again to 1, just after the startup period get passed
     tmrUpdateListView.Interval = 200
     '//200 to improve the performance and to decrease CPU Usage
     
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub tmrNoTopMost_Timer()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     
    '//This Timer run only once to disable the systray list horz bar just after loading frmmain
    ShowScrollBar Lstwindows.hwnd, SB_HORZ, False
    ShowScrollBar lstSystrayIcons.hwnd, SB_HORZ, False
             
    '//if Show as TopMost then don't proceed
    If Not IsShowTopMost Then
        '//we add the "If IsWindowVisible(hwnd) Then" to avoid displaying this window if the user hide it using the shortcut key or even with the normal way
        If IsWindowVisible(hwnd) Then
            '//"No top most" to delete Top Most state we created in Lstwindows_MouseUp and cmdshowwindow_Click and cmdShowAll_Click _
            'and cmdHide_Click and CmdHideAll_Click and CmdRevert_Click avoid nagging the user
            SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
        End If
    End If
    
    CheckShortcutKeys
    
    Dim i As Integer
    
    For i = 1 To Lstwindows.ListItems.count
        
        '//Check the listview for erros , if a window is listed as Hide and it is visible
        If Lstwindows.ListItems.Item(i).SubItems(1) = "Hide" Then
             
           If IsWindowVisible(Lstwindows.ListItems.Item(i).Tag) Then
              Lstwindows.ListItems(i).Checked = True
              cmdshowwindow_Click
              Lstwindows.ListItems(i).Checked = False
           End If
            
        End If
        
    Next i
    
     
    If lstSystrayIcons.ColumnHeaders(3).Width > 0 Then
       
       lstSystrayIcons.ColumnHeaders(3).Width = 0
    
    ElseIf lstSystrayIcons.ColumnHeaders(4).Width > 0 Then
    
       lstSystrayIcons.ColumnHeaders(4).Width = 0
      
    ElseIf lstSystrayIcons.ColumnHeaders(5).Width > 0 Then
    
       lstSystrayIcons.ColumnHeaders(5).Width = 0
       
    ElseIf lstSystrayIcons.ColumnHeaders(6).Width > 0 Then
    
       lstSystrayIcons.ColumnHeaders(6).Width = 0
       
    ElseIf lstSystrayIcons.ColumnHeaders(7).Width > 0 Then
    
       lstSystrayIcons.ColumnHeaders(7).Width = 0
            
    End If
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub TrayIcon_LeftMouseDown()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If IsWindowVisible(hwnd) = 0 Then
       ShowMe
       Me.WindowState = 0
       Me.SetFocus
    End If
      
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub TrayIcon_RightMouseDown()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    TrayIcon.TrackPopMenu = True
    Me.PopupMenu mnuSystray
    
End Sub

'//Shows me and also show me in the listview
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ShowMe()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: ShowMe
' Description: ShowMe Description
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo ShowMeErr
        
      Dim MyItemInTheListView As ListItem
      Set MyItemInTheListView = frmmain.Lstwindows.FindItem(frmmain.Caption) '//Search for This Software Item in the listview
      MyItemInTheListView.Checked = True '//Check This Software Item in the listview in order to call cmdshowwindow_Click
      frmmain.cmdshowwindow_Click '//Show This Software
      MyItemInTheListView.Checked = False '//Uncheck This Software Item in the listview in order to not appear checked to the user
    
ShowMeExit:
    On Error Resume Next

    Exit Sub
    
ShowMeErr:

    '//0x0f3ee124 number is a unique number for MyShowMe because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee124"
    
GoTo ShowMeExit

End Sub

'//Hide me and also hide me in the listview
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub HideMe()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: HideMe
' Description: HideMe Description
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006 - 2007,  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo HideMeErr
        
      Dim MyItemInTheListView As ListItem
      Set MyItemInTheListView = Lstwindows.FindItem(frmmain.Caption) '//Search for This Software Item in the listview
      MyItemInTheListView.Checked = True '//Check This Software Item in the listview in order to call cmdshowwindow_Click
      cmdHide_Click  '//Hide This Software
      MyItemInTheListView.Checked = False '//Uncheck This Software Item in the listview in order to not appear checked to the user
    
HideMeExit:
    On Error Resume Next

    Exit Sub
    
HideMeErr:

    '//0x0f3ee123 number is a unique number for MyHideMe because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    If OnErrorSysError Then SysError Err, "0x0f3ee123"
    
GoTo HideMeExit

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub cmdOptions_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Unload frmOptions '//must unload if it invisbally loaded
    
    If IsShowTopMost Then
       frmOptions.Show 0, frmmain
       '//why we add the next line after the prevoius line??
       '//to avoid showing frmoption at poz 0,0 then move it to the center screen
       '//now it will be displayed in the center screen 3ly tooooooool
       SetWindowPos frmOptions.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Else
       frmOptions.Show 1, frmmain
    End If
   
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub cmdHelp_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    PopupMenu menuHelp2, 0, cmdHelp.Left, cmdHelp.Top + cmdHelp.Height, mnuHelp3

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub ShowSysTrayIcon(Index As Integer)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Const NIF_TIP = &H4
    Const NIM_ADD = &H0
    Const NIF_ICON = &H2
    Const NIF_MESSAGE = &H1
    Dim rea As Boolean
    Dim TrayI As NOTIFYICONDATA
         
    lstSystrayIcons.ListItems(Index).SubItems(1) = "Show"
    TrayI.cbSize = Len(TrayI)
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.hwnd = lstSystrayIcons.ListItems(Index).SubItems(2)
    TrayI.uId = lstSystrayIcons.ListItems(Index).SubItems(3)
    TrayI.hIcon = lstSystrayIcons.ListItems(Index).SubItems(4)  'ILSysIcons.ListImages.Item(lstSystrayIcons.ListItems(Index).SmallIcon).ExtractIcon
    '//if i didn't add Chr$(0) to the next line, then tooltip will be 64 char in long
    '//putting Chr$(0) terminate it "quoted from ASTC Source"
    TrayI.szTip = lstSystrayIcons.ListItems(Index).SubItems(5) & Chr$(0)
    TrayI.ucallbackMessage = lstSystrayIcons.ListItems(Index).SubItems(6)
    Dim res As Boolean
    res = Shell_NotifyIcon(NIM_ADD, TrayI)
     
    If Not res Then
       TrayI.hIcon = ILSysIcons.ListImages.Item(lstSystrayIcons.ListItems(Index).SmallIcon).ExtractIcon
       res = Shell_NotifyIcon(NIM_ADD, TrayI)
    End If
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
   
   '//If the user click mnuExit then don't process this function
   '//why??
   '//because this coz the software to not to be unclosed
   If ExitMeNow Then Exit Sub
   
   '//because windows can't shut down if chkMinInsteadOfClosing is true
   If UnloadMode = vbAppWindows Then Exit Sub
   
   '//Handle the option of "Minimize to system tray instead of closing"
   Dim res As Integer
   
   res = GetINIProfileInt("chkMinInsteadOfClosing", 1)
   
   If res Then
       
      Cancel = 1
      
      If Not TrayIcon.IsDisplayed Then TrayIcon.Show Icon, SoftwareName
       
      HideMe
      
   End If

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Resize()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Dim res As Integer
    
    res = GetINIProfileInt("optmini", 0)
    
    If res Then
    
       If Me.WindowState = 1 Then
          
          If Not TrayIcon.IsDisplayed Then TrayIcon.Show Icon, SoftwareName
          
          HideMe
       
       ElseIf Me.WindowState = 0 Then
          
          If TrayIcon.IsDisplayed Then TrayIcon.Hide
       
       End If
        
    End If
    
    '//if the icon is displayed and optmini unselected
    If TrayIcon.IsDisplayed And res = 0 Then
       
       If Me.WindowState = 1 Then
          
          HideMe
       
       End If
    
    End If
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Unload(Cancel As Integer)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Cleanup:
    
    '//clear the virutal memory we used
    VirtualFreeEx hProc, ByVal ptrVM, 0&, MEM_RELEASE
    
    CloseHandle hProc

    Erase TaskList
    
    Set CVisWindows = Nothing
        
    Unload frmOptions
    
    TrayIcon.Hide

End Sub
