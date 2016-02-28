VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "jbet33.ocx"
Object = "{30E5838E-8E26-11DA-A1C2-807A50C10000}#1.0#0"; "acomte445.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmmain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmmain.frx":08CA
   ScaleHeight     =   4665
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin JwldButn2b.JeweledButton JewHelp 
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Help"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin VB.Timer tmrEnabletmrGenChecks 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   11640
      Top             =   1440
   End
   Begin VB.Timer tmrGenChecks 
      Interval        =   1000
      Left            =   11160
      Top             =   1440
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   11160
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
      Begin VB.PictureBox PIco 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   11040
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin ASTC.TrayIcon TrayIcon 
      Left            =   11160
      Top             =   1920
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin MSComctlLib.ImageList ILAppIcons 
      Left            =   11160
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin JwldButn2b.JeweledButton JewRemove 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "R&emove"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin JwldButn2b.JeweledButton JewTerminate 
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Terminate"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin MSComctlLib.ListView lstStartupPrograms 
      Height          =   3735
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Application Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Location"
         Object.Width           =   8996
      EndProperty
   End
   Begin JwldButn2b.JeweledButton JewOptions 
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Options"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin JwldButn2b.JeweledButton JewRefresh 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "Refresh startup applications list"
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Refresh"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin JwldButn2b.JeweledButton JewRun 
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      ToolTipText     =   "Run selected startup application"
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "R&un"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin JwldButn2b.JeweledButton JewBuy 
      Height          =   375
      Left            =   9480
      TabIndex        =   9
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Buy Now!"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin JwldButn2b.JeweledButton JewAdd 
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Add..."
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin VB.Shape shpMain 
      Height          =   4455
      Left            =   120
      Top             =   120
      Width           =   10695
   End
   Begin VB.Menu mnuSystray 
      Caption         =   "Systray Menu"
      Visible         =   0   'False
      Begin VB.Menu mnushow 
         Caption         =   "Show Handy Startup Monitor"
      End
      Begin VB.Menu mnuEnable 
         Caption         =   "Enable Startup Monitor"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWebSite 
         Caption         =   "WebSite..."
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu menupopup 
      Caption         =   "popumenu"
      Visible         =   0   'False
      Begin VB.Menu mnuRun 
         Caption         =   "Run/Terminate"
      End
      Begin VB.Menu mnuCopyFilePath 
         Caption         =   "Copy File Path"
      End
      Begin VB.Menu mnuOpenFileLocation 
         Caption         =   "&Open File Location"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove From Startup"
      End
   End
   Begin VB.Menu mnuHelpButton 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu mnuContents 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnuWebSiteFromHelp 
         Caption         =   "Visit Website"
      End
      Begin VB.Menu mnuPurchase 
         Caption         =   "Purchase Online"
      End
      Begin VB.Menu mnuEnterRegistrationKey 
         Caption         =   "Enter Registration Key..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAboutFromHelpButton 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

                                        '//Be2sm ALLAH elRa7man ElRa7eem'\\
                                
                                        '____oooooooooo_______oooooooo
                                        '___ooooooooooooo___ooooooooooooo
                                        '___oooooooooooooo_oooooooooooooo
                                        '___ooooooooooooooooooooooooooooo
                                        '____ooooooooooooooooooooooooooo
                                        '_____ooooooooooooooooooooooooo
                                        '______ooooooooooooooooooooooo
                                        '_________oooooooooooooooooo
                                        '___________ooooooooooooo
                                        '_____________ooooooooo
                                        '______________oooooo
                                        '_______________oooo
                                        '_______________ooo
                                        '______________oo
                                        '_____________o
                                        '___________o
                                        '________oo
                                        '______ooo
                                        '_____oooo
                                        '___oooooo
                                        '___ooooooo
                                        '____oooooooo
                                        '______oooooooo
                                        '________oooooooo
                                        '_________ooooooo
                                        '_________ooooooo
                                        '________oooooo........FrIeNdShIp
                                        '_______ooooo.......Is Like A Flower;
                                        '_____ooooo........Growing In Its Glory,
                                        '___ooooo...........Telling its own story;
                                        '__ooo.............FrIeNdShIp is Precious,
                                        '__oo..................Not Only In Shade,
                                        '__ o............But In D Sunshine Of Life.
                                        '....Thanks For Being My Friend
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '"√Õ» ⁄‰œ„« ÌﬁÊ· ·Ì «·‰«” ·« Ì„ﬂ‰ﬂ  ÕﬁÌﬁ –·ﬂ , ·√‰Â ÿÊ«· ÕÌ« Ì Ê«·‰«” ÌﬁÊ·Ê‰ ·Ì ·‰  ‰ÃÕ √»œ« ."~ œ  Ì—‰—
            '365 „ﬁÊ·… ··‰Ã«Õ~—ƒÊ› ‘»«Ìﬂ
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    
Option Explicit

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Load()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    'SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    
    Dim res As Integer
    Dim C1W As Integer
    Dim C2W As Integer
    Dim C3W As Integer
    Dim C4W As Integer
    Dim TickCount As Long
    Dim MainFormWidth As Integer
    Dim MainFormHeight As Integer
    If App.PrevInstance = True Then End
     
    JewRun.Caption = "R&un"
    JewHelp.Caption = "&Help"
    JewAdd.Caption = "&Add..."
    JewRemove.Caption = "R&emove"
    JewBuy.Caption = "&Buy Now!"
    JewRefresh.Caption = "&Refresh"
    JewOptions.Caption = "&Options"
    JewTerminate.Caption = "&Terminate"
     
    GetStartupPaths
    
    ShowSystrayIcon
     
    res = GetINIProfileInt("chkRemoved", 1)
    If res Then rchkRemoved = True
    
    res = GetINIProfileInt("chkNewlyAdded", 1)
    If res Then rchkNewlyAdded = True
    
    res = GetINIProfileInt("chkCheckatStartup", 0)
    If res Then rCheckatStartup = True
    
    res = GetINIProfileInt("chkAllowMonitor", 1)
    If res Then rMonitorAllowed = True
    
    TickCount = GetTickCount
    
    If Not IsNeverTrayIcon And TickCount <= 600000 Then
        
       If IsAlwaysTrayIcon Then
       
          Me.Hide
       
       ElseIf IsTrayIconOnlyWhenMini Then
          
          'Me.Hide
          
          If Not TrayIcon.IsDisplayed Then TrayIcon.Show Me.Icon, SoftwareName
          
       End If
       
    End If
     
    FillList
     
    SortListsFromini
    
    'Dim DaysLeft As String

    'DaysLeft = GetDaysLeft
      
    'If DaysLeft = "" Then
       'Me.Caption = App.Title
       'JewBuy.Visible = False
       'mnuEnterRegistrationKey.Visible = False
    'Else
       'Me.Caption = App.Title & " - Unregistered Trial Version"
    'End If
    '// 22/01/2010: now i will disable armadillo and package the software as is to avoid virus
    Me.Caption = App.Title & " - Unregistered Trial Version"
    
    MainFormWidth = GetINIProfileInt("W", Me.Width)
    MainFormHeight = GetINIProfileInt("H", Me.Height)
    C1W = GetINIProfileInt("C1W", lstStartupPrograms.ColumnHeaders(1).Width)
    C2W = GetINIProfileInt("C2W", lstStartupPrograms.ColumnHeaders(2).Width)
    C3W = GetINIProfileInt("C3W", lstStartupPrograms.ColumnHeaders(3).Width)
    C4W = GetINIProfileInt("C4W", lstStartupPrograms.ColumnHeaders(4).Width)
    
    Me.Width = MainFormWidth
    Me.Height = MainFormHeight
    lstStartupPrograms.ColumnHeaders(1).Width = C1W
    lstStartupPrograms.ColumnHeaders(2).Width = C2W
    lstStartupPrograms.ColumnHeaders(3).Width = C3W
    lstStartupPrograms.ColumnHeaders(4).Width = C4W
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub FadeNow()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Dim FadeSettings As Integer '//Buffer that will store a value from the ini that tell us if this form Faded before or not
       
    FadeSettings = GetINIProfileInt("Fade", 0)  '//Check if this form Faded before or not
    
    If FadeSettings = 0 Then FadeForm Me '//Fade this S if it doesn't faded before

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub GetStartupPaths()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    '//because the paths of startup folders is constant, no need to put it in the refreashing loop or even if Fill function
    CUStartupFolderPath = GetSpecialFolderPathbyAPI(frmmain.hwnd, CSIDL_STARTUP)
    If CUStartupFolderPath = "null" Then
       CUStartupFolderPath = GetSpecialFolderPathbyShell("StartUp")
    End If
     
    AllUStartupFolderPath = GetSpecialFolderPathbyAPI(frmmain.hwnd, CSIDL_COMMON_STARTUP)
    If AllUStartupFolderPath = "null" Then
       AllUStartupFolderPath = GetSpecialFolderPathbyShell("AllUsersStartup")
    End If
   
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function GetDaysLeft() As String
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Dim DaysLeft As String
    DaysLeft = ArmEnviron$("DAYSLEFT")
    GetDaysLeft = DaysLeft
    
End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub TerminateApplication(ItemIndex As Integer)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Dim message As String
    Dim res As VbMsgBoxResult
    Dim ProgramPath As String
    Dim ProgramName As String
    Dim IsWarningMessageDisplay As Boolean
    
    message = "WARNING: Terminating a process can cause undesired" & Chr(13) & Chr(10) & "results including loss of data and system instability. The" & Chr(13) & Chr(10) & "process will not be given the chance to save its state or" & Chr(13) & Chr(10) & "data before its terminated." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Are you sure you want to terminate the process??"
        
    If Not IsWarningMessageDisplay Then res = MsgBox(message, vbExclamation + vbYesNo)
    IsWarningMessageDisplay = True
    If res = vbYes Then
       ProgramPath = ExtractPath(lstStartupPrograms.ListItems(ItemIndex).SubItems(1))
       ProgramName = ExtractFileNameFromPath(ProgramPath)
       If IsProcessRunning(ProgramName) Then
          Killapp (ExtractFileNameFromPath(ProgramPath))
       End If
    End If
        
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub RemoveFromStartup(ItemIndex As Integer)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Dim hKey As Long
Dim strPath As String
Dim Location As String
Dim ProgramPath As String

    Location = lstStartupPrograms.ListItems(ItemIndex).SubItems(3)
           
    Select Case Left(Location, 4)
           Case "HKCR"
                hKey = HKEY_CLASSES_ROOT
           Case "HKCU"
                hKey = HKEY_CURRENT_USER
           Case "KHCR"
                hKey = HKEY_CLASSES_ROOT
           Case "HKLM"
                hKey = HKEY_LOCAL_MACHINE
    End Select
           
          
    If Location = "Current User Startup" Or Location = "All Users Startup" Then
    
       ProgramPath = ExtractPath(lstStartupPrograms.ListItems(ItemIndex).SubItems(1))
       
       If Location = "Current User Startup" Then
          
          DeleteShortcut ProgramPath, "Startup", lstStartupPrograms.ListItems(ItemIndex).Tag
       
       ElseIf Location = "All Users Startup" Then
          
          DeleteShortcut ProgramPath, "AllUsersStartup", lstStartupPrograms.ListItems(ItemIndex).Tag
          
       End If
    
    Else
       
       strPath = Mid(Location, 6, Len(Location) - 5)
           
       Call DeleteValue(hKey, strPath, lstStartupPrograms.ListItems(ItemIndex))
      
    End If
           
    JewRefresh_Click
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ShowSystrayIcon()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Dim res As Integer
    
    '//Check if always display a TrayIcon
    res = GetINIProfileInt("optalways", 0)
    If res Then IsAlwaysTrayIcon = True
    
    res = GetINIProfileInt("optnever", 0)
    If res Then IsNeverTrayIcon = True
   
    res = GetINIProfileInt("optmini", 1)
    If res Then IsTrayIconOnlyWhenMini = True
    
    If IsAlwaysTrayIcon Then
       If Not TrayIcon.IsDisplayed Then
              TrayIcon.Show Icon, SoftwareName
       End If
    ElseIf IsNeverTrayIcon Then
        If TrayIcon.IsDisplayed Then TrayIcon.Hide
    ElseIf IsTrayIconOnlyWhenMini Then
        If TrayIcon.IsDisplayed Then
           If IsWindowVisible(hwnd) <> 0 Then TrayIcon.Hide
        End If
    End If
    
    TrayIcon.TrackPopMenu = True
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub SortListsFromini()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
'//This Function used to sort the list at software startup
    
    Dim ColumnIndex As Integer
    Dim SortOrder As Integer
     
    SortOrder = GetINIProfileInt("sotord", 0)
    ColumnIndex = GetINIProfileInt("colind", 1)
    
    If SortOrder = 0 Then
       
       Call SortColumn(lstStartupPrograms, ColumnIndex, sortAscending, sortAlpha)

    Else
       
       Call SortColumn(lstStartupPrograms, ColumnIndex, sortDescending, sortAlpha)
    
    End If
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Resize()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     
    On Error Resume Next
    '//Auto Size The Form and listview'''''''''''''''''''''''''''''''''''''''''''''''
    lstStartupPrograms.Height = frmmain.Height - 1425
    shpMain.Height = lstStartupPrograms.Top + lstStartupPrograms.Height
    
    shpMain.Width = frmmain.Width - 345
    lstStartupPrograms.Width = frmmain.Width - 585
    On Error GoTo 0
    
   ' lstStartupPrograms.ColumnHeaders.Item(1).Width = 0.25 * lstStartupPrograms.Width
   ' lstStartupPrograms.ColumnHeaders.Item(2).Width = 0.38 * lstStartupPrograms.Width
   ' lstStartupPrograms.ColumnHeaders.Item(3).Width = 0.25 * lstStartupPrograms.Width
   ' lstStartupPrograms.ColumnHeaders.Item(4).Width = 0.65 * lstStartupPrograms.Width
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim res As Integer
    
    res = GetINIProfileInt("optmini", 1)
    
    If res Then
    
       If Me.WindowState = 1 Then
          
          If Not TrayIcon.IsDisplayed Then TrayIcon.Show Icon, SoftwareName
          
          Me.Hide
          
       ElseIf Me.WindowState = 0 Then
          
          If TrayIcon.IsDisplayed Then TrayIcon.Hide ': Me.Show
       
       End If
        
    End If
    
    '//if the icon is displayed and optmini is unselected
    If TrayIcon.IsDisplayed And res = 0 Then
       
       If Me.WindowState = 1 Then
                 
             Me.Hide
       
       End If
    
    End If
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Unload(Cancel As Integer)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    TrayIcon.Hide
    
    WriteToINIProfile "C1W", lstStartupPrograms.ColumnHeaders(1).Width
    WriteToINIProfile "C2W", lstStartupPrograms.ColumnHeaders(2).Width
    WriteToINIProfile "C3W", lstStartupPrograms.ColumnHeaders(3).Width
    WriteToINIProfile "C4W", lstStartupPrograms.ColumnHeaders(4).Width
    
    WriteToINIProfile "W", frmmain.Width
    WriteToINIProfile "H", frmmain.Height
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewAdd_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    '//Disable CheckList function to avoid this error "Invalid procedure call or argument"
    frmmain.tmrGenChecks.Enabled = False
    frmmain.tmrEnabletmrGenChecks.Enabled = True
       
    Dim res As Integer
    
    res = GetINIProfileInt("chkAllowMonitor", 1)
    
    If res Then rMonitorAllowed = False
        
    frmAdd.Show 1, frmmain
    
    If res Then rMonitorAllowed = True
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewBuy_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    ShellExecute Me.hwnd, vbNullString, OrderPage, vbNullString, "", SW_SHOWNORMAL

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewHelp_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    'Dim DaysLeft As String
    
    'DaysLeft = GetDaysLeft
        
    'If DaysLeft = "" Then
    '   Me.Caption = App.Title
    '   JewBuy.Visible = False
    '   mnuEnterRegistrationKey.Visible = False
    'End If
    
    PopupMenu mnuHelpButton, 0, JewHelp.Left, JewHelp.Top + JewHelp.Height, mnuContents
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewRemove_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Dim hKey As Long
Dim I As Integer
Dim strPath As String
Dim Location As String
Dim ProgramPath As String
        
        Dim cnt As Boolean
        
        cnt = True
           
        '//Disable the monitor until i get finish this to avoid displaying the balloon
    
        Dim res As Integer
        res = GetINIProfileInt("chkAllowMonitor", 1)
        If res Then rMonitorAllowed = False
    
        Do While cnt
        
                I = I + 1
                
                If I = lstStartupPrograms.ListItems.Count Then
                      
                   If lstStartupPrograms.ListItems(lstStartupPrograms.ListItems.Count).Checked Then
                      RemoveFromStartup (I)
                   End If
                   
                   Exit Do
                
                End If
                   
                If lstStartupPrograms.ListItems(I).Checked Then
                   
                   RemoveFromStartup (I)
                   
                   I = I - 1
                   
                End If
        Loop
        
        '//Enable the monitor again to display the balloons again
        If res Then rMonitorAllowed = True
        
'    For I = 1 To lstStartupPrograms.ListItems.Count
'
'        '//if you choose to delete any item or a bulk of items (rather than the last one)
'        If I > lstStartupPrograms.ListItems.Count Then
'           Exit For '//then the count decrease from the one the loop started with
'           '//and I is still count to reach that one therefore we must check if
'           '//I > count to exit for in order to avoid "Index out of bound error"
'        End If
'        If lstStartupPrograms.ListItems(I).Checked Then
'
'           RemoveFromStartup (I)
'
'       End If
'   Next

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewRun_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Dim I As Integer
Dim res As Double
Dim ProgramPath As String

    For I = 1 To lstStartupPrograms.ListItems.Count
        If lstStartupPrograms.ListItems(I).Checked Then
           ProgramPath = ExtractPath(lstStartupPrograms.ListItems(I).SubItems(1))
           On Error Resume Next
           If Not IsProcessRunning(ExtractFileNameFromPath(ProgramPath)) Then
              res = Shell(ProgramPath)
              If res = 0 Then
                 MsgBox "Cannot run the selected application", vbExclamation, "Error"
              End If
           End If
        End If
    Next
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewTerminate_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Dim I As Integer
    Dim message As String
    Dim res As VbMsgBoxResult
    Dim ProgramPath As String
    Dim ProgramName As String
    Dim IsWarningMessageDisplay As Boolean
    
    message = "WARNING: Terminating a process can cause undesired" & Chr(13) & Chr(10) & "results including loss of data and system instability. The" & Chr(13) & Chr(10) & "process will not be given the chance to save its state or" & Chr(13) & Chr(10) & "data before its terminated." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Are you sure you want to terminate the process??"
    
    For I = 1 To lstStartupPrograms.ListItems.Count
        If lstStartupPrograms.ListItems(I).Checked Then
           If Not IsWarningMessageDisplay Then res = MsgBox(message, vbExclamation + vbYesNo)
                IsWarningMessageDisplay = True
                If res = vbYes Then
                   ProgramPath = ExtractPath(lstStartupPrograms.ListItems(I).SubItems(1))
                   ProgramName = ExtractFileNameFromPath(ProgramPath)
                   If IsProcessRunning(ProgramName) Then
                      Killapp (ExtractFileNameFromPath(ProgramPath))
                   End If
                End If
        End If
    Next

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub lstStartupPrograms_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     
     Static SortSwitch As Boolean
     
     If Not SortSwitch Then
        
        Call SortColumn(lstStartupPrograms, ColumnHeader.Index, sortAscending, sortAlpha)
        
        WriteToINIProfile "sotord", sortAscending
        WriteToINIProfile "colind", ColumnHeader.Index
        
        SortSwitch = True
     
     ElseIf SortSwitch Then
        
        Call SortColumn(lstStartupPrograms, ColumnHeader.Index, sortDescending, sortAlpha)
        
        WriteToINIProfile "sotord", sortDescending
        WriteToINIProfile "colind", ColumnHeader.Index

        SortSwitch = False
     
     End If
       
  
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub lstStartupPrograms_ItemClick(ByVal Item As MSComctlLib.ListItem)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Dim ProgramPath As String
    Dim ProgramName As String
    
    If GetAsyncKeyState(VK_RBUTTON) Then
       
       lstStartupPrograms.SelectedItem = Item
                     
       ProgramPath = ExtractPath(Item.SubItems(1))
        
       ProgramName = ExtractFileNameFromPath(ProgramPath)
                     
       If IsProcessRunning(ProgramName) Then
           mnuRun.Caption = "Terminate"
       Else
           mnuRun.Caption = "Run"
       End If
       
       PopupMenu menupopup
    
    End If
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuAbout_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    frmabout.Show 1
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuCopyFilePath_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Dim ProgramPath As String
    
    ProgramPath = ExtractPath(lstStartupPrograms.SelectedItem.SubItems(1))
    
    Clipboard.Clear
    
    Clipboard.SetText ProgramPath
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuEnable_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If mnuEnable.Caption = "Enable Startup Monitor" Then
       
       rMonitorAllowed = True
       mnuEnable.Caption = "Disable Startup Monitor"
       
    ElseIf mnuEnable.Caption = "Disable Startup Monitor" Then
       
       rMonitorAllowed = False
       mnuEnable.Caption = "Enable Startup Monitor"
    
    End If
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuHelp_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
   
   Shell "hh " & AppPath & "help.chm", vbMaximizedFocus

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuContents_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
   
   Shell "hh " & AppPath & "help.chm", vbMaximizedFocus

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuAboutFromHelpButton_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    frmabout.Show 1

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuEnterRegistrationKey_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    ShowEnterKeyDialog Me.hwnd
    
    '//pt 08Mar2008: Hide "Enter Regsiterision Key..." if the key is enterd
    Dim KeyCode As String
    KeyCode = ArmEnviron$("USERKEY")
    If KeyCode <> "" Then
       mnuEnterRegistrationKey.Visible = False
       WriteToINIProfile "cor", 1
       JewBuy.Visible = False
    End If

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnushow_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    If IsWindowVisible(hwnd) = 0 Then
       ShowWindow hwnd, SW_RESTORE
    Else
       ShowWindow hwnd, SW_HIDE
    End If

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuWebSiteFromHelp_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

     ShellExecute Me.hwnd, vbNullString, SiteName, vbNullString, "", SW_SHOWNORMAL

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuPurchase_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    ShellExecute Me.hwnd, vbNullString, OrderPage, vbNullString, "", SW_SHOWNORMAL

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuWebSite_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    ShellExecute Me.hwnd, vbNullString, SiteName, vbNullString, "", SW_SHOWNORMAL

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuRun_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     
    Dim res As Double
    Dim ProgramPath As String
    
    Select Case mnuRun.Caption
           
           Case "Run"
                ProgramPath = ExtractPath(lstStartupPrograms.SelectedItem.SubItems(1))
                On Error Resume Next
                res = Shell(ProgramPath)
                If res = 0 Then
                   MsgBox "Cannot run the selected application", vbExclamation, "Error"
                End If
                
           Case "Terminate"
                TerminateApplication (lstStartupPrograms.SelectedItem.Index)
                
    End Select
   
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuRemove_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    RemoveFromStartup lstStartupPrograms.SelectedItem.Index
      
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuOpenFileLocation_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Dim ProgramPath As String
    Dim FileLocationPath As String
    
    ProgramPath = ExtractPath(lstStartupPrograms.SelectedItem.SubItems(1))
    
    FileLocationPath = ExtractFileLocationFromPath(ProgramPath)
    
    Shell "explorer " & FileLocationPath, vbMaximizedFocus
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub tmrEnabletmrGenChecks_Timer()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    '//The interval of this timer is 5000 == balloon timeout period in "ShowRemovedAlert" function.
    '//We Enable this timer whenever ShowRemovedAlert called and HSM systray icon isn't displayed
    '//so we use this timer to enable tmrGenChecks after 5000 to remove the systray icon again
    
    tmrGenChecks.Enabled = True
    tmrEnabletmrGenChecks.Enabled = False

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub tmrGenChecks_Timer()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     
    If mnuEnable.Caption = "Enable Startup Monitor" Or mnuEnable.Caption = "Disable Startup Monitor" Then
       If rMonitorAllowed Then
           mnuEnable.Caption = "Disable Startup Monitor"
       Else
           mnuEnable.Caption = "Enable Startup Monitor"
       End If
    End If
        
    If IsWindowVisible(hwnd) Then
       If TrayIcon.IsDisplayed And mnushow.Caption <> "Hide Startup Monitor" Then
          mnushow.Caption = "Hide Startup Monitor"
       End If
    Else
       If TrayIcon.IsDisplayed Then
          If TrayIcon.IsDisplayed And mnushow.Caption <> "Show Startup Monitor" Then
             mnushow.Caption = "Show Startup Monitor"
          End If
       End If
    End If
    
    If IsAlwaysTrayIcon Then
      If Not TrayIcon.IsDisplayed Then
              TrayIcon.Show Icon, SoftwareName
         End If
     ElseIf IsNeverTrayIcon Then
        If TrayIcon.IsDisplayed Then TrayIcon.Hide
     ElseIf IsTrayIconOnlyWhenMini Then
        If TrayIcon.IsDisplayed Then
           If IsWindowVisible(hwnd) <> 0 Then TrayIcon.Hide
        End If
     End If
             
     CheckList
     
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub TrayIcon_LeftMouseDBLCLK()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
       
    If IsWindowVisible(hwnd) = 0 Then
       '//using ShowWindow solves this problem
       ShowWindow hwnd, SW_RESTORE
    End If

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub TrayIcon_RightMouseUp()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Me.PopupMenu mnuSystray
     
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewOptions_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    frmOptions.Show 1, frmmain

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub JewRefresh_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    CheckList

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

   res = GetINIProfileInt("chkMinInsteadOfClosing", 0)

   If res Then

      Cancel = 1

      If Not TrayIcon.IsDisplayed Then TrayIcon.Show Icon, SoftwareName

      Me.Hide

   End If

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuExit_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    tmrGenChecks.Enabled = False
    
    ExitMeNow = True
    
    TrayIcon.Hide
    
    Unload frmOptions
    
    Unload Me
    
End Sub
