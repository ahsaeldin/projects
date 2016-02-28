VERSION 5.00
Object = "{30E5838E-8E26-11DA-A1C2-807A50C10000}#1.0#0"; "acomte445.ocx"
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "jbet33.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ILCustomIcons 
      Left            =   6600
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0E42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   7800
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   4905
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Easy Text Inserter is enabled."
            TextSave        =   "Easy Text Inserter is enabled."
            Key             =   "a1"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrCheckKeyboard 
      Interval        =   1
      Left            =   7320
      Top             =   240
   End
   Begin VB.TextBox txtPharse 
      Height          =   2295
      Left            =   3720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   1440
      Width           =   4455
   End
   Begin VB.TextBox txtKeyword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   600
      Width           =   2415
   End
   Begin ASTC.TrayIcon TrayIcon 
      Left            =   8280
      Top             =   240
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin MSComctlLib.TreeView trvTexts 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   7858
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ILTreeView"
      Appearance      =   1
   End
   Begin JwldButn2b.JeweledButton JewKeyNew 
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   4320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Ne&w"
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
   Begin JwldButn2b.JeweledButton JewKeyDelete 
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   4320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Dele&te"
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
   Begin JwldButn2b.JeweledButton JewFoldRename 
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Rename"
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
   Begin JwldButn2b.JeweledButton JewFoldNew 
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&New"
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
   Begin JwldButn2b.JeweledButton JewFoldDelete 
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   3840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Delete"
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
   Begin JwldButn2b.JeweledButton JewKeyEdit 
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Edit"
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
   Begin JwldButn2b.JeweledButton JewEnable 
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Di&sable"
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
   Begin JwldButn2b.JeweledButton JewOptions 
      Height          =   375
      Left            =   8280
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
      _ExtentX        =   1931
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
   Begin JwldButn2b.JeweledButton JewHelp 
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "A&bout"
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
   Begin JwldButn2b.JeweledButton Jewimport 
      Height          =   375
      Left            =   8280
      TabIndex        =   16
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "I&mport"
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
   Begin JwldButn2b.JeweledButton JewExport 
      Height          =   375
      Left            =   8280
      TabIndex        =   17
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "E&xport"
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
   Begin MSComctlLib.ImageList ILTreeView 
      Left            =   8880
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1C94
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1DA8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblphrase 
      BackStyle       =   0  'Transparent
      Caption         =   "Text to be inserted"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   15
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblkeyword 
      BackStyle       =   0  'Transparent
      Caption         =   "Keyword"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   14
      Top             =   240
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000003&
      Height          =   4695
      Left            =   120
      Top             =   120
      Width           =   9375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Keywords"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Folders"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   3960
      Width           =   975
   End
   Begin VB.Menu mnuHelpButton 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu mnuContents 
         Caption         =   "Contents"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWebSiteFromHelp 
         Caption         =   "Visit Website"
      End
      Begin VB.Menu mnuAboutFromHelpButton 
         Caption         =   "About..."
      End
   End
   Begin VB.Menu mnuSystray 
      Caption         =   "Systray"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Hide Easy Text Inserter"
      End
      Begin VB.Menu mnuEnable 
         Caption         =   "Disable"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuFolder 
      Caption         =   "New"
      Visible         =   0   'False
      Begin VB.Menu mnuNewFolder 
         Caption         =   "New Folder"
      End
      Begin VB.Menu mnuRenameFolder 
         Caption         =   "Rename"
      End
      Begin VB.Menu mnuDeleteFolder 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnuKeyword 
      Caption         =   "New Keyword"
      Visible         =   0   'False
      Begin VB.Menu mnuNewKeyword 
         Caption         =   "New Keyword"
      End
      Begin VB.Menu mnuEditKeyword 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuDeleteKeyword 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
                                '         ;M";::;;
                                '        ,':;: ""'.
                                '       ,M;. ;MM;;M:
                                '       ;MMM::MMMMM:
                                '      ,MMMMM'MMMMM:
                                '      ;MMMMM MMMMMM
                                '      MMMMM::MMMMMM:
                                '      :MM:',;MMMMMM'
                                '      ':: 'MMMMMMM:
                                '        '; :MMMMM"
                                '           ''"""'
                                '            .
                                '            M
                                '            M
                                '.           M           .
                                ''M..        M        ,;M'
                                ' 'MM;.      M       ;MM:
                                '  :MMM.     M      ;MM:
                                '  'MMM;     M     :MMM:
                                '   MMMM.    M     MMMM:
                                '  : MMMM:   M     MMMM:
                                '  : MMMM:   M    :MMMM:
                                '  :MMMMM    M    ;MMMM:
                                '  'MMMMM;   M   ,MMMMM:
                                '   :MMMMM.  M   ;MMMMM'
                                '    :MMMM;  M  :MMMMM"
                                '     'MMMM  M  ;MMMM"
                                '     ':MM  M ,MMM:'
                                '        "": M :""'
        
           
           '// «·⁄ﬁ»«  ÂÌ  ·ﬂ «·√‘Ì«¡ «·„ŒÌ›… «· Ì  —«Â« ⁄‰œ„«  ÕÊ· »’—ﬂ ⁄‰ «·Âœ› - Â‰—Ì ›Ê—œ

Dim KeysSender As New CKeys

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Load()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    '//want to finish Mext after 9 days by 10 June 2008 isA
    
    If App.PrevInstance = True Then End
    
    JewHelp.Caption = "A&bout"
    JewKeyNew.Caption = "Ne&w"
    JewFoldNew.Caption = "&New"
    JewKeyEdit.Caption = "&Edit"
    JewEnable.Caption = "Di&sable"
    JewOptions.Caption = "&Options"
    JewKeyDelete.Caption = "Dele&te"
    JewFoldRename.Caption = "&Rename"
    JewFoldDelete.Caption = "&Delete"
    JewExport.Caption = "E&xport"
    Jewimport.Caption = "I&mport"
    '////////////////////////////////////////////////////////////////////////////////////////
    Dim DaysLeft As String

    DaysLeft = GetDaysLeft
      
    If DaysLeft = "" Then
       Me.Caption = App.Title
       'JewBuy.Visible = False
       'mnuEnterRegistrationKey.Visible = False
    Else
       Me.Caption = App.Title & " - Unregistered Trial Version For " & GetDaysLeft & " Days"
    End If
    '////////////////////////////////////////////////////////////////////////////////////////
    
    'Dim FadeSettings As Integer '//Buffer that will store a value from the ini that tell us if this form Faded before or not

    'FadeSettings = GetINIProfileInt("Fade", 0)  '//Check if this form Faded before or not
    
    'If FadeSettings = 0 Then FadeForm Me '//Fade this S if it doesn't faded before

    ShowSystrayIcon

    Dim FirstRun As Integer
    
    FirstRun = GetINIProfileInt("FR", 1)  '//Check if this form Faded before or not
    
    If FirstRun = 1 Then
        
       LoadDefaultFoldersandValues
       
    ElseIf FirstRun = 2 Then
        
        FillTreeView
    
    End If
    
    CheckOptions
    
    If Not IsMextEnabledByDefault Then
    
       IsMextEnable = False
    
    Else
       
       IsMextEnable = True
    
    End If
    
    ToggleSysTrayIcon
    
    '//if you don't enculde this line :
    '//1.if the user open the program then
    '//2.delete a folder just before doing anything, an error will be arrised coz no node selected to be deleted
    If trvTexts.Nodes.Count <> 0 Then trvTexts.Nodes.Item(1).Selected = True
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub CheckOptions()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Dim res As Integer
    
    res = GetINIProfileInt("chkenabled", 1)

    If res Then IsMextEnabledByDefault = True
   
    res = GetINIProfileInt("chkminimize", 0)

    If res Then IsMinmizedAtStartup = True
   
    res = GetINIProfileInt("chkshowballoon", 1)

    If res Then IsShowBalloonOption = True

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub LoadDefaultFoldersandValues()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    WriteToINIProfile "FR", 2
           
    AddFolderAtStartUp "My Keywords"
    
    AddNewKeywordValue "My Keywords", "tha", "Thank you very much for your time and consideration"
    AddNewKeywordValue "My Keywords", "dhcf", "don't hesitate to contact us if you have any further questions"
    AddNewKeywordValue "My Keywords", "pt", "http://www.perfectiontools.com" & vbNewLine
    
    trvTexts.Nodes.Item(1).Expanded = True
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub AddFolder(FolderName As Variant, Optional DonotSaveSetting As Boolean, Optional OverWrite As Boolean)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    '//why Optional DonotSaveSetting As Boolean)??
    '//to avoid saving settings in a functions like FillTreeView
    
    '//Note You can't add a duplicated Folder Name
    
    On Error GoTo error:
    
    FolderName = LCase(FolderName)
    
    '//Hesham told me that it is redicules that you can't add a folder name
    '//if there are a keyword with the same name exists or visa versa, hence
    '//i added "AddSpecialValueToKey" function to develop a way in which i could
    '//add an extra chars to the "KEY" of the new node in order to avoid the silly message "key must be unique"
    Dim Key As Variant
    Key = AddSpecialValueToFolderKey(FolderName)
    If IsEmpty(Key) Then Key = FolderName

    Dim i As Integer
    
    '//tried to handle the logic where the user add a folder which is already exists as a keyword i.e. folder name = th, keyword = th
    'If trvTexts.Nodes.Count >= 1 Then
    '
    '   For i = 1 To trvTexts.Nodes.Count
    '
    '       If FolderName = trvTexts.Nodes(i).Text Then
    '
    '          If trvTexts.Nodes(i).Tag <> "Folder" Then
    '
    '             trvTexts.Nodes.Add , , FolderName, FolderName, 1
    '
    '             trvTexts.Nodes.Item(trvTexts.Nodes.Count).Tag = "Folder"   '//used to disingueshed from keywords in JewKeyNew_Click and may be some other function
    '
    '             GoTo Proceed
    '
    '          End If
    '
    '       End If
    '
    '   Next i
    '
    'End If
     
    trvTexts.Nodes.Add , , Key, FolderName, 1
    trvTexts.Nodes.Item(trvTexts.Nodes.Count).Tag = "Folder"   '//used to disingueshed from keywords in JewKeyNew_Click and may be some other function
            
Proceed:

    '//"FN" means Folder Names : In This Section i will put all Folders Names as an index to use
    '//later in GetAllSettings Function in FillTreeView function
    If DonotSaveSetting Then Exit Sub
    
    SaveSetting App.EXEName, "FN", FolderName, FolderName
    
    Exit Sub

error:
    
    If Err.Number = 35602 And Not OverWrite Then
    
       MsgBox "A folder with the name you specified already exists. Specify a different name."
       
    End If
    
End Sub

'//Well, why i added this function??
'//I have a small problem when loading folders and keywords at program startup
'//if there is a keyword and folder with the same name, then the normal AddFolder function will
'//add the chars "*&#%*F" to the folder key, hence no keyword will be loaded ya man
'//that's why i cloned AddFolder function without calling AddSpecialValueToFolderKey function to avoid this problem
'//Note that i will call this function only from LoadDefaultFoldersandValues & FillTreeView
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub AddFolderAtStartUp(FolderName As Variant, Optional DonotSaveSetting As Boolean, Optional OverWrite As Boolean)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    '//why Optional DonotSaveSetting As Boolean)??
    '//to avoid saving settings in a functions like FillTreeView
    
    '//Note You can't add a duplicated Folder Name
    
    On Error GoTo error:
    
    FolderName = LCase(FolderName)
   
    Dim i As Integer
    
    '//tried to handle the logic where the user add a folder which is already exists as a keyword i.e. folder name = th, keyword = th
    'If trvTexts.Nodes.Count >= 1 Then
    '
    '   For i = 1 To trvTexts.Nodes.Count
    '
    '       If FolderName = trvTexts.Nodes(i).Text Then
    '
    '          If trvTexts.Nodes(i).Tag <> "Folder" Then
    '
    '             trvTexts.Nodes.Add , , FolderName, FolderName, 1
    '
    '             trvTexts.Nodes.Item(trvTexts.Nodes.Count).Tag = "Folder"   '//used to disingueshed from keywords in JewKeyNew_Click and may be some other function
    '
    '             GoTo Proceed
    '
    '          End If
    '
    '       End If
    '
    '   Next i
    '
    'End If
     
    trvTexts.Nodes.Add , , FolderName, FolderName, 1
    trvTexts.Nodes.Item(trvTexts.Nodes.Count).Tag = "Folder"   '//used to disingueshed from keywords in JewKeyNew_Click and may be some other function
            
Proceed:

    '//"FN" means Folder Names : In This Section i will put all Folders Names as an index to use
    '//later in GetAllSettings Function in FillTreeView function
    If DonotSaveSetting Then Exit Sub
    
    SaveSetting App.EXEName, "FN", FolderName, FolderName
    
    Exit Sub

error:
    
    If Err.Number = 35602 And Not OverWrite Then
    
       MsgBox "A folder with the name you specified already exists. Specify a different name."
       
    End If
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function EditFolderName(FolderName As Variant) As Boolean
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    '//why Optional DonotSaveSetting As Boolean)??
    '//to avoid saving settings in a functions like FillTreeView
    
    '//Note You can't add a duplicated Folder Name
    
    FolderName = LCase(FolderName)
    
    On Error GoTo error:
     
    '//Hesham told me that it is redicules that you can't add a folder name
    '//if there are a keyword with the same name exists or visa versa, hence
    '//i added "AddSpecialValueToKey" function to develop a way in which i could
    '//add an extra chars to the "KEY" of the new node in order to avoid the silly message "key must be unique"
    Dim Key As Variant
    Key = AddSpecialValueToFolderKey(FolderName)
    If IsEmpty(Key) Then Key = FolderName
    
    Dim i As Integer
    Dim KeyWords As Variant
    Dim OldFolderName As String
    Dim KeyWordValues As Variant
   
    If trvTexts.SelectedItem.Tag = "Folder" Then
       
       OldFolderName = trvTexts.SelectedItem.Text
       
       trvTexts.SelectedItem.Key = Key
       
       trvTexts.SelectedItem.Text = FolderName
       
    ElseIf trvTexts.SelectedItem.Parent.Tag = "Folder" Then
        
        OldFolderName = trvTexts.SelectedItem.Parent.Text
        
        trvTexts.SelectedItem.Parent.Key = Key
        
        trvTexts.SelectedItem.Parent.Text = FolderName
    
    End If
    
    DeleteSetting App.EXEName, "FN", OldFolderName '//Delete the old folder name from "FN" section
    
    SaveSetting App.EXEName, "FN", FolderName, FolderName  '//Save the New Folder name in "FN" section
    
    KeyWords = GetAllSettings(App.EXEName, OldFolderName)
    
    If IsEmpty(KeyWords) Then Exit Function '//exit sub to avoid error if no keywords in this folder
    
    For i = LBound(KeyWords, 1) To UBound(KeyWords, 1)
           
          SaveSetting App.EXEName, FolderName, KeyWords(i, 0), KeyWords(i, 1)
           
    Next i
     
    DeleteSetting App.EXEName, OldFolderName '//Delete the Whole Old Section since we already rename it
    
    EditFolderName = True
    
    Exit Function

error:
    
    If Err.Number = 35602 Then
    
       MsgBox "A folder with the name you specified already exists. Specify a different name."
       
       EditFolderName = False
       
    End If
    
End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub DeleteFolder(FolderName As Variant)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     
    '//why i added On Error Resume Next here
    '//because deleting empty folder causes error here if the folder has no section
    '//i think that adding an empty section in "AddFolder" method but i prefer not to do zt
    On Error Resume Next
    DeleteSetting App.EXEName, FolderName '//Delete The Whole Section
    On Error GoTo 0
    
    DeleteSetting App.EXEName, "FN", FolderName
    
    If trvTexts.SelectedItem.Tag = "Folder" Then
       
       trvTexts.Nodes.Remove (trvTexts.SelectedItem.Index)
       
    ElseIf trvTexts.SelectedItem.Parent.Tag = "Folder" Then
        
       trvTexts.Nodes.Remove (trvTexts.SelectedItem.Parent.Index)
    
    End If
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function AddNewKeywordValue(FolderName As Variant, KeywordName As Variant, KeywordValue As Variant, Optional DonotSaveSetting As Boolean, Optional OverWrite As Boolean) As Boolean
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
      
    '//why Optional DonotSaveSetting As Boolean)??
    '//to avoid saving settings in a functions like FillTreeView
    
    On Error GoTo error:
    
    KeywordName = LCase(KeywordName)
    FolderName = LCase(FolderName)
       
    '//Hesham told me that it is redicules that you can't add a folder name
    '//if there are a keyword with the same name exists or visa versa, hence
    '//i added "AddSpecialValueToKey" function to develop a way in which i could
    '//add an extra chars to the "KEY" of the new node in order to avoid the silly message "key must be unique"
    Dim Key As Variant
    Key = AddSpecialValueToKey(KeywordName)
    If IsEmpty(Key) Then Key = KeywordName
    
    '//Note You can't add a dublicated "Keyword Name"
    trvTexts.Nodes.Add FolderName, tvwChild, Key, KeywordName, 2
    
    If DonotSaveSetting Then Exit Function
    
    SaveSetting App.EXEName, FolderName, KeywordName, KeywordValue
    
    AddNewKeywordValue = True
    
error:
    
    If Err.Number = 35602 And Not OverWrite Then
    
       MsgBox "A keyword with the name you specified already exists. Specify a different keyword name."

       AddNewKeywordValue = False
       
    End If
    
End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function EditKeywordValue(Keyword As String, Value As String) As Boolean
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
          
    On Error GoTo error:
    
    Keyword = LCase(Keyword)
   ' FolderName = LCase(FolderName)
    
    '//Hesham told me that it is redicules that you can't add a folder name
    '//if there are a keyword with the same name exists or visa versa, hence
    '//i added "AddSpecialValueToKey" function to develop a way in which i could
    '//add an extra chars to the "KEY" of the new node in order to avoid the silly message "key must be unique"
    Dim Key As Variant
    Key = AddSpecialValueToKey(Keyword)
    If IsEmpty(Key) Then Key = Keyword
    
    If txtPharse = "" Then Exit Function
    
    If txtKeyword = "" Then Exit Function
   
    trvTexts.SelectedItem.Key = Key
    
    trvTexts.SelectedItem.Text = Keyword
      
    txtKeyword = Keyword
    
    frmmain.txtPharse = Value
    
    SaveSetting App.EXEName, trvTexts.SelectedItem.Parent.Text, Keyword, Value
    
    EditKeywordValue = True
    
error:

    If Err.Number = 35602 Then
    
       MsgBox "A keyword with the name you specified already exists. Specify a different keyword name."
           
       EditKeywordValue = False
       
    End If
    
End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub DeleteKeywordValue()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
          
    Dim Result As VbMsgBoxResult
    
    If txtPharse = "" Then Exit Sub
    
    If txtKeyword = "" Then Exit Sub
   
    Result = MsgBox("Are you sure you want to delete the " & txtKeyword & " keyword ?", vbYesNo + vbQuestion, SoftwareName)
    
    If Result = vbNo Then Exit Sub
    
    trvTexts.SelectedItem.Key = txtKeyword
    
    trvTexts.SelectedItem.Text = txtKeyword
      
    DeleteSetting App.EXEName, trvTexts.SelectedItem.Parent.Key, txtKeyword
      
    trvTexts.Nodes.Remove (trvTexts.SelectedItem.Index)
          
    txtKeyword = ""
    
    txtPharse = ""
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub FillTreeView()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
        
    Dim i As Integer
    Dim g As Integer
    Dim FoldersNames As Variant
    Dim FolderValues As Variant
    
    FoldersNames = GetAllSettings(App.EXEName, "FN")
     
    'If IsEmpty(FoldersNames) Then '// if "FN" isn't there for any reason then LoadDefaultFoldersandValues
    '
    '   LoadDefaultFoldersandValues
    '
    '   FoldersNames = GetAllSettings(App.EXEName, "FN")
    '
    'End If
    
    If IsEmpty(FoldersNames) Then Exit Sub '//if the user delete all Folders then exit sub to avid error in the next loop
    
    '//walk through Folder Names
    For i = LBound(FoldersNames, 1) To UBound(FoldersNames, 1)
      
      AddFolderAtStartUp FoldersNames(i, 1), True
    
      FolderValues = GetAllSettings(App.EXEName, FoldersNames(i, 1))
      
      g = 0
       
      If IsEmpty(FolderValues) Then GoTo nexto:
      
      For g = LBound(FolderValues, 1) To UBound(FolderValues, 1)
          '//FolderValues(g, 0) The Keyword Name
          '//FolderValues(g, 1) The Keyword Value
          AddNewKeywordValue FoldersNames(i, 1), FolderValues(g, 0), FolderValues(g, 1), True
      Next g
      
nexto:

    Next i
     
    trvTexts.Nodes.Item(1).Expanded = True

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function GetDaysLeft() As String
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Dim DaysLeft As String

  '  DaysLeft = ArmEnviron$("DAYSLEFT")
  '  GetDaysLeft = DaysLeft
    
End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Resize()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    If Me.WindowState = 1 Then
    
        If mnuShow.Caption = "Show " & SoftwareName Then
            
            mnuShow.Caption = "Hide " & SoftwareName
        
        ElseIf mnuShow.Caption = "Hide " & SoftwareName Then
            
            Me.Visible = False
            
            mnuShow.Caption = "Show " & SoftwareName
         
        End If
    
    End If

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Unload(Cancel As Integer)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    TrayIcon.Hide
    
    Set KeysSender = Nothing
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewEnable_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If JewEnable.Caption = "Di&sable" Then
       JewEnable.Caption = "En&able"
       '//tmrCheckKeyboard.Enabled = False
       IsMextEnable = False
       mnuEnable.Caption = "Enable"
    Else
       JewEnable.Caption = "Di&sable"
      '// tmrCheckKeyboard.Enabled = True
       IsMextEnable = True
       mnuEnable.Caption = "Disable"
    End If
    
    ToggleSysTrayIcon
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewExport_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    On Error GoTo quitme:
     
    With CommonDialog
         .Filter = "XML files (*.xml)|*.xml" '"Application files (*.exe)"
         .CancelError = True
         .flags = cdlOFNPathMustExist + cdlOFNOverwritePrompt '+ cdlOFNExplorer '+ cdlOFNAllowMultiselect
         'Note for the pervious line: we must add cdlOFNExplorer to the flags when use cdlOFNAllowMultiselect
         'otherwise we showopen will be the silly old ShowOpen Dialog Box.
         .FileName = ""
         .DialogTitle = "Save As"
         .ShowSave
    End With
    
    SaveTreeToXML CommonDialog.FileName

Exit Sub
quitme:

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewFoldDelete_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Dim FolderName As String
    Dim Result As VbMsgBoxResult
    
    If trvTexts.Nodes.Count = 0 Then Exit Sub
    
    If trvTexts.SelectedItem.Tag = "Folder" Then
       
       FolderName = trvTexts.SelectedItem.Text
        
    ElseIf trvTexts.SelectedItem.Parent.Tag = "Folder" Then
         
       FolderName = trvTexts.SelectedItem.Parent.Text
       
    End If
     
    Result = MsgBox("Are you sure you want to delete the " & FolderName & " folder ?", vbYesNo + vbQuestion, SoftwareName)
    
    If Result = vbNo Then Exit Sub
    
    DeleteFolder FolderName
    
    txtPharse = ""
    
    txtKeyword = ""
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewFoldNew_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    frmFolders.Show 1

End Sub

Private Sub JewFoldRename_Click()
    
    Load frmFolders
    
    frmFolders.JewSave.Tag = "Rename" '//Tell FrmFolders the mode is rename a folder
    
    frmFolders.Show 1
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewHelp_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    PopupMenu mnuHelpButton, 0, JewHelp.Left, JewHelp.Top + JewHelp.Height

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Jewimport_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If trvTexts.Nodes.Count > 0 Then
        
        Dim MsgRes As VbMsgBoxResult
        
        MsgRes = MsgBox("Are you sure you want to import? Some existing keywords may be overwritten.", vbQuestion + vbYesNo, "Confirm")
        
        If MsgRes = vbNo Then Exit Sub
    
    End If
    
    On Error GoTo quitme:
    
    With CommonDialog
         .Filter = "XML files (*.xml)|*.xml" '"Application files (*.exe)"
         .CancelError = True
         .flags = cdlOFNPathMustExist '+ cdlOFNExplorer '+ cdlOFNAllowMultiselect
         'Note for the pervious line: we must add cdlOFNExplorer to the flags when use cdlOFNAllowMultiselect
         'otherwise we showopen will be the silly old ShowOpen Dialog Box.
         .FileName = ""
         .DialogTitle = "Select File"
         .ShowOpen
    End With
        
    FromXMLtoTree CommonDialog.FileName
    
Exit Sub
quitme:

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewKeyDelete_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If txtKeyword <> "" And txtPharse <> "" Then
        
       If trvTexts.SelectedItem.Tag = "Folder" Then '//If the selected item is folder then
       
            '//MsgBox "as" messagebox here to tell the user that he didnot select a keyword to delete
          
       Else '//else if the selected item is a keyword get the parent
         
          DeleteKeywordValue
       
       End If
       
    End If

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewKeyEdit_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If txtKeyword = "" Then Exit Sub
    If txtPharse = "" Then Exit Sub
    
    Load frmKeyword
    
    frmKeyword.Caption = "Edit Keyword"
     
    frmKeyword.txtKeyword = txtKeyword
    
    frmKeyword.txtPharse = txtPharse
    
    frmKeyword.txtKeyword.SelLength = Len(txtKeyword)
    
    frmKeyword.Tag = "Edit"
    
    frmKeyword.Show 1
    
    'If txtKeyword <> "" And txtPharse <> "" Then
    '
    '   If trvTexts.SelectedItem.Tag = "Folder" Then '//If the selected item is folder then
    '
    '      '//MsgBox "as" messagebox here to tell the user that he didnot select a keyword to edit
    '
    '   Else '//else if the selected item is a keyword get the parent
    '
    '      EditKeywordValue
    '
    '   End If
    '
    'End If
         
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewKeyNew_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Load frmKeyword
    
    frmKeyword.Caption = "Add New Keyword"
     
    frmKeyword.Tag = "Add"
    
    frmKeyword.Show 1
    
  '  If txtKeyword <> "" And txtPharse <> "" Then
  '
  '     If trvTexts.SelectedItem.Tag = "Folder" Then '//If the selected item is folder then
  '
  '       AddNewKeywordValue trvTexts.SelectedItem.Key, txtKeyword, txtPharse
  '
  '     Else '//else if the selected item is a keyword get the parent
  '
  '       AddNewKeywordValue trvTexts.SelectedItem.Parent.Key, txtKeyword, txtPharse
  '
  '     End If
  '
  '  End If
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewOptions_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    frmOptions.Show 1

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuAboutFromHelpButton_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    frmabout.Show 1
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuContents_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Shell "hh " & AppPath & "help.chm", vbMaximizedFocus

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuEnable_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If mnuEnable.Caption = "Disable" Then
       JewEnable.Caption = "En&able"
       '//tmrCheckKeyboard.Enabled = False
       IsMextEnable = False
       mnuEnable.Caption = "Enable"
    Else
       JewEnable.Caption = "Di&sable"
       '//tmrCheckKeyboard.Enabled = True
       IsMextEnable = True
       mnuEnable.Caption = "Disable"
    End If
    
    ToggleSysTrayIcon
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuShow_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    If mnuShow.Caption = "Show " & SoftwareName Then
          
        Me.Visible = True
        
        Me.WindowState = 0
       
        mnuShow.Caption = "Hide " & SoftwareName
        
    ElseIf mnuShow.Caption = "Hide " & SoftwareName Then
         
         Me.Visible = False
     
         mnuShow.Caption = "Show " & SoftwareName
       
    End If
 
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuExit_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    TrayIcon.Hide
    
    Unload frmabout

    Unload Me

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub mnuWebSiteFromHelp_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
        
    ShellExecute Me.hwnd, vbNullString, SiteName, vbNullString, "", SW_SHOWNORMAL
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ShowSystrayIcon()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If Not TrayIcon.IsDisplayed Then
    
       TrayIcon.Show Icon, SoftwareName
    
    End If
    
    TrayIcon.TrackPopMenu = True
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub tmrCheckKeyboard_Timer()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Static IsStartedUp As Boolean
    
    If Not IsStartedUp Then
       If IsMinmizedAtStartup Then Me.Visible = False
       IsStartedUp = True
    End If
    
    If Not IsMextEnable Then GoTo CheckShortCut:
    
    Const KeyDown = -32767
    
    Dim i As Integer
        
    Dim KeyCode As Integer
    
    Dim PressedKeys As String
       
    For KeyCode = 0 To 255
        
        If GetAsyncKeyState(KeyCode) = KeyDown Then
           
           PressedKeys = LogKey(KeyCode)
        
           DoEvents
        
        End If
        
    Next

    For i = 1 To trvTexts.Nodes.Count
        '//If trvTexts.Nodes(I).Text = PressedKeys Then
        If InStr(1, PressedKeys, trvTexts.Nodes(i).Text) And PressedKeys <> "" Then
           Debug.Print trvTexts.Nodes(i).Text, PressedKeys
           If trvTexts.Nodes(i).Tag <> "Folder" Then '// only check with keywords texts not folder texts
              'comment old way and use KeysSender rather Than the crap SendKeys
              ''//SendKeys "{BACKSPACE " & Len(trvTexts.Nodes(I).Text) & "}"
              '//SendKeys GetSetting(App.EXEName, trvTexts.Nodes(I).Parent.Text, trvTexts.Nodes(I).Text)
              Dim KeyToBeSent As String
              KeyToBeSent = GetSetting(App.EXEName, trvTexts.Nodes(i).Parent.Text, trvTexts.Nodes(i).Text) '//note:  "Len(trvTexts.Nodes(I).Text)" means the backspace that we will send first
              KeysSender.SendKeys KeyToBeSent, Len(trvTexts.Nodes(i).Text)
              PressedKeys = ""
              Cache = ""
           End If
        End If
    Next
    
CheckShortCut:
    
    If IsMextEnable Then
          
       stbar.SimpleText = "Easy Text Inserter is enabled."
       JewEnable.Caption = "Di&sable"
       mnuEnable.Caption = "Disable"
    
    Else
       
       stbar.SimpleText = "Easy Text Inserter is disabled."
       JewEnable.Caption = "En&able"
       mnuEnable.Caption = "Enable"

    End If
    
    If trvTexts.Nodes.Count <= 1 Then
       JewExport.Enabled = False
    Else
       JewExport.Enabled = True
    End If
    
    CheckShortcutKeys
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub TrayIcon_LeftMouseDBLCLK()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    mnuShow_Click
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub TrayIcon_RightMouseUp()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Me.PopupMenu mnuSystray
     
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub trvTexts_AfterLabelEdit(Cancel As Integer, NewString As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Dim res As Boolean
        
    If NewString = "" Then
       Cancel = 1
       Exit Sub '//prevent empty strings
    End If
    
    If trvTexts.SelectedItem.Tag = "Folder" Then
       
       res = EditFolderName(NewString)
       
       If Not res Then Cancel = 1
       
    Else
    
       res = EditKeywordValue(NewString, txtPharse)
       
       If Not res Then Cancel = 1
       
       'trvTexts.SelectedItem.Key = NewString
       'trvTexts.SelectedItem.Text = NewString
       'txtKeyword = NewString
       'SaveSetting App.EXEName, trvTexts.SelectedItem.Parent.Text, txtKeyword, txtPharse
    
    End If
 
End Sub

Private Sub trvTexts_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyDelete Then
        
       If trvTexts.SelectedItem.Tag = "Folder" Then
            
          JewFoldDelete_Click
          
       Else
       
          JewKeyDelete_Click
       
       End If
    
    End If
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub trvTexts_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    'If Button = 2 Then
    '
    '   PopupMenu mnuKeyword
    '
    'End If

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub trvTexts_NodeClick(ByVal Node As MSComctlLib.Node)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    On Error GoTo error:

    txtKeyword = Node.Text
    
    txtPharse = GetSetting(App.EXEName, Node.Parent.Text, Node.Text, "")
    
Exit Sub

error:
     
    txtPharse = ""
    txtKeyword = ""
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub txtKeyword_KeyDown(KeyCode As Integer, Shift As Integer)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If KeyCode = 65 And GetAsyncKeyState(vbKeyControl) Then
       
       txtKeyword.SelStart = 0
       
       txtKeyword.SelLength = Len(txtKeyword)
       
    End If

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub txtPharse_KeyDown(KeyCode As Integer, Shift As Integer)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If KeyCode = 65 And GetAsyncKeyState(vbKeyControl) Then
        
       txtPharse.SelStart = 0
       
       txtPharse.SelLength = Len(txtPharse)
    
    End If
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub SaveTreeToXML(FilePath As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Dim xmldoc As DOMDocument30
    Dim aElement As IXMLDOMElement
    Dim cElement As IXMLDOMElement
    Dim dElement As IXMLDOMElement
    Dim eElement As IXMLDOMElement
    Dim fElement As IXMLDOMElement
    Dim RootElement As IXMLDOMElement
    Dim ProcInstr As IXMLDOMProcessingInstruction

    Set xmldoc = New DOMDocument30 'Creating DOM Document object
    'this adds the processing instruction, the first line in an XML document
    Set ProcInstr = xmldoc.createProcessingInstruction("xml", "version=""1.0""")
    xmldoc.appendChild ProcInstr
    Set RootElement = xmldoc.createElement("mext") 'Create the root element
    Set xmldoc.documentElement = RootElement
    '//Set comElement = xmldoc.createComment("My students who took web programming ")'Creating comment node
    
    Dim i As Integer
    Dim g As Integer
    Dim FoldersNames As Variant
    Dim FolderValues As Variant
    
    FoldersNames = GetAllSettings(App.EXEName, "FN")
     
    If IsEmpty(FoldersNames) Then Exit Sub '//if the user delete all Folders then exit sub to avid error in the next loop
    
    '//walk through Folder Names
    For i = LBound(FoldersNames, 1) To UBound(FoldersNames, 1)
      
      'add the comment node after the root
      'Create the Folder Name
      Dim FolderIndex As String
      FolderIndex = "folder" & i '//Folder1 for ex:
      Set aElement = xmldoc.createElement(FolderIndex)
      'add the Folder node to the root
      RootElement.appendChild aElement

      Set fElement = xmldoc.createElement("folderName") '//Create Folder Node
      fElement.nodeTypedValue = FoldersNames(i, 1)
      aElement.appendChild fElement
  
      FolderValues = GetAllSettings(App.EXEName, FoldersNames(i, 1))
      
      g = 0
       
      If IsEmpty(FolderValues) Then GoTo nexto:
      
      For g = LBound(FolderValues, 1) To UBound(FolderValues, 1)
          '//FolderValues(g, 0) The Keyword Name
          '//FolderValues(g, 1) The Keyword Value
           Set cElement = xmldoc.createElement("key")
           aElement.appendChild cElement
           Set dElement = xmldoc.createElement("name")
           dElement.nodeTypedValue = FolderValues(g, 0)
           cElement.appendChild dElement

           Set eElement = xmldoc.createElement("value")
           eElement.nodeTypedValue = FolderValues(g, 1)
           cElement.appendChild eElement
      Next g
      
nexto:

    Next i
    
    xmldoc.save FilePath

    Set xmldoc = Nothing
    Set RootElement = Nothing
    Set ProcInstr = Nothing
    
'    Dim i As Integer
'
'    For i = 1 To trvTexts.Nodes.Count
'
'        If trvTexts.Nodes.Item(i).Tag = "Folder" Then
'
'           'add the comment node after the root
'           'Create the Folder Name
'           Dim FolderIndex As String
'           FolderIndex = "folder" & i '//Folder1 for ex:
'           Set aElement = xmldoc.createElement(FolderIndex)
'           'add the Folder node to the root
'           RootElement.appendChild aElement
'
'           Set fElement = xmldoc.createElement("folderName") '//Create Folder Node
'           fElement.nodeTypedValue = trvTexts.Nodes.Item(i).Text
'           aElement.appendChild fElement
'
'        Else
'
'           Set cElement = xmldoc.createElement("key")
'           aElement.appendChild cElement
'           Set dElement = xmldoc.createElement("name")
'           dElement.nodeTypedValue = trvTexts.Nodes.Item(i).Text
'           cElement.appendChild dElement
'
'           Set eElement = xmldoc.createElement("value")
'           eElement.nodeTypedValue = GetSetting(App.EXEName, trvTexts.Nodes.Item(i).Parent.Text, trvTexts.Nodes.Item(i).Text, "")
'           cElement.appendChild eElement
'
'        End If
'
'    Next i
'

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub FromXMLtoTree(FilePath As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
     
   Dim FolderName As String
   Dim KeywordName As String
   Dim oNode As MSXML2.IXMLDOMNode
   Dim FolderNode As MSXML2.IXMLDOMNode
   Dim oNodeValue As MSXML2.IXMLDOMNode
   Dim oDom As New MSXML2.DOMDocument30
   Dim oElement As MSXML2.IXMLDOMElement
   Dim oNodeList As MSXML2.IXMLDOMNodeList
   
   oDom.Load FilePath
     
   Set oNodeList = oDom.selectNodes("//Mext/")
    
   For Each FolderNode In oNodeList '//loop through the nodelist
   
       If FolderNode.parentNode.nodeName = "mext" Then
       
          For Each oNode In FolderNode.childNodes
              
              For Each oNodeValue In oNode.childNodes
                  
                  Debug.Print FolderNode.nodeName, oNode.nodeName, oNodeValue.nodeName, oNodeValue.Text
                              
                  If oNode.nodeName = "folderName" Then
                     AddFolder oNodeValue.Text, , True
                     FolderName = oNodeValue.Text
                  ElseIf oNodeValue.nodeName = "name" Then
                     AddNewKeywordValue FolderName, oNodeValue.Text, "", , True
                     KeywordName = oNodeValue.Text
                  ElseIf oNodeValue.nodeName = "value" Then
                     SaveSetting App.EXEName, FolderName, KeywordName, oNodeValue.Text
                  End If
              
              Next
          
          Next
       
       End If

   Next
   
   Set oDom = Nothing
   Set oNodeList = Nothing
   Set oNode = Nothing

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub ToggleSysTrayIcon()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If IsMextEnable Then
    
       TrayIcon.ChangeIcon Me.Icon
    
    Else
       
       TrayIcon.ChangeIcon ILCustomIcons.ListImages(1).ExtractIcon
    
    End If
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function AddSpecialValueToKey(KeywordName As Variant) As Variant
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Dim i As Integer
    Dim Key As String
    Dim FoldersNames As Variant
   
    FoldersNames = GetAllSettings(App.EXEName, "FN")
       
    If IsEmpty(FoldersNames) Then GoTo nexto '//if the user delete all Folders then exit sub to avid error in the next loop
        
    '//walk through Folder Names
    For i = LBound(FoldersNames, 1) To UBound(FoldersNames, 1)
           
        If FoldersNames(i, 1) = KeywordName Then
           
           Key = KeywordName & "*&#%*k"
           
           AddSpecialValueToKey = Key
           
           Exit Function
           
        End If
          
     Next i
     
nexto:

End Function

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function AddSpecialValueToFolderKey(FolderName As Variant) As Variant
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    Dim i As Integer
    Dim g As Integer
    Dim FoldersNames As Variant
    Dim FolderValues As Variant
    
    FoldersNames = GetAllSettings(App.EXEName, "FN")
    
    If IsEmpty(FoldersNames) Then Exit Function '//if the user delete all Folders then exit sub to avid error in the next loop
    
    '//walk through Folder Names
    For i = LBound(FoldersNames, 1) To UBound(FoldersNames, 1)
      
      FolderValues = GetAllSettings(App.EXEName, FoldersNames(i, 1))
      
      If FolderName = FoldersNames(i, 1) Then
      
          Exit Function '//exit coz there is already a folder with that name
      
      End If
      
      g = 0
       
      If IsEmpty(FolderValues) Then GoTo nexto:
      
      For g = LBound(FolderValues, 1) To UBound(FolderValues, 1)
          '//FolderValues(g, 0) The Keyword Name
          '//FolderValues(g, 1) The Keyword Value
          
          If FolderValues(g, 0) = FolderName Then
          
             AddSpecialValueToFolderKey = FolderName & "*&#%*F"
          
          End If

      Next g
      
nexto:

    Next i

End Function

