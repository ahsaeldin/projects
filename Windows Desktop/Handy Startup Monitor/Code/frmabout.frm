VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "jbet33.ocx"
Begin VB.Form frmabout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7185
   Icon            =   "frmabout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Perfection Tools Software Homepage"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2040
         MouseIcon       =   "frmabout.frx":08CA
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Perfection Tools Software Support Email  Address"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1560
         MouseIcon       =   "frmabout.frx":0A1C
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   6120
         Picture         =   "frmabout.frx":0B6E
         Top             =   360
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   120
         Picture         =   "frmabout.frx":1A38
         Top             =   360
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Software Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   3675
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "© Copyright"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   1
         Top             =   600
         Width           =   5175
      End
   End
   Begin JwldButn2b.JeweledButton cmdOK 
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&OK"
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
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

            '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
            '“Success is not the key to happiness. Happiness is the key to success.'
            'If you love what you are doing, you will be successful.”              '
            '                                                          Herman Cain '
            '* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
                            ''''''''''''''''''''''''''''''''''''''
                            '___888888888888888__________________'
                            '__888888822222228888________________'
                            '_888888222I2222288888_______________'
                            '888888222222222228888822228888______'
                            '888882222222222222288222222222888___'
                            '88888222222222LUV22222222222222288__'
                            '_8888822222222222222222222222222_88_'
                            '__88888222222222222222U22222222__888'
                            '___888822222222222222222222222___888'
                            '____8888222222222222222222222____888'
                            '_____8888222222222222222222_____888_'
                            '______8882222222222222222_____8888__'
                            '_______888822222222222______888888__'
                            '________8888882222______88888888____'
                            '_________888888_____888888888_______'
                            '__________88888888888888____________'
                            '___________888888888________________'
                            '____________888888__________________'
                            '_____________8888___________________'
                            '______________88____________________'''''
                            '______________8 luv u more than anything'
                            ''''''''''''''''''''''''''''''''''''''''''

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Load()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
        
    Me.Caption = "About " & SoftwareName
    
    Label7.Caption = SoftwareName & " " & App.Major & "." & App.Minor & "." & App.Revision
    
    'Label8.Caption = "© Copyright 2006 - " & Year(Date) & ", Perfection Tools Software. All rights reserved."
  
    Label8.Caption = "© Copyright 2008, Perfection Tools Software. All rights reserved."

    cmdOK.Caption = "&OK"

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Label3_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    ShellExecute Me.hwnd, vbNullString, SiteName, vbNullString, "", SW_SHOWNORMAL

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Label5_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    ShellExecute Me.hwnd, vbNullString, "mailto:" & Email, vbNullString, "", SW_SHOWNORMAL

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Label9_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    ShellExecute Me.hwnd, vbNullString, "mailto:" & Email, vbNullString, "", SW_SHOWNORMAL

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub cmdOk_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
  
    Unload Me
  
End Sub
