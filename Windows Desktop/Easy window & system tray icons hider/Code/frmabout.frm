VERSION 5.00
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
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
         MouseIcon       =   "frmabout.frx":1782
         MousePointer    =   99  'Custom
         TabIndex        =   5
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
         MouseIcon       =   "frmabout.frx":18D4
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   6480
         Picture         =   "frmabout.frx":1A26
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   120
         Picture         =   "frmabout.frx":31A8
         Top             =   360
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
         Left            =   960
         TabIndex        =   2
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   5355
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
    
'    Label8.Caption = "© Copyright 2006 - " & Year(Date) & ", Perfection Tools Software. All rights reserved."
  
  Label8.Caption = "© Copyright 2007, Perfection Tools Software. All rights reserved."

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
