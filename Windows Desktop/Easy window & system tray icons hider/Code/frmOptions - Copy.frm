VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   5580
      Index           =   0
      Left            =   225
      ScaleHeight     =   5580
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraGeneral 
         Height          =   2415
         Left            =   0
         TabIndex        =   4
         Top             =   2640
         Visible         =   0   'False
         Width           =   4575
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   975
            Left            =   840
            TabIndex        =   5
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame fraKeyboardShortcuts 
         Height          =   2385
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   5655
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   1095
            Left            =   1200
            TabIndex        =   3
            Top             =   360
            Width           =   2775
         End
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   6645
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   11721
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Keyboard Shortcuts"
            Key             =   "KeyboardShortcuts"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
