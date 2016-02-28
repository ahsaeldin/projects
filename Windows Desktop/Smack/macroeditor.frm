VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScript 
   Caption         =   "Macro Editor"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   Icon            =   "macroeditor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "macroeditor.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H8000000E&
      Height          =   4935
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton cmdDown 
         Caption         =   "Dow&n"
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdDuplicate 
         Caption         =   "D&uplicate"
         Height          =   375
         Left            =   4560
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3600
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "&Up"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin MSComctlLib.ListView lstmain 
         Height          =   4215
         Left            =   0
         TabIndex        =   12
         Top             =   720
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   7435
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Help"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2000
      Begin MSComctlLib.ListView lstCommands 
         Height          =   4215
         Left            =   0
         TabIndex        =   5
         Top             =   720
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   7435
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   3351
         EndProperty
      End
      Begin VB.CommandButton cmdAddCommand 
         Caption         =   "&Add Command"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 
 lstCommands.ListItems.Add 1, "ie", "Open Web Site", , 1
 
End Sub
