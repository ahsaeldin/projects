VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "jbet33.ocx"
Begin VB.Form frmBackUpAll 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup All Macros"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmBackUpAll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   3585
   StartUpPosition =   2  'CenterScreen
   Begin JwldButn2b.JeweledButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Cancel"
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
      BackColor       =   12632256
   End
   Begin JwldButn2b.JeweledButton cmdOk 
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Ok"
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
      BackColor       =   12632256
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.OptionButton OptBundleAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bundle all macros in one file."
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.OptionButton OptSeparate 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmBackUpAll.frx":058A
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmBackUpAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    '//cprin 26Jul2006: Tell frmJMacro that User choosed to nothing
    Tags.BuckupAllType = ""
    'cprin 26Jul2006: Unload the form if Cancel
    Unload Me
End Sub

Private Sub cmdOk_Click()
    'cprin 26Jul2006: Unload the form if Ok
    Unload Me
End Sub

Private Sub Form_Load()
    'cprin 26Jul2006: added next code block
    'to ensure that jeweled buttons will remain
    'as i want.
    cmdOk.Caption = "&Ok"
    cmdCancel.Caption = "&Cancel"
     
    '//cprin 26Jul2006: check which default one selected
    If OptSeparate.Value = True Then Tags.BuckupAllType = "bundle"
    '//cprin 26Jul2006: check which default one selected
    If OptBundleAll.Value = True Then Tags.BuckupAllType = "bundle"
    
End Sub

Private Sub OptBundleAll_Click()
    '//cprin 26Jul2006: Tell frmJMacro that User choosed to bundle.
    Tags.BuckupAllType = "bundle"
End Sub

Private Sub OptSeparate_Click()
    '//cprin 26Jul2006: Tell frmJMacro that User choosed to separate.
    Tags.BuckupAllType = "separate"
End Sub
