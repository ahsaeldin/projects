VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "jbet33.ocx"
Begin VB.Form frmtip 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tip"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8760
   ControlBox      =   0   'False
   Icon            =   "frmtip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin JwldButn2b.JeweledButton JewOK 
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   4200
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
   Begin VB.Image Image1 
      Height          =   3825
      Left            =   240
      Picture         =   "frmtip.frx":0E42
      Top             =   240
      Width           =   8310
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000003&
      Height          =   3975
      Left            =   120
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "frmtip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Load()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    JewOK.Caption = "&OK"

    Dim FirstRun As Integer
    
    FirstRun = GetINIProfileInt("FR", 1)  '//Check if this form Faded before or not
    
    If FirstRun = 1 Then
        
       Load frmmain
       
    ElseIf FirstRun = 2 Then
        
       frmmain.Show
    
       Unload Me
    
    End If
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewOK_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    frmmain.Show
    
    Unload Me

End Sub
