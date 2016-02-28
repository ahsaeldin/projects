VERSION 5.00
Begin VB.Form frmtip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5850
   Icon            =   "frmtip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.CheckBox chkDontShowAgain 
      Caption         =   "Don't Show this again"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label lblTip 
      Caption         =   $"frmtip.frx":1782
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   5535
   End
End
Attribute VB_Name = "frmtip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '“The engineers that develop the front end tend to like to do it on ASP.Net'
            'but the back-end is usually on Java so we want to be pragmatic and not    '
            'religious.”                                                               '
            '                                                            Sacha Labourey'
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Load()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Dim DefChar As String
    
    Dim dkey As String * 15
    
    dkey = ""
        
    DefChar = Str(Asc("N"))
           
    GetINIProfileString Show_Me, "1,1,0#" & DefChar & "*", dkey, Len(dkey)
        
    Me.Caption = SoftwareName
        
    lblTip = "Press " & "( " & DecodeKey(dkey) & " )" & " to show " & SoftwareName
         
    '//added next line to check "IsTipWinodw" in "IsTask" to ignore and avoiding adding frmtip to the listview
    '//because if frmtip added to the listview it will be hided if the user press Hide All button
    SetProp Me.hwnd, "IsTipWindow", 111
     
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub cmdOk_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    If chkDontShowAgain.Value = 1 Then
       
       WriteToINIProfile "Tip1", 1
    
    End If
    
    Unload Me

End Sub

