VERSION 5.00
Begin VB.Form frmRenameMacro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rename Macro"
   ClientHeight    =   2550
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2925
   Icon            =   "frmRenameMacro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmRenameMacro.frx":08CA
   ScaleHeight     =   2550
   ScaleWidth      =   2925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDescription 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox txtMacroName 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "New Macro Name"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmRenameMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()

Me.Tag = "cancel"

Me.Hide

End Sub

Private Sub OKButton_Click()

Dim projectpath
Dim i As Integer
Dim MacDate As String
Dim MacTime As String

If txtMacroName = "" Then
   
   MsgBox "Enter a name for your Macro or click cancel", vbCritical, "Error"

Else
                  
   For i = 0 To frmJMacro.CobMacros.ListCount - 1
   
       If txtMacroName = frmJMacro.CobMacros.List(i) Then
          
          GoTo AlreadyChoosed:
       
       End If
   
   Next i
         
         
   If txtDescription.Text <> "" Then txtDescription.Tag = "UserUpdateDecription"
   
   Me.Tag = "ok"
   
   Me.Hide

End If

Exit Sub

AlreadyChoosed:


frmRenameMacro.caption = "Name Already Exist... "
      
txtMacroName.SelStart = 0

txtMacroName.SelLength = Len(txtMacroName)

MsgBox "Name Already Exist... ", vbCritical, App.Title

End Sub
