VERSION 5.00
Begin VB.Form frmLanguage 
   Caption         =   "Choose Language"
   ClientHeight    =   1320
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   3480
   Icon            =   "frmLanguage.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "frmLanguage.frx":08CA
   ScaleHeight     =   1320
   ScaleWidth      =   3480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Default 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Defualt"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Manage your Recorded Macros"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Manage your Recorded Macros"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdSet 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Set"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Manage your Recorded Macros"
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstLanguage 
      Height          =   645
      ItemData        =   "frmLanguage.frx":655EC
      Left            =   120
      List            =   "frmLanguage.frx":65605
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "frmLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdCancel_Click()

Unload Me

End Sub

Private Sub cmdSet_Click()

Me.Hide

Unload Me

End Sub

Private Sub Default_Click()

Dim i As Integer

Dim where As String

Rem set the language to the default
Tags.IsDefault = True

Rem Delete any pre setting for any language as a default

For i = 0 To 6
     
     
    where = InStr(1, lstLanguage.List(i), "Default")
    
    If where > 0 Then
                     
       Rem here the language name return in listbox without the word Default
       lstLanguage.List(i) = Left(lstLanguage.List(i), Len(lstLanguage.List(i)) - 10)
            
    End If
    
Next i

Rem write the default lang
lstLanguage.List(lstLanguage.ListIndex) = lstLanguage.List(lstLanguage.ListIndex) & Space(3) & "Default"

End Sub

Private Sub Form_Load()
Dim LangNumber As Integer


Rem we have 7 language
Rem Index of Listbox start form 0 to 6

Rem Defualt Lang
LangNumber = Val(GetSetting(App.EXEName, "h", "la", 2))

Rem write the default lang
lstLanguage.List(LangNumber - 1) = lstLanguage.List(LangNumber - 1) & Space(3) & "Default"

End Sub


