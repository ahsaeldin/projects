VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "jbet33.ocx"
Begin VB.Form frmRenameMacro 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rename Macro"
   ClientHeight    =   2025
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4305
   Icon            =   "frmRenameMacro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4305
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtMacroName 
         Height          =   375
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtDescription 
         Height          =   405
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   2895
      End
      Begin JwldButn2b.JeweledButton OKButton 
         Default         =   -1  'True
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
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
         BackColor       =   12632256
      End
      Begin JwldButn2b.JeweledButton CancelButton 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "New Name"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblDescription 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
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

Private Sub Form_Load()
'cprin 18July2006: added next code block
'to ensure that jeweled buttons will remain
'as i want.
''''''''''''''''''''''''''''''''''''''''''''
OKButton.Caption = "&OK"
CancelButton.Caption = "&Cancel"
''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub OKButton_Click()

Dim projectpath
Dim i As Integer
Dim MacDate As String
Dim MacTime As String

'//cprin 29Jul2006: added next if block to prevent user from Enter a macro name starts with number
'//because If the user backup a macro we will insert the len of Macro Name "NameLen" after cprinsosonono hence we want to check if the macro name len exceeds
'//Only one digits means exceeds 9 or 99 Hence We must prevent the macro name to start with numbers because later we will check in the import functions
'//[if the first char in the macro name from the left is nubmer then]
If IsNumeric(Left(txtMacroName, 1)) Then
   '//Tell the user that
   MsgBox "Macro Name can't Starts with numbers", vbCritical, SoftwareName
   '//Set the SelStart to the first char in txtMacroName
   txtMacroName.SelStart = 0
   '//Highlight the name entered by the user
   txtMacroName.SelLength = Len(txtMacroName)
   Exit Sub
End If

'//pt 5Sep2007:prevent these chars \ / : * ? " < > |
If InStr(1, txtMacroName, "\") Or InStr(1, txtMacroName, "/") Or InStr(1, txtMacroName, ":") Or InStr(1, txtMacroName, "*") Or InStr(1, txtMacroName, "?") Or InStr(1, txtMacroName, Chr(34)) Or InStr(1, txtMacroName, "<") Or InStr(1, txtMacroName, ">") Or InStr(1, txtMacroName, "|") Then
   '//Tell the user that
   MsgBox "A macro name cannot contain any of the following characters:" & Chr(13) & Chr(10) & "\ / : * ?  " & Chr(34) & " < > |", vbCritical, SoftwareName
   '//Set the SelStart to the first char in txtMacroName
   txtMacroName.SelStart = 0
   '//Highlight the name entered by the user
   txtMacroName.SelLength = Len(txtMacroName)
   Exit Sub
End If


'//cprin 31Jul2006: added next if block to prevent file names more than 200 char
If Len(txtMacroName) > 200 Then
   '//Tell the user that
   MsgBox "Macro Name can't exceeds 200 Char", vbCritical, SoftwareName
   '//Clear the text box
   txtMacroName = ""
   Exit Sub
End If


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


frmRenameMacro.Caption = "Name Already Exist... "
      
txtMacroName.SelStart = 0

txtMacroName.SelLength = Len(txtMacroName)

MsgBox "Name Already Exist... ", vbCritical, App.Title

End Sub

Private Sub txtDescription_Change()
txtDescription.MaxLength = 30
End Sub

Private Sub txtMacroName_Change()
txtMacroName.MaxLength = 22
End Sub
