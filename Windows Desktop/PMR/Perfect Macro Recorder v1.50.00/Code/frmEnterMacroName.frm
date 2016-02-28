VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "jbet33.ocx"
Begin VB.Form frmEnterMacroName 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save Macro"
   ClientHeight    =   3345
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4545
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtMacroName 
         Height          =   375
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtDescription 
         Height          =   375
         Left            =   1320
         MaxLength       =   36
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   2895
      End
      Begin VB.ComboBox comKeys 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox chkCtrl 
         BackColor       =   &H00FFFFFF&
         Caption         =   "C&trl"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Value           =   1  'Checked
         Width           =   200
      End
      Begin VB.CheckBox chkAlt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Alt"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1080
         TabIndex        =   4
         Top             =   1800
         Width           =   200
      End
      Begin VB.CheckBox chkShift 
         Caption         =   "S&hift"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1920
         TabIndex        =   5
         Top             =   1800
         Width           =   200
      End
      Begin JwldButn2b.JeweledButton OKButton 
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "&Save"
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
         Left            =   1320
         TabIndex        =   8
         Top             =   2640
         Width           =   1095
         _ExtentX        =   1931
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
      Begin VB.CheckBox ChkEnable 
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   230
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use Shortcut Key to play back this Macro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00E0E0E0&
         Height          =   975
         Left            =   120
         Top             =   1560
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Macro Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Alt"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   1800
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmEnterMacroName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()

Me.Tag = "cancel"

Me.Hide

End Sub

Private Sub CancelButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

CancelButton.Default = True

End Sub

Private Sub ChkEnable_Click()

If ChkEnable.Value = 1 Then

   chkAlt.Enabled = True
   
   chkCtrl.Enabled = True
   
   chkShift.Enabled = True
   
   comKeys.Enabled = True
   
 
Else
   

   chkAlt.Enabled = False
   
   chkCtrl.Enabled = False
   
   chkShift.Enabled = False
   
   comKeys.Enabled = False

End If

End Sub


Private Sub Form_Load()

'cprin 18July2006: added next code block
'to ensure that jeweled buttons will remain
'as i want.
''''''''''''''''''''''''''''''''''''''''''''
OKButton.Caption = "&Save"
CancelButton.Caption = "&Cancel"
''''''''''''''''''''''''''''''''''''''''''''
comKeys.AddItem "1"
comKeys.AddItem "2"
comKeys.AddItem "3"
comKeys.AddItem "4"
comKeys.AddItem "5"
comKeys.AddItem "6"
comKeys.AddItem "7"
comKeys.AddItem "8"
comKeys.AddItem "9"
comKeys.AddItem "F1"
comKeys.AddItem "F2"
comKeys.AddItem "F3"
comKeys.AddItem "F4"
comKeys.AddItem "F5"
comKeys.AddItem "F6"
comKeys.AddItem "F7"
comKeys.AddItem "F8"
comKeys.AddItem "F9"
comKeys.AddItem "F10"
comKeys.AddItem "F11"
comKeys.AddItem "F12"
comKeys.AddItem "a"
comKeys.AddItem "b"
comKeys.AddItem "c"
comKeys.AddItem "d"
comKeys.AddItem "e"
comKeys.AddItem "f"
comKeys.AddItem "g"
comKeys.AddItem "h"
comKeys.AddItem "i"
comKeys.AddItem "j"
comKeys.AddItem "k"
comKeys.AddItem "l"
comKeys.AddItem "m"
comKeys.AddItem "n"
comKeys.AddItem "o"
comKeys.AddItem "p"
comKeys.AddItem "q"
comKeys.AddItem "r"
comKeys.AddItem "s"
comKeys.AddItem "t"
comKeys.AddItem "u"
comKeys.AddItem "v"
comKeys.AddItem "w"
comKeys.AddItem "x"
comKeys.AddItem "y"
comKeys.AddItem "z"

comKeys.AddItem "Insert"

comKeys.AddItem "Delete"
comKeys.AddItem "`"
comKeys.AddItem "-"
comKeys.AddItem "="
comKeys.AddItem "Home"
comKeys.AddItem "Page Down"
comKeys.AddItem "Page Up"
comKeys.AddItem "Pause"
comKeys.AddItem "["
comKeys.AddItem "]"

comKeys.AddItem "\"
comKeys.AddItem ";"
comKeys.AddItem "'"
comKeys.AddItem "End"
comKeys.AddItem "\"
comKeys.AddItem ","
comKeys.AddItem "."

comKeys.AddItem "/"
comKeys.AddItem "Up"
comKeys.AddItem "Down"
comKeys.AddItem "Right"
comKeys.AddItem "Left"
comKeys.AddItem "Tab"
comKeys.AddItem "Num lock"
comKeys.AddItem "Num 0"
comKeys.AddItem "Num 1"
comKeys.AddItem "Num 2"
comKeys.AddItem "Num 3"
comKeys.AddItem "Num 4"
comKeys.AddItem "Num 5"
comKeys.AddItem "Num 6"
comKeys.AddItem "Num 7"
comKeys.AddItem "Num 8"
comKeys.AddItem "Num 9"
comKeys.AddItem "Num +"
comKeys.AddItem "Num -"
comKeys.AddItem "Num *"
comKeys.AddItem "Num /"
comKeys.AddItem "."

comKeys.ListIndex = 9
 
 
End Sub


Private Sub Label2_Click()

If ChkEnable.Value = 1 Then ChkEnable.Value = 0: Exit Sub

If ChkEnable.Value = 0 Then ChkEnable.Value = 1

End Sub

Private Sub Label3_Click()

If chkCtrl.Enabled = False Then Exit Sub:

If chkCtrl.Value = 1 Then chkCtrl.Value = 0: Exit Sub

If chkCtrl.Value = 0 Then chkCtrl.Value = 1

End Sub

Private Sub Label4_Click()

If chkAlt.Enabled = False Then Exit Sub

If chkAlt.Value = 1 Then chkAlt.Value = 0: Exit Sub

If chkAlt.Value = 0 Then chkAlt.Value = 1

End Sub

Private Sub Label5_Click()

If chkShift.Enabled = False Then Exit Sub:

If chkShift.Value = 1 Then chkShift.Value = 0: Exit Sub

If chkShift.Value = 0 Then chkShift.Value = 1

End Sub

Private Sub OKButton_Click()
Dim i As Long
Dim projectpath
Dim OF As OFSTRUCT
Dim MacDate As String
Dim MacTime As String
Dim SelectedShortKey As String
Dim IsAlreadyChoosed As Integer
Dim IsListedInComboBox As Boolean

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
       
   
   If ChkEnable.Value = 1 Then
      
      SelectedShortKey = GetSelectedKey(txtMacroName)
      
      If SelectedShortKey = "failed" Then Exit Sub
   
   Else
      
      SelectedShortKey = GetSelectedKey(txtMacroName, True)
   
   End If
   
   If right(App.Path, 1) <> "\" Then projectpath = App.Path + "\" Else projectpath = App.Path

   projectpath = projectpath & "Data\" & txtMacroName & ".ptf"
        
   MacDate = Date
   
   MacTime = time
      
   On Error Resume Next
   
          MkDir AppPath & "Data"
          
   On Error GoTo 0
                      
   For i = 0 To frmJMacro.CobMacros.ListCount - 1
   
       If txtMacroName = frmJMacro.CobMacros.List(i) Then
          
          IsListedInComboBox = True
          
         ' GoTo AlreadyChoosed:
       
       End If
   
   Next i

Rem i made a mistake that i check if the file name exist by
Rem checking a macro file with the same name , the true way is to
Rem check the frmJMacro.CobMacros.List, but beacuse i get tired to make
Rem the following code so i will not delete it.

   IsAlreadyChoosed = OpenFile(projectpath, OF, OF_READ)
   CloseHandle IsAlreadyChoosed
   
   If IsListedInComboBox And IsAlreadyChoosed = -1 Then
      'this case of the macro listed in combo box
      'but the macro file isn't in macro folder
      'may be deleted by a nag user
     ' MsgBox "sdfadsf"
   End If
   
   If IsListedInComboBox Then GoTo AlreadyChoosed:
     
OverWriteOldOne:
   
   Open projectpath For Output As #8
     
        Write #8, MacTime, MacDate, txtDescription
       
   Close #8
    
   Me.Tag = "ok"
      
   Me.Hide
  
End If

Exit Sub

AlreadyChoosed:

'/cprin 17July2006: changed next line to display a well form messagebox
'MsgBox "A macro with the name " & txtMacroName & " Already Exist, please choose another name.", vbCritical, App.Title
Dim res As Integer
res = MsgBox(txtMacroName & " macro already exists. do you want to overwrite it?", vbYesNo, App.Title)
If res = vbYes Then GoTo OverWriteOldOne:

frmEnterMacroName.Caption = "Name Already Exist... "
txtMacroName.SelStart = 0
txtMacroName.SelLength = Len(txtMacroName)

End Sub
Private Function GetSelectedKey(MacroName As String, Optional cancel As Boolean) As String
'============================================================

' Routine Name: GetSelectedKey
' Description:
' Author: Author's Name
' Date: 26/08/2004 07:06:52 Õ
' Copyright © 2004
' Notes:

' Modification History:

'============================================================

    On Error GoTo GetSelectedKeyErr
    Dim i As Long
    Dim Alt As String
    Dim Ctrl As String
    Dim Char As String
    Dim Shift As String
    Dim KeySelected As String
    Dim SavedMacros As Variant
    Dim SelectedMacroIndex As Long
    Dim SelectedMacroName As String
    
    If cancel Then GoTo DontAssignShortKey
    
    If chkCtrl.Value = 1 Then
       
       Ctrl = "Ctrl+"
       
    End If
    
    If chkAlt.Value = 1 Then
    
       Alt = "Alt+"
       
    End If
     
    If chkShift.Value = 1 Then
    
       Shift = "Shift+"
       
    End If
    
    
    KeySelected = Ctrl & Alt & Shift & comKeys.List(comKeys.ListIndex)

DontAssignShortKey:
    

Rem check if the key is realy selected
    
   SavedMacros = GetAllSettings(App.EXEName, "sm")
   
   For i = LBound(SavedMacros, 1) To UBound(SavedMacros, 1)
        
       If KeySelected = SavedMacros(i, 1) And Not KeySelected = "" Then
          
          MsgBox KeySelected & " is already selected for ( " & SavedMacros(i, 0) & " ) Macro", vbCritical, App.Title
               
          GetSelectedKey = "failed"
          
          Exit Function
         
       End If
      
   Next i
   
GetSelectedKeyExit:

   SaveSetting App.EXEName, "sm", MacroName, KeySelected

   GetSelectedKey = KeySelected
   
   

    On Error Resume Next

    Exit Function

GetSelectedKeyErr:

    Dim sMsg As String
    'Add your error display message here, or use this simple MsgBox display...

    With Err
        sMsg = "Error: " & .Number & vbCrLf
        sMsg = sMsg & "Description: " & .Description & vbCrLf
        sMsg = sMsg & "Code Location: " & App.Title & ":: EnterName_GetSelectedKey" & vbCrLf
        sMsg = sMsg & "Source: " & .Source & vbCrLf
        sMsg = sMsg & "Last DLL Error: " & .LastDllError & vbCrLf
        'MsgBox sMsg, vbOKOnly + vbCritical, "Error" & vbCrLf
    End With
GoTo GetSelectedKeyExit

End Function


Private Sub OKButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   OKButton.Default = True

End Sub

Private Sub txtDescription_Change()
txtDescription.MaxLength = 30
End Sub

Private Sub txtMacroName_Change()
txtMacroName.MaxLength = 22
End Sub
