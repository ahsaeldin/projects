VERSION 5.00
Begin VB.Form frmEnterMacroName 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save New Macro"
   ClientHeight    =   4470
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3825
   ControlBox      =   0   'False
   Icon            =   "frmEnterMacroName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmEnterMacroName.frx":030A
   ScaleHeight     =   4470
   ScaleWidth      =   3825
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ChkEnable 
      Caption         =   "S&hift"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   195
   End
   Begin VB.CheckBox chkShift 
      Caption         =   "S&hift"
      Enabled         =   0   'False
      Height          =   195
      Left            =   2280
      TabIndex        =   6
      Top             =   2880
      Width           =   200
   End
   Begin VB.CheckBox chkAlt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Alt"
      Enabled         =   0   'False
      Height          =   195
      Left            =   1440
      TabIndex        =   5
      Top             =   2880
      Width           =   200
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
      Left            =   600
      TabIndex        =   4
      Top             =   2880
      Value           =   1  'Checked
      Width           =   200
   End
   Begin VB.ComboBox comKeys 
      Enabled         =   0   'False
      Height          =   315
      Left            =   600
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtDescription 
      Height          =   855
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox txtMacroName 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFC0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Alt"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ctrl"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Use Shortcut Key to playback this Macro"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   4215
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Name for this Macro"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   2535
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
comKeys.AddItem "1"
comKeys.AddItem "2"
comKeys.AddItem "3"
comKeys.AddItem "4"
comKeys.AddItem "5"
comKeys.AddItem "6"
comKeys.AddItem "7"
comKeys.AddItem "8"
comKeys.AddItem "9"
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
comKeys.ListIndex = 0
End Sub

Private Sub Label2_Click()

If ChkEnable.Value = 1 Then ChkEnable.Value = 0: Exit Sub

If ChkEnable.Value = 0 Then ChkEnable.Value = 1

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

   projectpath = projectpath & "mac\" & txtMacroName & ".dec"
        
   MacDate = Date
   
   MacTime = time
      
   On Error Resume Next
   
          MkDir AppPath & "mac"
          
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
     
   Open projectpath For Output As #8
     
        Write #8, MacTime, MacDate, txtDescription
       
   Close #8
    
   Me.Tag = "ok"
      
   Me.Hide
  
End If

Exit Sub

AlreadyChoosed:


frmEnterMacroName.caption = "Name Already Exist... "
      
txtMacroName.SelStart = 0

txtMacroName.SelLength = Len(txtMacroName)

MsgBox "Name Already Exist... ", vbCritical, App.Title

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
