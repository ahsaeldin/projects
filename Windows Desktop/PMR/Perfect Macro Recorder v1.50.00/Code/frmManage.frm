VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "jbet33.ocx"
Begin VB.Form frmManage 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Macros"
   ClientHeight    =   2610
   ClientLeft      =   555
   ClientTop       =   3525
   ClientWidth     =   6810
   Icon            =   "frmManage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6615
      Begin JwldButn2b.JeweledButton cmdshortkey 
         Height          =   375
         Left            =   5040
         TabIndex        =   3
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Shortcut key"
         PictureSize     =   0
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin JwldButn2b.JeweledButton cmdDeleteAll 
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Delete &All"
         PictureSize     =   0
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin JwldButn2b.JeweledButton cmdDeleteMacro 
         Height          =   375
         Left            =   5040
         TabIndex        =   5
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Delete"
         PictureSize     =   0
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin JwldButn2b.JeweledButton cmdRenameMacro 
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Rename"
         PictureSize     =   0
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin MSComctlLib.ListView lstviewMacros 
         Height          =   2055
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdRunStartup 
      BackColor       =   &H00FFFFFF&
      Caption         =   "St&artup Macro"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   0
      ToolTipText     =   "choose or unchoose  macro for running when Smart Macro starts"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer tmrBoldMe 
      Interval        =   100
      Left            =   8880
      Top             =   4440
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   5760
      Width           =   1575
   End
   Begin JwldButn2b.JeweledButton JeweledButton1 
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Edit Macro"
      PictureSize     =   0
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
End
Attribute VB_Name = "frmManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Rem Here we Fill the listview with Macros Data



Private Sub cmdDeleteAll_Click()
Dim i As Long



For i = 1 To lstviewMacros.ListItems.Count
    
    DeleteSelectedMacro
 
Next i

    

End Sub

Private Sub cmdDeleteAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    cmdDeleteAll.Default = True
    
End Sub

Private Sub cmdDeleteMacro_Click()
      
      DeleteSelectedMacro
      
End Sub

Private Sub cmdDeleteMacro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
                   
      cmdDeleteMacro.Default = True

End Sub

Private Sub cmdRenameMacro_Click()
        
        RenameMacroByButton
 
End Sub

Private Sub cmdRenameMacro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
       
       cmdRenameMacro.Default = True

End Sub

Private Sub cmdRunStartup_Click()

Dim o As ListItem
Dim SelectedMacroName As String
  

If cmdRunStartup.Caption = "Startup Macro" Then

   SelectedMacroName = GetSelectedMacro

   SaveSetting App.EXEName, "fr", "st", SelectedMacroName
   
   Set o = lstviewMacros.FindItem(SelectedMacroName, 0, 1, 0)
   
   o.SubItems(5) = "Run at StartUp"
   
ElseIf cmdRunStartup.Caption = "De-Startup" Then
   
   SelectedMacroName = GetSelectedMacro

   SaveSetting App.EXEName, "fr", "st", ""
   
   Set o = lstviewMacros.FindItem(SelectedMacroName, 0, 1, 0)
   
   o.SubItems(5) = ""
  
End If

End Sub

Private Sub cmdRunStartup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdRunStartup.Default = True

End Sub

Private Sub cmdshortkey_Click()
    
    SetShortKey

End Sub


Private Sub cmdshortkey_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    cmdshortkey.Default = True
    
End Sub

Private Sub Form_Load()

'cprin 18July2006: added next code block
'to ensure that jeweled buttons will remain
'as i want.
''''''''''''''''''''''''''''''''''''''''''''
cmdRenameMacro.Caption = "&Rename"
cmdDeleteMacro.Caption = "&Delete"
cmdDeleteAll.Caption = "Delete &All"
cmdshortkey.Caption = "&Shortcut key"
''''''''''''''''''''''''''''''''''''''''''''

Dim i
Dim o As ListItem
Dim MacDate As String
Dim MacTime As String
Dim projectpath As String
Dim MacDecription As String
Dim MacroShortKey As String
Dim MacroAtStartUp As String



With lstviewMacros

Rem     Macro Name Column

        .ColumnHeaders.Add , , "   Macro"
        
        .ColumnHeaders.Item(1).Width = 1500
         
Rem     Macro Description Column
         
        .ColumnHeaders.Add , , "                      Description"
        
        .ColumnHeaders.Item(2).Width = 3000
        
Rem     Macro Date Column
        
        .ColumnHeaders.Add , , "Creation Time"
        
        .ColumnHeaders.Item(3).Alignment = lvwColumnCenter
          
Rem     Macro Time Column
        
        .ColumnHeaders.Add , , "Creation Date"
        
        .ColumnHeaders.Item(4).Alignment = lvwColumnCenter

Rem     ShortKey Column

        .ColumnHeaders.Add , , "Shortcut Key"
                
        .ColumnHeaders.Item(5).Alignment = lvwColumnLeft
 
Rem     StartUp State
        
        .ColumnHeaders.Add , , "Notes"
                
        .ColumnHeaders.Item(6).Alignment = lvwColumnLeft
       
Rem     if no macro File Detected

        If frmJMacro.CobMacros.ListCount = 0 And frmManage.Visible = True Then
    
           MsgBox "You didn't have any recorded macros, to record new macro click Record New Macro", vbInformation, SoftwareName
           
        End If
        
Rem     Fill Loop
        
        For i = 0 To frmJMacro.CobMacros.ListCount - 1
                        
             
           
             .ListItems.Add , , frmJMacro.CobMacros.List(i)
        
Rem          Add Macro Name

             .ListItems.Item(i + 1).SubItems(1) = "Macro Deleted"
           
             If right(App.Path, 1) <> "\" Then projectpath = App.Path + "\" Else projectpath = App.Path

             projectpath = projectpath & "Data\" & frmJMacro.CobMacros.List(i) & ".ptf"
        
Rem          Get Decription,Time,Date from the .ptf File
              
             On Error GoTo error:

             Open projectpath For Input As #11
              
                   Input #11, MacTime, MacDate, MacDecription
 
             Close #11
                                            
             On Error Resume Next
             
             .ListItems.Item(i + 1).SubItems(1) = MacDecription
                
             .ListItems.Item(i + 1).SubItems(2) = MacTime
             
             .ListItems.Item(i + 1).SubItems(3) = MacDate
                 
              MacroShortKey = GetSetting(App.EXEName, "sm", frmJMacro.CobMacros.List(i), "")
              
             .ListItems.Item(i + 1).SubItems(4) = MacroShortKey
        
        Next i
        
        MacroAtStartUp = GetSetting(App.EXEName, "fr", "st", "nononomacros")
        
        If MacroAtStartUp <> "nononomacros" Then
                       
           On Error Resume Next
           
           Set o = lstviewMacros.FindItem(MacroAtStartUp, 0, 1, 0)
   
           o.SubItems(5) = "Run at StartUp"
        
        End If
        
End With

Exit Sub

error:

'//cprin 17July2006: changed next line to display a well form of MsgBox
'MsgBox "Data missed or manually deleted by the user."
MsgBox "Can't Found Macros Data Located in " & Chr(13) & Chr(10) & projectpath & "Data\"
End Sub

Private Sub lstviewMacros_AfterLabelEdit(cancel As Integer, NewString As String)
       
      Call RenameMacro(NewString)
       
End Sub


Private Sub RenameMacro(NewString As String, Optional byRenameButton As Boolean)
'============================================================

' Routine Name: RenameMacro
' Description: Rename the Macro and preform the realaed changes
' Author: Ahmed Saad
' Date: 23/08/2004 06:18:30 Õ
' Copyright © 2004 CprinGold Software
' Notes: we will use it in lstviewMacros_AfterLaberEdit and  cmdRename_Click

' Modification History:

'============================================================

    On Error GoTo RenameMacroErr

Dim i As Long
Dim OldName As String
Dim MoveResult As Long
Dim SelectedItem As Integer
Dim MacroShortKey As String
Dim oldMacFilePath As String
Dim NewMacFilePath As String
Dim OldDecFilePath As String
Dim NewDecFilePath As String
Rem  1.Get the OldName
Rem  2.change the Macro File Name
Rem  3.change the Macro Description ".ptf"
Rem  4.change the name in regedit
Rem  5.FillComboBox

Rem  ===========================================================================

Rem  1.Get the OldName

For i = 0 To lstviewMacros.ListItems.Count - 1

  
    If lstviewMacros.ListItems.Item(i + 1).Selected = True Then
    
       OldName = lstviewMacros.ListItems.Item(i + 1).Text
       
       SelectedItem = i + 1
       
    End If
    

Next i

Rem  2.change the Macro File Name
Rem  3.change the Macro Description ".ptf"

oldMacFilePath = AppPath & "Data\" & OldName & ".pmf"

NewMacFilePath = AppPath & "Data\" & NewString & ".pmf"

OldDecFilePath = AppPath & "Data\" & OldName & ".ptf"

NewDecFilePath = AppPath & "Data\" & NewString & ".ptf"

MoveResult = MoveFile(oldMacFilePath, NewMacFilePath)

Rem Moveing Error

If MoveResult = 0 Then
   'koko1
   MsgBox "Can't Rename Macro"
   
   lstviewMacros.ListItems.Item(SelectedItem).Text = OldName
   
   Exit Sub
   
End If

MoveResult = MoveFile(OldDecFilePath, NewDecFilePath)

If MoveResult = 0 Then
   'koko1
   MsgBox "Can't Rename Macro"
   
   lstviewMacros.ListItems.Item(SelectedItem).Text = OldName
   
   Exit Sub
   
End If

Rem  4.change the name in regedit

MacroShortKey = GetSetting(App.EXEName, "sm", OldName, "")

DeleteSetting App.EXEName, "sm", OldName

SaveSetting App.EXEName, "sm", NewString, MacroShortKey

Rem  5.FillComboBox
             
frmJMacro.FillComboBox
 
If byRenameButton Then
    
    lstviewMacros.ListItems.Item(SelectedItem).Text = NewString

End If
 

RenameMacroExit:
    On Error Resume Next

    Exit Sub

RenameMacroErr:

    Dim sMsg As String
    'Add your error display message here, or use this simple MsgBox display...

    With Err
        sMsg = "Error: " & .Number & vbCrLf
        sMsg = sMsg & "Description: " & .Description & vbCrLf
        sMsg = sMsg & "Code Location: " & App.Title & ":: Manage_RenameMacro" & vbCrLf
        sMsg = sMsg & "Source: " & .Source & vbCrLf
        sMsg = sMsg & "Last DLL Error: " & .LastDllError & vbCrLf
        MsgBox sMsg, vbOKOnly + vbCritical, "Error" & vbCrLf
    End With
GoTo RenameMacroExit

End Sub



Private Sub lstviewMacros_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim i As Long

Rem UnBold all

For i = 0 To lstviewMacros.ListItems.Count - 1

  
        lstviewMacros.ListItems.Item(i + 1).Bold = False
           
           
    
Next i


Rem if the user click cmdRenameMacro or cmdDeleteMacro we

Rem will use them to know the clicked Macro

cmdRenameMacro.Tag = Item.Text

cmdRenameMacro.Tag = Item.Text

Rem Bold me only
Item.Bold = True

End Sub


Private Function GetSelectedMacro() As String
'============================================================

' Routine Name: GetSelectedMacro
' Description: Get Selected Macro Name
' Author: Ahmed Saad
' Date: 23/08/2004 06:46:06 Õ
' Copyright © 2004 CprinGold Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo GetSelectedMacroErr

    'The bulk of your routine here...Dim i As Long
Dim i As Long
Dim MacName As String

For i = 0 To lstviewMacros.ListItems.Count - 1

    
    If lstviewMacros.ListItems.Item(i + 1).Selected = True Then
    
       MacName = lstviewMacros.ListItems.Item(i + 1).Text
        
       lstviewMacros.ListItems.Item(i + 1).Bold = True
       
Rem    we need only the first one
       
    End If
    
  

Next i
    
'debug.Print MacName, Len(MacName)
GetSelectedMacro = MacName
 
GetSelectedMacroExit:
    On Error Resume Next

    Exit Function

GetSelectedMacroErr:

    Dim sMsg As String
    'Add your error display message here, or use this simple MsgBox display...

    With Err
        sMsg = "Error: " & .Number & vbCrLf
        sMsg = sMsg & "Description: " & .Description & vbCrLf
        sMsg = sMsg & "Code Location: " & App.Title & ":: Manage_GetSelectedMacro" & vbCrLf
        sMsg = sMsg & "Source: " & .Source & vbCrLf
        sMsg = sMsg & "Last DLL Error: " & .LastDllError & vbCrLf
        MsgBox sMsg, vbOKOnly + vbCritical, "Error" & vbCrLf
    End With
GoTo GetSelectedMacroExit

End Function

Private Sub lstviewMacros_KeyDown(KeyCode As Integer, Shift As Integer)

If GetAsyncKeyState(vbKeyF2) Then
    
   lstviewMacros.StartLabelEdit
    

ElseIf GetAsyncKeyState(vbKeyDelete) Then

   DeleteSelectedMacro

End If

End Sub

Private Sub OKButton_Click()
 
 Unload Me
 
End Sub

Private Sub OKButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  OKButton.Default = True
  
End Sub

Private Sub tmrBoldMe_Timer()

Dim BoldMacroName As String

Dim PreSelectedMacroName As String

PreSelectedMacroName = GetSetting(App.EXEName, "fr", "st", "")

Rem when the user click cmdmange in frmJmacro
Rem the frmmange loaded and first item in the list view is selected by default
Rem but the item isn't BOLD so we call GetSelectedMacro Function to Make this item BOLD
BoldMacroName = GetSelectedMacro

If BoldMacroName = PreSelectedMacroName And BoldMacroName <> "" Then
   
   If cmdRunStartup.Caption <> "De-Startup" Then cmdRunStartup.Caption = "De-Startup"

Else
   
   If cmdRunStartup.Caption <> "Startup Macro" Then cmdRunStartup.Caption = "Startup Macro"
   
End If

If lstviewMacros.ListItems.Count = 0 Then
      
   cmdshortkey.Enabled = False
   
   cmdDeleteAll.Enabled = False
   
   cmdRunStartup.Enabled = False

   cmdDeleteMacro.Enabled = False
   
   cmdRenameMacro.Enabled = False

Else

   cmdshortkey.Enabled = True
   
   cmdDeleteAll.Enabled = True
   
   cmdRunStartup.Enabled = True

   cmdDeleteMacro.Enabled = True
   
   cmdRenameMacro.Enabled = True
   
   
End If


tmrBoldMe.Interval = 600

Rem why i use 600
Rem the story that
Rem if i make it 600 from the first run i will see just a while the
Rem the selected macro trun from normal font to bold that's why i load the interval
Rem with 1 millyseconds Rem if it still 1 you will see the listview refreshed always so we make it 600

End Sub

Private Sub RenameMacroByButton()
'============================================================

' Routine Name: RenameMacroByButton
' Description:
' Author: Ahmed Saad
' Date: 23/08/2004 08:28:39 Õ
' Copyright © 2004 CprinGold Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo RenameMacroByButtonErr


Dim MacFileTime As String
Dim MacFileDate As String
Dim DecFilePath As String
Dim NewMacroName As String
Dim OldMacroName As String
Dim MacroListItem As ListItem
Dim NewMacDecription As String

OldMacroName = GetSelectedMacro

Load frmRenameMacro

'frmRenameMacro.caption = "Rename " & cmdRenameMacro.Tag & " Macro"

frmRenameMacro.Show 1

If frmRenameMacro.Tag = "ok" Then
   
   NewMacroName = frmRenameMacro.txtMacroName

   If frmRenameMacro.txtDescription.Tag = "UserUpdateDecription" Then
          
       NewMacDecription = frmRenameMacro.txtDescription
       
   End If
   
   Unload frmRenameMacro
      
   Call RenameMacro(NewMacroName, True)
   
   If NewMacDecription <> "" Then
   
          DecFilePath = AppPath & "Data\" & NewMacroName & ".ptf"
          

Rem        if we write the Decription directly with append mode we will put
Rem        the new Decription with the old one.
Rem        and if we write the Decription with output mode we will delete
Rem        the old time and date of macro so we will make the following.

          Open DecFilePath For Input As #5
     
                Input #5, MacFileTime, MacFileDate
       
          Close #5
      
         Open DecFilePath For Output As #5
     
              Write #5, MacFileTime, MacFileDate, NewMacDecription
              
Rem              Update lstviewMacros with new Decription
              
              Set MacroListItem = lstviewMacros.FindItem(NewMacroName, 0, 1, 0)
              
              lstviewMacros.ListItems.Item(MacroListItem.index).SubItems(1) = NewMacDecription
         
         Close #5
    
   End If
   

End If

RenameMacroByButtonExit:
    On Error Resume Next

    Exit Sub

RenameMacroByButtonErr:

    Dim sMsg As String
    'Add your error display message here, or use this simple MsgBox display...

    With Err
        sMsg = "Error: " & .Number & vbCrLf
        sMsg = sMsg & "Description: " & .Description & vbCrLf
        sMsg = sMsg & "Code Location: " & App.Title & ":: Manage_RenameMacroByButton" & vbCrLf
        sMsg = sMsg & "Source: " & .Source & vbCrLf
        sMsg = sMsg & "Last DLL Error: " & .LastDllError & vbCrLf
        MsgBox sMsg, vbOKOnly + vbCritical, "Error" & vbCrLf
    End With
GoTo RenameMacroByButtonExit

End Sub

Private Sub DeleteSelectedMacro()
'============================================================

' Routine Name: DeleteSelectedMacro
' Description:
' Author: Ahmed Saad
' Date: 23/08/2004 08:34:02 Õ
' Copyright © 2004 CprinGold Software
' Notes:

' Modification History:

'============================================================

     On Error Resume Next
Dim MacroAtStartUp As String
Dim SelMacFilePath As String
Dim SelDecFilePath As String
Dim SelectedMacroName As String
      
      
      SelectedMacroName = GetSelectedMacro
        
      MacroAtStartUp = GetSetting(App.EXEName, "fr", "st", "nononomacros")
        
      If MacroAtStartUp <> "nononomacros" Then
         If MacroAtStartUp = SelectedMacroName Then
             SaveSetting App.EXEName, "fr", "st", ""
         End If
      End If
      SelMacFilePath = AppPath & "Data\" & SelectedMacroName & ".pmf"
  
      SelDecFilePath = AppPath & "Data\" & SelectedMacroName & ".ptf"
      
      Kill SelMacFilePath
      
      Kill SelDecFilePath
          
      On Error GoTo DeleteSelectedMacroErr
       
      DeleteSetting App.EXEName, "sm", SelectedMacroName
      
      frmJMacro.FillComboBox
      
      Unload Me
      
      Me.Show

DeleteSelectedMacroExit:
    On Error Resume Next

    Exit Sub

DeleteSelectedMacroErr:

    Dim sMsg As String
    'Add your error display message here, or use this simple MsgBox display...

    With Err
        sMsg = "Error: " & .Number & vbCrLf
        sMsg = sMsg & "Description: " & .Description & vbCrLf
        sMsg = sMsg & "Code Location: " & App.Title & ":: Manage_DeleteSelectedMacro" & vbCrLf
        sMsg = sMsg & "Source: " & .Source & vbCrLf
        sMsg = sMsg & "Last DLL Error: " & .LastDllError & vbCrLf
        MsgBox sMsg, vbOKOnly + vbCritical, "Error" & vbCrLf
    End With
GoTo DeleteSelectedMacroExit

End Sub

Private Sub FillListView()
'============================================================

' Routine Name: FillListView
' Description:
' Author: Ahmed Saad
' Date: 23/08/2004 08:45:55 Õ
' Copyright © 2004 CprinGold Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo FillListViewErr
     
    Dim i As Long
    Dim MacDate As String
    Dim MacTime As String
    Dim DecFilePath As String
    Dim MacDecription As String
      
    With lstviewMacros
    
        For i = 0 To frmJMacro.CobMacros.ListCount - 1
                        
             
             
            
                 .ListItems.Add , , frmJMacro.CobMacros.List(i)
        
Rem          Add Macro Name

                 .ListItems.Item(i + 1).SubItems(1) = "Macro Deleted"
           
             DecFilePath = AppPath & "Data\" & frmJMacro.CobMacros.List(i) & ".ptf"
        
Rem          Get Decription,Time,Date from the .ptf File
              
             

             Open DecFilePath For Input As #11
              
                   Input #11, MacTime, MacDate, MacDecription
 
             Close #11
                                            
             
             .ListItems.Item(i + 1).SubItems(1) = MacDecription
                
             .ListItems.Item(i + 1).SubItems(2) = MacTime
             
             .ListItems.Item(i + 1).SubItems(3) = MacDate
                 
  
            
        
        Next i
    
    End With
FillListViewExit:
    On Error Resume Next

    Exit Sub

FillListViewErr:

    Dim sMsg As String
    'Add your error display message here, or use this simple MsgBox display...

    With Err
        sMsg = "Error: " & .Number & vbCrLf
        sMsg = sMsg & "Description: " & .Description & vbCrLf
        sMsg = sMsg & "Code Location: " & App.Title & ":: Manage_FillListView" & vbCrLf
        sMsg = sMsg & "Source: " & .Source & vbCrLf
        sMsg = sMsg & "Last DLL Error: " & .LastDllError & vbCrLf
        MsgBox sMsg, vbOKOnly + vbCritical, "Error" & vbCrLf
    End With
GoTo FillListViewExit

End Sub


Private Sub SetShortKey()
'============================================================

' Routine Name: SetShortKey
' Description:
' Author: Ahmed Saad
' Date: 24/08/2004 01:43:49 Õ
' Copyright © 2004 CprinGold Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo SetShortKeyErr
    
    Dim SelectedMacro As String
    Dim MacroListItem As ListItem
    
    SelectedMacro = GetSelectedMacro
    
    Set MacroListItem = lstviewMacros.FindItem(SelectedMacro, 0, 1, 0)

    Load frmSetShortKey
    
    'frmSetShortKey.caption = "Shortcut Key for " & SelectedMacro & " Macro"
  
    frmSetShortKey.Show 1, frmManage
    
SetShortKeyExit:
    On Error Resume Next

    Exit Sub

SetShortKeyErr:

    Dim sMsg As String
    'Add your error display message here, or use this simple MsgBox display...

    With Err
        sMsg = "Error: " & .Number & vbCrLf
        sMsg = sMsg & "Description: " & .Description & vbCrLf
        sMsg = sMsg & "Code Location: " & App.Title & ":: ObjectName_ProcName" & vbCrLf
        sMsg = sMsg & "Source: " & .Source & vbCrLf
        sMsg = sMsg & "Last DLL Error: " & .LastDllError & vbCrLf
        MsgBox sMsg, vbOKOnly + vbCritical, "Error" & vbCrLf
    End With
GoTo SetShortKeyExit

End Sub
