VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "jbet33.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAdd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add an application to Startup"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog ComDiaAdd 
      Left            =   3480
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraAdd 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6735
      Begin VB.ComboBox comWhere 
         Height          =   315
         ItemData        =   "frmAdd.frx":08CA
         Left            =   120
         List            =   "frmAdd.frx":08DA
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Timer tmrCheckTextBoxs 
         Interval        =   1
         Left            =   4320
         Top             =   1680
      End
      Begin JwldButn2b.JeweledButton JewOK 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&OK"
         PictureSize     =   0
         OriginalPicSizeW=   0
         OriginalPicSizeH=   0
         Enabled         =   0   'False
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
      Begin VB.TextBox txtname 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   6495
      End
      Begin VB.TextBox txtPath 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   6495
      End
      Begin JwldButn2b.JeweledButton JewCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   2160
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
      End
      Begin JwldButn2b.JeweledButton JewBrowse 
         Height          =   375
         Left            =   5400
         TabIndex        =   2
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&Browse..."
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File Path :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "File Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'«««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««
'«««⁄„· ﬂ„« ·Ê ﬂ‰  ·”  »Õ«Ã… ≈·Ï «·„«· , Õˆ» ﬂ„« ·Ê ﬂ‰  ·„   √–Ï ›Ï «·Õ» „‰ ﬁ»· , «—ﬁ’ ﬂ„« ·Ê ·ﬂ Ìﬂ‰ √Õœ« Ì—«ﬁ»ﬂ"~„«—ﬂ  ÊÌ‰
'«««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««««

Option Explicit

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Load()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    JewOK.Caption = "&OK"
    JewCancel.Caption = "&Cancel"
    JewBrowse.Caption = "&Browse..."
    
    comWhere.ListIndex = 0 'Registry\Current User\Run

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub tmrCheckTextBoxs_Timer()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If txtname <> "" Then
       If txtPath <> "" Then
          If InStr(1, txtPath, ":\") >= 1 Then '//make a very simple check that this is a path
             JewOK.Enabled = True
          End If
       Else
          JewOK.Enabled = False
       End If
    Else
       JewOK.Enabled = False
    End If
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewOK_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    '### 25Feb2008 07:12 PM:
    '«·‰Â«—œ… ⁄«Ì“ «‰‘—  œÊÌ‰… ÃœÌœ… »⁄œ  œÊÌ‰… ‘«Ê‘«‰ﬂ
    '###
          
    Dim FilePath As String
    Dim FileName As String
    Dim FullFilePath As String
    
    FullFilePath = txtPath
    
    FilePath = txtPath
    
    FilePath = ExtractPath(FilePath)
    
    FileName = ExtractFileNameFromPath(FilePath)
    
    '//prevent these chars \ / : * ? " < > | in the filename
    If InStr(1, FileName, "\") Or InStr(1, FileName, "/") Or _
    InStr(1, FileName, ":") Or InStr(1, FileName, "*") Or InStr(1, FileName, "?") _
    Or InStr(1, FileName, Chr(34)) Or InStr(1, FileName, "<") Or InStr(1, FileName, ">") _
    Or InStr(1, FileName, "|") Then
       '//Tell the user that
       MsgBox "A file name cannot contain any of the following characters:" & Chr(13) & Chr(10) & "\ / : * ?  " & Chr(34) & " < > |", vbCritical, SoftwareName
       '//Set the SelStart to the first char in txtMacroName
       txtPath.SelStart = 0
       '//Highlight the name entered by the user
       txtPath.SelLength = Len(txtPath)
       Exit Sub
    End If
    
    '//ﬂœ… „‘ »«ﬁÏ ≈·« 3 —„Ê“ „«⁄„· ‘  Õﬁﬁ ⁄·Ï „ÊÃÊœÂ„ ›Ï „”«— «·„·›
    'InStr(1, FilePath, "\") Or InStr(1, FilePath, "/") Or InStr(1, FilePath, ":")
    
    '//prevent these chars * ? " < > | in the file path
    If InStr(1, FilePath, "*") Or InStr(1, FilePath, "?") _
    Or InStr(1, FilePath, Chr(34)) Or InStr(1, FilePath, "<") Or InStr(1, FilePath, ">") _
    Or InStr(1, FilePath, "|") Then
       '//Tell the user that
       MsgBox "A file path cannot contain any of the following characters:" & Chr(13) & Chr(10) & "* ?  " & Chr(34) & " < > |", vbCritical, SoftwareName
       '//Set the SelStart to the first char in txtMacroName
       txtPath.SelStart = 0
       '//Highlight the name entered by the user
       txtPath.SelLength = Len(txtPath)
       Exit Sub
    End If
    
    Select Case comWhere.ListIndex
           
           Case 0 'Registry\Current User\Run
           
                Savestring HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", txtname, FullFilePath

           Case 1 'Registry\Current User\RunOnce
           
                Savestring HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", txtname, FullFilePath

           Case 2 'Registry\AllUsers\Run
                
                Savestring HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", txtname, FullFilePath
               
           Case 3 'Registry\AllUsers\RunOnce
                
                Savestring HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", txtname, FullFilePath

    End Select

    frmmain.JewRefresh_Click
    
    Unload Me

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewBrowse_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    On Error GoTo quitme:
    
    With ComDiaAdd
         .Filter = "Application files (*.exe)|*.exe|All files (*.*)|*.*" '"Application files (*.exe)"
         .CancelError = True
         .flags = cdlOFNPathMustExist '+ cdlOFNExplorer '+ cdlOFNAllowMultiselect
         'Note for the pervious line: we must add cdlOFNExplorer to the flags when use cdlOFNAllowMultiselect
         'otherwise we showopen will be the silly old ShowOpen Dialog Box.
         .FileName = ""
         .DialogTitle = "Select File"
         .ShowOpen
    End With
   
    txtPath = ComDiaAdd.FileName
    txtname = ComDiaAdd.FileTitle
    
Exit Sub
quitme:

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewCancel_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Unload Me

End Sub
