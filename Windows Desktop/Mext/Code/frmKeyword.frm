VERSION 5.00
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "jbet33.ocx"
Begin VB.Form frmKeyword 
   BackColor       =   &H00F9F9F9&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4920
   Icon            =   "frmKeyword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtKeyword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox txtPharse 
      Height          =   2295
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
   Begin JwldButn2b.JeweledButton JewSave 
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3120
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
   End
   Begin JwldButn2b.JeweledButton JewCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3120
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
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000003&
      Height          =   3495
      Left            =   120
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label lblkeyword 
      BackStyle       =   0  'Transparent
      Caption         =   "Keyword"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblphrase 
      BackStyle       =   0  'Transparent
      Caption         =   "Text to be inserted"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frmKeyword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Load()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

    JewSave.Caption = "&Save"
    JewCancel.Caption = "&Cancel"

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Jewcancel_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Unload Me

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub JewSave_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    Dim res As Boolean
    
    If txtKeyword <> "" And txtPharse <> "" Then
        
       If Me.Tag = "Add" Then
          
          
          If frmmain.trvTexts.SelectedItem.Tag = "Folder" Then '//If the selected item is folder then

             res = frmmain.AddNewKeywordValue(frmmain.trvTexts.SelectedItem.Key, txtKeyword, txtPharse)

          Else '//else if the selected item is a keyword get the parent
         
             res = frmmain.AddNewKeywordValue(frmmain.trvTexts.SelectedItem.Parent.Key, txtKeyword, txtPharse)
       
          End If
          
          If Not res Then
            
            txtKeyword.SetFocus
            
            txtKeyword.SelStart = 0
            
            txtKeyword.SelLength = Len(txtKeyword)
            
            Exit Sub  '//exit sub DONOT unload the form to enable the user to re-enter a new keyword and value
          
          End If
       
       ElseIf Me.Tag = "Edit" Then
       
          If frmmain.trvTexts.SelectedItem.Tag = "Folder" Then '//If the selected item is folder then
             
             '//MsgBox "as" messagebox here to tell the user that he didnot select a keyword to edit
    
          Else '//else if the selected item is a keyword get the parent
    
             res = frmmain.EditKeywordValue(txtKeyword, txtPharse)
             
             If Not res Then
               
                txtKeyword.SetFocus
            
                txtKeyword.SelStart = 0
            
                txtKeyword.SelLength = Len(txtKeyword)
            
                Exit Sub  '//exit sub DONOT unload the form to enable the user to re-enter a new keyword and value
                      
             End If
     
          End If
    
       End If
       
    Else
       
       If txtPharse = "" And txtKeyword = "" Then
          MsgBox "Please, Insert a keyword and the text to be inserted.", vbCritical: Exit Sub
       ElseIf txtKeyword = "" Then
          MsgBox "Please, Insert a keyword.", vbCritical: txtKeyword.SetFocus:   Exit Sub
       ElseIf txtPharse = "" Then
          MsgBox "Please, Insert the text to be inserted.", vbCritical: txtPharse.SetFocus: Exit Sub
       End If
       
    End If
            
    Unload Me
    
End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub txtKeyword_KeyDown(KeyCode As Integer, Shift As Integer)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If KeyCode = 65 And GetAsyncKeyState(vbKeyControl) Then
        
       txtKeyword.SelStart = 0
       
       txtKeyword.SelLength = Len(txtKeyword)
    
    End If

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub txtPharse_Click()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

  'If Me.Tag = "Edit" Then txtPharse.SelStart = 1: txtPharse.SelLength = Len(txtPharse)

End Sub

'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub txtPharse_KeyDown(KeyCode As Integer, Shift As Integer)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    If KeyCode = 65 And GetAsyncKeyState(vbKeyControl) Then
        
       txtPharse.SelStart = 0
       
       txtPharse.SelLength = Len(txtPharse)
    
    End If
    
End Sub
