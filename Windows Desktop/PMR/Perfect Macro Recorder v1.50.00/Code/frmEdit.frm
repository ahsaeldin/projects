VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F247AF03-2671-4421-A87A-846ED80CD2A9}#1.0#0"; "jbet33.ocx"
Begin VB.Form frmEdit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9135
   Icon            =   "frmEdit.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList SmallImageList1 
      Left            =   9480
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":0950
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":0CA2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrCheckColumnsWidth 
      Interval        =   100
      Left            =   5280
      Top             =   1920
   End
   Begin MSComctlLib.ImageList LargeImageList 
      Left            =   9960
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView CommandsListView 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   12726
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "SmallImageList1"
      SmallIcons      =   "SmallImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin JwldButn2b.JeweledButton cmdCancel 
      Height          =   375
      Left            =   12240
      TabIndex        =   1
      Top             =   1920
      Width           =   375
      _ExtentX        =   661
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
   Begin JwldButn2b.JeweledButton cmdOk 
      Height          =   375
      Left            =   12240
      TabIndex        =   2
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   "&Save Changes"
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
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//Write How to Hide a listview's column

'//1. Cell Colour is While

'//2. Column Width is zero

'//3. Check Column Width in a timer for changing

'//4. Disalbe Full Row Select

'//cprin 7Aug2006: added next function to init frmedit
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Load()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'============================================================

' Routine Name: Form_Load
' Description: init frmedit
' Author: Ahmed Saad
' Date: 07/08/06 11:51:29 ã
' Copyright © 2006  Perfection Tools Software
' Notes:

' Modification History:

'============================================================

    On Error GoTo Form_LoadErr

    '//cprin 7Aug2006: added next code block
    'to ensure that jeweled buttons will remain
    'as i want.
    ''''''''''''''''''''''''''''''''''''''''''''
    cmdCancel.Caption = "&Cancel"
    cmdOk.Caption = "&Save Changes"
    ''''''''''''''''''''''''''''''''''''''''''''
    
    '//cprin 7Aug2006: added next 5 lines to add 5 columns to the listview
    CommandsListView.ColumnHeaders.Add 1, "Command", "", 1990
    CommandsListView.ColumnHeaders.Add 2, "Type", "", 1900
    CommandsListView.ColumnHeaders.Add 3, "Command Parameter 1", "", 1900
    CommandsListView.ColumnHeaders.Add 4, "Command Parameter 2", "", 1900
    CommandsListView.ColumnHeaders.Add 5, "Timeout", "", 0
    
    '//cprin 7Aug2006: added next line to make the width of ListView is equal to sum of Colunms
    '//i.e. We have 5 columns * 2000 [Colunm width] hence we won't have any unneeded space of the listview
    '//Note The first 4 Columns will have a same width equal to 2000 except the last one is equal to 1950
    '//why we make this?=>the answer is to avoid display the horizontal Scroll Bar at the bottom of the Listview
    CommandsListView.Width = CommandsListView.ColumnHeaders.Item(1).Width * 4
    
    '//cprin 7Aug2006: added next line because sometimes
    '//frmedit for unknown reason increase its width from the right
    frmEdit.Width = CommandsListView.Width + 300
       
    '//cprin 28Aug2006: added next 3 lines to get the selected macro path in the combo box
    Dim MacName As String
    Dim SelectedMacroPath As String
    SelectedMacroPath = GetListedIndexMacroInTheCombo(MacName)
    
    '//cprin 28Aug2006: added next line to assign a title for frmEdit include the name of the current edited macro
    Me.Caption = "Edit " & MacName & " Macro"
     
    '//cprin 28Aug2006: added next line open the macro for editing and to load the CommandsListView
    OpenMacroForEditing (SelectedMacroPath)


Form_LoadExit:
    On Error Resume Next

    Exit Sub
    
Form_LoadErr:
    
    '//0x0ffee11 is a unique number for FrmEdit_Form_Load because i don't want to write
    '//the function name in the error message display to the User for 2 reasons
    '//1.It is my internal structure
    '//2.May be Function name will be useful in Reverse engineering "may be"
    SysError Err, "0x0ffee11"

GoTo Form_LoadExit

End Sub

Private Sub tmrCheckColumnsWidth_Timer()
    
    If CommandsListView.ColumnHeaders(5).Width > 0 Then
       
       CommandsListView.ColumnHeaders(5).Width = 0
    
    End If

End Sub

