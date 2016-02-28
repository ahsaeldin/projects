VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{2F49899B-2A36-11DA-A1C1-90BB54C10000}#1.0#0"; "cprinsti.ocx"
Begin VB.Form Form1 
   Caption         =   "VBTrayIcons"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin STI.TrayIcons TrayIcons1 
      Left            =   2760
      Top             =   3720
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   2535
      Begin VB.CommandButton cmdHide 
         Caption         =   "&Hide Icon"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdFeedback 
         Caption         =   "Feedback"
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3615
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   100
      Cols            =   7
      ForeColorSel    =   16711680
      GridColorFixed  =   255
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
 
 'Fill the FlexGrid control with System Tray Data
 FillFlexGrid

End Sub

Private Sub initFlexGrid()


  MSFlexGrid1.ColWidth(0) = 300
  MSFlexGrid1.TextMatrix(0, 0) = "I"
  MSFlexGrid1.ColWidth(1) = 3000
  MSFlexGrid1.TextMatrix(0, 1) = "Application Path"
  MSFlexGrid1.ColWidth(2) = 1300
  MSFlexGrid1.TextMatrix(0, 2) = "uID"
  MSFlexGrid1.ColWidth(3) = 1300
  MSFlexGrid1.TextMatrix(0, 3) = "hWnd"
  MSFlexGrid1.ColWidth(4) = 1300
  MSFlexGrid1.TextMatrix(0, 4) = "hIcon"
  MSFlexGrid1.ColWidth(5) = 1700
  MSFlexGrid1.TextMatrix(0, 5) = "ToolTip"
  MSFlexGrid1.ColWidth(6) = 1500
  MSFlexGrid1.TextMatrix(0, 6) = "uCallbackMessage"
  MSFlexGrid1.ForeColorFixed = vbBlue
  
   
End Sub

Private Sub cmdFeedback_Click()
  TrayIcons1.sendreview
End Sub

Private Sub cmdHide_Click(Index As Integer)
 
 If MSFlexGrid1.Visible = True Then
    'Hide the selected Item from the System Tray area
    TrayIcons1.Hide MSFlexGrid1.RowSel
    MSFlexGrid1.Clear
    FillFlexGrid
 End If

End Sub

Public Sub FillFlexGrid()

Dim i As Integer
Dim TrayList() As TrayIcon

initFlexGrid

'Get the System Tray Data
TrayList = TrayIcons1.GetSysTrayIcons

'Fill the FlexGrid with Data
For i = 0 To TrayIcons1.IconCount - 1
    
    MSFlexGrid1.TextMatrix(i + 1, 0) = i + 1
    If TrayList(i).APath = "" Then
       MSFlexGrid1.TextMatrix(i + 1, 1) = "N/A"
    Else
       MSFlexGrid1.TextMatrix(i + 1, 1) = TrayList(i).APath
    End If
    MSFlexGrid1.TextMatrix(i + 1, 2) = TrayList(i).uId
    MSFlexGrid1.TextMatrix(i + 1, 3) = TrayList(i).hWnd
    MSFlexGrid1.TextMatrix(i + 1, 4) = TrayList(i).hIcon
    MSFlexGrid1.TextMatrix(i + 1, 5) = TrayList(i).ToolTip
    MSFlexGrid1.TextMatrix(i + 1, 6) = "&H" & Hex$(TrayList(i).ucallbackMessage)
   
   
Next i

End Sub
