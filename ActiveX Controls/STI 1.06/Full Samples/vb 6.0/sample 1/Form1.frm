VERSION 5.00
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
   StartUpPosition =   2  'CenterScreen
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TrayList As MSFlexGrid
Attribute TrayList.VB_VarHelpID = -1

Private Sub cmdFeedback_Click()
  'Send Feedback to CprinGold Software
  TrayIcons1.sendreview
End Sub

Private Sub cmdHide_Click(Index As Integer)
 
 If TrayList.Visible = True Then
    'Hide the selected Item from the System Tray area
    TrayIcons1.Hide TrayList.RowSel
    TrayIcons1.FillMSFlexGrid TrayList
 End If

End Sub

Private Sub Form_Load()
 'dynamiclly add a FlexGrid Control at runtime
 Set TrayList = Controls.Add("MSFlexGridLib.MSFlexGrid", "SysTrayList")
 'fill the FlexGrid
 TrayIcons1.FillMSFlexGrid TrayList
 'Show the FlexGrid
 TrayList.Visible = True
End Sub

