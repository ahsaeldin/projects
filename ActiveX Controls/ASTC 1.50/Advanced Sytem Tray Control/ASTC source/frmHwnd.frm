VERSION 5.00
Begin VB.Form frmHwnd 
   Caption         =   "Internal window for Hwnd"
   ClientHeight    =   0
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   1725
   LinkTopic       =   "Form1"
   ScaleHeight     =   0
   ScaleWidth      =   1725
End
Attribute VB_Name = "frmHwnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Caption = hwnd
End Sub
